import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
import tempfile
import os
from datetime import datetime

# --- KONFIGURASI HALAMAN ---
st.set_page_config(
    page_title="E-Statement Converter",
    page_icon="🏦",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- CUSTOM CSS ---
st.markdown("""
<style>
    .main-header { font-size: 2.5rem; font-weight: bold; color: #1F4E79; text-align: center; margin-bottom: 1rem; }
    .sub-header { font-size: 1.2rem; color: #666; text-align: center; margin-bottom: 2rem; }
    .stat-card { background: linear-gradient(135deg, #1F4E79 0%, #2E75B6 100%); color: white; padding: 1.5rem; border-radius: 10px; text-align: center; margin-bottom: 1rem; }
    .stat-value { font-size: 1.8rem; font-weight: bold; }
    .stat-label { font-size: 0.9rem; opacity: 0.9; }
</style>
""", unsafe_allow_html=True)

# --- HELPER FUNCTIONS ---
def clean_currency(text):
    if not text: return 0.0
    clean = re.sub(r'[^\d.]', '', str(text).replace(',', ''))
    try: return float(clean)
    except: return 0.0

def extract_bri_logic(pdf_path):
    """Logika Khusus Mutasi BRI BritAma Bisnis"""
    all_data = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            words = page.extract_words(x_tolerance=3, y_tolerance=3)
            if not words: continue

            # Batas Header (Mencari 'Tanggal Transaksi')
            header_y = 0
            for w in words:
                if "Tanggal" in w['text'] and "Transaksi" in w['text']:
                    header_y = w['bottom']
                    break
            
            # Filter data di bawah header & urutkan
            words = [w for w in words if w['top'] > header_y]
            words.sort(key=lambda x: (x['top'], x['x0']))

            transaction_blocks = []
            current_block = []
            
            for w in words:
                # Anchor BRI: Tanggal format DD/MM/YY di kolom kiri (x < 80)
                if w['x0'] < 80 and re.match(r'^\d{2}/\d{2}/\d{2}$', w['text']):
                    if current_block: transaction_blocks.append(current_block)
                    current_block = [w]
                else:
                    if current_block: current_block.append(w)
            
            if current_block: transaction_blocks.append(current_block)

            for block in transaction_blocks:
                tx = {
                    "date": "", "desc": [], "teller": "", 
                    "debit": 0.0, "kredit": 0.0, "balance": 0.0
                }
                
                # Temp storage untuk angka
                money_vals = []
                
                for w in block:
                    x, txt = w['x0'], w['text']
                    
                    # 1. Date (Abaikan jam HH:MM:SS)
                    if x < 80 and re.match(r'^\d{2}/\d{2}/\d{2}$', txt):
                        tx["date"] = txt
                    
                    # 2. Teller/User ID (Sekitar x: 280-350)
                    elif 280 <= x < 350:
                        tx["teller"] = txt
                        
                    # 3. Description (Antara tanggal dan teller)
                    elif 80 <= x < 280:
                        if not re.match(r'^\d{2}:\d{2}:\d{2}$', txt): # Buang jam
                            tx["desc"].append(txt)
                    
                    # 4. Money Columns (Debet, Kredit, Saldo)
                    elif x >= 350:
                        if re.search(r'[\d,]+\.\d{2}', txt):
                            money_vals.append(clean_currency(txt))

                if tx["date"]:
                    # BRI Pattern: [Debet] [Kredit] [Saldo]
                    if len(money_vals) >= 3:
                        tx["debit"] = money_vals[0]
                        tx["kredit"] = money_vals[1]
                        tx["balance"] = money_vals[2]
                    
                    all_data.append({
                        "tanggal": tx["date"],
                        "keterangan": " ".join(tx["desc"]).strip(),
                        "teller": tx["teller"],
                        "debit": tx["debit"],
                        "kredit": tx["kredit"],
                        "saldo": tx["balance"]
                    })
    return pd.DataFrame(all_data)

# --- MAIN APP ---
def mainBriEstatement():
    st.markdown('<div class="main-header">🏦 E-Statement Converter</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-header">Konversi Mutasi BRI ke format Excel/CSV</div>', unsafe_allow_html=True)

    st.sidebar.header("📁 Pengaturan")
    bank_option = st.sidebar.selectbox("Pilih Bank", ["BRI", "BNI"])
    
    uploaded_file = st.file_uploader(f"Upload PDF {bank_option}", type="pdf")

    if uploaded_file:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
            tmp.write(uploaded_file.getvalue())
            tmp_path = tmp.name

        try:
            with st.spinner(f"Memproses Mutasi {bank_option}..."):
                if bank_option == "BRI":
                    df = extract_bri_logic(tmp_path)
                else:
                    # Anda bisa memasukkan extract_bni_logic di sini jika ingin digabung
                    df = pd.DataFrame()

            if not df.empty:
                # --- STAT CARDS ---
                total_debit = df['debit'].sum()
                total_kredit = df['kredit'].sum()
                
                c1, c2, c3 = st.columns(3)
                with c1:
                    st.markdown(f'<div class="stat-card"><div class="stat-label">Total Transaksi</div><div class="stat-value">{len(df)}</div></div>', unsafe_allow_html=True)
                with c2:
                    st.markdown(f'<div class="stat-card"><div class="stat-label">Total Debet</div><div class="stat-value">Rp {total_debit:,.2f}</div></div>', unsafe_allow_html=True)
                with c3:
                    st.markdown(f'<div class="stat-card"><div class="stat-label">Total Kredit</div><div class="stat-value">Rp {total_kredit:,.2f}</div></div>', unsafe_allow_html=True)

                st.subheader("📋 Preview Mutasi")
                st.dataframe(df, use_container_width=True, hide_index=True)

                # Export
                st.subheader("📥 Download")
                col_dl1, col_dl2 = st.columns(2)
                
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='Mutasi_BRI')
                
                col_dl1.download_button(
                    label="💾 Download Excel (.xlsx)",
                    data=buffer.getvalue(),
                    file_name=f"BRI_Export_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                csv = df.to_csv(index=False).encode('utf-8')
                col_dl2.download_button(label="📄 Download CSV", data=csv, file_name="BRI_Export.csv", mime="text/csv")

            else:
                st.warning("Data tidak ditemukan. Pastikan file adalah PDF Mutasi BRI yang valid.")

        except Exception as e:
            st.error(f"Gagal memproses: {str(e)}")
        finally:
            if os.path.exists(tmp_path): os.remove(tmp_path)

    st.markdown("---")
    st.markdown("<div style='text-align: center; color: #666;'>E-Statement Converter | Sukabumi 2026</div>", unsafe_allow_html=True)

if __name__ == "__main__":
    mainBriEstatement()