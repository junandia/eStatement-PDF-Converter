import PyPDF2
import pandas as pd
import re
import streamlit as st
import io
import os

from common_ui import render_page_header, render_upload_section, render_download_section

def extract_mandiri_estatement(pdf_file):
    extracted_data = []
    
    reader = PyPDF2.PdfReader(pdf_file)
    table_started = False
    
    for page in reader.pages:
        text = page.extract_text()
        if not text:
            continue
        
        lines = text.split('\n')
        i = 0
        while i < len(lines):
            line = lines[i].strip()
            
            # Melewati header sampai ketemu tabel transaksi
            if "Posting Date" in line and "Remark" in line:
                table_started = True
                i += 1
                continue
            
            if not table_started:
                i += 1
                continue

            # 1. DETEKSI TANGGAL (Contoh: 01 Feb 2026,)
            if re.match(r'^\d{2} \w{3} \d{4},', line):
                date_part = line.replace(',', '').strip()
                full_date = date_part
                remark_parts = []
                
                # 2. DETEKSI JAM (Baris tepat di bawah tanggal)
                i += 1
                if i < len(lines):
                    next_line = lines[i].strip()
                    # Pola jam HH:mm:ss
                    time_match = re.match(r'^(\d{2}:\d{2}:\d{2})(.*)', next_line)
                    if time_match:
                        full_date = f"{date_part} {time_match.group(1)}"
                        # Sisa teks setelah jam
                        if time_match.group(2).strip():
                            remark_parts.append(time_match.group(2).strip())
                    else:
                        remark_parts.append(next_line)
                
                # 3. KUMPULKAN REMARK SAMPAI KETEMU BARIS NOMINAL (Greedy)
                j = i + 1
                debit = credit = balance = 0.0
                found_money = False
                
                while j < len(lines):
                    curr_line = lines[j].strip()
                    
                    # Berhenti jika bertemu transaksi baru (failsafe)
                    if re.match(r'^\d{2} \w{3} \d{4},', curr_line):
                        break
                        
                    # Cari baris yang mengandung minimal 3 angka desimal (Debit, Credit, Balance)
                    num_pattern = r'(\d{1,3}(?:,\d{3})*(?:\.\d{2}))'
                    all_nums = re.findall(num_pattern, curr_line)
                    
                    if len(all_nums) >= 3:
                        debit = float(all_nums[-3].replace(',', ''))
                        credit = float(all_nums[-2].replace(',', ''))
                        balance = float(all_nums[-1].replace(',', ''))
                        
                        # Ambil teks sebelum angka pertama (sisa Remark/Ref No)
                        text_before_nums = re.split(num_pattern, curr_line)[0].strip()
                        if text_before_nums and text_before_nums != "-":
                            remark_parts.append(text_before_nums)
                        
                        found_money = True
                        j += 1
                        break
                    else:
                        # Masukkan ke remark jika bukan footer
                        if curr_line and not any(x in curr_line for x in ["Page", "Created", "Account Statement"]):
                            remark_parts.append(curr_line)
                    j += 1
                
                # 4. GABUNGKAN REMARK
                final_remark = " ".join(remark_parts).strip()
                final_remark = re.sub(' +', ' ', final_remark) # Bersihkan spasi ganda
                
                if found_money:
                    extracted_data.append([
                        full_date, final_remark, debit, credit, balance
                    ])
                
                i = j - 1
            i += 1
                
    df = pd.DataFrame(extracted_data, columns=["Posting Date", "Remark", "Debit", "Credit", "Balance"])
    return df

def mainMandiriEstatement():
    render_page_header(
        "Mandiri e-Statement Converter",
        "Upload file e-statement PDF Mandiri untuk mengkonversi ke Excel.",
    )

    # Upload section
    render_upload_section(
        "Upload File PDF Mandiri",
        "Pilih satu atau lebih file PDF e-statement Mandiri. Sistem akan otomatis memproses dan menggabungkan data transaksi."
    )

    uploaded_files = st.file_uploader(
        "Upload PDF Mandiri",
        type="pdf",
        accept_multiple_files=True,
        help="Pilih file PDF e-statement Mandiri",
        label_visibility="collapsed",
    )

    if uploaded_files:
        progress_bar = st.progress(0)
        status_text = st.empty()

        all_transactions = []
        for i, uploaded_file in enumerate(uploaded_files):
            status_text.text(f"Memproses file {i+1} dari {len(uploaded_files)}: {uploaded_file.name}")
            progress_bar.progress((i) / len(uploaded_files))

            try:
                df = extract_mandiri_estatement(uploaded_file)
                if not df.empty:
                    df['source_file'] = uploaded_file.name
                    all_transactions.append(df)
                else:
                    st.warning(f"Tidak ada data transaksi ditemukan pada {uploaded_file.name}")
            except Exception as e:
                st.error(f"Gagal memproses {uploaded_file.name}: {str(e)}")

        progress_bar.progress(1.0)
        status_text.text("✅ Semua file berhasil diproses!")

        if all_transactions:
            global_dataframe = pd.concat(all_transactions, ignore_index=True)

            # Summary metrics
            total_transactions = len(global_dataframe)
            total_debit = global_dataframe['Debit'].sum()
            total_credit = global_dataframe['Credit'].sum()

            st.markdown("---")
            st.subheader("📊 Ringkasan Transaksi")

            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Total Transaksi", f"{total_transactions:,}")
            with col2:
                st.metric("Total Debit", f"Rp {total_debit:,.0f}")
            with col3:
                st.metric("Total Kredit", f"Rp {total_credit:,.0f}")

            st.markdown("---")
            st.subheader("📋 Data Transaksi")
            st.dataframe(global_dataframe, use_container_width=True)

            # Download section
            render_download_section()

            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
                global_dataframe.to_excel(writer, sheet_name="All Transactions", index=False)
            excel_buffer.seek(0)

            col_dl1, col_dl2 = st.columns(2)
            with col_dl1:
                st.download_button(
                    label="📥 Download Excel File",
                    data=excel_buffer.getvalue(),
                    file_name="Mandiri_Statement_Combined.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            with col_dl2:
                st.download_button(
                    label="📥 Download CSV",
                    data=global_dataframe.to_csv(index=False).encode('utf-8'),
                    file_name="Mandiri_Statement_Combined.csv",
                    mime="text/csv",
                    use_container_width=True,
                )

if __name__ == "__main__":
    mainMandiriEstatement()
