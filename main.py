import streamlit as st
import pandas as pd
from io import BytesIO
import io

# Set page config as the first Streamlit command
st.set_page_config(
    page_title="Bank Statement Converter",
    layout="wide",
    initial_sidebar_state="expanded",
    page_icon="🏦"
)

from common_ui import render_page_header
from BcaEstatement import mainBcaEstatement
from SeaBankEstatement import mainSeaBankEstatement
from MandiriEstatement import mainMandiriEstatement
from BsiEstatement import mainBsiEstatement
from BniEstatement import mainBniEstatement
from BriEstatement import mainBriEstatement


def main():
    render_page_header(
        "Aplikasi Statement Bank Konverter",
        "Konversi statement bank PDF ke Excel/CSV dengan mudah.",
        "Pilih bank di sidebar lalu upload file PDF statement untuk memulai."
    )

    # Sidebar dengan menu pilihan
    st.sidebar.markdown("### 🏦 Pilih Bank")
    menu = st.sidebar.selectbox(
        "Pilih Bank",
        ["Beranda", "BCA", "SeaBank", "Mandiri", "BSI", "BNI","BRI"],
        label_visibility="collapsed"
    )

    # Pilihan menu
    if menu == "BCA":
        st.sidebar.success("✅ BCA dipilih")
        mainBcaEstatement()
    elif menu == "SeaBank":
        st.sidebar.success("✅ SeaBank dipilih")
        mainSeaBankEstatement()
    elif menu == "Mandiri":
        st.sidebar.success("✅ Mandiri dipilih")
        mainMandiriEstatement()
    elif menu == "BSI":
        st.sidebar.success("✅ BSI dipilih")
        mainBsiEstatement()
    elif menu == "BNI":
        st.sidebar.success("✅ BNI dipilih")
        mainBniEstatement()
    elif menu == "BRI":
        st.sidebar.success("✅ BRI dipilih")
        mainBriEstatement()
    else:
        st.subheader("Selamat datang di aplikasi konversi statement bank!")
        st.write("Silakan pilih bank di sidebar untuk memulai.")
        st.info("💡 **Tips:** Pastikan file PDF e-statement dalam format yang benar untuk hasil terbaik.")
        st.success("🔒 **Privasi & Keamanan:** Kami **tidak menyimpan** file PDF yang Anda unggah sama sekali di server. Seluruh pemrosesan dilakukan langsung di memori (*in-memory*) dan data akan langsung hilang saat sesi ditutup.")
        
        st.markdown("---")
        st.subheader("☕ Dukung Pengembang")
        st.write("Jika aplikasi ini bermanfaat untuk Anda, pertimbangkan untuk memberikan dukungan melalui Saweria:")
        st.image("https://saweria.co/widgets/qr?streamKey=f8b2e53b07f2a494739dca385e592a08", width=200, caption="Scan untuk dukung via Saweria")

    st.sidebar.markdown("---")
    st.sidebar.markdown("### ☕ Dukung Aplikasi")
    st.sidebar.image("https://saweria.co/widgets/qr?streamKey=f8b2e53b07f2a494739dca385e592a08", width=180)
    st.sidebar.markdown("### 📞 Kontak <div style='font-size: 0.9em; color: #888;'><a href='mailto:rismawan.email@gmail.com'>Rismawan Junandia</a></div>",
        unsafe_allow_html=True,
    )
    st.sidebar.markdown("### &copy; 2026 Rismawan Junandia. All rights reserved.",)

# Memanggil fungsi utama
if __name__ == "__main__":
    main()
