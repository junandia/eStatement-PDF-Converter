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
from mainBriEstatement import mainBriEstatement


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
        ["Beranda", "BCA", "SeaBank", "Mandiri", "BSI", "BNI"],
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

    st.sidebar.markdown("### 📞 Kontak <div style='margin-top: 20px; font-size: 0.9em; color: #888;'><a href='mailto:rismawan.email@gmail.com'>Rismawan Junandia</a></div>",
        unsafe_allow_html=True,
    )
    st.sidebar.markdown("### &copy; 2026 Rismawan Junandia. All rights reserved.",)

# Memanggil fungsi utama
if __name__ == "__main__":
    main()
