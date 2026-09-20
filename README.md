# 🏦 Bank e-Statement Converter

**Bank e-Statement Converter** adalah aplikasi web berbasis Streamlit untuk mengubah laporan rekening / mutasi bank (PDF e-Statement) menjadi file Excel (.xlsx) dan CSV secara otomatis, cepat, dan akurat.

---

## 🚀 Live Demo

Coba demo aplikasi secara langsung di sini:  
👉 **[https://estatementpdfnew.streamlit.app/](https://estatementpdfnew.streamlit.app/)**

---

## 🔒 Privasi & Keamanan Data
> 🛡️ **Penting**: Keamanan dan privasi data keuangan Anda adalah prioritas. Website demo dan aplikasi ini **tidak menyimpan (no-storage)** file PDF e-Statement yang Anda unggah sama sekali. Seluruh proses konversi berjalan di memori (*in-memory*) dan file Anda akan otomatis terhapus seketika setelah sesi konversi selesai atau halaman ditutup.

---

## 🏧 Bank yang Didukung

Aplikasi ini mendukung parsing format e-Statement dari berbagai bank di Indonesia:
- **BCA** (Bank Central Asia)
- **SeaBank**
- **Bank Mandiri**
- **BSI** (Bank Syariah Indonesia)
- **BNI** (Bank Negara Indonesia)
- **BRI** (Bank Rakyat Indonesia - BritAma Bisnis)

---

## ✨ Fitur Utama

- **Multi-Bank Converter**: Pilih bank yang diinginkan dari navigasi sidebar.
- **Konversi Otomatis**: Mengekstraksi tanggal, deskripsi transaksi, debet/kredit, saldo, dan nomor rekening.
- **Batch Processing**: Mendukung unggah multiple file PDF sekaligus untuk digabungkan menjadi satu rekapitulasi.
- **Ringkasan & Metrik Transaksi**: Menampilkan total transaksi, total debit, total kredit, serta net flow.
- **Ekspor Fleksibel**: Download hasil mutasi ke format **Excel (.xlsx)** atau **CSV**.
- **In-Memory & Safe Processing**: File diproses secara aman di memori tanpa meninggalkan file sisa yang tidak perlu di server. **Catatan**: Aplikasi tidak menyimpan file Anda di server kami; data akan terhapus secara otomatis setelah sesi pemrosesan selesai.

---

## 📖 Cara Penggunaan

1. Buka aplikasi (lokal atau melalui [Live Demo](https://estatementpdfnew.streamlit.app/)).
2. Pilih bank yang sesuai di menu **Sidebar**.
3. Unggah satu atau beberapa file PDF e-Statement Anda.
4. Aplikasi akan memproses mutasi dan menampilkan pratinjau data serta ringkasan transaksi.
5. Klik tombol **Download Excel File** atau **Download CSV** untuk mengunduh hasilnya.

---

## 💻 Panduan Instalasi Lokal

### 1. Clone Repository
```bash
git clone https://github.com/junandia/bcaeStatement.git
cd bcaeStatement
```

### 2. Buat Virtual Environment (Opsional tapi Direkomendasikan)
```bash
# Windows
python -m venv venv
venv\Scripts\activate

# Linux / macOS
python3 -m venv venv
source venv/bin/activate
```

### 3. Install Dependensi
```bash
pip install -r requirements.txt
```

> ⚠️ **Catatan untuk BCA & BSI**: Parser BCA dan BSI menggunakan modul `tabula-py` yang membutuhkan **Java (JRE/JDK)** terpasang pada komputer Anda dan terdaftar di `PATH`. Bank lain (Mandiri, BNI, BRI, SeaBank) murni menggunakan Python (`pdfplumber` / `PyPDF2`).

### 4. Jalankan Aplikasi
```bash
streamlit run main.py
```

Buka peramban (browser) di alamat `http://localhost:8501`.

---

## ☕ Dukung Pengembangan
Jika aplikasi ini bermanfaat bagi Anda, dukung pengembang dengan memberikan donasi melalui Saweria:

![Dukung Kami](https://saweria.co/widgets/qr?streamKey=f8b2e53b07f2a494739dca385e592a08)

---


## 👨‍💻 Kontak & Pengembang

**Rismawan Junandia**  
- ✉️ Email: [rismawan.email@gmail.com](mailto:rismawan.email@gmail.com)  
- 🌐 Website: [https://ukkazu.biz.id](http://ukkazu.biz.id/)  
- 📝 Blog: [https://ukkazudigital.blogspot.com/](https://ukkazudigital.blogspot.com/)

