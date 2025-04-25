# 🛫 Collective Roster - Automatic DCR Filling App


## ✨ Deskripsi

**Collective Roster - Automatic CR Filling App** adalah aplikasi berbasis web menggunakan Streamlit yang dirancang untuk mengotomatisasi pengisian data *Daily Crew Report (DCR)* secara efisien berdasarkan file template dan input harian dari kru penerbangan.

Aplikasi ini membantu staf roster atau admin kru untuk menghindari proses manual yang memakan waktu dan rawan kesalahan saat mengisi laporan harian.

---

## ⚙️ Fitur Utama

- 📂 Upload file **template** dan **input harian**
- 🔍 **Preview otomatis** setelah upload untuk validasi
- 🛠️ Tombol **proses data** dengan validasi dan pengisian otomatis
- 📈 **Preview hasil akhir** sebelum diunduh
- 📅 Download hasil dalam format `.xlsx`
- 🔄 Tombol **refresh aplikasi**
- 🎨 Tampilan UI interaktif & profesional

---

## 🧰 Teknologi yang Digunakan

- Python 3.9+
- Streamlit
- pandas
- openpyxl

---

## 🚀 Cara Menjalankan Secara Lokal

1. **Clone repositori ini:**

```bash
git clone https://github.com/username/CollectiveRosterApp.git
cd CollectiveRosterApp
```

2. **(Opsional) Buat virtual environment dan aktifkan:**

```bash
python -m venv venv
# Windows
venv\Scripts\activate
# macOS/Linux
source venv/bin/activate
```

3. **Install dependensi:**

```bash
pip install -r requirements.txt
```

4. **Jalankan aplikasi:**

```bash
streamlit run app.py
```

---

## 🌐 Deployment (Opsional)

### Deploy ke Streamlit Cloud:

1. Pastikan file `app.py`, `requirements.txt`, dan `README.md` sudah ada di GitHub
2. Masuk ke [https://streamlit.io/cloud](https://streamlit.io/cloud)
3. Hubungkan ke repositorimu
4. Klik "Deploy" dan aplikasi siap digunakan

---

## 📁 Struktur Folder

```
├── app.py
├── requirements.txt
├── README.md
├── assets/
│   └── logo.png
```

---

## 📬 Kontak

Pengembang: [Maurino Audrian Putra](https://www.linkedin.com/in/maurino23)

---

## 📄 Lisensi

Proyek ini dilisensikan di bawah MIT License. Silakan lihat file `LICENSE` untuk informasi selengkapnya.
