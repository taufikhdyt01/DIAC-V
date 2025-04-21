# DIAC-V Enterprise Platform

DIAC-V adalah software ERP (Enterprise Resource Planning) berbasis Python untuk desktop yang menyediakan antarmuka untuk mengelola berbagai departemen/grup bisnis dengan integrasi Excel.

## Fitur Utama

- Login dan pengelolaan hak akses multi-level
- Dashboard interaktif
- Integrasi dengan Excel untuk setiap departemen
- Terorganisir berdasarkan grup/departemen:
  - ADE Group
  - BDU Group
  - MAR Group
  - MAN Group
  - PRJ Group
  - FIN Group
  - LEG Group
- Fitur backup otomatis
- Antarmuka modern dan responsif

## Persyaratan Sistem

- Python 3.7+
- Dependensi dalam requirements.txt

## Instalasi

1. Clone atau download repositori ini
2. Buat virtual environment:
   ```
   python -m venv venv
   ```
3. Aktifkan virtual environment:
   - Windows:
     ```
     venv\Scripts\activate
     ```
   - macOS/Linux:
     ```
     source venv/bin/activate
     ```
4. Install dependensi:
   ```
   pip install -r requirements.txt
   ```

## Cara Menggunakan

1. Jalankan aplikasi:
   ```
   python main.py
   ```
2. Pada pertama kali, aplikasi akan membuat struktur folder dan file database default
3. Login dengan kredensial berikut:
   - Username: `admin`
   - Password: `admin123`
   - Atau gunakan kredensial departemen seperti `mary_bdu/password123`
4. Setelah login, dashboard akan menampilkan semua departemen
5. Klik pada kartu departemen untuk mengakses modul-modul di departemen tersebut

## Struktur Folder

```
DIAC-V/
│
├── assets/            # Aset aplikasi (logo, ikon, dll)
│   ├── icons/
│   ├── styles/
│   └── fonts/
│
├── data/              # Data Excel untuk setiap departemen
│   ├── users.xlsx     # Data pengguna
│   ├── ade/
│   ├── bdu/
│   └── ...
│
├── modules/           # Modul aplikasi
│   ├── auth.py
│   ├── excel_connector.py
│   └── ...
│
├── views/             # UI komponen
│   ├── login_view.py
│   ├── dashboard_view.py
│   └── ...
│
├── main.py            # Entry point aplikasi
├── config.py          # Konfigurasi aplikasi
└── requirements.txt   # Dependensi Python
```

## Hak Akses

Aplikasi memiliki beberapa level akses:

- **User**: Hanya dapat mengakses departemen mereka sendiri
- **Manager**: Dapat mengakses departemen mereka sendiri dan departemen terkait
- **Director**: Dapat mengakses banyak departemen
- **Admin**: Akses penuh ke semua departemen + fitur admin
- **CEO**: Akses penuh ke semua departemen

## Integrasi Excel

DIAC-V terintegrasi dengan Excel melalui konektor khusus yang memungkinkan:

- Membaca data dari file Excel
- Menyimpan data ke file Excel
- Menambahkan data baru
- Membuat backup otomatis sebelum perubahan

Setiap departemen memiliki modul Excel spesifik yang sesuai dengan fungsinya.

## Pengembangan Lebih Lanjut

Untuk mengembangkan modul-modul spesifik departemen:

1. Buat file view baru di folder `views/` untuk UI departemen
2. Implementasikan logika bisnis di modul baru di folder `modules/`
3. Tambahkan connector ke Excel jika diperlukan
4. Update `dashboard_view.py` untuk menghubungkan ke modul baru

## Troubleshooting

1. **Error saat menjalankan aplikasi**:

   - Pastikan semua dependensi terinstal (`pip install -r requirements.txt`)
   - Periksa error log untuk detail masalah

2. **Kesalahan akses file**:

   - Pastikan folder data/ dan subfoldernya memiliki izin tulis

3. **Login gagal**:
   - Reset database pengguna dengan menghapus file `data/users.xlsx` (aplikasi akan membuat yang baru dengan pengguna default)

## Dukungan

Jika Anda memiliki pertanyaan atau masalah, silakan buat issue di repositori ini atau hubungi pengembang.
