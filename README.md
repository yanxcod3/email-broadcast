# 📧 Email Broadcast System — SMTP

Script automasi pengiriman email massal (_broadcast_) berbasis Python dan SMTP dengan fitur personalisasi otomatis, menu terminal interaktif, pencegahan email ganda (_auto-resume_), serta penanganan cerdas terhadap batas kuota (_rate limit_).

---

## 🌟 Fitur Utama

- **Menu Interaktif Terminal**: Memudahkan pengujian koneksi, pengecekan pratinjau template, hingga eksekusi pengiriman dengan konfirmasi keamanan.
- **Template Terpisah ([template.txt](template.txt))**: Subjek, nama pengirim, dan isi email berada di file tersendiri sehingga tidak perlu mengotak-atik skrip Python.
- **Personalisasi Fleksibel**: Otomatis mengganti placeholder `{Nama_Kolom}` sesuai data CSV/Excel (_case-insensitive_ & aman terhadap rumus/karakter khusus).
- **Auto-Resume & Anti-Duplikasi**: Jika pengiriman terhenti di tengah jalan (misal kena limit atau koneksi putus), skrip dapat otomatis melanjutkan pengiriman ke sisa antrean tanpa mengirim ulang ke email yang sudah berhasil.
- **Deteksi Limit / Kuota Server**: Mendeteksi otomatis jika server SMTP membatasi jumlah pengiriman (_rate limit / quota exceeded_) dan menawarkan opsi jeda aman.
- **Auto-Reconnect & Fallback Port**: Otomatis menyambungkan ulang jika koneksi terputus saat jeda antar email, serta mendukung Port 587 (STARTTLS) maupun Port 465 (SSL).
- **Mendukung CSV & Excel**: Mendukung format `.csv`, `.xlsx`, dan `.xls` dengan deteksi otomatis pemisah koma (`,`) maupun titik-koma (`;`).
- **Pencatatan Log Lengkap ([send_log.csv](send_log.csv))**: Mencatat setiap status pengiriman (_success_, _failed_, atau _skipped_) lengkap dengan nomor baris asli CSV dan waktu pengiriman.

---

## 📁 Struktur Direktori

```text
email-broadcast/
├── data/
│   └── recipients.csv        # File data target penerima (CSV/Excel)
├── template.txt              # File template email (Pengirim, Subjek, Isi)
├── broadcast.py              # Skrip utama email broadcast
├── send_log.csv              # Log hasil pengiriman (otomatis dibuat)
└── README.md                 # Dokumentasi proyek
```

---

## ⚙️ Persyaratan Sistem

- **Python 3.9+** (sudah kompatibel dengan Python modern)
- **Library Tambahan** (opsional, hanya jika menggunakan file `.xlsx`):
  ```powershell
  pip install openpyxl python-dotenv
  ```

---

## 🛠️ Panduan Konfigurasi

Buka file **[.env](.env)** (salin dari `.env.example`) atau sesuaikan di **[broadcast.py](broadcast.py)**:

```python
# SMTP Settings
SMTP_HOST = "mail.yourdomain.com"           # Outgoing mail server
SMTP_PORT = 587                             # 587 untuk TLS (direkomendasikan), 465 untuk SSL
SMTP_EMAIL = "your-email@yourdomain.com"    # Alamat email / username Anda
SMTP_PASSWORD = "PasswordEmailAnda"         # Password email akun Anda
SMTP_USE_TLS = True                         # True untuk port 587, False untuk 465

# File Settings
CSV_FILE = "data/recipients.csv"            # Lokasi file data target
TEMPLATE_FILE = "template.txt"              # Lokasi file template
OUTPUT_LOG = "send_log.csv"                 # Nama file log hasil pengiriman

# Delay (detik) antar email untuk menghindari filter spam
DELAY = 5
```

---

## 📝 Format File Template (`template.txt`)

Isi file **[template.txt](template.txt)** diatur dengan format sederhana:

```text
SENDER_NAME: Tim Informasi
SUBJECT: Informasi Penting untuk {Nama}

Halo {Nama},

Semoga Anda dalam keadaan sehat dan sukses selalu.

Email ini dikirimkan khusus untuk Anda terkait {Topik}. Kami ingin menginformasikan bahwa informasi terbaru untuk Anda telah diperbarui.

Berikut detail informasi Anda:
- Nama Penerima: {Nama}
- Keterangan: {Keterangan}

Jika Anda memiliki pertanyaan lebih lanjut, jangan ragu untuk membalas email ini.

Terima kasih atas perhatian dan kerja sama Anda.

Salam hangat,

Tim Layanan & Operasional
Website: https://example.com
```

> **Catatan & Cara Kerja Placeholder:**
>
> - **Header Template**:
>   - `SENDER_NAME:` mengatur nama pengirim email (dan otomatis mengisi tag `{sender_name}` di dalam isi email).
>   - `SUBJECT:` mengatur judul/subjek email (juga mendukung placeholder seperti `{Nama}`).
> - **Placeholder Dinamis `{Nama_Kolom}`**:
>   - Teks di dalam tanda `{...}` bersifat dinamis dan akan otomatis digantikan dengan data pada **baris dan nama kolom CSV/Excel** yang bersangkutan (*case-insensitive*).
>   - *Contoh*: Jika di CSV Anda ada kolom `Nama` dan `Keterangan`, gunakan `{Nama}` dan `{Keterangan}` di template. Jika ada kolom `Instansi`, Anda bisa langsung menulis `{Instansi}`.
>   - Deteksi kolom email penerima dilakukan otomatis mencari kolom bernama `Email`, `E-mail`, `Mail`, atau `Surel`.
> - **Format Body**: Baris setelah baris kosong pertama adalah isi email (_body_). Mendukung teks polos (*plain text*) maupun tag HTML (seperti `<p>`, `<br>`, `<b>`, `<a>`).

---

## 🚀 Cara Menjalankan

Jalankan skrip melalui terminal atau PowerShell:

```powershell
python broadcast.py
```

Terminal akan menampilkan menu interaktif:

```text
============================================================
           EMAIL BROADCAST SYSTEM — SMTP
============================================================
PILIHAN MENU:
  [1] Cek Koneksi SMTP
  [2] Detail Template & Data Target
  [3] Send Broadcast
  [0] Keluar
============================================================
Pilih menu (0-3):
```

### 1. `[1] Cek Koneksi SMTP`

Menguji konektivitas dan autentikasi username & password ke server SMTP tanpa mengirimkan email. Sangat berguna untuk memastikan konfigurasi sudah benar sebelum mulai _broadcast_.

### 2. `[2] Detail Template & Data Target`

Menampilkan rincian template email, nama pengirim, status file data target, berapa email yang sudah pernah terkirim sebelumnya, serta **pratinjau (_live preview_) hasil render email pertama**.

### 3. `[3] Send Broadcast`

Memulai proses pengiriman:

- Menampilkan konfirmasi jumlah penerima terlebih dahulu.
- Jika ditemukan riwayat di `send_log.csv`, sistem akan memberi pilihan untuk **hanya melanjutkan sisa antrean yang belum terkirim** atau mengulang dari awal.
- Mengirim email satu per satu sesuai jeda waktu `DELAY`.
- Jika terdeteksi _limit/quota_ dari server, script otomatis menjeda proses dan memberi opsi untuk menunggu atau menghentikan secara aman.

---

## 📊 File Log Pengiriman (`send_log.csv`)

Setiap email yang diproses akan dicatat secara _real-time_ ke file **[send_log.csv](send_log.csv)** dengan format kolom standar yang rapi dibuka di Microsoft Excel:

| row | email                     |             status              | subject                                         |      timestamp      |
| :-: | :------------------------ | :-----------------------------: | :---------------------------------------------- | :-----------------: |
|  2  | alex@example.com          |             success             | Informasi Penting untuk Alex                    | 2026-09-28 15:10:05 |
|  3  | budi@example.com          |             success             | Informasi Penting untuk Budi                    | 2026-09-28 15:10:10 |
|  4  | user.invalid@example.com  | failed: (550, 'User not found') | Informasi Penting untuk Citra                   | 2026-09-28 15:10:15 |

---

## 💡 Tips & Rekomendasi Penggunaan

1. **Jeda Antar Email (`DELAY`)**: Disarankan minimal `5` detik per email agar server pengirim maupun penyedia penerima (Gmail, Yahoo, dll.) tidak menandai email Anda sebagai spam.
2. **Kena Limit Kuota Server**: Jika terkena limit (misal kuota harian hosting habis), pilih opsi berhenti aman. Begitu kuota di-reset esok hari, cukup jalankan `python broadcast.py` dan pilih **`[3] Send Broadcast` ➔ `[1] Lanjutkan sisa antrean`**.
3. **Mengulang Broadcast dari Nol**: Jika ingin mengulang pengiriman ke semua orang dari awal, Anda cukup menghapus atau mengganti nama file `send_log.csv`.
