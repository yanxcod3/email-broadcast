#!/usr/bin/env python3
"""
Email Broadcast Script — SMTP
Kirim email massal dari CSV/Excel dengan personalisasi otomatis.
"""

from __future__ import annotations

import csv
import os
import re
import smtplib
import socket
import ssl
import time
import random
from email.header import Header
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formataddr, formatdate, make_msgid
from pathlib import Path

# ============================================================
# Coba muat konfigurasi dari file .env jika ada
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

# SMTP Settings
# Konfigurasi dapat disimpan di file .env atau diset langsung di bawah ini
SMTP_HOST = os.getenv("SMTP_HOST", "")
SMTP_PORT = int(os.getenv("SMTP_PORT") or 587)
SMTP_EMAIL = os.getenv("SMTP_EMAIL", "")
SMTP_PASSWORD = os.getenv("SMTP_PASSWORD", "")
SMTP_USE_TLS = os.getenv("SMTP_USE_TLS", "True").lower() in ("true", "1", "yes")

# File Settings (File data otomatis dideteksi dari folder data/)
TEMPLATE_FILE = "template.txt"                  # File template email (subjek & isi)
OUTPUT_LOG = "send_log.csv"                     # Log pengiriman

# Delay (detik) antar email
DELAY = 5

# ============================================================
# FUNCTIONS
# ============================================================

def load_template(filepath: str) -> tuple[str, str, str]:
    """
    Load sender_name, subject, dan body template dari file.
    
    Format file template (template.txt):
      SENDER_NAME: Nama Pengirim
      SUBJECT: Judul email {Keyword}
      
      Isi email / body...
    """
    path = Path(filepath)
    if not path.exists():
        raise FileNotFoundError(f"File template '{filepath}' tidak ditemukan!")
    
    with open(path, "r", encoding="utf-8-sig") as f:
        content = f.read()

    lines = content.splitlines()
    sender_name = ""
    subject = ""
    body_start_idx = len(lines)
    
    for i, line in enumerate(lines):
        line_stripped = line.strip()
        
        # Baris kosong pertama setelah header menandai awal body
        if not line_stripped:
            if sender_name or subject:
                body_start_idx = i + 1
                break
            else:
                continue
        
        lower_line = line_stripped.lower()
        if lower_line.startswith(("sender_name:", "sender:", "from:", "nama_pengirim:", "nama:")):
            sender_name = line_stripped.split(":", 1)[1].strip()
        elif lower_line.startswith(("subject:", "subjek:")):
            subject = line_stripped.split(":", 1)[1].strip()
        elif line_stripped == "---":
            body_start_idx = i + 1
            break
        else:
            if not subject:
                subject = line_stripped
                body_start_idx = i + 1
            else:
                body_start_idx = i
            break

    body_lines = lines[body_start_idx:]
    if body_lines and body_lines[0].strip() == "---":
        body_lines = body_lines[1:]

    body = "\n".join(body_lines).strip()
    return sender_name, subject, body


def get_available_data_files() -> list[Path]:
    """Mencari semua file CSV / Excel di folder data/."""
    data_dir = Path("data")
    if not data_dir.exists() or not data_dir.is_dir():
        return []
    valid_exts = {".csv", ".xlsx", ".xls"}
    files = [
        f for f in data_dir.iterdir()
        if f.is_file() and f.suffix.lower() in valid_exts and not f.name.startswith("~") and not f.name.startswith(".")
    ]
    files.sort(key=lambda f: f.stat().st_mtime, reverse=True)
    return files


def resolve_data_file(configured_file: str = "") -> str:
    """Deteksi file data target secara otomatis di folder data/."""
    target = (configured_file or os.getenv("CSV_FILE", "")).strip()
    
    # 1. Jika konfigurasi spesifik diisi dan file-nya ada
    if target:
        p = Path(target)
        if p.exists() and p.is_file():
            return str(p)
        alt = Path("data") / target
        if alt.exists() and alt.is_file():
            return str(alt)
    
    # 2. Auto-deteksi file di dalam folder data/
    available = get_available_data_files()
    if not available:
        return target or "data/recipients.csv"
        
    if len(available) == 1:
        auto_file = available[0]
        return str(auto_file)
        
    # Jika ada beberapa file di folder data/, tampilkan pilihan interaktif
    print(f"\n  📁 Ditemukan {len(available)} file data di folder 'data/':")
    for idx, f in enumerate(available, start=1):
        print(f"     [{idx}] {f.name} ({f.stat().st_size} bytes)")
    try:
        choice = input(f"  Pilih file data yang ingin digunakan (1-{len(available)}) [Default 1]: ").strip()
        if choice.isdigit() and 1 <= int(choice) <= len(available):
            return str(available[int(choice) - 1])
    except Exception:
        pass
    return str(available[0])


def load_data(filepath: str = "") -> list[dict]:
    """Load data dari CSV atau Excel dengan auto-detect delimiter dan BOM."""
    actual_path_str = resolve_data_file(filepath) if not filepath or not Path(filepath).exists() else filepath
    path = Path(actual_path_str)
    
    # Fallback pencarian ke folder data/
    if not path.exists():
        alt_path = Path("data") / actual_path_str
        if alt_path.exists():
            path = alt_path
        else:
            print(f"  ❌ ERROR: File data '{actual_path_str}' tidak ditemukan di folder data/!")
            return []
    
    if path.suffix.lower() == ".csv":
        try:
            with open(path, "r", encoding="utf-8-sig", errors="replace") as f:
                sample = f.read(4096)
                f.seek(0)
                # Auto-deteksi pemisah koma vs titik-koma (format Excel regional)
                delim = ";" if sample.count(";") > sample.count(",") else ","
                reader = csv.DictReader(f, delimiter=delim)
                # Bersihkan spasi pada nama kolom dan abaikan baris kosong
                cleaned_data = []
                for file_row_idx, row in enumerate(reader, start=2):
                    if not any(v and str(v).strip() for v in row.values()):
                        continue  # Abaikan baris kosong
                    cleaned_row = {
                        (k.strip() if k else f"col_{idx}"): (v.strip() if isinstance(v, str) else v)
                        for idx, (k, v) in enumerate(row.items())
                    }
                    cleaned_row["_row_num"] = file_row_idx
                    cleaned_data.append(cleaned_row)
                return cleaned_data
        except Exception as e:
            print(f"  ❌ Gagal membaca CSV: {e}")
            return []

    elif path.suffix.lower() in (".xlsx", ".xls"):
        try:
            import openpyxl
            wb = openpyxl.load_workbook(path, data_only=True)
            ws = wb.active
            headers = [
                str(cell.value).strip() if cell.value is not None else f"col_{i}"
                for i, cell in enumerate(ws[1])
            ]
            data = []
            for file_row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
                if any(v is not None and str(v).strip() for v in row):
                    row_dict = {
                        headers[i]: (str(row[i]).strip() if row[i] is not None else "")
                        for i in range(len(headers)) if i < len(row)
                    }
                    row_dict["_row_num"] = file_row_idx
                    data.append(row_dict)
            return data
        except ImportError:
            print("  ❌ ERROR: Modul 'openpyxl' belum terinstall. Jalankan: pip install openpyxl")
            return []
        except Exception as e:
            print(f"  ❌ Gagal membaca file Excel: {e}")
            return []
    else:
        print(f"  ❌ ERROR: Format file tidak didukung: {path.suffix}")
        return []


def create_email(row: dict, sender_name: str, subject_template: str, body_template: str) -> tuple[str, str, str]:
    """Buat email dari data row dan template dengan penggantian placeholder case-insensitive."""
    row_data = dict(row)
    row_data["sender_name"] = sender_name
    
    subject = subject_template
    body = body_template
    
    # Penggantian placeholder case-insensitive (misal {penulis} vs {Penulis})
    # Menggunakan lambda agar aman terhadap backslash dalam data
    for key, value in row_data.items():
        if key is not None and not str(key).startswith("_"):
            clean_k = str(key).strip()
            clean_v = str(value).strip() if value is not None else ""
            pattern = re.compile(rf"\{{{re.escape(clean_k)}\}}", re.IGNORECASE)
            subject = pattern.sub(lambda _: clean_v, subject)
            body = pattern.sub(lambda _: clean_v, body)
    
    # Pencarian kolom email case-insensitive
    raw_email = ""
    for k, v in row_data.items():
        if k and str(k).strip().lower() in ("email", "e-mail", "mail", "surel"):
            raw_email = str(v or "").strip()
            break
    if not raw_email:
        raw_email = str(row_data.get("Email", "")).strip()
        
    # Standarisasi pemisah koma jika ada beberapa email
    to_email = ", ".join(e.strip() for e in raw_email.replace(";", ",").split(",") if e.strip())
    return to_email, subject, body


def send_email(smtp_server: smtplib.SMTP | smtplib.SMTP_SSL, from_email: str, to_email: str, subject: str, body: str, sender_name: str) -> tuple[bool, str]:
    """Kirim 1 email dengan header standar RFC lengkap (Message-ID, Date, MIME) agar lolos filter spam."""
    domain = from_email.split("@")[-1] if "@" in from_email else "example.com"
    
    # Deteksi HTML atau Plain Text
    is_html = any(tag in body.lower() for tag in ("<html", "<p>", "<br>", "<div", "<table", "<body"))
    
    if is_html:
        # Jika HTML, gunakan multipart/alternative dengan teks biasa sebagai fallback
        msg = MIMEMultipart("alternative")
        plain_fallback = re.sub(r"<[^>]+>", "", body).strip()
        msg.attach(MIMEText(plain_fallback, "plain", "utf-8"))
        msg.attach(MIMEText(body, "html", "utf-8"))
    else:
        # Jika Plain Text murni, gunakan MIMEText langsung (hindari empty multipart yang ditandai spam)
        msg = MIMEText(body, "plain", "utf-8")
    
    # Header wajib sesuai standar internet (RFC 5322) untuk lolos filter spam
    msg["From"] = formataddr((sender_name, from_email))
    msg["To"] = to_email
    msg["Subject"] = Header(subject, "utf-8")
    msg["Date"] = formatdate(localtime=True)
    msg["Message-ID"] = make_msgid(domain=domain)
    msg["User-Agent"] = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Thunderbird/115.0"
    msg["X-Mailer"] = "Thunderbird 115.0"
    
    try:
        smtp_server.send_message(msg)
        return True, ""
    except (smtplib.SMTPServerDisconnected, smtplib.SMTPConnectError, socket.error):
        # Re-raise agar blok auto-reconnect di pemanggil aktif
        raise
    except Exception as e:
        return False, str(e)


def init_smtp() -> smtplib.SMTP | smtplib.SMTP_SSL:
    """Inisialisasi koneksi SMTP dengan smart fallback."""
    context = ssl.create_default_context()
    
    if SMTP_PORT == 465 or not SMTP_USE_TLS:
        try:
            server = smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, context=context, timeout=12)
        except (ConnectionResetError, ConnectionRefusedError, ssl.SSLError, socket.error) as conn_err:
            print(f"    ℹ️  Port 465 SSL ditolak jaringan/server ({conn_err}).")
            print("       Mencoba beralih otomatis ke Port 587 (STARTTLS)...")
            server = smtplib.SMTP(SMTP_HOST, 587, timeout=12)
            server.ehlo()
            server.starttls(context=context)
            server.ehlo()
    else:
        server = smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=12)
        if SMTP_USE_TLS:
            server.ehlo()
            server.starttls(context=context)
            server.ehlo()
            
    server.login(SMTP_EMAIL, SMTP_PASSWORD)
    return server


def log_result(log_file: str, row_num: int, email: str, status: str, subject: str):
    """Log hasil pengiriman ke CSV."""
    file_path = Path(log_file)
    needs_header = not file_path.exists() or file_path.stat().st_size == 0
    with open(file_path, "a", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        if needs_header:
            writer.writerow(["row", "email", "status", "subject", "timestamp"])
        writer.writerow([row_num, email, status, subject, time.strftime("%Y-%m-%d %H:%M:%S")])


def get_sent_emails(log_file: str) -> set[str]:
    """Membaca log pengiriman untuk mendeteksi email yang sudah sukses terkirim."""
    path = Path(log_file)
    if not path.exists() or path.stat().st_size == 0:
        return set()
    
    sent = set()
    try:
        with open(path, "r", encoding="utf-8-sig", errors="replace") as f:
            reader = csv.DictReader(f)
            for row in reader:
                if row.get("status", "").strip().lower() == "success":
                    email = row.get("email", "").strip().lower()
                    if email:
                        sent.add(email)
    except Exception:
        pass
    return sent


def get_row_email(row: dict) -> str:
    """Ambil alamat email dari baris data secara case-insensitive."""
    for k, v in row.items():
        if k and str(k).strip().lower() in ("email", "e-mail", "mail", "surel"):
            return str(v or "").strip().lower()
    return str(row.get("Email", "")).strip().lower()


def is_rate_limit_error(error_msg: str) -> bool:
    """Deteksi apakah pesan error disebabkan oleh batas kuota / rate limit server SMTP."""
    keywords = [
        "limit", "quota", "too many", "try again later", "exceeded", 
        "rate", "max allowed", "hourly", "daily", "451", "452", "421", "550 5.7"
    ]
    lower = error_msg.lower()
    return any(kw in lower for kw in keywords)


# ============================================================
# MENU ACTIONS
# ============================================================

def menu_check_connection():
    """Menu 1: Cek Koneksi SMTP."""
    protocol = "SSL" if (SMTP_PORT == 465 or not SMTP_USE_TLS) else "STARTTLS"
    print("\n" + "=" * 60)
    print("  [1] CEK KONEKSI SMTP")
    print("=" * 60)
    print(f"  Host      : {SMTP_HOST}:{SMTP_PORT} ({protocol})")
    print(f"  Akun      : {SMTP_EMAIL}")
    print("  Status    : Menghubungkan dan memverifikasi autentikasi...")
    
    try:
        server = init_smtp()
        print("  Hasil     : ✅ Koneksi & Autentikasi Berhasil!")
        server.quit()
    except Exception as e:
        print("  Hasil     : ❌ Gagal terhubung!")
        print(f"  Detail    : {e}")


def menu_detail_template():
    """Menu 2: Detail Template & Data Target."""
    print("\n" + "=" * 60)
    print("  [2] DETAIL TEMPLATE & DATA TARGET")
    print("=" * 60)
    
    # 1. Detail Template
    print(f"  File Template : {TEMPLATE_FILE}")
    try:
        sender_name, subject_template, body_template = load_template(TEMPLATE_FILE)
        print(f"  Pengirim      : {sender_name or '(tidak diset)'}")
        print(f"  Subjek        : {subject_template}")
        print("\n  --- Cuplikan Isi Email ---")
        lines = body_template.strip().splitlines()
        for line in lines[:8]:
            print(f"  | {line}")
        if len(lines) > 8:
            print(f"  | ... ({len(lines) - 8} baris lainnya)")
        print("  " + "-" * 40)
    except Exception as e:
        print(f"  ❌ ERROR Template: {e}")
        return

    # 2. Detail Data & Status Log
    data_file = resolve_data_file()
    print(f"\n  File Data     : {data_file}")
    data = load_data(data_file)
    if not data:
        print("  ❌ ERROR: Data CSV kosong atau file tidak ditemukan!")
        return
        
    sent_emails = get_sent_emails(OUTPUT_LOG)
    already_sent = sum(1 for r in data if get_row_email(r) in sent_emails)
    pending_count = len(data) - already_sent

    columns_display = [k for k in data[0].keys() if not k.startswith("_")]
    print(f"  Total Target  : {len(data)} penerima")
    print(f"  Kolom CSV     : {', '.join(columns_display)}")
    print(f"  Sudah Terkirim: {already_sent} penerima (status 'success' di {OUTPUT_LOG})")
    print(f"  Sisa Antrean  : {pending_count} penerima belum terkirim")

    # 3. Contoh Render Email Pertama
    print("\n  --- Preview Email Pertama (Hasil Personalisasi) ---")
    to_email, subject, body = create_email(data[0], sender_name, subject_template, body_template)
    print(f"  Kepada  : {to_email}")
    print(f"  Subjek  : {subject}")
    print("  Isi     :")
    for line in body.splitlines()[:6]:
        print(f"    {line}")
    if len(body.splitlines()) > 6:
        print("    ...")
    print("  " + "-" * 50)


def menu_send_broadcast():
    """Menu 3: Send Broadcast dengan Auto-Resume & Deteksi Limit."""
    print("\n" + "=" * 60)
    print("  [3] SEND BROADCAST")
    print("=" * 60)
    
    # Load template
    try:
        sender_name, subject_template, body_template = load_template(TEMPLATE_FILE)
    except Exception as e:
        print(f"  ❌ Gagal membaca template: {e}")
        return

    # Load data (otomatis deteksi dari folder data/)
    data_file = resolve_data_file()
    data = load_data(data_file)
    if not data:
        print("  ❌ Gagal membaca data CSV!")
        return

    # Cek riwayat pengiriman sebelumnya di send_log.csv
    sent_emails = get_sent_emails(OUTPUT_LOG)
    pending_data = [row for row in data if get_row_email(row) not in sent_emails]
    already_sent_count = len(data) - len(pending_data)

    print(f"  File Target  : {data_file} (Total: {len(data)} baris)")
    print(f"  Pengirim     : {sender_name}")
    print(f"  Subjek       : {subject_template}")
    print(f"  Delay        : {DELAY} detik / email")
    print(f"  Log Hasil    : {OUTPUT_LOG}")
    print("-" * 60)

    target_queue = data

    # Jika ada data yang sudah pernah sukses terkirim sebelumnya
    if already_sent_count > 0:
        print(f"  ℹ️  Ditemukan riwayat pengiriman di '{OUTPUT_LOG}':")
        print(f"     • Sudah sukses terkirim : {already_sent_count} penerima")
        print(f"     • Sisa antrean baru     : {len(pending_data)} penerima")
        print("-" * 60)

        if not pending_data:
            print("  🎉 SEMUA penerima dalam file data sudah berhasil terkirim sebelumnya!")
            print("  Pilihan:")
            print("    [1] Kirim ulang semua data dari awal")
            print("    [0] Kembali ke menu utama")
            ulang = input("  Pilih tindakan (1/0): ").strip()
            if ulang == "1":
                target_queue = data
            else:
                return
        else:
            print("  Pilihan Pengiriman:")
            print(f"    [1] Lanjutkan sisa antrean ({len(pending_data)} penerima) [Rekomendasi]")
            print(f"    [2] Kirim ulang seluruh data ({len(data)} penerima)")
            print("    [0] Batal")
            pilih_mode = input("  Pilih mode (1/2/0): ").strip()
            if pilih_mode == "1":
                target_queue = pending_data
            elif pilih_mode == "2":
                target_queue = data
            else:
                print("  ❌ Pengiriman dibatalkan.")
                return
    else:
        konfirmasi = input(f"  Yakin ingin mengirim email ke {len(data)} penerima? (y/n): ").strip().lower()
        if konfirmasi != "y":
            print("  ❌ Pengiriman dibatalkan oleh pengguna.")
            return

    print(f"\n  Menghubungkan ke SMTP server...")
    try:
        server = init_smtp()
        print(f"  ✅ Terhubung ke SMTP. Memulai proses broadcast ({len(target_queue)} email)...")
    except Exception as e:
        print(f"  ❌ Gagal terhubung ke SMTP: {e}")
        return

    print("-" * 60)
    success_count = 0
    fail_count = 0
    
    try:
        for i, row in enumerate(target_queue, start=1):
            to_email, subject, body = create_email(row, sender_name, subject_template, body_template)
            recipient_info = row.get("Penulis") or row.get("Nama") or row.get("Name") or ""
            recipient_display = f"{to_email} ({recipient_info})" if recipient_info else to_email
            
            row_num = row.get("_row_num", i)
            
            if not to_email:
                print(f"  [{i}/{len(target_queue)}] ⚠️ SKIP — Alamat email kosong (Baris #{row_num})")
                fail_count += 1
                log_result(OUTPUT_LOG, row_num, "", "skipped (empty email)", subject)
                continue
            
            print(f"  [{i}/{len(target_queue)}] (Baris CSV #{row_num}) Mengirim ke : {recipient_display}")
            print(f"        Subjek      : {subject}")
            
            # Coba kirim email dengan auto-reconnect jika koneksi idle/drop
            try:
                success, error_msg = send_email(server, SMTP_EMAIL, to_email, subject, body, sender_name)
            except (smtplib.SMTPServerDisconnected, smtplib.SMTPConnectError, socket.error):
                print("        ⚠️ Koneksi terputus. Mencoba reconnect otomatis...")
                try:
                    server = init_smtp()
                    success, error_msg = send_email(server, SMTP_EMAIL, to_email, subject, body, sender_name)
                except Exception as rec_err:
                    success, error_msg = False, f"Reconnect failed: {rec_err}"
            
            if success:
                print("        Status      : ✅ Berhasil terkirim")
                success_count += 1
                log_result(OUTPUT_LOG, row_num, to_email, "success", subject)
            else:
                print(f"        Status      : ❌ Gagal ({error_msg})")
                fail_count += 1
                log_result(OUTPUT_LOG, row_num, to_email, f"failed: {error_msg}", subject)
                
                # Deteksi jika server menolak karena kena limit/kuota
                if is_rate_limit_error(error_msg):
                    print("\n  " + "!" * 58)
                    print("  ⚠️ TERDETEKSI LIMIT / KUOTA PENGIRIMAN DARI SERVER!")
                    print(f"     Pesan: {error_msg}")
                    print("  " + "!" * 58)
                    print("  Pengiriman dijeda otomatis untuk melindungi reputasi email Anda.")
                    print("  Data yang sudah berhasil terkirim TETAP TERSIMPAN AMAN di log.")
                    print("  Pilihan:")
                    print("    [1] Tunggu 10 menit lalu coba lanjutkan otomatis")
                    print("    [2] Hentikan sekarang (Nanti bisa lanjut otomatis kapan saja)")
                    pilihan_limit = input("  Pilih tindakan (1/2): ").strip()
                    
                    if pilihan_limit == "1":
                        print("  Menunggu 10 menit (600 detik)...")
                        for rem in range(600, 0, -30):
                            print(f"  Tersisa {rem // 60} menit {rem % 60} detik...")
                            time.sleep(30)
                        try:
                            server = init_smtp()
                            print("  ✅ Re-koneksi berhasil. Melanjutkan pengiriman...")
                        except Exception as e_recon:
                            print(f"  ❌ Re-koneksi gagal ({e_recon}). Menghentikan proses.")
                            break
                    else:
                        print("  Proses dihentikan secara aman.")
                        break
            
            # Delay antar email (kecuali email terakhir) dengan sedikit jeda acak agar natural
            if i < len(target_queue):
                sleep_sec = DELAY + random.uniform(1.0, 3.5)
                time.sleep(sleep_sec)
            print()
            
    except KeyboardInterrupt:
        print("\n  ⚠️ Pengiriman dihentikan oleh pengguna (Ctrl+C).")
    finally:
        try:
            server.quit()
        except Exception:
            pass
    
    # Summary
    print("=" * 60)
    print("  SUMMARY PENGIRIMAN")
    print("=" * 60)
    print(f"  Diproses sesi ini : {len(target_queue)}")
    print(f"  ✅ Berhasil       : {success_count}")
    print(f"  ❌ Gagal / Dilewati: {fail_count}")
    print(f"  📁 File Log       : {OUTPUT_LOG}")
    print("=" * 60)


# ============================================================
# MAIN
# ============================================================

def main():
    while True:
        print("\n" + "=" * 60)
        print("           EMAIL BROADCAST SYSTEM — SMTP")
        print("=" * 60)
        print("PILIHAN MENU:")
        print("  [1] Cek Koneksi SMTP")
        print("  [2] Detail Template & Data Target")
        print("  [3] Send Broadcast")
        print("  [0] Keluar")
        print("=" * 60)
        
        try:
            pilihan = input("Pilih menu (0-3): ").strip()
        except (KeyboardInterrupt, EOFError):
            print("\nProgram dihentikan.")
            break
            
        if pilihan == "1":
            menu_check_connection()
        elif pilihan == "2":
            menu_detail_template()
        elif pilihan == "3":
            menu_send_broadcast()
        elif pilihan in ("0", "exit", "q"):
            print("\nTerima kasih. Program selesai.")
            break
        else:
            print("\n⚠️ Pilihan tidak valid! Masukkan angka 0 - 3.")


if __name__ == "__main__":
    main()
