from flask import Blueprint, request, jsonify, Response, render_template_string
from google_services import GoogleServiceProvider
import config
import datetime
from datetime import timezone, timedelta
import base64
import io
import re
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# Inisialisasi Blueprint
dokumentasi_bp = Blueprint('dokumentasi_bp', __name__)

# Kita akan menggunakan instance provider yang diteruskan nanti atau buat baru jika perlu
# Namun, cara terbaik di Flask pattern ini adalah import instance dari app atau buat baru di sini.
# Untuk konsistensi dengan sparta-backend, kita buat instance lokal atau import.
# Asumsi: Kita buat instance baru disini atau gunakan yang ada.

# --- PERBAIKAN: Inisialisasi Provider SEKALI SAJA di sini ---
try:
    print("🔄 Inisialisasi GoogleServiceProvider untuk Dokumentasi API...")
    provider = GoogleServiceProvider()
    print("✅ GoogleServiceProvider Dokumentasi API Siap.")
except Exception as e:
    print(f"❌ Gagal inisialisasi GoogleServiceProvider di Dokumentasi API: {e}")
    provider = None


# Helper date
def to_ymd(val):
    if not val: return ""
    try:
        # Hapus Z jika ada
        val_str = str(val).replace("Z","")
        # Parse ISO
        dt = datetime.datetime.fromisoformat(val_str)
        return dt.strftime("%Y-%m-%d")
    except:
        # Fallback regex yyyy-mm-dd
        m = re.match(r"^(\d{4}-\d{2}-\d{2})", str(val))
        return m.group(1) if m else str(val)

def extract_file_id_from_url(url: str):
    if not url: return None
    m = re.search(r"/d/([A-Za-z0-9_\-]+)", url)
    if m: return m.group(1)
    m = re.search(r"[?&]id=([A-Za-z0-9_\-]+)", url)
    if m: return m.group(1)
    if re.fullmatch(r"[A-Za-z0-9_\-]{20,}", url):
        return url
    return None

def drive_file_public_url(file_id):
    return f"https://drive.google.com/uc?export=view&id={file_id}"

# --- ROUTES ---

@dokumentasi_bp.route('/doc/auth/login', methods=['POST'])
def doc_login():
    
    data = request.get_json()
    username = (data.get("username") or "").strip()
    password = (data.get("password") or "").strip()

    # Baca sheet Cabang (menggunakan fungsi existing sparta karena sheet Cabang sama)
    # Tapi main.py baca manual. Kita gunakan provider.
    try:
        # Menggunakan sheet Cabang dari config utama Sparta
        rows = provider.sheet.worksheet(config.CABANG_SHEET_NAME).get_all_values()
    except Exception as e:
        return jsonify({"ok": False, "message": f"Error Sheet: {str(e)}"}), 500

    if not rows:
        return jsonify({"ok": False, "message": "Sheet Cabang Kosong"}), 400

    headers = rows[0]
    try:
        email_idx = headers.index("EMAIL_SAT")
        cabang_idx = headers.index("CABANG")
        jabatan_idx = headers.index("JABATAN")
        nama_idx = headers.index("NAMA LENGKAP")
    except ValueError:
        return jsonify({"ok": False, "message": "Header Sheet Cabang tidak sesuai"}), 500

    ALLOWED_ROLES = ["BRANCH BUILDING SUPPORT", "BRANCH BUILDING COORDINATOR", "BRANCH BUILDING & MAINTENANCE MANAGER"]

    for r in rows[1:]:
        if len(r) > max(email_idx, cabang_idx, jabatan_idx, nama_idx):
            if str(r[email_idx]).strip() == username and str(r[cabang_idx]).strip() == password:
                user_jabatan = str(r[jabatan_idx]).strip().upper()
                if user_jabatan in ALLOWED_ROLES:
                    # Log login bisa ditambahkan logicnya ke provider jika perlu
                    return jsonify({"ok": True, "message": "Login berhasil", "user": {
                        "email": r[email_idx], "cabang": r[cabang_idx], 
                        "nama": r[nama_idx], "jabatan": r[jabatan_idx]
                    }})
                else:
                    return jsonify({"ok": False, "message": "Jabatan tidak diizinkan"}), 403
    
    return jsonify({"ok": False, "message": "Username atau password salah"}), 401

@dokumentasi_bp.route('/doc/spk-data', methods=['POST'])
def doc_spk_data():
    data = request.get_json()
    cabang_filter = (data.get("cabang") or "").strip()
    
    # Baca Sheet SPK_DATA (Gunakan fungsi existing Sparta / Provider)
    rows = provider.dokumentasi_read_sheet(config.SPK_DATA_SHEET_NAME) # Jika SPK_DATA ada di sheet Dokumentasi
    # JIKA SPK_DATA ada di Spreadsheet UTAMA Sparta, ganti doc_read_sheet jadi:
    # rows = provider.sheet.worksheet(config.SPK_DATA_SHEET_NAME).get_all_values()
    
    if not rows: return jsonify({"ok": True, "data": []})
    
    headers = rows[0]
    # Helper index safe
    def idx(name): return headers.index(name) if name in headers else None
    
    nomor_idx = idx("Nomor Ulok")
    cabang_idx = idx("Cabang")
    sipil_idx = idx("Nama Kontraktor")
    
    # Logic cari index ME (seperti di main.py)
    me_idx = sipil_idx
    contractor_cols = [i for i, h in enumerate(headers) if h == "Nama Kontraktor"]
    if len(contractor_cols) >= 2:
        me_idx = contractor_cols[-1] # Ambil yang terakhir
    
    awal_idx = idx("Waktu Mulai")
    akhir_idx = idx("Waktu Selesai")
    nama_toko_idx = idx("Nama_Toko")

    out = []
    for r in rows[1:]:
        if not r: continue
        # Filter cabang
        curr_cabang = r[cabang_idx] if cabang_idx is not None and len(r) > cabang_idx else ""
        if cabang_filter and str(curr_cabang).strip() != cabang_filter:
            continue
            
        out.append({
            "nomorUlok": r[nomor_idx] if nomor_idx is not None and len(r) > nomor_idx else "",
            "cabang": curr_cabang,
            "kontraktorSipil": r[sipil_idx] if sipil_idx is not None and len(r) > sipil_idx else "",
            "kontraktorMe": r[me_idx] if me_idx is not None and len(r) > me_idx else "",
            "spkAwal": to_ymd(r[awal_idx]) if awal_idx is not None and len(r) > awal_idx else "",
            "spkAkhir": to_ymd(r[akhir_idx]) if akhir_idx is not None and len(r) > akhir_idx else "",
            "namaToko": r[nama_toko_idx] if nama_toko_idx is not None and len(r) > nama_toko_idx else "",
        })
    return jsonify({"ok": True, "data": out})

@dokumentasi_bp.route('/doc/view-photo/<file_id>', methods=['GET'])
def doc_view_photo(file_id):
    stream = provider.dokumentasi_get_file_stream(file_id)
    if stream:
        return Response(stream.read(), mimetype="image/jpeg")
    return "Not Found", 404

@dokumentasi_bp.route('/doc/save-temp', methods=['POST'])
def doc_save_temp():
    data = request.get_json()
    nomor_ulok = (data.get("nomorUlok") or "").strip()
    if not nomor_ulok: return jsonify({"ok": False, "error": "nomorUlok required"}), 400

    sheet_name = config.DOC_SHEET_NAME_TEMP
    rows = provider.dokumentasi_read_sheet(sheet_name)
    
    # Header handling (Jika kosong buat header)
    if not rows:
        headers = [
            "Timestamp", "Nomor Ulok", "Nama Toko", "Kode Toko", "Cabang",
            "Tanggal GO", "Tanggal ST", "Tanggal Ambil Foto", "SPK Awal", "SPK Akhir",
            "Kontraktor Sipil", "Kontraktor ME", "Email Pengirim"
        ] + [f"Photo{i}" for i in range(1, 39)]
        provider.dokumentasi_append_row(sheet_name, headers)
        rows = [headers]

    headers = rows[0]
    try:
        nomor_idx = headers.index("Nomor Ulok")
    except:
        return jsonify({"ok": False, "error": "Header Sheet Temp Salah"}), 500

    # Cari baris lama
    found_row_idx = None
    old_data = []
    
    # Ingat: rows[0] adalah header. rows[1] adalah data baris ke-2 di excel.
    for i, r in enumerate(rows):
        if i == 0: continue
        if len(r) > nomor_idx and str(r[nomor_idx]).strip() == nomor_ulok:
            found_row_idx = i + 1 # Excel index (1-based)
            old_data = r
            break
            
    if not old_data:
        old_data = [""] * len(headers)

    # Logic Foto
    photo_id = data.get("photoId")
    base64_photo = data.get("photoBase64")
    photo_note = (data.get("photoNote") or "").upper().strip()
    
    photo_cell_value = ""
    
    # A. Upload Base64 (User ambil foto normal)
    if base64_photo and photo_id is not None:
        filename = f"{nomor_ulok}_foto_{photo_id}.jpg"
        fid = provider.dokumentasi_upload_image(base64_photo, filename)
        if fid:
            photo_cell_value = drive_file_public_url(fid)
            
    # =========================================================================
    # B. TIDAK BISA DIFOTO (Logic Baru: Upload file lokal ke Drive)
    # =========================================================================
    elif photo_note == "TIDAK BISA DIFOTO" and photo_id is not None:
        try:
            # Cari path file default di folder static
            # Asumsi struktur: server/dokumentasi_api.py dan server/static/fototidakbisadiambil.jpeg
            current_dir = os.path.dirname(os.path.abspath(__file__))
            default_path = os.path.join(current_dir, "static", "fototidakbisadiambil.jpeg")

            if os.path.exists(default_path):
                # 1. Baca file lokal
                with open(default_path, "rb") as f:
                    raw = f.read()
                
                # 2. Convert ke base64 string
                base64_default = "data:image/jpeg;base64," + base64.b64encode(raw).decode()
                
                # 3. Upload sebagai file baru ke Drive
                filename = f"{nomor_ulok}_foto_{photo_id}.jpg"
                fid = provider.dokumentasi_upload_image(base64_default, filename)
                
                if fid:
                    photo_cell_value = drive_file_public_url(fid)
                else:
                    print("Gagal upload foto default ke Drive")
            else:
                print(f"File default tidak ditemukan di: {default_path}")
                # Fallback ke config jika file lokal hilang
                photo_cell_value = drive_file_public_url(config.DOC_DEFAULT_PHOTO_ID)

        except Exception as e:
            print(f"Error proses foto default: {e}")
            # Fallback terakhir
            photo_cell_value = drive_file_public_url(config.DOC_DEFAULT_PHOTO_ID)
        
    # C. Tidak berubah -> Pakai nilai lama
    elif photo_id is not None:
        idx_old = 12 + int(photo_id) # Offset kolom foto
        if len(old_data) > idx_old:
            photo_cell_value = old_data[idx_old]

    # Construct Row Baru
    new_row = [""] * len(headers)
    
    # Map data fixed
    fixed_map = {
        "Timestamp": datetime.datetime.now().isoformat(),
        "Nomor Ulok": nomor_ulok,
        "Nama Toko": data.get("namaToko", ""),
        "Kode Toko": data.get("kodeToko", ""),
        "Cabang": (data.get("cabang") or "").strip(),
        "Tanggal GO": data.get("tanggalGo", ""),
        "Tanggal ST": data.get("tanggalSt", ""),
        "Tanggal Ambil Foto": data.get("tanggalAmbilFoto", ""),
        "SPK Awal": data.get("spkAwal", ""),
        "SPK Akhir": data.get("spkAkhir", ""),
        "Kontraktor Sipil": data.get("kontraktorSipil", ""),
        "Kontraktor ME": data.get("kontraktorMe", ""),
        "Email Pengirim": data.get("emailPengirim", ""),
    }

    # Isi kolom fixed
    for h, v in fixed_map.items():
        if h in headers:
            idx = headers.index(h)
            # Jika v kosong, pakai data lama
            new_row[idx] = v if v else (old_data[idx] if len(old_data) > idx else "")

    # Isi kolom foto
    for i in range(1, 39):
        idx = 12 + i
        if len(new_row) <= idx: new_row.extend([""] * (idx - len(new_row) + 1))
        
        if photo_id is not None and i == int(photo_id):
            new_row[idx] = photo_cell_value
        else:
             new_row[idx] = old_data[idx] if len(old_data) > idx else ""

    # Simpan
    if found_row_idx:
        provider.dokumentasi_update_row(sheet_name, found_row_idx, new_row)
    else:
        provider.dokumentasi_append_row(sheet_name, new_row)

    return jsonify({"ok": True, "message": "Temp saved"})

@dokumentasi_bp.route('/doc/get-temp', methods=['POST'])
def doc_get_temp():
    data = request.get_json()
    nomor_ulok = (data.get("nomorUlok") or "").strip()
    
    rows = provider.dokumentasi_read_sheet(config.DOC_SHEET_NAME_TEMP)
    if not rows: return jsonify({"ok": False, "message": "Sheet kosong"})
    
    headers = rows[0]
    try:
        nomor_idx = headers.index("Nomor Ulok")
    except:
        return jsonify({"ok": False, "message": "Kolom Nomor Ulok tidak ada"})

    found = None
    for r in rows[1:]:
        if len(r) > nomor_idx and str(r[nomor_idx]).strip() == nomor_ulok:
            found = r
            break
            
    if not found: return jsonify({"ok": True, "data": None})

    # Map result
    data_map = {h: (found[i] if i < len(found) else "") for i, h in enumerate(headers)}
    
    # Ambil Photos
    photo_ids = []
    start_idx = headers.index("Photo1") if "Photo1" in headers else 13
    for i in range(38):
        col_idx = start_idx + i
        raw_val = found[col_idx] if col_idx < len(found) else ""
        if raw_val:
            fid = extract_file_id_from_url(raw_val) or raw_val
            photo_ids.append(fid)
        else:
            photo_ids.append("")
            
    result = {
        "nomorUlok": data_map.get("Nomor Ulok", ""),
        "namaToko": data_map.get("Nama Toko", ""),
        "kodeToko": data_map.get("Kode Toko", ""),
        "cabang": data_map.get("Cabang", ""),
        "tanggalGo": to_ymd(data_map.get("Tanggal GO", "")),
        "tanggalSt": to_ymd(data_map.get("Tanggal ST", "")),
        "tanggalAmbilFoto": to_ymd(data_map.get("Tanggal Ambil Foto", "")),
        "spkAwal": to_ymd(data_map.get("SPK Awal", "")),
        "spkAkhir": to_ymd(data_map.get("SPK Akhir", "")),
        "kontraktorSipil": data_map.get("Kontraktor Sipil", ""),
        "kontraktorMe": data_map.get("Kontraktor ME", ""),
        "emailPengirim": data_map.get("Email Pengirim", ""),
        "photos": photo_ids,
    }
    # Kembalikan full data map juga supaya aman
    result.update(data_map)
    result["photos"] = photo_ids  # Override photos dengan list array

    return jsonify({"ok": True, "data": result})

@dokumentasi_bp.route('/doc/cek-status', methods=['POST'])
def doc_cek_status():
    data = request.get_json()
    ulok = (data.get("nomorUlok") or "").strip()
    
    rows = provider.dokumentasi_read_sheet(config.DOC_SHEET_NAME_FINAL)
    if not rows: return jsonify({"ok": True, "status": "BELUM ADA"})
    
    headers = rows[0]
    try:
        ulok_idx = headers.index("Nomor Ulok")
        status_idx = headers.index("Status Validasi")
    except:
        return jsonify({"ok": True, "status": "BELUM ADA"}) # Default aman

    for r in rows[1:]:
        if len(r) > ulok_idx and str(r[ulok_idx]).strip() == ulok:
            status = r[status_idx] if len(r) > status_idx else ""
            return jsonify({"ok": True, "status": status.upper().strip()})
            
    return jsonify({"ok": True, "status": "BELUM ADA"})

@dokumentasi_bp.route('/doc/save-toko', methods=['POST'])
def doc_save_toko():
    """
    Menyimpan data dokumentasi final ke sheet 'dokumentasi_bangunan'.
    Menerima data toko, foto-foto, dan PDF base64 untuk diupload ke Drive.
    """
    data = request.get_json()
    nomor_ulok = (data.get("nomorUlok") or "").strip()
    if not nomor_ulok:
        return jsonify({"ok": False, "error": "nomorUlok required"}), 400

    sheet_name = config.DOC_SHEET_NAME_FINAL
    rows = provider.dokumentasi_read_sheet(sheet_name)
    
    # Header handling (Jika kosong buat header)
    if not rows:
        headers = [
            "Timestamp", "Nomor Ulok", "Nama Toko", "Kode Toko", "Cabang",
            "Tanggal GO", "Tanggal ST", "Tanggal Ambil Foto", "SPK Awal", "SPK Akhir",
            "Kontraktor Sipil", "Kontraktor ME", "Email Pengirim", "Link PDF",
            "Status Validasi", "Validator", "Waktu Validasi", "Catatan Revisi"
        ] + [f"Photo{i}" for i in range(1, 39)]
        provider.dokumentasi_append_row(sheet_name, headers)
        rows = [headers]

    headers = rows[0]
    try:
        nomor_idx = headers.index("Nomor Ulok")
    except:
        return jsonify({"ok": False, "error": "Header Sheet Final tidak valid"}), 500

    # Cari baris lama
    found_row_idx = None
    old_data = []
    
    for i, r in enumerate(rows):
        if i == 0: 
            continue
        if len(r) > nomor_idx and str(r[nomor_idx]).strip() == nomor_ulok:
            found_row_idx = i + 1  # Excel index (1-based)
            old_data = r
            break
            
    if not old_data:
        old_data = [""] * len(headers)

    # Upload PDF jika ada
    pdf_url = ""
    pdf_base64 = data.get("pdfBase64")
    if pdf_base64:
        pdf_filename = f"{nomor_ulok}_dokumentasi.pdf"
        pdf_file_id = provider.doc_upload_file(pdf_base64, pdf_filename, "application/pdf")
        if pdf_file_id:
            pdf_url = f"https://drive.google.com/file/d/{pdf_file_id}/view"
        else:
            # Fallback: gunakan doc_upload_image (jika doc_upload_file tidak ada)
            pdf_file_id = provider.dokumentasi_upload_image(pdf_base64, pdf_filename)
            if pdf_file_id:
                pdf_url = f"https://drive.google.com/file/d/{pdf_file_id}/view"

    # Upload foto-foto jika ada (array of base64)
    photos_base64 = data.get("photosBase64", [])
    photo_urls = []
    
    for idx, photo_b64 in enumerate(photos_base64):
        if photo_b64:
            photo_filename = f"{nomor_ulok}_foto_{idx + 1}.jpg"
            photo_fid = provider.dokumentasi_upload_image(photo_b64, photo_filename)
            if photo_fid:
                photo_urls.append(drive_file_public_url(photo_fid))
            else:
                photo_urls.append("")
        else:
            photo_urls.append("")

    # Ambil photo URLs dari data jika dikirim sebagai URLs (bukan base64)
    existing_photo_urls = data.get("photoUrls", [])
    
    # Construct Row Baru
    new_row = [""] * len(headers)
    
    # Map data fixed
    fixed_map = {
        "Timestamp": datetime.datetime.now().isoformat(),
        "Nomor Ulok": nomor_ulok,
        "Nama Toko": data.get("namaToko", ""),
        "Kode Toko": data.get("kodeToko", ""),
        "Cabang": (data.get("cabang") or "").strip(),
        "Tanggal GO": data.get("tanggalGo", ""),
        "Tanggal ST": data.get("tanggalSt", ""),
        "Tanggal Ambil Foto": data.get("tanggalAmbilFoto", ""),
        "SPK Awal": data.get("spkAwal", ""),
        "SPK Akhir": data.get("spkAkhir", ""),
        "Kontraktor Sipil": data.get("kontraktorSipil", ""),
        "Kontraktor ME": data.get("kontraktorMe", ""),
        "Email Pengirim": data.get("emailPengirim", ""),
        "Link PDF": pdf_url,
        "Status Validasi": data.get("statusValidasi", "MENUNGGU VALIDASI"),
        "Validator": data.get("validator", ""),
        "Waktu Validasi": data.get("waktuValidasi", ""),
        "Catatan Revisi": data.get("catatanRevisi", ""),
    }

    # Isi kolom fixed
    for h, v in fixed_map.items():
        if h in headers:
            idx = headers.index(h)
            # Jika v kosong, pakai data lama (kecuali Timestamp selalu update)
            if h == "Timestamp":
                new_row[idx] = v
            else:
                new_row[idx] = v if v else (old_data[idx] if len(old_data) > idx else "")

    # Isi kolom foto (mulai dari kolom ke-18 atau sesuai posisi Photo1)
    photo_start_idx = headers.index("Photo1") if "Photo1" in headers else 18
    
    for i in range(38):
        col_idx = photo_start_idx + i
        if col_idx >= len(new_row):
            new_row.extend([""] * (col_idx - len(new_row) + 1))
        
        # Prioritas: photo_urls dari upload baru -> existing_photo_urls -> old_data
        if i < len(photo_urls) and photo_urls[i]:
            new_row[col_idx] = photo_urls[i]
        elif i < len(existing_photo_urls) and existing_photo_urls[i]:
            new_row[col_idx] = existing_photo_urls[i]
        elif len(old_data) > col_idx:
            new_row[col_idx] = old_data[col_idx]
        else:
            new_row[col_idx] = ""

    # Simpan ke sheet
    if found_row_idx:
        provider.dokumentasi_update_row(sheet_name, found_row_idx, new_row)
    else:
        provider.dokumentasi_append_row(sheet_name, new_row)

    # Hapus data temp setelah berhasil simpan final (opsional)
    if data.get("deleteTemp", False):
        try:
            temp_rows = provider.dokumentasi_read_sheet(config.DOC_SHEET_NAME_TEMP)
            for i, r in enumerate(temp_rows):
                if i == 0:
                    continue
                if len(r) > 1 and str(r[1]).strip() == nomor_ulok:
                    # Hapus dengan clear row atau update kosong
                    provider.dokumentasi_update_row(config.DOC_SHEET_NAME_TEMP, i + 1, [""] * len(r))
                    break
        except Exception as e:
            print(f"Warning: Gagal hapus temp data: {e}")

    return jsonify({
        "ok": True, 
        "message": "Data dokumentasi berhasil disimpan",
        "pdfUrl": pdf_url,
        "nomorUlok": nomor_ulok
    })

@dokumentasi_bp.route('/doc/send-pdf-email', methods=['POST'])
def doc_send_email():
    """
    Mengirim email notifikasi dengan PDF dokumentasi ke validator.
    """
    data = request.get_json()
    nomor_ulok = (data.get("nomorUlok") or "").strip()
    cabang = (data.get("cabang") or "").strip()
    nama_toko = data.get("namaToko", "")
    pdf_url = data.get("pdfUrl", "")
    pdf_base64 = data.get("pdfBase64")
    email_pengirim = data.get("emailPengirim", "")
    
    if not nomor_ulok:
        return jsonify({"ok": False, "error": "nomorUlok required"}), 400

    # 1. Ambil list email validator dari Sheet Cabang
    try:
        cabang_rows = provider.sheet.worksheet(config.CABANG_SHEET_NAME).get_all_values()
    except Exception as e:
        return jsonify({"ok": False, "error": f"Gagal membaca Sheet Cabang: {str(e)}"}), 500

    if not cabang_rows:
        return jsonify({"ok": False, "error": "Sheet Cabang kosong"}), 400

    headers = cabang_rows[0]
    try:
        email_idx = headers.index("EMAIL_SAT")
        cabang_idx = headers.index("CABANG")
        jabatan_idx = headers.index("JABATAN")
        nama_idx = headers.index("NAMA LENGKAP")
    except ValueError as e:
        return jsonify({"ok": False, "error": f"Header Sheet Cabang tidak lengkap: {str(e)}"}), 500

    # Cari validator (Building Support Dokumentasi atau Koordinator) di cabang yang sama
    VALIDATOR_ROLES = ["BRANCH BUILDING COORDINATOR", "BRANCH BUILDING & MAINTENANCE MANAGER"]
    validator_emails = []
    
    for r in cabang_rows[1:]:
        if len(r) > max(email_idx, cabang_idx, jabatan_idx):
            row_cabang = str(r[cabang_idx]).strip().upper()
            row_jabatan = str(r[jabatan_idx]).strip().upper()
            row_email = str(r[email_idx]).strip()
            
            # Filter berdasarkan cabang dan jabatan validator
            if row_cabang == cabang.upper() and row_jabatan in VALIDATOR_ROLES:
                if row_email and "@" in row_email:
                    validator_emails.append(row_email)

    if not validator_emails:
        return jsonify({"ok": False, "error": f"Tidak ditemukan validator untuk cabang {cabang}"}), 400

    # 2. Siapkan konten email
    subject = f"[Dokumentasi Bangunan] Permintaan Validasi - {nomor_ulok} - {nama_toko}"
    
    # URL validasi
    base_url = request.host_url.rstrip('/')
    validate_url = f"{base_url}/doc/validate?ulok={nomor_ulok}&status=VALID"
    revisi_url = f"{base_url}/doc/validate?ulok={nomor_ulok}&status=REVISI"
    
    html_body = f"""
    <html>
    <body style="font-family: Arial, sans-serif; line-height: 1.6;">
        <h2 style="color: #2c3e50;">Permintaan Validasi Dokumentasi Bangunan</h2>
        <p>Yth. Validator,</p>
        <p>Berikut adalah detail dokumentasi bangunan yang memerlukan validasi:</p>
        
        <table style="border-collapse: collapse; margin: 20px 0;">
            <tr>
                <td style="padding: 8px; border: 1px solid #ddd; background: #f9f9f9;"><strong>Nomor Ulok</strong></td>
                <td style="padding: 8px; border: 1px solid #ddd;">{nomor_ulok}</td>
            </tr>
            <tr>
                <td style="padding: 8px; border: 1px solid #ddd; background: #f9f9f9;"><strong>Nama Toko</strong></td>
                <td style="padding: 8px; border: 1px solid #ddd;">{nama_toko}</td>
            </tr>
            <tr>
                <td style="padding: 8px; border: 1px solid #ddd; background: #f9f9f9;"><strong>Cabang</strong></td>
                <td style="padding: 8px; border: 1px solid #ddd;">{cabang}</td>
            </tr>
            <tr>
                <td style="padding: 8px; border: 1px solid #ddd; background: #f9f9f9;"><strong>Pengirim</strong></td>
                <td style="padding: 8px; border: 1px solid #ddd;">{email_pengirim}</td>
            </tr>
            <tr>
                <td style="padding: 8px; border: 1px solid #ddd; background: #f9f9f9;"><strong>Link PDF</strong></td>
                <td style="padding: 8px; border: 1px solid #ddd;"><a href="{pdf_url}">{pdf_url}</a></td>
            </tr>
        </table>
        
        <p>Silakan klik tombol di bawah untuk melakukan validasi:</p>
        
        <div style="margin: 20px 0;">
            <a href="{validate_url}" style="background-color: #27ae60; color: white; padding: 12px 24px; text-decoration: none; border-radius: 5px; margin-right: 10px;">
                ✓ VALIDASI
            </a>
            <a href="{revisi_url}" style="background-color: #e74c3c; color: white; padding: 12px 24px; text-decoration: none; border-radius: 5px;">
                ✗ MINTA REVISI
            </a>
        </div>
        
        <p style="color: #7f8c8d; font-size: 12px; margin-top: 30px;">
            Email ini dikirim secara otomatis oleh sistem Dokumentasi Bangunan SPARTA.<br>
            Harap tidak membalas email ini.
        </p>
    </body>
    </html>
    """
    
    # 3. Kirim email ke semua validator
    sent_to = []
    errors = []
    
    for email_to in validator_emails:
        try:
            # Siapkan attachment PDF jika ada base64
            attachment = None
            if pdf_base64:
                clean_b64 = pdf_base64.split(",")[-1] if "," in pdf_base64 else pdf_base64
                attachment = {
                    "filename": f"Dokumentasi_{nomor_ulok}.pdf",
                    "content": clean_b64,
                    "mime_type": "application/pdf"
                }
            
            # Gunakan provider.send_email atau langsung gmail_service
            success = send_email_with_attachment(
                provider.gmail_service,
                email_to,
                subject,
                html_body,
                attachment
            )
            
            if success:
                sent_to.append(email_to)
            else:
                errors.append(f"Gagal kirim ke {email_to}")
                
        except Exception as e:
            errors.append(f"Error kirim ke {email_to}: {str(e)}")

    if sent_to:
        return jsonify({
            "ok": True, 
            "message": f"Email berhasil dikirim ke {len(sent_to)} validator",
            "sentTo": sent_to,
            "errors": errors if errors else None
        })
    else:
        return jsonify({
            "ok": False, 
            "error": "Gagal mengirim email ke semua validator",
            "details": errors
        }), 500


def send_email_with_attachment(gmail_service, to_email, subject, html_body, attachment=None):
    """
    Helper function untuk mengirim email dengan attachment opsional.
    """
    
    try:
        message = MIMEMultipart()
        message['to'] = to_email
        message['subject'] = subject
        
        # Body HTML
        message.attach(MIMEText(html_body, 'html'))
        
        # Attachment jika ada
        if attachment:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(base64.b64decode(attachment['content']))
            encoders.encode_base64(part)
            part.add_header(
                'Content-Disposition',
                f'attachment; filename="{attachment["filename"]}"'
            )
            message.attach(part)
        
        # Encode dan kirim
        raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode('utf-8')
        
        gmail_service.users().messages().send(
            userId='me',
            body={'raw': raw_message}
        ).execute()
        
        return True
    except Exception as e:
        print(f"Error sending email: {e}")
        return False

# Endpoint HTML Validate & Revisi (Render String HTML)
@dokumentasi_bp.route('/doc/validate', methods=['GET'])
def doc_validate():
    """
    Endpoint untuk validasi dokumentasi bangunan.
    Dipanggil dari link di email validator.
    """
    ulok = request.args.get('ulok', '').strip()
    status = request.args.get('status', '').strip().upper()
    catatan = request.args.get('catatan', '').strip()
    
    if not ulok:
        html = """
        <html>
        <body style="font-family: Arial, sans-serif; text-align: center; padding: 50px;">
            <h1 style="color: #e74c3c;">❌ Error</h1>
            <p>Parameter 'ulok' tidak ditemukan.</p>
        </body>
        </html>
        """
        return render_template_string(html), 400

    if status not in ["VALID", "REVISI", "DITOLAK"]:
        html = f"""
        <html>
        <body style="font-family: Arial, sans-serif; text-align: center; padding: 50px;">
            <h1 style="color: #e74c3c;">❌ Error</h1>
            <p>Status '{status}' tidak valid. Gunakan: VALID, REVISI, atau DITOLAK.</p>
        </body>
        </html>
        """
        return render_template_string(html), 400

    # Jika status REVISI, tampilkan form untuk catatan
    if status == "REVISI" and not catatan:
        html = f"""
        <html>
        <head>
            <title>Form Revisi - {ulok}</title>
            <style>
                body {{ font-family: Arial, sans-serif; padding: 50px; max-width: 600px; margin: auto; }}
                h1 {{ color: #2c3e50; }}
                .form-group {{ margin: 20px 0; }}
                label {{ display: block; margin-bottom: 5px; font-weight: bold; }}
                textarea {{ width: 100%; height: 150px; padding: 10px; border: 1px solid #ddd; border-radius: 5px; }}
                button {{ background-color: #e74c3c; color: white; padding: 12px 30px; border: none; border-radius: 5px; cursor: pointer; font-size: 16px; }}
                button:hover {{ background-color: #c0392b; }}
            </style>
        </head>
        <body>
            <h1>📝 Form Permintaan Revisi</h1>
            <p><strong>Nomor Ulok:</strong> {ulok}</p>
            <form method="GET" action="/doc/validate">
                <input type="hidden" name="ulok" value="{ulok}">
                <input type="hidden" name="status" value="REVISI">
                <div class="form-group">
                    <label for="catatan">Catatan Revisi (wajib diisi):</label>
                    <textarea name="catatan" id="catatan" required placeholder="Jelaskan apa yang perlu direvisi..."></textarea>
                </div>
                <button type="submit">Kirim Permintaan Revisi</button>
            </form>
        </body>
        </html>
        """
        return render_template_string(html)

    # Update status di sheet
    try:
        sheet_name = config.DOC_SHEET_NAME_FINAL
        rows = provider.dokumentasi_read_sheet(sheet_name)
        
        if not rows:
            raise Exception("Sheet dokumentasi_bangunan kosong")

        headers = rows[0]
        ulok_idx = headers.index("Nomor Ulok") if "Nomor Ulok" in headers else None
        status_idx = headers.index("Status Validasi") if "Status Validasi" in headers else None
        validator_idx = headers.index("Validator") if "Validator" in headers else None
        waktu_idx = headers.index("Waktu Validasi") if "Waktu Validasi" in headers else None
        catatan_idx = headers.index("Catatan Revisi") if "Catatan Revisi" in headers else None

        if ulok_idx is None or status_idx is None:
            raise Exception("Kolom 'Nomor Ulok' atau 'Status Validasi' tidak ditemukan")

        # Cari baris dengan ulok yang sesuai
        found_row_idx = None
        found_row = None
        
        for i, r in enumerate(rows):
            if i == 0:
                continue
            if len(r) > ulok_idx and str(r[ulok_idx]).strip() == ulok:
                found_row_idx = i + 1  # 1-based index
                found_row = r.copy()
                break

        if not found_row_idx:
            raise Exception(f"Data dengan Nomor Ulok '{ulok}' tidak ditemukan")

        # Pastikan row cukup panjang
        while len(found_row) < len(headers):
            found_row.append("")

        # Update kolom status
        found_row[status_idx] = status
        
        if validator_idx is not None:
            found_row[validator_idx] = request.args.get('validator', 'Email Validator')
            
        if waktu_idx is not None:
            found_row[waktu_idx] = datetime.datetime.now().isoformat()
            
        if catatan_idx is not None and catatan:
            found_row[catatan_idx] = catatan

        # Simpan perubahan
        provider.dokumentasi_update_row(sheet_name, found_row_idx, found_row)

        # Tampilkan halaman sukses
        if status == "VALID":
            html = f"""
            <html>
            <head><title>Validasi Berhasil</title></head>
            <body style="font-family: Arial, sans-serif; text-align: center; padding: 50px;">
                <h1 style="color: #27ae60;">✓ Validasi Berhasil</h1>
                <p>Dokumentasi bangunan dengan Nomor Ulok <strong>{ulok}</strong> telah divalidasi.</p>
                <p>Status: <span style="background: #27ae60; color: white; padding: 5px 15px; border-radius: 3px;">VALID</span></p>
                <p style="margin-top: 30px; color: #7f8c8d;">Anda dapat menutup halaman ini.</p>
            </body>
            </html>
            """
        else:  # REVISI atau DITOLAK
            html = f"""
            <html>
            <head><title>Permintaan Revisi</title></head>
            <body style="font-family: Arial, sans-serif; text-align: center; padding: 50px;">
                <h1 style="color: #e67e22;">📝 Permintaan Revisi Dikirim</h1>
                <p>Dokumentasi bangunan dengan Nomor Ulok <strong>{ulok}</strong> memerlukan revisi.</p>
                <p>Status: <span style="background: #e67e22; color: white; padding: 5px 15px; border-radius: 3px;">{status}</span></p>
                <p><strong>Catatan:</strong> {catatan or '-'}</p>
                <p style="margin-top: 30px; color: #7f8c8d;">Pengirim akan menerima notifikasi untuk melakukan revisi.</p>
            </body>
            </html>
            """
        
        return render_template_string(html)

    except Exception as e:
        html = f"""
        <html>
        <body style="font-family: Arial, sans-serif; text-align: center; padding: 50px;">
            <h1 style="color: #e74c3c;">❌ Error</h1>
            <p>Terjadi kesalahan saat memproses validasi:</p>
            <p style="color: #c0392b;">{str(e)}</p>
        </body>
        </html>
        """
        return render_template_string(html), 500