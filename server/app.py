from gevent import monkey
monkey.patch_all()

import datetime
import os
import traceback
import json
import base64
import requests # Pastikan sudah install: pip install requests
from flask import Flask, request, jsonify, render_template, url_for
from dotenv import load_dotenv
from flask_cors import CORS
from datetime import timezone, timedelta
from num2words import num2words

import config
from google_services import GoogleServiceProvider
from pdf_generator import create_pdf_from_data, create_recap_pdf
from spk_generator import create_spk_pdf
from pengawasan_email_logic import get_email_details, FORM_LINKS
# Tambahkan import ini di bagian paling atas file app.py jika belum ada
from werkzeug.utils import secure_filename
from document_api import doc_bp
from data_api import data_bp
from dokumentasi_api import dokumentasi_bp # <--- dokumentasi bangunan

load_dotenv()
app = Flask(__name__, static_folder='static', template_folder='templates')

app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50 Megabytes

CORS(app,
     origins=[
         "http://127.0.0.1:5500",
         "http://localhost:5500",
         "http://localhost:3000",
         "https://building-alfamart.vercel.app",
         "https://building-alfamart.vercel.app/",
         "https://instruksi-lapangan.vercel.app",
         "https://instruksi-lapangan.vercel.app/",
         "https://gantt-chart-bnm.vercel.app",
         "https://gantt-chart-bnm.vercel.app/",
         "https://sparta-alfamart.vercel.app",
         "https://sparta-alfamart.vercel.app/",
         "https://frontend-form-virid.vercel.app",
         "https://frontend-form-virid.vercel.app/",
         "https://script.google.com",
         "https://opnamebnm.vercel.app",
         "https://opnamebnm.vercel.app/",
         "https://penyimpanan-dokumen.vercel.app",
         "https://penyimpanan-dokumen.vercel.app/",

     ],
     methods=["GET", "POST", "OPTIONS", "PUT", "PATCH", "DELETE"],
     allow_headers=["*"],
     supports_credentials=True
)

google_provider = GoogleServiceProvider()

app.register_blueprint(data_bp)
app.register_blueprint(doc_bp)
app.register_blueprint(dokumentasi_bp) # <--- dokumentasi bangunan
app.json.sort_keys = False

# --- KONFIGURASI PROXY GAS (DARI BACKEND LAMA) ---
GAS_URLS = {
  "input-pic": "https://script.google.com/macros/s/AKfycbzMWfroqPvtZXA1gz5VqdUzJhtyV_q8hWH92gl7JqFct5_dTVI2mcwmDHY6Rac5vmu-ww/exec",
  "login": "https://script.google.com/macros/s/AKfycbzCWExZ5r__w0viXeC1o5FXerwsqaC8y5XZg_W8zPMozlnLILHOJ1pPT4N-JDOFN6Jy/exec",
  "h2": "https://script.google.com/macros/s/AKfycbyHaiwKENoWsOEEgj2KHr3LQW-PwfkF-Fob7fgvUV52AusSAWaY8etSmeSZeiotK7Jvhw/exec",
  "h5": "https://script.google.com/macros/s/AKfycbzVdc7Uz2SFopdbcaWerO5UK7t6PAc7cPJVrV2s45iwe5uFTGtLzLRP7ZLv4T7kWus/exec",
  "h7": "https://script.google.com/macros/s/AKfycbxI5kdLFj4_45qqN63DZXJQ0Bv5CfSCjMcuMX7xWsaEbURUDhUhtrrfEa0eM8Jsq-sJHQ/exec",
  "h8": "https://script.google.com/macros/s/AKfycbySrgPzCpyUYhHKCGXrtIWfMtSTdLXdQeonzfoI1ro_R4rIZ9EewKPS4T6vM4ZoQuftUQ/exec",
  "h10": "https://script.google.com/macros/s/AKfycbzKmUVsDXzoTd39R37lVbkOhBPwuwasJQk2VVnLxD0jmy_yymdWpTPwa8YN87TLLloF8A/exec",
  "h12": "https://script.google.com/macros/s/AKfycbxwEkg5bSqPqXs2j6xnGqZZ-vU_ao5SpUCs6wlSrj5UOr5KL5UqM2Fjpoev7fyVlv29/exec",
  "h14": "https://script.google.com/macros/s/AKfycby50bSmRPZv7dizGTP-en9HmJ4wHzVX_aO0iwbXpU_2P6T1dlhENzk6XldL7LdbvaKL_A/exec",
  "h16": "https://script.google.com/macros/s/AKfycbzSFg-L9Kfu2EjAR2R70c6RE6vm9mFDItGq5JbuFIJhQIJHVcA20fpYlC4xSncpcbYKMA/exec",
  "h17": "https://script.google.com/macros/s/AKfycbwJdY8vlAbzy8_iEoqTZFyrE_ZlnHjdNy777eVUeOf8sc7aV0V_5bwwbPserG0hMCafnA/exec",
  "h18": "https://script.google.com/macros/s/AKfycbwRo61WD-mjalS5StjfzpnNA2peTXMXcb6mTAjfrxPU93kmtjYJrp_uecumr3qQPnB4mQ/exec",
  "h22": "https://script.google.com/macros/s/AKfycbx1TgvEvXqSehwU6h3GVsmuJ49gPs0NsWZ9NsJZUJd7W30Qa97tPBrfKJnSQHV_Fhfr/exec",
  "h23": "https://script.google.com/macros/s/AKfycbz-XKWo1WEHve_a6KIVNNcgy8ZtlnSNGKTfzsa9CF_La6i88VzFi-kEgOfpT2f6PIc/exec",
  "h25": "https://script.google.com/macros/s/AKfycbxZ-4vyNXPqcpUyORsJqhBxEKZgBEVMbmmXmLQNoZbmgK1Dxda-MG8l2UQhTEjXw0fhOw/exec",
  "h28": "https://script.google.com/macros/s/AKfycby3VlggJaznV7GHpp86p_eu8lvEm82uAoh5IVTURdHCprBWFFFh5cg7QnBSMtEVdESf/exec",
  "h32": "https://script.google.com/macros/s/AKfycbxqGR62PV_X86mFFJpMclgNYtGvkx8gVboONt62ynnFdOr25xKINkelEBdqnrqP7SHN/exec",
  "h33": "https://script.google.com/macros/s/AKfycbzuYFFG018O54U4nsp6iEtJ4kg57no3Juan22FICwf_VZkpNLjKMW7ZLu3Z_6WoWYsTOA/exec",
  "h41": "https://script.google.com/macros/s/AKfycbwgIbxOWzUeVwMnzqUsrggqTivg-_mtUXEnNYVKs8aIRlmcJo4JZz5dPlXDYfn4Fbib/exec",
  "serah_terima": "https://script.google.com/macros/s/AKfycbzYIuPoVoaD6HT7GjmGI2flSKzepirT9KztAX0Qg8vZbaixOBw3yhXe70qea8KCSuogtA/exec",
  "perpanjangan_spk": "https://script.google.com/macros/s/AKfycbyQRhiyX-zAyIyyHHU8OASCTj9O2tCSmnNesiX9o3q9ipQjKp5Qkx_LN6UmARtWMqJe/exec",
  "login_perpanjanganspk": "https://script.google.com/macros/s/AKfycbyQRhiyX-zAyIyyHHU8OASCTj9O2tCSmnNesiX9o3q9ipQjKp5Qkx_LN6UmARtWMqJe/exec"
}

# --- HELPER GENERATOR EMAIL (MIGRASI DARI BACKEND LAMA) ---
def generate_perpanjangan_email_body(data):
    # Logika konversi dari template JS 'perpanjangan_spk_notification'
    status = data.get('status_persetujuan', '').upper()
    is_approved = status == 'DISETUJUI'
    is_rejected = status == 'DITOLAK'
    
    color = 'green' if is_approved else ('red' if is_rejected else 'black')
    
    rejection_link = ""
    if is_rejected:
        rejection_link = """
        <p style="margin-top: 15px;">
            Silakan ajukan ulang permohonan perpanjangan SPK melalui link berikut:
            <br/>
            <a href="https://frontend-form-virid.vercel.app/login-perpanjanganspk.html" target="_blank">Ajukan Ulang Perpanjangan SPK</a>
        </p>
        """
    
    lampiran_html = ""
    if data.get('link_lampiran_user') and data.get('link_lampiran_user') != '-':
        lampiran_html = f'<li><strong>Lampiran Pendukung:</strong> <a href="{data.get("link_lampiran_user")}" target="_blank">Lihat Lampiran</a></li>'

    html_body = f"""
    <div style="font-family: Arial, sans-serif; line-height: 1.6;">
        <p>Pemberitahuan,</p>
        <p>Permintaan perpanjangan SPK untuk <strong>Nomor Ulok {data.get('nomor_ulok')}</strong> telah direspon.</p>
        <ul>
            <li><strong>Status:</strong> <strong style="color:{color};">{status}</strong></li>
            <li><strong>Ditinjau Oleh:</strong> {data.get('disetujui_oleh', 'N/A')}</li>
            <li><strong>Waktu Respon:</strong> {data.get('waktu_persetujuan', 'N/A')}</li>
            {f'<li><strong>Alasan Penolakan:</strong> {data.get("alasan_penolakan")}</li>' if is_rejected else ''}
            
            <li style="margin-top: 10px;"><strong>Tanggal SPK Akhir Lama:</strong> {data.get('tanggal_spk_akhir')}</li>
            <li><strong>Pertambahan Hari:</strong> {data.get('pertambahan_hari')} hari</li>
            <li><strong>Tanggal SPK Akhir Baru:</strong> {data.get('tanggal_spk_akhir_setelah_perpanjangan')}</li>
            <li style="margin-top: 10px;"><strong>Dokumen:</strong> <a href="{data.get('link_pdf')}" target="_blank">Lihat PDF</a></li>
            {lampiran_html}
        </ul>
        {rejection_link}
        <p>Terima kasih.</p>
    </div>
    """
    subject = f"[{status}] Perpanjangan SPK untuk Toko: {data.get('nomor_ulok')}"
    return subject, html_body

def generate_materai_email_body(data, role='manager'):
    # Logika konversi dari template JS 'materai_upload'
    recipients_list = data.get('managerRecipients' if role == 'manager' else 'otherRecipients', [])
    names = [r.get('name') for r in recipients_list if r.get('name')]
    sapaan = ", ".join(names) if names else "Tim Terkait"
    
    extra_manager_link = ""
    if role == 'manager':
        extra_manager_link = """
        <p style="margin-top: 20px;">
            Untuk mengisi form selanjutnya (SPK), silakan akses tautan berikut:
            <br/>
            📌 <a href="https://building-alfamart.vercel.app/login_spk.html" target="_blank">Isi Form SPK</a>
        </p>
        """

    html_body = f"""
    <div style="font-family: Arial, sans-serif; line-height: 1.6;">
      <p>Semangat Pagi,</p>
      <p>Bapak/Ibu <strong>{sapaan}</strong>,</p>
      <p>Email ini merupakan notifikasi bahwa dokumen materai baru telah diunggah dengan rincian:</p>
      <ul>
        <li><strong>Tanggal Upload:</strong> {data.get('tanggal_upload', 'N/A')}</li>
        <li><strong>Cabang:</strong> {data.get('cabang', 'N/A')}</li>
        <li><strong>Kode Ulok:</strong> {data.get('ulok', 'N/A')}</li>
        <li><strong>Lingkup Kerja:</strong> {data.get('lingkup_kerja', 'N/A')}</li>
      </ul>
      <p>Silakan lihat dokumen yang diunggah:</p>
      <p>📄 <a href="{data.get('pdfUrl', '#')}" target="_blank">Lihat Dokumen PDF</a></p>
      {extra_manager_link}
      <p>Terima kasih.</p>
    </div>
    """
    subject = f"Dokumen Final RAB Penawaran Termaterai - {data.get('ulok')}"
    return subject, html_body

def format_ulok(nomor_ulok_raw: str) -> str:
    """Format Nomor Ulok to pattern XXXX-XXXX-XXXX[-R].
    Examples:
    - 'Z0012512TEST' -> 'Z001-2512-TEST'
    - 'Z0012512TESTR' -> 'Z001-2512-TEST-R'
    - 'Z001-2512-TEST-R' -> 'Z001-2512-TEST-R' (idempotent)
    """
    if not nomor_ulok_raw:
        return ''
    s = str(nomor_ulok_raw).strip()
    # If already contains dashes in expected places and optional '-R', return as is
    if '-' in s:
        # Normalize multiple dashes; ensure final suffix preserved
        parts = s.split('-')
        if len(parts) >= 3:
            head = parts[0]
            mid = parts[1]
            tail = parts[2]
            suffix = parts[3] if len(parts) > 3 else None
            base = f"{head}-{mid}-{tail}"
            return f"{base}-{suffix}" if suffix else base
        # Fall through to rebuild from cleaned string
        s = ''.join(parts)
    # Handle optional trailing 'R'
    has_R = s.endswith('R')
    core = s[:-1] if has_R else s
    # Ensure core has at least 12 chars
    core = core[:12]
    if len(core) < 12:
        # Pad or just return original if insufficient
        return nomor_ulok_raw
    formatted = f"{core[:4]}-{core[4:8]}-{core[8:12]}"
    return f"{formatted}-R" if has_R else formatted

def get_tanggal_h(start_date, jumlah_hari_kerja):
    tanggal = start_date
    count = 0
    if not jumlah_hari_kerja: return tanggal
    while count < jumlah_hari_kerja:
        tanggal += timedelta(days=1)
        if tanggal.weekday() < 5:
            count += 1
    return tanggal

@app.route('/')
def index():
    return "Backend server is running and healthy.", 200

# --- ENDPOINTS OTENTIKASI & DATA UMUM ---
@app.route('/api/login', methods=['POST'])
def login():
    data = request.get_json()
    email = data.get('email')
    cabang = data.get('cabang')
    if not email or not cabang:
        return jsonify({"status": "error", "message": "Email and cabang are required"}), 400
    try:
        is_valid, role = google_provider.validate_user(email, cabang)
        if is_valid:
            return jsonify({"status": "success", "message": "Login successful", "role": role}), 200
        else:
            return jsonify({"status": "error", "message": "Invalid credentials"}), 401
    except Exception as e:
        traceback.print_exc()
        return jsonify({"status": "error", "message": "An internal server error occurred"}), 500

@app.route('/api/check_status', methods=['GET'])
def check_status():
    email = request.args.get('email')
    cabang = request.args.get('cabang')
    if not email or not cabang:
        return jsonify({"error": "Email and cabang parameters are missing"}), 400
    try:
        status_data = google_provider.check_user_submissions(email, cabang)
        return jsonify(status_data), 200
    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500

@app.route('/api/check_status_rab_2', methods=['GET'])
def check_status_rab_2():
    email = request.args.get('email')
    cabang = request.args.get('cabang')
    if not email or not cabang:
        return jsonify({"error": "Email and cabang parameters are missing"}), 400
    try:
        status_data = google_provider.check_user_submissions_rab_2(email, cabang)
        return jsonify(status_data), 200
    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500

def get_pt_name_by_email(provider, email):
    if not email: return "NAMA PT TIDAK DITEMUKAN"
    try:
        # Mengakses sheet Cabang
        cabang_sheet = provider.sheet.worksheet(config.CABANG_SHEET_NAME)
        records = cabang_sheet.get_all_records()
        
        # Loop cari email yang cocok
        for record in records:
            # Pastikan nama kolom email di sheet Cabang sesuai (misal: 'EMAIL_SAT')
            record_email = str(record.get('EMAIL_SAT', '')).strip().lower()
            if record_email == str(email).strip().lower():
                # Pastikan nama kolom PT di sheet Cabang sesuai (misal: 'Nama_PT' atau 'NAMA PT')
                return record.get('Nama_PT', '').strip() or record.get('NAMA PT', '').strip()
    except Exception as e:
        print(f"Error fetching PT Name: {e}")
    return "NAMA PT TIDAK DITEMUKAN"

@app.route('/api/check_ulok_rab_2', methods=['GET'])
def check_ulok_rab_2():
    ulok = request.args.get('ulok')
    
    if not ulok:
        return jsonify({"status": "error", "message": "Parameter ulok dibutuhkan"}), 400

    try:
        # Panggil fungsi helper yang baru kita buat
        result = google_provider.check_ulok_exists_rab_2(ulok)
        
        return jsonify({
            "status": "success",
            "data": result
        }), 200
        
    except Exception as e:
        traceback.print_exc()
        return jsonify({"status": "error", "message": str(e)}), 500

# --- ENDPOINTS UNTUK ALUR KERJA RAB ---
@app.route('/api/submit_rab', methods=['POST'])
def submit_rab():
    data = request.get_json()
    if not data:
        return jsonify({"status": "error", "message": "Invalid JSON data"}), 400

    new_row_index = None
    try:
        nomor_ulok_raw = data.get(config.COLUMN_NAMES.LOKASI, '')
        lingkup_pekerjaan = data.get(config.COLUMN_NAMES.LINGKUP_PEKERJAAN, '')

        if not nomor_ulok_raw:
            return jsonify({
                "status": "error",
                "message": "Nomor Ulok tidak boleh kosong."
            }), 400

        # 1. Cek Duplikasi AKTIF (Waiting / Approved)
        # Jika ditemukan data yang sedang berjalan dengan ULOK & LINGKUP sama, BLOKIR.
        # Fungsi ini sekarang sudah support format Renovasi (13 digit) & New Store (12 digit).
        if google_provider.check_ulok_exists(nomor_ulok_raw, lingkup_pekerjaan):
            return jsonify({
                "status": "error",
                "message": (
                    f"RAB untuk Nomor Ulok {format_ulok(nomor_ulok_raw)} dengan lingkup {lingkup_pekerjaan} "
                    "sudah pernah diajukan dan sedang diproses atau sudah disetujui."
                )
            }), 409

        # 2. Cek apakah ini REVISI (Data Ditolak)
        # Jika lolos cek nomor 1, kita cari apakah ada data DITOLAK yang bisa ditimpa.
        rejected_row_index = google_provider.find_rejected_row_index(
            nomor_ulok_raw, 
            lingkup_pekerjaan
        )

        # 1. Ambil Email Pembuat
        email_pembuat = data.get('Email_Pembuat')
        
        # 2. Cari Nama PT berdasarkan Email
        nama_pt_kontraktor = get_pt_name_by_email(google_provider, email_pembuat)
        
        # 3. Masukkan ke dalam dictionary 'data' agar tersimpan di Form2 & muncul di PDF
        data[config.COLUMN_NAMES.NAMA_PT] = nama_pt_kontraktor

        # Set status awal & timestamp
        WIB = timezone(timedelta(hours=7))
        data[config.COLUMN_NAMES.STATUS] = config.STATUS.WAITING_FOR_COORDINATOR
        data[config.COLUMN_NAMES.TIMESTAMP] = datetime.datetime.now(WIB).isoformat()

        # PENTING: Jika Revisi, kita harus MENGOSONGKAN kolom persetujuan lama & alasan penolakan
        # agar flow dimulai dari awal lagi di Google Sheet
        data[config.COLUMN_NAMES.KOORDINATOR_APPROVER] = ""
        data[config.COLUMN_NAMES.KOORDINATOR_APPROVAL_TIME] = ""
        data[config.COLUMN_NAMES.MANAGER_APPROVER] = ""
        data[config.COLUMN_NAMES.MANAGER_APPROVAL_TIME] = ""
        data['Alasan Penolakan'] = ""

        # --- 1) HITUNG GRAND TOTAL NON-SBO (sudah ada) ---
        total_non_sbo = 0.0
        for i in range(1, 201):
            kategori = data.get(f'Kategori_Pekerjaan_{i}')
            total_item_str = data.get(f'Total_Harga_Item_{i}')
            if kategori and kategori != 'PEKERJAAN SBO' and total_item_str:
                try:
                    total_non_sbo += float(total_item_str)
                except ValueError:
                    pass
        data[config.COLUMN_NAMES.GRAND_TOTAL_NONSBO] = total_non_sbo

        # --- 2) HITUNG GRAND TOTAL (SEMUA ITEM, TERMASUK SBO) ---
        total_semua_item = 0.0
        for i in range(1, 201):
            total_item_str = data.get(f'Total_Harga_Item_{i}')
            if total_item_str:
                try:
                    total_semua_item += float(total_item_str)
                except ValueError:
                    pass

        # Simpan "data lama" Grand Total (sebelum pembulatan & PPN)
        data[config.COLUMN_NAMES.GRAND_TOTAL] = total_semua_item

        # --- 3) HITUNG GRAND TOTAL FINAL (UNTUK SPK) ---
        #   - dibulatkan ke bawah kelipatan 10.000
        #   - + PPN 11%
        pembulatan = (total_semua_item // 10000) * 10000  # kelipatan 10.000
        ppn = pembulatan * 0.11
        final_grand_total = pembulatan + ppn

        # Simpan ke kolom baru "Grand Total Final"
        data[config.COLUMN_NAMES.GRAND_TOTAL_FINAL] = final_grand_total

        # --- 4) ARSIPKAN DETAIL ITEM KE JSON ---
        item_keys_to_archive = (
            'Kategori_Pekerjaan_', 'Jenis_Pekerjaan_', 'Satuan_Item_',
            'Volume_Item_', 'Harga_Material_Item_', 'Harga_Upah_Item_',
            'Total_Material_Item_', 'Total_Upah_Item_', 'Total_Harga_Item_'
        )
        item_details = {
            k: v for k, v in data.items()
            if k.startswith(item_keys_to_archive)
        }
        data['Item_Details_JSON'] = json.dumps(item_details)

        jenis_toko = data.get('Proyek', 'N/A')
        nama_toko = data.get('Nama_Toko', data.get('nama_toko', 'N/A'))
        lingkup_pekerjaan = data.get('Lingkup_Pekerjaan', data.get('lingkup_pekerjaan', 'N/A'))

        # --- LOGIC FORMATTING ULOK DIPERBARUI ---
        nomor_ulok_formatted = nomor_ulok_raw
        if isinstance(nomor_ulok_raw, str):
            clean_ulok = nomor_ulok_raw.replace("-", "") # Bersihkan dash dulu
            
            if len(clean_ulok) == 13 and clean_ulok.endswith('R'):
                # Format Renovasi: XXXX-YYYY-ZZZZ-R
                nomor_ulok_formatted = (
                    f"{clean_ulok[:4]}-"
                    f"{clean_ulok[4:8]}-"
                    f"{clean_ulok[8:12]}-"
                    f"{clean_ulok[12:]}"
                )
            elif len(clean_ulok) == 12:
                # Format Normal: XXXX-YYYY-ZZZZ
                nomor_ulok_formatted = (
                    f"{clean_ulok[:4]}-"
                    f"{clean_ulok[4:8]}-"
                    f"{clean_ulok[8:]}"
                )

        # --- 5) GENERATE PDF (NON-SBO & REKAP) ---
        pdf_nonsbo_bytes = create_pdf_from_data(
            google_provider, data, exclude_sbo=True
        )
        pdf_recap_bytes = create_recap_pdf(google_provider, data)

        pdf_nonsbo_filename = f"RAB_NON-SBO_{jenis_toko}_{nomor_ulok_formatted}.pdf"
        pdf_recap_filename = f"REKAP_RAB_{jenis_toko}_{nomor_ulok_formatted}.pdf"

        link_pdf_nonsbo = google_provider.upload_file_to_drive(
            pdf_nonsbo_bytes,
            pdf_nonsbo_filename,
            'application/pdf',
            config.PDF_STORAGE_FOLDER_ID
        )
        link_pdf_rekap = google_provider.upload_file_to_drive(
            pdf_recap_bytes,
            pdf_recap_filename,
            'application/pdf',
            config.PDF_STORAGE_FOLDER_ID
        )

        data[config.COLUMN_NAMES.LINK_PDF_NONSBO] = link_pdf_nonsbo
        data[config.COLUMN_NAMES.LINK_PDF_REKAP] = link_pdf_rekap
        data[config.COLUMN_NAMES.LOKASI] = nomor_ulok_formatted

        # --- 6) SIMPAN KE SHEET ---
        if rejected_row_index:
            # === KASUS REVISI: Update Baris Lama ===
            google_provider.update_row(
                config.DATA_ENTRY_SHEET_NAME,
                rejected_row_index,
                data
            )
            final_row_index = rejected_row_index
            print(f"Revisi RAB: Mengupdate baris {final_row_index}")
        else:
            # === KASUS BARU: Tambah Baris Baru ===
            new_row_index = google_provider.append_to_sheet(
                data,
                config.DATA_ENTRY_SHEET_NAME
            )
            final_row_index = new_row_index
            print(f"RAB Baru: Menambah baris {final_row_index}")

        # --- 7) KIRIM EMAIL KE KOORDINATOR ---
        cabang = data.get('Cabang')
        if not cabang:
            raise Exception("Field 'Cabang' is empty. Cannot find Coordinator.")

        coordinator_emails = google_provider.get_emails_by_jabatan(
            cabang,
            config.JABATAN.KOORDINATOR
        )
        if not coordinator_emails:
            raise Exception(
                f"Tidak ada email Koordinator yang ditemukan untuk cabang '{cabang}'."
            )

        base_url = "https://sparta-backend-5hdj.onrender.com"
        approver_for_link = coordinator_emails[0]
        approval_url = (
            f"{base_url}/api/handle_rab_approval"
            f"?action=approve&row={final_row_index}"
            f"&level=coordinator&approver={approver_for_link}"
        )
        rejection_url = (
            f"{base_url}/api/reject_form/rab"
            f"?row={final_row_index}&level=coordinator"
            f"&approver={approver_for_link}"
        )

        email_html = render_template(
            'email_template.html',
            doc_type="RAB",
            level='Koordinator',
            form_data=data,
            approval_url=approval_url,
            rejection_url=rejection_url
        )

        google_provider.send_email(
            to=coordinator_emails,
            subject=f"[TAHAP 1: PERLU PERSETUJUAN] RAB Proyek {nama_toko}: {jenis_toko} - {lingkup_pekerjaan}",
            html_body=email_html,
            attachments=[
                (pdf_nonsbo_filename, pdf_nonsbo_bytes, 'application/pdf'),
                (pdf_recap_filename, pdf_recap_bytes, 'application/pdf')
            ]
        )

        return jsonify({
            "status": "success",
            "message": "Data successfully submitted/updated and approval email sent."
        }), 200

    except Exception as e:
        # Hanya hapus baris jika itu adalah input BARU (bukan revisi)
        if new_row_index and not rejected_row_index:
            google_provider.delete_row(
                config.DATA_ENTRY_SHEET_NAME,
                new_row_index
            )
        traceback.print_exc()
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route('/api/submit_rab_kedua', methods=['POST'])
def submit_rab_kedua():
    # --- PERBAIKAN 1: Inisialisasi variabel di awal ---
    new_row_index = None
    file_pdf_IL_upload = None
    data = {} 
    
    try:
        # --- PERBAIKAN 2: Pengecekan Content Type yang lebih aman ---
        content_type = request.content_type or ""

        if content_type.startswith('multipart/form-data'):
            # Jika request adalah Form Data (ada file upload)
            data = dict(request.form) # Konversi form data ke dictionary
            
            # Konversi tipe data numerik dari string
            for key, value in data.items():
                if key.startswith(('Volume_Item_', 'Harga_Material_Item_', 'Harga_Upah_Item_', 'Total_Harga_Item_')):
                    try:
                        data[key] = float(value)
                    except (ValueError, TypeError):
                        data[key] = 0.0
            
            # Ambil file PDF jika ada
            if 'file_pdf' in request.files:
                file_pdf_IL_upload = request.files['file_pdf']
                
        else:
            # Jika request adalah JSON biasa
            data = request.get_json()
            if not data:
                return jsonify({"status": "error", "message": "Invalid JSON data"}), 400

        # --- LOGIKA UTAMA (SAMA SEPERTI SEBELUMNYA) ---
        
        nomor_ulok_raw = data.get(config.COLUMN_NAMES.LOKASI, '')
        lingkup_pekerjaan = data.get(config.COLUMN_NAMES.LINGKUP_PEKERJAAN, '')

        if not nomor_ulok_raw:
            return jsonify({
                "status": "error",
                "message": "Nomor Ulok tidak boleh kosong."
            }), 400

        # Cek revisi / duplikasi
        is_revising = google_provider.is_revision(
            nomor_ulok_raw,
            data.get('Email_Pembuat')
        )

        if not is_revising and google_provider.check_ulok_exists(nomor_ulok_raw, lingkup_pekerjaan):
            return jsonify({
                "status": "error",
                "message": (
                    f"Nomor Ulok {nomor_ulok_raw} dengan lingkup {lingkup_pekerjaan} "
                    "sudah pernah diajukan dan sedang diproses atau sudah disetujui."
                )
            }), 409

        # Set status awal & timestamp
        WIB = timezone(timedelta(hours=7))
        data[config.COLUMN_NAMES.STATUS] = config.STATUS.WAITING_FOR_COORDINATOR
        data[config.COLUMN_NAMES.TIMESTAMP] = datetime.datetime.now(WIB).isoformat()

        # --- HITUNG TOTAL ---
        total_non_sbo = 0.0
        total_semua_item = 0.0
        
        for i in range(1, 201):
            kategori = data.get(f'Kategori_Pekerjaan_{i}')
            try:
                total_item_val = float(data.get(f'Total_Harga_Item_{i}', 0))
            except (ValueError, TypeError):
                total_item_val = 0.0

            if total_item_val > 0:
                total_semua_item += total_item_val
                if kategori and kategori != 'PEKERJAAN SBO':
                    total_non_sbo += total_item_val

        data[config.COLUMN_NAMES.GRAND_TOTAL_NONSBO] = total_non_sbo
        data[config.COLUMN_NAMES.GRAND_TOTAL] = total_semua_item

        # Pembulatan & PPN
        pembulatan = (total_semua_item // 10000) * 10000
        ppn = pembulatan * 0.11
        final_grand_total = pembulatan + ppn
        data[config.COLUMN_NAMES.GRAND_TOTAL_FINAL] = final_grand_total

        # Arsip JSON
        item_keys_to_archive = (
            'Kategori_Pekerjaan_', 'Jenis_Pekerjaan_', 'Satuan_Item_',
            'Volume_Item_', 'Harga_Material_Item_', 'Harga_Upah_Item_',
            'Total_Material_Item_', 'Total_Upah_Item_', 'Total_Harga_Item_'
        )
        item_details = {k: v for k, v in data.items() if k.startswith(item_keys_to_archive)}
        data['Item_Details_JSON'] = json.dumps(item_details)

        jenis_toko = data.get('Proyek', 'N/A')
        nama_toko = data.get('Nama_Toko', data.get('nama_toko', 'N/A'))
        
        
        nomor_ulok_formatted = format_ulok(nomor_ulok_raw)
        data[config.COLUMN_NAMES.LOKASI] = nomor_ulok_formatted
        # --- ENDLOGIKA UTAMA ---

        # --- GENERATE PDF OTOMATIS ---
        pdf_nonsbo_bytes = create_pdf_from_data(google_provider, data, exclude_sbo=True)
        pdf_recap_bytes = create_recap_pdf(google_provider, data)

        pdf_nonsbo_filename = f"IL_NON-SBO_{jenis_toko}_{nomor_ulok_formatted}.pdf"
        pdf_recap_filename = f"REKAP_IL_{jenis_toko}_{nomor_ulok_formatted}.pdf"

        link_pdf_nonsbo = google_provider.upload_file_to_drive(
            pdf_nonsbo_bytes, pdf_nonsbo_filename, 'application/pdf', config.PDF_STORAGE_FOLDER_ID
        )
        link_pdf_rekap = google_provider.upload_file_to_drive(
            pdf_recap_bytes, pdf_recap_filename, 'application/pdf', config.PDF_STORAGE_FOLDER_ID
        )

        data[config.COLUMN_NAMES.LINK_PDF_NONSBO] = link_pdf_nonsbo
        data[config.COLUMN_NAMES.LINK_PDF_REKAP] = link_pdf_rekap

        # --- HANDLE UPLOAD PDF MANUAL ---
        link_pdf_IL = ""
        manual_filename = ""
        
        if  file_pdf_IL_upload:
            manual_filename = secure_filename(f"IL_{file_pdf_IL_upload.filename}")
            manual_file_bytes = file_pdf_IL_upload.read()
            
            link_pdf_IL = google_provider.upload_file_to_drive(
                manual_file_bytes,
                manual_filename,
                file_pdf_IL_upload.content_type,
                config.PDF_STORAGE_FOLDER_ID
            )
            data[config.COLUMN_NAMES.LINK_PDF_IL] = link_pdf_IL

        # --- CEGAH DOUBLE DATA: CEK ULOK DI SHEET RAB 2 UNTUK STATUS DITOLAK ---
        try:
            spreadsheet = google_provider.gspread_client.open_by_key(config.SPREADSHEET_ID_RAB_2)
            worksheet = spreadsheet.worksheet(config.DATA_ENTRY_SHEET_NAME_RAB_2)

            header = worksheet.row_values(1)
            # Cari index kolom penting
            def col_index(col_name):
                try:
                    return header.index(col_name) + 1
                except ValueError:
                    return None

            col_ulok = col_index(config.COLUMN_NAMES.LOKASI)
            col_status = col_index(config.COLUMN_NAMES.STATUS)
            col_timestamp = col_index(config.COLUMN_NAMES.TIMESTAMP)

            existing_row_index = None
            rejected_statuses = {
                config.STATUS.REJECTED_BY_COORDINATOR,
                config.STATUS.REJECTED_BY_MANAGER,
            }

            # Jika kolom ada, iterasi baris untuk mencari ULOK yang sama dengan status ditolak
            if col_ulok and col_status:
                all_rows = worksheet.get_all_values()
                for r_idx in range(2, len(all_rows) + 1):
                    row_vals = all_rows[r_idx - 1]
                    current_ulok = row_vals[col_ulok - 1] if len(row_vals) >= col_ulok else ""
                    current_status = row_vals[col_status - 1] if len(row_vals) >= col_status else ""
                    if current_ulok == nomor_ulok_formatted and current_status in rejected_statuses:
                        existing_row_index = r_idx
                        break

            # Jika ditemukan baris yang ditolak sebelumnya, jangan append. Ubah statusnya ke waiting sesuai asal penolakan
            if existing_row_index:
                previous_status = worksheet.cell(existing_row_index, col_status).value if col_status else ""
                if previous_status == config.STATUS.REJECTED_BY_COORDINATOR:
                    new_status_for_existing = config.STATUS.WAITING_FOR_COORDINATOR
                elif previous_status == config.STATUS.REJECTED_BY_MANAGER:
                    new_status_for_existing = config.STATUS.WAITING_FOR_MANAGER
                else:
                    new_status_for_existing = config.STATUS.WAITING_FOR_COORDINATOR

                if col_status:
                    worksheet.update_cell(existing_row_index, col_status, new_status_for_existing)
                if col_timestamp:
                    worksheet.update_cell(existing_row_index, col_timestamp, data.get(config.COLUMN_NAMES.TIMESTAMP, ""))

                # --- OVERWRITE DATA ITEM & FIELD LAIN BERDASARKAN SUBMISSION BARU ---
                item_prefixes = (
                    'Kategori_Pekerjaan_', 'Jenis_Pekerjaan_', 'Satuan_Item_',
                    'Volume_Item_', 'Harga_Material_Item_', 'Harga_Upah_Item_',
                    'Total_Material_Item_', 'Total_Upah_Item_', 'Total_Harga_Item_'
                )

                # Pastikan field penting lainnya juga diperbarui dari data baru
                # Termasuk: totals, links, item_details_json, lingkup, cabang, dll jika ada di data
                for header_name in header:
                    try:
                        if header_name in data:
                            google_provider.update_cell_by_sheet(
                                worksheet, existing_row_index, header_name, data.get(header_name, "")
                            )
                        elif header_name.startswith(item_prefixes):
                            # Kosongkan kolom item lama yang tidak ada di submission baru
                            google_provider.update_cell_by_sheet(
                                worksheet, existing_row_index, header_name, ""
                            )
                    except Exception as upd_err:
                        print(f"Warning: gagal update kolom {header_name} pada baris {existing_row_index}: {upd_err}")

                new_row_index = existing_row_index
            else:
                # --- SIMPAN KE SHEET ---
                # Menggunakan fungsi append_to_dynamic_sheet agar masuk ke Spreadsheet RAB 2
                new_row_index = google_provider.append_to_dynamic_sheet(
                    config.SPREADSHEET_ID_RAB_2,
                    config.DATA_ENTRY_SHEET_NAME_RAB_2,
                    data
                )
        except Exception as sheet_check_err:
            # Jika terjadi error saat cek, fallback ke append agar tidak memblokir alur
            print(f"Warning: gagal cek double data RAB 2: {sheet_check_err}")
            new_row_index = google_provider.append_to_dynamic_sheet(
                config.SPREADSHEET_ID_RAB_2,
                config.DATA_ENTRY_SHEET_NAME_RAB_2,
                data
            )

        # --- KIRIM EMAIL ---
        cabang = data.get('Cabang')
        if not cabang:
            raise Exception("Field 'Cabang' is empty.")

        coordinator_emails = google_provider.get_emails_by_jabatan(cabang, config.JABATAN.KOORDINATOR)
        if not coordinator_emails:
            raise Exception(f"Tidak ada email Koordinator untuk cabang '{cabang}'.")

        base_url = "https://sparta-backend-5hdj.onrender.com"
        approver_for_link = coordinator_emails[0]
        
        # Arahkan ke handler approval RAB 2
        approval_url = f"{base_url}/api/handle_rab_2_approval?action=approve&row={new_row_index}&level=coordinator&approver={approver_for_link}"
        rejection_url = f"{base_url}/api/reject_form/rab_kedua?action=reject&row={new_row_index}&level=coordinator&approver={approver_for_link}"

        email_html = render_template(
            'email_template.html',
            doc_type="RAB",
            level='Koordinator',
            form_data=data,
            approval_url=approval_url,
            rejection_url=rejection_url
        )

        attachments_list = [
            (pdf_nonsbo_filename, pdf_nonsbo_bytes, 'application/pdf'),
            (pdf_recap_filename, pdf_recap_bytes, 'application/pdf')
        ]
        
        if file_pdf_IL_upload and link_pdf_IL:
             attachments_list.append((manual_filename, manual_file_bytes, file_pdf_IL_upload.content_type))

        google_provider.send_email(
            to=coordinator_emails,
            subject=f"[TAHAP 1: PERLU PERSETUJUAN] IL Proyek {nama_toko} - {lingkup_pekerjaan}",
            html_body=email_html,
            attachments=attachments_list
        )

        return jsonify({
            "status": "success",
            "message": "Data submitted successfully (RAB 2).",
            "pdf_auto_links": {
                "nonsbo": link_pdf_nonsbo,
                "recap": link_pdf_rekap
            },
            "pdf_IL": link_pdf_IL
        }), 200

    except Exception as e:
        traceback.print_exc()
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route('/api/reject_form/rab', methods=['GET'])
def reject_form_rab():
    row = request.args.get('row')
    level = request.args.get('level')
    approver = request.args.get('approver')
    
    if not all([row, level, approver]):
        return "Parameter tidak lengkap.", 400

    row_data = google_provider.get_row_data(int(row))
    if not row_data:
        return "Data permintaan tidak ditemukan.", 404

    logo_url = url_for('static', filename='Alfamart-Emblem.png', _external=True)
    item_identifier = row_data.get(config.COLUMN_NAMES.LOKASI, 'N/A')
    
    return render_template(
        'rejection_form.html',
        form_action=url_for('handle_rab_approval', _external=True),
        row=row,
        level=level,
        approver=approver,
        item_type="RAB",
        item_identifier=item_identifier,
        logo_url=logo_url
    )

@app.route('/api/reject_form/rab_kedua', methods=['GET'])
def reject_form_rab_kedua():
    row = request.args.get('row')
    level = request.args.get('level')
    approver = request.args.get('approver')
    
    if not all([row, level, approver]):
        return "Parameter tidak lengkap.", 400

    try:
        # --- PERBAIKAN DI SINI ---
        # 1. Buka Spreadsheet & Sheet KHUSUS RAB 2
        spreadsheet = google_provider.gspread_client.open_by_key(config.SPREADSHEET_ID_RAB_2)
        worksheet = spreadsheet.worksheet(config.DATA_ENTRY_SHEET_NAME_RAB_2)
        
        # 2. Ambil data menggunakan helper 'get_row_data_by_sheet' (bukan get_row_data biasa)
        row_data = google_provider.get_row_data_by_sheet(worksheet, int(row))
        
        if not row_data:
            return "Data permintaan tidak ditemukan di Sheet RAB 2.", 404

        logo_url = url_for('static', filename='Alfamart-Emblem.png', _external=True)
        item_identifier = row_data.get(config.COLUMN_NAMES.LOKASI, 'N/A')
        
        return render_template(
            'rejection_form.html',
            # --- PERBAIKAN JUGA DI SINI ---
            # Arahkan form action ke handler RAB 2, bukan RAB 1
            form_action=url_for('handle_rab_2_approval', _external=True), 
            row=row,
            level=level,
            approver=approver,
            item_type="IL",
            item_identifier=item_identifier,
            logo_url=logo_url
        )
    except Exception as e:
        traceback.print_exc()
        return f"Terjadi kesalahan server: {str(e)}", 500


@app.route('/api/handle_rab_approval', methods=['GET', 'POST'])
def handle_rab_approval():
    if request.method == 'POST':
        data = request.form
    else:
        data = request.args

    action = data.get('action')
    row_str = data.get('row')
    level = data.get('level')
    approver = data.get('approver')
    reason = data.get('reason', 'Tidak ada alasan yang diberikan.') # Ambil alasan jika ada
    
    logo_url = url_for('static', filename='Alfamart-Emblem.png', _external=True)

    if not all([action, row_str, level, approver]):
        return render_template('response_page.html', title='Incomplete Parameters', message='URL parameters are incomplete.', logo_url=logo_url), 400
    try:
        row = int(row_str)
        row_data = google_provider.get_row_data(row)
        if not row_data:
            return render_template('response_page.html', title='Data Not Found', message='This request may have been deleted.', logo_url=logo_url)
        
        item_details_json = row_data.get('Item_Details_JSON', '{}')
        if item_details_json:
            try:
                item_details = json.loads(item_details_json)
                row_data.update(item_details)
            except json.JSONDecodeError:
                print(f"Warning: Could not decode Item_Details_JSON for row {row}")
        
        current_status = row_data.get(config.COLUMN_NAMES.STATUS, "").strip()
        expected_status_map = {'coordinator': config.STATUS.WAITING_FOR_COORDINATOR, 'manager': config.STATUS.WAITING_FOR_MANAGER}
        
        if current_status != expected_status_map.get(level):
            msg = f'This action has already been processed. Current status: <strong>{current_status}</strong>.'
            return render_template('response_page.html', title='Action Already Processed', message=msg, logo_url=logo_url)
        
        WIB = timezone(timedelta(hours=7))
        current_time = datetime.datetime.now(WIB).isoformat()
        
        cabang = row_data.get(config.COLUMN_NAMES.CABANG)
        jenis_toko = row_data.get(config.COLUMN_NAMES.PROYEK, 'N/A')
        nama_toko = row_data.get('Nama_Toko', row_data.get('nama_toko', 'N/A'))
        lingkup_pekerjaan = row_data.get(config.COLUMN_NAMES.LINGKUP_PEKERJAAN, 'N/A')

        creator_email = row_data.get(config.COLUMN_NAMES.EMAIL_PEMBUAT)

        if action == 'reject':
            new_status = ""
            if level == 'coordinator':
                new_status = config.STATUS.REJECTED_BY_COORDINATOR
                google_provider.update_cell(row, config.COLUMN_NAMES.KOORDINATOR_APPROVER, approver)
                google_provider.update_cell(row, config.COLUMN_NAMES.KOORDINATOR_APPROVAL_TIME, current_time)
            elif level == 'manager':
                new_status = config.STATUS.REJECTED_BY_MANAGER
                google_provider.update_cell(row, config.COLUMN_NAMES.MANAGER_APPROVER, approver)
                google_provider.update_cell(row, config.COLUMN_NAMES.MANAGER_APPROVAL_TIME, current_time)
            
            google_provider.update_cell(row, config.COLUMN_NAMES.STATUS, new_status)
            google_provider.update_cell(row, 'Alasan Penolakan', reason)
            if creator_email:
                subject = f"[DITOLAK] Pengajuan RAB Proyek {nama_toko}: {jenis_toko} - {lingkup_pekerjaan}"
                body = (f"<p>Pengajuan RAB Toko <b>{nama_toko}</b> untuk proyek <b>{jenis_toko} - {lingkup_pekerjaan}</b> telah <b>DITOLAK</b>.</p>"
                        f"<p><b>Alasan Penolakan:</b></p>"
                        f"<p><i>{reason}</i></p>"
                        f"<p>Silakan ajukan revisi RAB Anda melalui link berikut:</p>"
                        f"<p><a href='https://sparta-alfamart.vercel.app' target='_blank' rel='noopener noreferrer'>Input Ulang RAB</a></p>")
                google_provider.send_email(to=creator_email, subject=subject, html_body=body)
            return render_template('response_page.html', title='Permintaan Ditolak', message='Status permintaan telah diperbarui.', logo_url=logo_url)

        elif level == 'coordinator' and action == 'approve':
            google_provider.update_cell(row, config.COLUMN_NAMES.STATUS, config.STATUS.WAITING_FOR_MANAGER)
            google_provider.update_cell(row, config.COLUMN_NAMES.KOORDINATOR_APPROVER, approver)
            google_provider.update_cell(row, config.COLUMN_NAMES.KOORDINATOR_APPROVAL_TIME, current_time)
            manager_email = google_provider.get_email_by_jabatan(cabang, config.JABATAN.MANAGER)
            if manager_email:
                row_data[config.COLUMN_NAMES.KOORDINATOR_APPROVER] = approver
                row_data[config.COLUMN_NAMES.KOORDINATOR_APPROVAL_TIME] = current_time
                base_url = "https://sparta-backend-5hdj.onrender.com"
                approval_url_manager = f"{base_url}/api/handle_rab_approval?action=approve&row={row}&level=manager&approver={manager_email}"
                rejection_url_manager = f"{base_url}/api/reject_form/rab?row={row}&level=manager&approver={manager_email}"
                email_html_manager = render_template('email_template.html', doc_type="RAB", level='Manajer', form_data=row_data, approval_url=approval_url_manager, rejection_url=rejection_url_manager, additional_info=f"Telah disetujui oleh Koordinator: {approver}")
                pdf_nonsbo_bytes = create_pdf_from_data(google_provider, row_data, exclude_sbo=True)
                pdf_recap_bytes = create_recap_pdf(google_provider, row_data)
                pdf_nonsbo_filename = f"RAB_NON-SBO_{jenis_toko}_{row_data.get('Nomor Ulok')}.pdf"
                pdf_recap_filename = f"REKAP_RAB_{jenis_toko}_{row_data.get('Nomor Ulok')}.pdf"
                google_provider.send_email(manager_email, f"[TAHAP 2: PERLU PERSETUJUAN] RAB Proyek {nama_toko}: {jenis_toko} - {lingkup_pekerjaan}", email_html_manager, attachments=[(pdf_nonsbo_filename, pdf_nonsbo_bytes, 'application/pdf'),(pdf_recap_filename, pdf_recap_bytes, 'application/pdf')])
            return render_template('response_page.html', title='Persetujuan Diteruskan', message='Terima kasih. Persetujuan Anda telah dicatat.', logo_url=logo_url)
        
        elif level == 'manager' and action == 'approve':
            google_provider.update_cell(row, config.COLUMN_NAMES.STATUS, config.STATUS.APPROVED)
            google_provider.update_cell(row, config.COLUMN_NAMES.MANAGER_APPROVER, approver)
            google_provider.update_cell(row, config.COLUMN_NAMES.MANAGER_APPROVAL_TIME, current_time)
            
            row_data[config.COLUMN_NAMES.STATUS] = config.STATUS.APPROVED
            row_data[config.COLUMN_NAMES.MANAGER_APPROVER] = approver
            row_data[config.COLUMN_NAMES.MANAGER_APPROVAL_TIME] = current_time
            
            # ====== Generate PDFs ======
            pdf_nonsbo_bytes = create_pdf_from_data(google_provider, row_data, exclude_sbo=True)
            pdf_nonsbo_filename = f"DISETUJUI_RAB_NON-SBO_{jenis_toko}_{row_data.get('Nomor Ulok')}.pdf"

            pdf_recap_bytes = create_recap_pdf(google_provider, row_data)
            pdf_recap_filename = f"DISETUJUI_REKAP_RAB_{jenis_toko}_{row_data.get('Nomor Ulok')}.pdf"

            # Upload ke Drive
            link_pdf_nonsbo = google_provider.upload_file_to_drive(
                pdf_nonsbo_bytes, pdf_nonsbo_filename, 'application/pdf', config.PDF_STORAGE_FOLDER_ID
            )
            link_pdf_rekap = google_provider.upload_file_to_drive(
                pdf_recap_bytes, pdf_recap_filename, 'application/pdf', config.PDF_STORAGE_FOLDER_ID
            )

            # Update sheet
            google_provider.update_cell(row, config.COLUMN_NAMES.LINK_PDF_NONSBO, link_pdf_nonsbo)
            google_provider.update_cell(row, config.COLUMN_NAMES.LINK_PDF_REKAP, link_pdf_rekap)

            row_data[config.COLUMN_NAMES.LINK_PDF_NONSBO] = link_pdf_nonsbo
            row_data[config.COLUMN_NAMES.LINK_PDF_REKAP] = link_pdf_rekap

            google_provider.copy_to_approved_sheet(row_data)
            
            # ====== Kirim data ke Summary Sheet ======
            google_provider.copy_to_summary_sheet(row_data)

            # ====== Kumpulkan email dari jabatan ======
            email_pembuat = row_data.get(config.COLUMN_NAMES.EMAIL_PEMBUAT)
            kontraktor_emails = [email_pembuat] if email_pembuat else []
            coordinator_emails = google_provider.get_emails_by_jabatan(cabang, config.JABATAN.KOORDINATOR)
            manager_email = approver  # manager yang menyetujui

            # ====== Attachment bersama ======
            email_attachments = [
                (pdf_nonsbo_filename, pdf_nonsbo_bytes, 'application/pdf'),
                (pdf_recap_filename, pdf_recap_bytes, 'application/pdf')
            ]

            subject = f"[FINAL - DISETUJUI] Pengajuan RAB Proyek {nama_toko}: {jenis_toko} - {lingkup_pekerjaan}"

            # Template body utama (bold + link Google Drive)
            base_body = (
                f"<p>Pengajuan RAB Toko <b>{nama_toko}</b> untuk proyek <b>{jenis_toko} - {lingkup_pekerjaan}</b> "
                f"telah disetujui sepenuhnya.</p>"
                f"<p>Tiga versi file PDF RAB telah dilampirkan:</p>"
                f"<ul>"
                f"<li><b>{pdf_nonsbo_filename}</b>: Hanya berisi item pekerjaan di luar SBO.</li>"
                f"<li><b>{pdf_recap_filename}</b>: Rekapitulasi Total Biaya.</li>"
                f"</ul>"
                f"<p>Link Google Drive:</p>"
                f"<ul>"
                f"<li><a href='{link_pdf_nonsbo}'>Link PDF Non-SBO</a></li>"
                f"<li><a href='{link_pdf_rekap}'>Link PDF Rekapitulasi</a></li>"
                f"</ul>"
            )

            # 1) KONTRAKTOR → body utama + link upload materai/SPH
            if kontraktor_emails:
                kontraktor_body = (
                    base_body +
                    f"<p>Silakan upload Rekapitulasi RAB Termaterai & SPH melalui link berikut:</p>"
                    f"<p><a href='https://materai-rab-pi.vercel.app/login' "
                    f"target='_blank'>UPLOAD REKAP RAB TERMATERAI & SPH</a></p>"
                )

                google_provider.send_email(
                    to=kontraktor_emails,
                    subject=subject,
                    html_body=kontraktor_body,
                    attachments=email_attachments
                )

            # 2) KOORDINATOR → hanya body utama
            if coordinator_emails:
                google_provider.send_email(
                    to=coordinator_emails,
                    subject=subject,
                    html_body=base_body,
                    attachments=email_attachments
                )

            # 3) MANAGER → hanya body utama
            if manager_email:
                google_provider.send_email(
                    to=[manager_email],
                    subject=subject,
                    html_body=base_body,
                    attachments=email_attachments
                )

            return render_template('response_page.html', 
                title='Persetujuan Berhasil', 
                message='Tindakan Anda telah berhasil diproses.', 
                logo_url=logo_url
            )

    except Exception as e:
        traceback.print_exc()
        return render_template('response_page.html', title='Internal Error', message=f'An internal error occurred: {str(e)}', logo_url=logo_url), 500

@app.route('/api/handle_rab_2_approval', methods=['GET', 'POST'])
def handle_rab_2_approval():
    if request.method == 'POST':
        data = request.form
    else:
        data = request.args

    action = data.get('action')
    row_str = data.get('row')
    level = data.get('level')
    approver = data.get('approver')
    reason = data.get('reason', '-')
    
    logo_url = url_for('static', filename='Alfamart-Emblem.png', _external=True)

    if not all([action, row_str, level, approver]):
        return "Parameter tidak lengkap", 400

    try:
        row = int(row_str)
        # BUKA SPREADSHEET RAB 2
        spreadsheet = google_provider.gspread_client.open_by_key(config.SPREADSHEET_ID_RAB_2)
        worksheet = spreadsheet.worksheet(config.DATA_ENTRY_SHEET_NAME_RAB_2)
        
        # Baca data baris
        row_data = google_provider.get_row_data_by_sheet(worksheet, row)
        
        if not row_data:
            return "Data tidak ditemukan", 404

        WIB = timezone(timedelta(hours=7))
        current_time = datetime.datetime.now(WIB).isoformat()
        
        # --- LOGIKA APPROVAL ---
        if action == 'approve':
            # =================================================================
            # LOGIKA KOORDINATOR (Sama seperti sebelumnya)
            # =================================================================
            if level == 'coordinator':
                # 1. Update Status Spreadsheet
                google_provider.update_cell_by_sheet(worksheet, row, config.COLUMN_NAMES.STATUS, config.STATUS.WAITING_FOR_MANAGER)
                google_provider.update_cell_by_sheet(worksheet, row, config.COLUMN_NAMES.KOORDINATOR_APPROVER, approver)
                google_provider.update_cell_by_sheet(worksheet, row, config.COLUMN_NAMES.KOORDINATOR_APPROVAL_TIME, current_time)
                
                # 2. Update data lokal
                row_data[config.COLUMN_NAMES.KOORDINATOR_APPROVER] = approver
                row_data[config.COLUMN_NAMES.KOORDINATOR_APPROVAL_TIME] = current_time

                # 3. Cari Email Manager
                cabang = row_data.get('Cabang')
                manager_email = google_provider.get_email_by_jabatan(cabang, config.JABATAN.MANAGER)
                
                if manager_email:
                    base_url = "https://sparta-backend-5hdj.onrender.com" 
                    approval_url_manager = f"{base_url}/api/handle_rab_2_approval?action=approve&row={row}&level=manager&approver={manager_email}"
                    rejection_url_manager = f"{base_url}/api/reject_form/rab_kedua?action=reject&row={row}&level=manager&approver={manager_email}"
                    
                    # 4. Render Email HTML
                    email_html_manager = render_template(
                        'email_template.html', 
                        doc_type="RAB", 
                        level='Manajer', 
                        form_data=row_data, 
                        approval_url=approval_url_manager,
                        rejection_url=rejection_url_manager, 
                        additional_info=f"Telah disetujui oleh Koordinator: {approver}"
                    )
                    
                    # 5. Generate PDF Otomatis
                    jenis_toko = row_data.get('Proyek', 'N/A')
                    nama_toko = row_data.get('Nama_Toko', row_data.get('nama_toko', 'N/A'))
                    lingkup_pekerjaan = row_data.get('Lingkup_Pekerjaan', row_data.get('lingkup_pekerjaan', 'N/A'))

                    pdf_nonsbo_bytes = create_pdf_from_data(google_provider, row_data, exclude_sbo=True)
                    pdf_recap_bytes = create_recap_pdf(google_provider, row_data)
                    
                    pdf_nonsbo_filename = f"IL_NON-SBO_{jenis_toko}_{row_data.get('Nomor Ulok')}.pdf"
                    pdf_recap_filename = f"REKAP_IL_{jenis_toko}_{row_data.get('Nomor Ulok')}.pdf"
                    
                    final_attachments = [
                        (pdf_nonsbo_filename, pdf_nonsbo_bytes, 'application/pdf'),
                        (pdf_recap_filename, pdf_recap_bytes, 'application/pdf')
                    ]

                    # 6. Download File Manual (IL) untuk Koordinator -> Manager
                    link_pdf_manual = row_data.get(config.COLUMN_NAMES.LINK_PDF_IL, '')
                    if link_pdf_manual and len(link_pdf_manual) > 5:
                        m_name, m_bytes, m_type = google_provider.download_file_from_link(link_pdf_manual)
                        if m_name and m_bytes:
                            final_attachments.append((m_name, m_bytes, m_type))

                    # 7. Kirim Email
                    google_provider.send_email(
                        manager_email, 
                        f"[TAHAP 2: PERLU PERSETUJUAN] Instruksi Lapangan Proyek {nama_toko} - {lingkup_pekerjaan}", 
                        email_html_manager, 
                        attachments=final_attachments
                    )
                
                return render_template('response_page.html', title='Sukses', message='Disetujui Koordinator (IL)', logo_url=logo_url)

            # =================================================================
            # LOGIKA MANAGER
            # =================================================================
            elif level == 'manager':
                # 1. Update Status & Kolom Manager di Spreadsheet
                google_provider.update_cell_by_sheet(worksheet, row, config.COLUMN_NAMES.STATUS, config.STATUS.APPROVED)
                google_provider.update_cell_by_sheet(worksheet, row, config.COLUMN_NAMES.MANAGER_APPROVER, approver)
                google_provider.update_cell_by_sheet(worksheet, row, config.COLUMN_NAMES.MANAGER_APPROVAL_TIME, current_time)
                
                # 2. Update data lokal variable row_data
                row_data[config.COLUMN_NAMES.STATUS] = config.STATUS.APPROVED
                row_data[config.COLUMN_NAMES.MANAGER_APPROVER] = approver
                row_data[config.COLUMN_NAMES.MANAGER_APPROVAL_TIME] = current_time

                # 3. Generate PDF FINAL (Yang sudah ada nama Manager)
                jenis_toko = row_data.get('Proyek', 'N/A')
                nama_toko = row_data.get('Nama_Toko', row_data.get('nama_toko', 'N/A'))
                cabang = row_data.get('Cabang')
                lingkup_pekerjaan = row_data.get('Lingkup_Pekerjaan', row_data.get('lingkup_pekerjaan', 'N/A'))


                pdf_nonsbo_bytes = create_pdf_from_data(google_provider, row_data, exclude_sbo=True)
                pdf_nonsbo_filename = f"DISETUJUI_IL_NON-SBO_{jenis_toko}_{row_data.get('Nomor Ulok')}.pdf"

                pdf_recap_bytes = create_recap_pdf(google_provider, row_data)
                pdf_recap_filename = f"DISETUJUI_REKAP_IL_{jenis_toko}_{row_data.get('Nomor Ulok')}.pdf"

                # 4. Upload PDF Final ke Google Drive
                link_pdf_nonsbo = google_provider.upload_file_to_drive(
                    pdf_nonsbo_bytes, pdf_nonsbo_filename, 'application/pdf', config.PDF_STORAGE_FOLDER_ID
                )
                link_pdf_rekap = google_provider.upload_file_to_drive(
                    pdf_recap_bytes, pdf_recap_filename, 'application/pdf', config.PDF_STORAGE_FOLDER_ID
                )

                # 5. Update Link PDF Baru ke Spreadsheet RAB 2
                google_provider.update_cell_by_sheet(worksheet, row, config.COLUMN_NAMES.LINK_PDF_NONSBO, link_pdf_nonsbo)
                google_provider.update_cell_by_sheet(worksheet, row, config.COLUMN_NAMES.LINK_PDF_REKAP, link_pdf_rekap)

                row_data[config.COLUMN_NAMES.LINK_PDF_NONSBO] = link_pdf_nonsbo
                row_data[config.COLUMN_NAMES.LINK_PDF_REKAP] = link_pdf_rekap

                # 6. Copy ke Sheet Approved RAB 2 (Form3)
                google_provider.copy_to_approved_sheet_kedua(row_data)

                # 7. Kumpulkan Email Penerima
                email_pembuat = row_data.get(config.COLUMN_NAMES.EMAIL_PEMBUAT)
                support_emails = [email_pembuat] if email_pembuat else []
                coordinator_emails = google_provider.get_emails_by_jabatan(cabang, config.JABATAN.KOORDINATOR)
                manager_email = approver

                # 8. Siapkan Attachment Email (PDF Final RAB & Rekap)
                email_attachments = [
                    (pdf_nonsbo_filename, pdf_nonsbo_bytes, 'application/pdf'),
                    (pdf_recap_filename, pdf_recap_bytes, 'application/pdf')
                ]

                # --- BAGIAN DOWNLOAD FILE MANUAL (IL) DARI DRIVE UNTUK FINAL ---
                # Mengambil link dari spreadsheet dan mendownload ulang
                link_pdf_manual = row_data.get(config.COLUMN_NAMES.LINK_PDF_IL, '')
                
                if link_pdf_manual and len(link_pdf_manual) > 5:
                    print(f"System (Manager Approve): Mendownload file manual dari: {link_pdf_manual}")
                    m_name, m_bytes, m_type = google_provider.download_file_from_link(link_pdf_manual)
                    
                    if m_name and m_bytes:
                        email_attachments.append((m_name, m_bytes, m_type))
                        print("System: File manual berhasil dilampirkan di email final.")
                # ---------------------------------------------------------------

                subject = f"[FINAL - DISETUJUI] Pengajuan Instruksi Lapangan Proyek {nama_toko} - {lingkup_pekerjaan}"

                # 9. Buat Body Email Dasar
                base_body = (
                    f"<p>Pengajuan Instruksi Lapangan untuk proyek <b>{nama_toko}</b> di cabang <b>{cabang}</b> "
                    f"telah disetujui sepenuhnya.</p>"
                    f"<p>Tiga versi file PDF Instruksi Lapangan telah dilampirkan (Final):</p>"
                    f"<ul>"
                    f"<li><b>{pdf_nonsbo_filename}</b>: Hanya berisi item pekerjaan di luar SBO.</li>"
                    f"<li><b>{pdf_recap_filename}</b>: Rekapitulasi Total Biaya.</li>"
                    f"<li><b>{m_name}</b>: Rekapitulasi Lampiran Tambahan / Instruksi Lapangan.</li>"
                    
                    f"</ul>"
                    f"<p>Link Google Drive:</p>"
                    f"<ul>"
                    f"<li><a href='{link_pdf_nonsbo}'>Link PDF Non-SBO (Final)</a></li>"
                    f"<li><a href='{link_pdf_rekap}'>Link PDF Rekapitulasi (Final)</a></li>"
                )
                
                if link_pdf_manual:
                     base_body += f"<li><a href='{link_pdf_manual}'>Link Lampiran Tambahan / Instruksi Lapangan</a></li>"
                
                base_body += "</ul>"

                # 10. Kirim Email ke SUPPORT
                if support_emails:
                    support_body = (
                        base_body +
                        f"<hr>"
                        f"<p><b>TINDAKAN DIPERLUKAN:</b></p>"
                        f"<p>Silakan Opname pekerjaan instruksi lapangan melalui link berikut:</p>"
                        f"<p><a href='https://sparta-alfamart.vercel.app' "
                        f"target='_blank' style='background-color:#28a745; color:white; padding:10px 15px; text-decoration:none; border-radius:5px;'>INPUT OPNAME IL</a></p>"
                    )
                    google_provider.send_email(
                        to=support_emails,
                        subject=subject,
                        html_body=support_body,
                        attachments=email_attachments
                    )

                # 11. Kirim Email ke KOORDINATOR
                if coordinator_emails:
                    google_provider.send_email(
                        to=coordinator_emails,
                        subject=subject,
                        html_body=base_body,
                        attachments=email_attachments
                    )

                # 12. Kirim Email ke MANAGER
                if manager_email:
                    google_provider.send_email(
                        to=[manager_email],
                        subject=subject,
                        html_body=base_body,
                        attachments=email_attachments
                    )

                return render_template('response_page.html', title='Sukses', message='Disetujui Manager (IL)', logo_url=logo_url)

        elif action == 'reject':
            # 1. Tentukan Status Penolakan (Oleh Koordinator atau Manager)
            if level == 'coordinator':
                new_status = config.STATUS.REJECTED_BY_COORDINATOR
                # Simpan nama penolak di kolom Koordinator
                google_provider.update_cell_by_sheet(worksheet, row, config.COLUMN_NAMES.KOORDINATOR_APPROVER, approver)
                google_provider.update_cell_by_sheet(worksheet, row, config.COLUMN_NAMES.KOORDINATOR_APPROVAL_TIME, current_time)
            elif level == 'manager':
                new_status = config.STATUS.REJECTED_BY_MANAGER
                # Simpan nama penolak di kolom Manager
                google_provider.update_cell_by_sheet(worksheet, row, config.COLUMN_NAMES.MANAGER_APPROVER, approver)
                google_provider.update_cell_by_sheet(worksheet, row, config.COLUMN_NAMES.MANAGER_APPROVAL_TIME, current_time)

            # 2. Update Status & Alasan di Spreadsheet
            google_provider.update_cell_by_sheet(worksheet, row, config.COLUMN_NAMES.STATUS, new_status)
            google_provider.update_cell_by_sheet(worksheet, row, 'Alasan Penolakan', reason)
            
            # 3. Kirim Email Notifikasi ke Pembuat (Kontraktor/Support)
            creator_email = row_data.get(config.COLUMN_NAMES.EMAIL_PEMBUAT)
            nama_toko = row_data.get('Nama_Toko', row_data.get('nama_toko', 'N/A'))
            lingkup_pekerjaan = row_data.get('Lingkup_Pekerjaan', row_data.get('lingkup_pekerjaan', 'N/A'))
           
            
            if creator_email:
                subject = f"[DITOLAK] Pengajuan Instruksi Lapangan Proyek {nama_toko} - {lingkup_pekerjaan}"
                penolak = "Koordinator" if level == 'coordinator' else "Branch Manager"
                
                body = (
                    f"<p>Yth. Bapak/Ibu,</p>"
                    f"<p>Pengajuan Instruksi Lapangan untuk proyek <b>{nama_toko}</b> dengan lingkup pekerjaan <b>{lingkup_pekerjaan}</b> telah <b>DITOLAK</b> oleh {penolak} ({approver}).</p>"
                    f"<p><b>Alasan Penolakan:</b></p>"
                    f"<blockquote style='background-color:#ffebeb; border-left:5px solid #dc3545; padding:10px;'><i>{reason}</i></blockquote>"
                    f"<p>Silakan perbaiki dan ajukan revisi melalui sistem.</p>"
                    f"<p>https://sparta-alfamart.vercel.app</p>"
                    f"<p>Terima kasih.</p>"
                )
                
                # Kirim email tanpa attachment
                google_provider.send_email(to=creator_email, subject=subject, html_body=body)

            return render_template('response_page.html', title='Ditolak', message='Pengajuan IL Berhasil Ditolak.', logo_url=logo_url)

    except Exception as e:
        traceback.print_exc()
        return f"Error: {str(e)}", 500


# --- ENDPOINTS UNTUK ALUR KERJA SPK ---
@app.route('/api/get_approved_rab', methods=['GET'])
def get_approved_rab():
    user_cabang = request.args.get('cabang')
    if not user_cabang:
        return jsonify({"error": "Cabang parameter is missing"}), 400
    try:
        approved_rabs = google_provider.get_approved_rab_by_cabang(user_cabang)
        return jsonify(approved_rabs), 200
    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500

@app.route('/api/get_approved_rab_kedua', methods=['GET'])
def get_approved_rab_kedua():
    user_cabang = request.args.get('cabang')
    if not user_cabang:
        return jsonify({"error": "Cabang parameter is missing"}), 400
    try:
        approved_rabs = google_provider.get_approved_rab_by_cabang_kedua(user_cabang)
        return jsonify(approved_rabs), 200
    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route('/api/get_kontraktor', methods=['GET'])
def get_kontraktor():
    user_cabang = request.args.get('cabang')
    if not user_cabang:
        return jsonify({"error": "Cabang parameter is missing"}), 400
    try:
        kontraktor_list = google_provider.get_kontraktor_by_cabang(user_cabang)
        return jsonify(kontraktor_list), 200
    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500

@app.route('/api/get_all_ulok_rab', methods=['GET'])
def get_all_rab_ulok_data_list():
    try:
        # Mengembalikan list of objects langsung form 3
        ulok_list = google_provider.get_all_rab_ulok()
        return jsonify(ulok_list), 200
    except Exception as e:
        traceback.print_exc()
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route('/api/get_gantt_data', methods=['GET'])
def get_gantt_data():
    ulok = request.args.get('ulok')
    lingkup = request.args.get('lingkup') # Ambil parameter lingkup
    
    if not ulok or not lingkup:
        return jsonify({"status": "error", "message": "Parameter ulok dan lingkup diperlukan"}), 400
        
    try:
        # Panggil fungsi service dengan dua parameter
        result = google_provider.get_gantt_data_by_ulok(ulok, lingkup)
        
        # Validasi jika data tidak ditemukan
        if not result['rab']:
            return jsonify({"status": "error", "message": "Data tidak ditemukan untuk kombinasi Ulok dan Lingkup tersebut"}), 404
            
        return jsonify({
            "status": "success",
            "rab": result['rab'],
            "filtered_categories": result['filtered_categories'],
            "gantt_data": result['gantt'],
            "day_gantt_data": result['day_gantt'],
            "dependency_data": result['dependency']
        }), 200
        
    except Exception as e:
        traceback.print_exc()
        return jsonify({"status": "error", "message": str(e)}), 500


# get ulok by email 
@app.route('/api/get_ulok_by_email', methods=['GET'])
def get_ulok_by_email():
    email = request.args.get('email')
    if not email:
        return jsonify({"error": "Parameter email kosong"}), 400

    try:
        ulok_list = google_provider.get_ulok_by_email(email)
        return jsonify(ulok_list), 200
    except Exception as e:
        traceback.print_exc()
        return jsonify({"status": "error", "message": str(e)}), 500

# Get ulok by Cabang PIC

@app.route('/api/get_ulok_by_cabang_pic', methods=['GET'])
def get_ulok_by_cabang_pic():
    cabang = request.args.get('cabang')
    if not cabang:
        return jsonify({"error": "Parameter cabang kosong"}), 400
    try:
        ulok_list = google_provider.get_ulok_by_cabang_pic(cabang)
        return jsonify(ulok_list), 200
    except Exception as e:
        traceback.print_exc()
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route('/api/gantt/insert', methods=['POST'])
def insert_gantt_data():
    """
    Insert data baru ke sheet Gantt Chart.
    Request body berisi field-field sesuai header sheet.
    """
    data = request.get_json()
    if not data:
        return jsonify({"status": "error", "message": "Request body kosong"}), 400
    
    # Validasi field wajib
    nomor_ulok = data.get(config.COLUMN_NAMES.LOKASI)
    lingkup = data.get(config.COLUMN_NAMES.LINGKUP_PEKERJAAN)
    
    if not nomor_ulok or not lingkup:
        return jsonify({
            "status": "error", 
            "message": "Field 'Nomor Ulok' dan 'Lingkup_Pekerjaan' wajib diisi"
        }), 400
    
    try:
        # Set timestamp jika belum ada
        if not data.get(config.COLUMN_NAMES.TIMESTAMP):
            WIB = timezone(timedelta(hours=7))
            data[config.COLUMN_NAMES.TIMESTAMP] = datetime.datetime.now(WIB).isoformat()
        
        result = google_provider.insert_gantt_chart_data(data)
        
        if result["success"]:
            return jsonify({
                "status": "success",
                "message": result["message"],
                "row_index": result.get("row_index")
            }), 201
        else:
            return jsonify({
                "status": "error",
                "message": result["message"]
            }), 400
            
    except Exception as e:
        traceback.print_exc()
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route('/api/gantt/day/insert', methods=['POST'])
def insert_day_gantt_data():
    """
    Insert atau Remove data masif ke/dari sheet Day Gantt Chart.
    
    Mode Insert - Request body bisa dalam 2 format:
    
    Format 1 - List langsung:
    [
        {
            "Nomor Ulok": "asa-asa-sas",
            "Lingkup_Pekerjaan": "ME",
            "Kategori": "Persiapan",
            "h_awal": "19/12/2025",
            "h_akhir": "21/12/2025"
        },
        ...
    ]
    
    Format 2 - Object dengan kategori_data:
    {
        "nomor_ulok": "asa-asa-sas",
        "lingkup_pekerjaan": "ME",
        "kategori_data": [
            {"Kategori": "Persiapan", "h_awal": "19/12/2025", "h_akhir": "21/12/2025"},
            {"Kategori": "Pembersihan", "h_awal": "23/12/2025", "h_akhir": "25/12/2025"},
            {"Kategori": "Bobokan", "h_awal": "28/12/2025", "h_akhir": "30/12/2025"}
        ]
    }
    
    Mode Remove - Object dengan remove_kategori_data:
    {
        "nomor_ulok": "asa-asa-sas",
        "lingkup_pekerjaan": "ME",
        "remove_kategori_data": [
            {"Kategori": "Persiapan", "h_awal": "19/12/2025", "h_akhir": "21/12/2025"},
            {"Kategori": "Pembersihan", "h_awal": "23/12/2025", "h_akhir": "25/12/2025"}
        ]
    }
    
    Note: Jika h_awal dan h_akhir tidak diisi pada remove, akan menghapus semua baris
    dengan kategori tersebut untuk nomor_ulok dan lingkup_pekerjaan yang sama.
    """
    data = request.get_json()
    if not data:
        return jsonify({"status": "error", "message": "Request body kosong"}), 400
    
    try:
        # Cek format request
        if isinstance(data, list):
            # Format 1: List langsung (Insert mode)
            if len(data) == 0:
                return jsonify({
                    "status": "error",
                    "message": "List data kosong"
                }), 400
            
            result = google_provider.insert_day_gantt_chart_data(data)
        
        elif isinstance(data, dict):
            # Cek apakah mode Remove
            remove_kategori_data = data.get('remove_kategori_data')
            
            if remove_kategori_data is not None:
                # Mode Remove
                nomor_ulok = data.get('nomor_ulok') or data.get(config.COLUMN_NAMES.LOKASI)
                lingkup_pekerjaan = data.get('lingkup_pekerjaan') or data.get(config.COLUMN_NAMES.LINGKUP_PEKERJAAN)
                
                if not nomor_ulok or not lingkup_pekerjaan:
                    return jsonify({
                        "status": "error",
                        "message": "Field 'nomor_ulok' dan 'lingkup_pekerjaan' wajib diisi untuk mode remove"
                    }), 400
                
                if not isinstance(remove_kategori_data, list) or len(remove_kategori_data) == 0:
                    return jsonify({
                        "status": "error",
                        "message": "Field 'remove_kategori_data' harus berupa list yang tidak kosong"
                    }), 400
                
                result = google_provider.remove_day_gantt_chart_data(
                    nomor_ulok, 
                    lingkup_pekerjaan, 
                    remove_kategori_data
                )
                
                if result["success"]:
                    return jsonify({
                        "status": "success",
                        "message": result["message"],
                        "deleted_count": result.get("deleted_count", 0),
                        "deleted_items": result.get("deleted_items", [])
                    }), 200
                else:
                    return jsonify({
                        "status": "error",
                        "message": result["message"]
                    }), 400
            
            # Format 2: Object dengan kategori_data (Insert mode)
            nomor_ulok = data.get('nomor_ulok') or data.get(config.COLUMN_NAMES.LOKASI)
            lingkup_pekerjaan = data.get('lingkup_pekerjaan') or data.get(config.COLUMN_NAMES.LINGKUP_PEKERJAAN)
            kategori_data = data.get('kategori_data')
            
            if not nomor_ulok or not lingkup_pekerjaan:
                return jsonify({
                    "status": "error",
                    "message": "Field 'nomor_ulok' dan 'lingkup_pekerjaan' wajib diisi"
                }), 400
            
            if not kategori_data or not isinstance(kategori_data, list):
                return jsonify({
                    "status": "error",
                    "message": "Field 'kategori_data' harus berupa list yang tidak kosong"
                }), 400
            
            result = google_provider.insert_day_gantt_chart_single(nomor_ulok, lingkup_pekerjaan, kategori_data)
        
        else:
            return jsonify({
                "status": "error",
                "message": "Format request tidak valid. Gunakan list atau object."
            }), 400
        
        if result["success"]:
            return jsonify({
                "status": "success",
                "message": result["message"],
                "details": result.get("details", {})
            }), 201
        else:
            return jsonify({
                "status": "error",
                "message": result["message"]
            }), 400
            
    except Exception as e:
        traceback.print_exc()
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route('/api/gantt/dependency/insert', methods=['POST'])
def insert_dependency_gantt_data():
    """
    Insert atau Remove data masif ke/dari sheet Dependency Gantt.
    
    Mode Insert - Request body bisa dalam 2 format:
    
    Format 1 - List langsung:
    [
        {
            "Nomor Ulok": "asa-asa-sas",
            "Lingkup_Pekerjaan": "ME",
            "Kategori": "INSTALASI",
            "Kategori_Terikat": "FIXTURE"
        },
        ...
    ]
    
    Format 2 - Object dengan dependency_data:
    {
        "nomor_ulok": "asa-asa-sas",
        "lingkup_pekerjaan": "ME",
        "dependency_data": [
            {"Kategori": "INSTALASI", "Kategori_Terikat": "FIXTURE"},
            {"Kategori": "FIXTURE", "Kategori_Terikat": "PEKERJAAN TAMBAHAN"}
        ]
    }
    
    Mode Remove - Object dengan remove_dependency_data:
    {
        "nomor_ulok": "asa-asa-sas",
        "lingkup_pekerjaan": "ME",
        "remove_dependency_data": [
            {"Kategori": "INSTALASI", "Kategori_Terikat": "FIXTURE"},
            {"Kategori": "FIXTURE", "Kategori_Terikat": "PEKERJAAN TAMBAHAN"}
        ]
    }
    """
    data = request.get_json()
    if not data:
        return jsonify({"status": "error", "message": "Request body kosong"}), 400
    
    try:
        # Cek format request
        if isinstance(data, list):
            # Format 1: List langsung (Insert mode)
            if len(data) == 0:
                return jsonify({
                    "status": "error",
                    "message": "List data kosong"
                }), 400
            
            result = google_provider.insert_dependency_gantt_data(data)
        
        elif isinstance(data, dict):
            # Cek apakah mode Remove
            remove_dependency_data = data.get('remove_dependency_data')
            
            if remove_dependency_data is not None:
                # Mode Remove
                nomor_ulok = data.get('nomor_ulok') or data.get(config.COLUMN_NAMES.LOKASI)
                lingkup_pekerjaan = data.get('lingkup_pekerjaan') or data.get(config.COLUMN_NAMES.LINGKUP_PEKERJAAN)
                
                if not nomor_ulok or not lingkup_pekerjaan:
                    return jsonify({
                        "status": "error",
                        "message": "Field 'nomor_ulok' dan 'lingkup_pekerjaan' wajib diisi untuk mode remove"
                    }), 400
                
                if not isinstance(remove_dependency_data, list) or len(remove_dependency_data) == 0:
                    return jsonify({
                        "status": "error",
                        "message": "Field 'remove_dependency_data' harus berupa list yang tidak kosong"
                    }), 400
                
                result = google_provider.remove_dependency_gantt_data(
                    nomor_ulok, 
                    lingkup_pekerjaan, 
                    remove_dependency_data
                )
                
                if result["success"]:
                    return jsonify({
                        "status": "success",
                        "message": result["message"],
                        "deleted_count": result.get("deleted_count", 0),
                        "deleted_items": result.get("deleted_items", [])
                    }), 200
                else:
                    return jsonify({
                        "status": "error",
                        "message": result["message"]
                    }), 400
            
            # Format 2: Object dengan dependency_data (Insert mode)
            nomor_ulok = data.get('nomor_ulok') or data.get(config.COLUMN_NAMES.LOKASI)
            lingkup_pekerjaan = data.get('lingkup_pekerjaan') or data.get(config.COLUMN_NAMES.LINGKUP_PEKERJAAN)
            dependency_data = data.get('dependency_data')
            
            if not nomor_ulok or not lingkup_pekerjaan:
                return jsonify({
                    "status": "error",
                    "message": "Field 'nomor_ulok' dan 'lingkup_pekerjaan' wajib diisi"
                }), 400
            
            if not dependency_data or not isinstance(dependency_data, list):
                return jsonify({
                    "status": "error",
                    "message": "Field 'dependency_data' harus berupa list yang tidak kosong"
                }), 400
            
            result = google_provider.insert_dependency_gantt_single(nomor_ulok, lingkup_pekerjaan, dependency_data)
        
        else:
            return jsonify({
                "status": "error",
                "message": "Format request tidak valid. Gunakan list atau object."
            }), 400
        
        if result["success"]:
            return jsonify({
                "status": "success",
                "message": result["message"],
                "details": result.get("details", {})
            }), 201
        else:
            return jsonify({
                "status": "error",
                "message": result["message"]
            }), 400
            
    except Exception as e:
        traceback.print_exc()
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route('/api/gantt/pengawasan/insert', methods=['POST'])
def insert_pengawasan_to_gantt():
    """
    Insert atau Remove data Pengawasan (angka hari) ke/dari kolom Pengawasan_1 s/d Pengawasan_10.
    
    Mode Insert:
    - Cari baris berdasarkan Nomor Ulok dan Lingkup_Pekerjaan
    - Insert ke kolom Pengawasan yang pertama masih kosong
    
    Mode Remove:
    - Jika ada field 'remove_day', hapus nilai tersebut dari kolom Pengawasan
    - Setelah dihapus, nilai-nilai di kolom berikutnya akan digeser ke kiri
    - Contoh: Pengawasan_1=5, Pengawasan_2=10, Pengawasan_3=15, remove_day=5
              Hasil: Pengawasan_1=10, Pengawasan_2=15, Pengawasan_3=""
    
    Request body untuk Insert:
    {
        "nomor_ulok": "Z001-2512-TEST",
        "lingkup_pekerjaan": "SIPIL",
        "pengawasan_day": 5
    }
    
    Request body untuk Remove:
    {
        "nomor_ulok": "Z001-2512-TEST",
        "lingkup_pekerjaan": "SIPIL",
        "remove_day": 5
    }
    """
    data = request.get_json()
    if not data:
        return jsonify({"status": "error", "message": "Request body kosong"}), 400
    
    nomor_ulok = data.get('nomor_ulok') or data.get(config.COLUMN_NAMES.LOKASI)
    lingkup_pekerjaan = data.get('lingkup_pekerjaan') or data.get(config.COLUMN_NAMES.LINGKUP_PEKERJAAN)
    pengawasan_day = data.get('pengawasan_day')
    remove_day = data.get('remove_day')
    
    if not nomor_ulok:
        return jsonify({
            "status": "error",
            "message": "Field 'nomor_ulok' wajib diisi"
        }), 400
    
    if not lingkup_pekerjaan:
        return jsonify({
            "status": "error",
            "message": "Field 'lingkup_pekerjaan' wajib diisi"
        }), 400
    
    # Mode Remove: jika ada remove_day
    if remove_day is not None:
        try:
            result = google_provider.remove_pengawasan_from_gantt_chart(
                nomor_ulok, 
                lingkup_pekerjaan, 
                remove_day
            )
            
            if result["success"]:
                return jsonify({
                    "status": "success",
                    "message": result["message"],
                    "row_index": result.get("row_index"),
                    "removed_value": result.get("removed_value"),
                    "new_values": result.get("new_values")
                }), 200
            else:
                return jsonify({
                    "status": "error",
                    "message": result["message"]
                }), 400
                
        except Exception as e:
            traceback.print_exc()
            return jsonify({"status": "error", "message": str(e)}), 500
    
    # Mode Insert: jika ada pengawasan_day
    if pengawasan_day is None:
        return jsonify({
            "status": "error",
            "message": "Field 'pengawasan_day' atau 'remove_day' wajib diisi"
        }), 400
    
    try:
        result = google_provider.insert_pengawasan_to_gantt_chart(
            nomor_ulok, 
            lingkup_pekerjaan, 
            pengawasan_day
        )
        
        if result["success"]:
            return jsonify({
                "status": "success",
                "message": result["message"],
                "row_index": result.get("row_index"),
                "column_name": result.get("column_name"),
                "value": result.get("value")
            }), 200
        else:
            return jsonify({
                "status": "error",
                "message": result["message"]
            }), 400
            
    except Exception as e:
        traceback.print_exc()
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route('/api/gantt/day/keterlambatan', methods=['POST'])
def update_keterlambatan_day_gantt():
    """
    Insert atau Update kolom keterlambatan di sheet day_gantt_chart.
    Cari baris berdasarkan Nomor Ulok, Lingkup_Pekerjaan, Kategori, h_awal, dan h_akhir.
    Jika ditemukan, update kolom keterlambatan. Jika tidak, insert baris baru.
    
    Request body:
    {
        "nomor_ulok": "Z001-2512-TEST",
        "lingkup_pekerjaan": "ME",
        "kategori": "INSTALASI",
        "h_awal": "29/12/2025",
        "h_akhir": "31/12/2025",
        "keterlambatan": 3
    }
    """
    data = request.get_json()
    if not data:
        return jsonify({"status": "error", "message": "Request body kosong"}), 400
    
    nomor_ulok = data.get('nomor_ulok') or data.get(config.COLUMN_NAMES.LOKASI)
    lingkup_pekerjaan = data.get('lingkup_pekerjaan') or data.get(config.COLUMN_NAMES.LINGKUP_PEKERJAAN)
    kategori = data.get('kategori') or data.get('Kategori')
    h_awal = data.get('h_awal') or data.get(config.COLUMN_NAMES.HARI_AWAL)
    h_akhir = data.get('h_akhir') or data.get(config.COLUMN_NAMES.HARI_AKHIR)
    keterlambatan = data.get('keterlambatan')
    
    # Validasi field wajib
    if not nomor_ulok:
        return jsonify({"status": "error", "message": "Field 'nomor_ulok' wajib diisi"}), 400
    
    if not lingkup_pekerjaan:
        return jsonify({"status": "error", "message": "Field 'lingkup_pekerjaan' wajib diisi"}), 400
    
    if not kategori:
        return jsonify({"status": "error", "message": "Field 'kategori' wajib diisi"}), 400
    
    if not h_awal:
        return jsonify({"status": "error", "message": "Field 'h_awal' wajib diisi"}), 400
    
    if not h_akhir:
        return jsonify({"status": "error", "message": "Field 'h_akhir' wajib diisi"}), 400
    
    if keterlambatan is None:
        return jsonify({"status": "error", "message": "Field 'keterlambatan' wajib diisi"}), 400
    
    try:
        result = google_provider.update_keterlambatan_day_gantt(
            nomor_ulok,
            lingkup_pekerjaan,
            kategori,
            h_awal,
            h_akhir,
            keterlambatan
        )
        
        if result["success"]:
            return jsonify({
                "status": "success",
                "message": result["message"],
                "row_index": result.get("row_index"),
                "action": result.get("action"),
                "value": result.get("value")
            }), 200
        else:
            return jsonify({
                "status": "error",
                "message": result["message"]
            }), 400
            
    except Exception as e:
        traceback.print_exc()
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route('/api/get_spk_status', methods=['GET'])
def get_spk_status():
    ulok = request.args.get('ulok')
    lingkup = request.args.get('lingkup')  # 🔥 tambahan baru

    if not ulok:
        return jsonify({"error": "Parameter ulok kosong"}), 400

    if not lingkup:
        return jsonify({"error": "Parameter lingkup kosong"}), 400

    try:
        spk_sheet = google_provider.sheet.worksheet(config.SPK_DATA_SHEET_NAME)
        records = spk_sheet.get_all_records()

        # --- PERBAIKAN: Normalisasi Input ---
        target_ulok = str(ulok).replace("-", "").strip().upper()
        target_lingkup = str(lingkup).strip().lower()

        found_record = None
        found_index = -1

        for i, row in enumerate(records, start=2):  # row 2 = data pertama
            # Normalisasi data dari Sheet
            row_ulok_raw = str(row.get("Nomor Ulok", ""))
            row_ulok = row_ulok_raw.replace("-", "").strip().upper()

            # Handle variasi nama kolom Lingkup
            row_lingkup_raw = row.get("Lingkup Pekerjaan", row.get("Lingkup_Pekerjaan", ""))
            row_lingkup = str(row_lingkup_raw).strip().lower()

            if row_ulok == target_ulok and row_lingkup == target_lingkup:
                found_record = row
                found_index = i

        if found_record:
            return jsonify({
                "Status": found_record.get("Status"),
                "RowIndex": found_index,
                "Data": found_record
            }), 200

        # Tidak ada SPK untuk kombinasi ULok + Lingkup → boleh submit
        return jsonify(None), 200
        
    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500

@app.route('/api/submit_spk', methods=['POST'])
def submit_spk():
    data = request.get_json()
    if not data:
        return jsonify({"status": "error", "message": "Invalid JSON data"}), 400

    row_index_for_update = data.get("RowIndex")  # Jika revisi
    is_revision = data.get("Revisi") == "YES"

    new_row_index = None  # Untuk mode tambah baris

    try:
        if not is_revision:
            incoming_ulok = data.get("Nomor Ulok", "")
            incoming_lingkup = data.get("Lingkup Pekerjaan", "")

            is_duplicate = google_provider.check_spk_exists(incoming_ulok, incoming_lingkup)

            if is_duplicate:
                return jsonify({
                    "status": "error",
                    "message": (
                        f"SPK untuk Ulok {incoming_ulok} dengan lingkup {incoming_lingkup} "
                        "sudah pernah diajukan dan sedang diproses (Menunggu BM) atau sudah disetujui."
                    )
                }), 409

            # Ambil semua data SPK yang ada
            spk_sheet = google_provider.sheet.worksheet(config.SPK_DATA_SHEET_NAME)
            all_records = spk_sheet.get_all_records()

            for record in all_records:
                existing_ulok = str(record.get("Nomor Ulok", "")).replace("-", "").replace(" ", "").strip().upper()
                
                rec_lingkup_raw = record.get("Lingkup Pekerjaan", record.get("Lingkup_Pekerjaan", ""))
                existing_lingkup = str(rec_lingkup_raw).strip().lower()
                
                status = str(record.get("Status", "")).strip()

                if incoming_ulok == existing_ulok and incoming_lingkup == existing_lingkup:
                    if status != config.STATUS.SPK_REJECTED:
                        return jsonify({
                            "status": "error",
                            "message": (
                                f"SPK untuk Ulok {data.get('Nomor Ulok')} dengan lingkup {data.get('Lingkup Pekerjaan')} "
                                "sudah pernah diajukan dan sedang diproses atau sudah disetujui."
                            )
                        }), 409

        WIB = timezone(timedelta(hours=7))
        now = datetime.datetime.now(WIB)

        # Timestamp & status baru (selalu reset saat submit)
        data['Timestamp'] = now.isoformat()
        data['Status'] = config.STATUS.WAITING_FOR_BM_APPROVAL

        # ---- PERHITUNGAN DURASI ----
        start_date = datetime.datetime.fromisoformat(data['Waktu Mulai'])
        duration = int(data['Durasi'])
        end_date = start_date + timedelta(days=duration - 1)
        data['Waktu Selesai'] = end_date.isoformat()

        # ---- HITUNG BIAYA ----
        total_cost = float(data.get('Grand Total', 0))
        terbilang_text = num2words(total_cost, lang='id').title()
        data['Biaya'] = total_cost
        data['Terbilang'] = f"( {terbilang_text} Rupiah )"

        cabang = data.get('Cabang')
        nama_toko = data.get('Nama_Toko', data.get('nama_toko', 'N/A'))
        kode_toko = data.get('Kode Toko', 'N/A')
        jenis_toko = data.get('Jenis_Toko', data.get('Proyek', 'N/A'))
        lingkup_pekerjaan = data.get('Lingkup Pekerjaan', data.get('Lingkup_Pekerjaan', data.get('lingkup_pekerjaan', 'N/A')))

        data['Nama_Toko'] = nama_toko
        data['Lingkup Pekerjaan'] = lingkup_pekerjaan

        spk_manual_1 = data.get('spk_manual_1', '')
        spk_manual_2 = data.get('spk_manual_2', '')
        cabang_code = google_provider.get_cabang_code(cabang)

        if not is_revision:
            spk_sequence = google_provider.get_next_spk_sequence(cabang, now.year, now.month)
            full_spk_number = f"{spk_sequence:03d}/PROPNDEV-{cabang_code}/{spk_manual_1}/{spk_manual_2}"
        else:
            full_spk_number = data.get("Nomor SPK")

        data['Nomor SPK'] = full_spk_number
        data['PAR'] = data.get('PAR', '')

        # ---- BUAT PDF BARU ----
        pdf_bytes = create_spk_pdf(google_provider, data)
        pdf_filename = f"SPK_{data.get('Proyek')}_{data.get('Nomor Ulok')}.pdf"

        pdf_link = google_provider.upload_file_to_drive(
            pdf_bytes, pdf_filename, 'application/pdf', config.PDF_STORAGE_FOLDER_ID
        )
        data['Link PDF'] = pdf_link

        if is_revision and row_index_for_update:
            google_provider.update_row(
                config.SPK_DATA_SHEET_NAME,
                int(row_index_for_update),
                data
            )
            row_to_notify = int(row_index_for_update)
        else:
            new_row_index = google_provider.append_to_sheet(data, config.SPK_DATA_SHEET_NAME)
            row_to_notify = new_row_index

        # ---- Kirim Email ke Branch Manager ----
        branch_manager_email = google_provider.get_email_by_jabatan(
            cabang, config.JABATAN.BRANCH_MANAGER
        )
        if not branch_manager_email:
            raise Exception(f"Branch Manager email for branch '{cabang}' not found.")

        base_url = "https://sparta-backend-5hdj.onrender.com"
        approval_url = f"{base_url}/api/handle_spk_approval?action=approve&row={row_to_notify}&approver={branch_manager_email}"
        rejection_url = f"{base_url}/api/reject_form/spk?row={row_to_notify}&approver={branch_manager_email}"

        email_html = render_template(
            'email_template.html',
            doc_type="SPK",
            level='Branch Manager',
            form_data=data,
            approval_url=approval_url,
            rejection_url=rejection_url
        )

        google_provider.send_email(
            to=branch_manager_email,
            subject=f"[PERLU PERSETUJUAN BM] SPK Proyek {nama_toko} ({kode_toko}): {jenis_toko} - {lingkup_pekerjaan}",
            html_body=email_html,
            attachments=[(pdf_filename, pdf_bytes, 'application/pdf')]
        )

        return jsonify({"status": "success", "message": "SPK successfully submitted for approval."}), 200

    except Exception as e:
        if new_row_index:
            google_provider.delete_row(config.SPK_DATA_SHEET_NAME, new_row_index)

        traceback.print_exc()
        return jsonify({"status": "error", "message": str(e)}), 500

    
@app.route('/api/reject_form/spk', methods=['GET'])
def reject_form_spk():
    row = request.args.get('row')
    approver = request.args.get('approver')
    
    if not all([row, approver]):
        return "Parameter tidak lengkap.", 400

    spk_sheet = google_provider.sheet.worksheet(config.SPK_DATA_SHEET_NAME)
    row_data = google_provider.get_row_data_by_sheet(spk_sheet, int(row))
    if not row_data:
        return "Data permintaan tidak ditemukan.", 404

    logo_url = url_for('static', filename='Alfamart-Emblem.png', _external=True)
    item_identifier = row_data.get('Nomor Ulok', 'N/A')
    
    return render_template(
        'rejection_form.html',
        form_action=url_for('handle_spk_approval', _external=True),
        row=row,
        approver=approver,
        level=None, # SPK tidak memiliki 'level'
        item_type="SPK",
        item_identifier=item_identifier,
        logo_url=logo_url
    )

@app.route('/api/handle_spk_approval', methods=['GET', 'POST'])
def handle_spk_approval():
    if request.method == 'POST':
        data = request.form
    else:
        data = request.args
        
    action = data.get('action')
    row_str = data.get('row')
    approver = data.get('approver')
    reason = data.get('reason', 'Tidak ada alasan yang diberikan.') # Ambil alasan jika ada
    
    logo_url = url_for('static', filename='Alfamart-Emblem.png', _external=True)

    if not all([action, row_str, approver]):
        return render_template('response_page.html', title='Parameter Tidak Lengkap', message='URL tidak lengkap.', logo_url=logo_url), 400
    
    try:
        row_index = int(row_str)
        spk_sheet = google_provider.sheet.worksheet(config.SPK_DATA_SHEET_NAME)
        row_data = google_provider.get_row_data_by_sheet(spk_sheet, row_index)

        if not row_data:
            return render_template('response_page.html', title='Data Tidak Ditemukan', message='Permintaan ini mungkin sudah dihapus.', logo_url=logo_url)
        
        current_status = row_data.get('Status', '').strip()
        if current_status != config.STATUS.WAITING_FOR_BM_APPROVAL:
            msg = f'Tindakan ini sudah diproses. Status saat ini: <strong>{current_status}</strong>.'
            return render_template('response_page.html', title='Tindakan Sudah Diproses', message=msg, logo_url=logo_url)

        WIB = timezone(timedelta(hours=7))
        current_time = datetime.datetime.now(WIB).isoformat()
        
        initiator_email = row_data.get('Dibuat Oleh')
        
        if action == 'approve':
            new_status = config.STATUS.SPK_APPROVED
            google_provider.update_cell_by_sheet(spk_sheet, row_index, 'Status', new_status)
            google_provider.update_cell_by_sheet(spk_sheet, row_index, 'Disetujui Oleh', approver)
            google_provider.update_cell_by_sheet(spk_sheet, row_index, 'Waktu Persetujuan', current_time)
            
            row_data['Status'] = new_status
            row_data['Disetujui Oleh'] = approver
            row_data['Waktu Persetujuan'] = current_time


            final_pdf_bytes = create_spk_pdf(google_provider, row_data)
            final_pdf_filename = f"SPK_DISETUJUI_{row_data.get('Proyek')}_{row_data.get('Nomor Ulok')}.pdf"
            final_pdf_link = google_provider.upload_file_to_drive(final_pdf_bytes, final_pdf_filename, 'application/pdf', config.PDF_STORAGE_FOLDER_ID)
            google_provider.update_cell_by_sheet(spk_sheet, row_index, 'Link PDF', final_pdf_link)

            nomor_ulok_spk = row_data.get('Nomor Ulok')
            cabang = row_data.get('Cabang')

            # --- INSERT DATA KE SUMMARY OPNAME SHEET ---
            try:
                opname_data = {
                    config.COLUMN_NAMES.CABANG: cabang,
                    config.COLUMN_NAMES.LOKASI: nomor_ulok_spk,
                    config.COLUMN_NAMES.KODE_TOKO: row_data.get('Kode Toko', row_data.get('kode_toko', '')),
                    config.COLUMN_NAMES.NAMA_TOKO: row_data.get('Nama_Toko', row_data.get('nama_toko', '')),
                    config.COLUMN_NAMES.AWAL_SPK: row_data.get('Waktu Mulai', row_data.get('awal_spk', '')),
                    config.COLUMN_NAMES.AKHIR_SPK: row_data.get('Waktu Selesai', row_data.get('akhir_spk', '')),
                    config.COLUMN_NAMES.TAMBAH_SPK: row_data.get('Tambah SPK', row_data.get('tambah_spk', '')),
                    config.COLUMN_NAMES.TANGGAL_SERAH_TERIMA: row_data.get('Tanggal Serah Terima', row_data.get('tanggal_serah_terima', '')),
                    config.COLUMN_NAMES.TANGGAL_OPNAME_FINAL: row_data.get('Tanggal Opname Final', row_data.get('tanggal_opname_final', ''))
                }
                google_provider.append_to_dynamic_sheet(
                    config.OPNAME_SHEET_ID,
                    config.SUMMARY_OPNAME_SHEET_NAME,
                    opname_data
                )
                print(f"Data berhasil diinsert ke Summary Opname untuk Ulok: {nomor_ulok_spk}")
            except Exception as opname_error:
                print(f"Warning: Gagal insert ke Summary Opname: {opname_error}")
                # Tidak raise error agar proses approval tetap berjalan
            # --- END INSERT SUMMARY OPNAME ---
            
            # --- UPDATE DATA SPK KE SUMMARY DATA SHEET ---
            try:
                google_provider.copy_to_summary_sheet(row_data, source_type='SPK')
                print(f"Data SPK berhasil diupdate ke Summary Data untuk Ulok: {nomor_ulok_spk}")
            except Exception as summary_error:
                print(f"Warning: Gagal update ke Summary Data: {summary_error}")
            # --- END UPDATE SUMMARY DATA ---
            
            # --- INSERT DATA KE SUMMARY SERAH TERIMA SHEET ---
            try:
                serah_terima_data = {
                    'kode_toko': row_data.get('Kode Toko', row_data.get('kode_toko', '')),
                    'lingkup_pekerjaan': row_data.get('Lingkup Pekerjaan', row_data.get('Lingkup_Pekerjaan', row_data.get('lingkup_pekerjaan', ''))),
                    'kontraktor': row_data.get('Nama Kontraktor', row_data.get('nama_kontraktor', row_data.get('kontraktor', ''))),
                    'nama_toko': row_data.get('Nama_Toko', row_data.get('nama_toko', '')),
                    'kode_ulok': nomor_ulok_spk,
                    'cabang': cabang
                }
                print(f"DEBUG: Attempting to insert serah terima data: {serah_terima_data}")
                google_provider.append_to_dynamic_sheet(
                    config.PENGAWASAN_SPREADSHEET_ID,
                    config.SUMMARY_SERAH_TERIMA_SHEET_NAME,
                    serah_terima_data
                )
                print(f"Data berhasil diinsert ke Summary Serah Terima untuk Ulok: {nomor_ulok_spk}")
            except Exception as serah_terima_error:
                print(f"Warning: Gagal insert ke Summary Serah Terima: {serah_terima_error}")
                traceback.print_exc()
                # Tidak raise error agar proses approval tetap berjalan
            # --- END INSERT SUMMARY SERAH TERIMA ---
            
            jenis_toko = row_data.get('Jenis_Toko', row_data.get('Proyek', 'N/A'))
            nama_toko = row_data.get('Nama_Toko', row_data.get('nama_toko', 'N/A'))
            lingkup_pekerjaan = row_data.get('Lingkup Pekerjaan', row_data.get('Lingkup_Pekerjaan', row_data.get('lingkup_pekerjaan', 'N/A')))

            manager_email = google_provider.get_email_by_jabatan(cabang, config.JABATAN.MANAGER)
            support_emails = google_provider.get_emails_by_jabatan(cabang, config.JABATAN.SUPPORT)
            coordinator_emails = google_provider.get_emails_by_jabatan(cabang, config.JABATAN.KOORDINATOR)

            # Cari email spesifik pembuat RAB menggunakan Ulok DAN Lingkup Pekerjaan
            pembuat_rab_email = google_provider.get_rab_creator_by_ulok(nomor_ulok_spk, lingkup_pekerjaan) if nomor_ulok_spk else None

            bm_email = approver
            bbm_manager_email = manager_email
            
            kontraktor_list = []
            if pembuat_rab_email:
                kontraktor_list.append(pembuat_rab_email)

            other_recipients = set()
            if initiator_email: other_recipients.add(initiator_email.strip())

            if pembuat_rab_email: other_recipients.add(pembuat_rab_email.strip())

            jenis_toko = row_data.get('Jenis_Toko', row_data.get('Proyek', 'N/A'))
            nama_toko = row_data.get('Nama_Toko', row_data.get('nama_toko', 'N/A'))
            lingkup_pekerjaan = row_data.get('Lingkup Pekerjaan', row_data.get('Lingkup_Pekerjaan', row_data.get('lingkup_pekerjaan', 'N/A')))

            kode_toko = row_data.get('Kode Toko', 'N/A')
            subject = f"[DISETUJUI] SPK Proyek {nama_toko} ({kode_toko}): {jenis_toko} - {lingkup_pekerjaan}"
            
            email_attachments = [(final_pdf_filename, final_pdf_bytes, 'application/pdf')]

            body_bm = (f"<p>SPK yang Anda setujui untuk Toko <b>{nama_toko}</b> pada proyek <b>{jenis_toko} - {lingkup_pekerjaan}</b> ({row_data.get('Nomor Ulok')}) telah disetujui sepenuhnya dan final.</p>"
                       f"<p>File PDF final terlampir.</p>")
            google_provider.send_email(to=[bm_email], subject=subject, html_body=body_bm, attachments=email_attachments)

            if bbm_manager_email:
                 link_input_pic = f"<p>Silakan melakukan input PIC pengawasan melalui link berikut: <a href='https://frontend-form-virid.vercel.app/login-input_pic.html' target='_blank' rel='noopener noreferrer'>Input PIC Pengawasan</a></p>"
                 body_bbm = (f"<p>SPK yang diajukan untuk Toko <b>{nama_toko}</b> pada proyek <b>{jenis_toko} - {lingkup_pekerjaan}</b> ({row_data.get('Nomor Ulok')}) telah disetujui oleh Branch Manager.</p>"
                             f"{link_input_pic}"
                             f"<p>File PDF final terlampir.</p>")
                 google_provider.send_email(to=[bbm_manager_email], subject=subject, html_body=body_bbm, attachments=email_attachments)
                 other_recipients.discard(bbm_manager_email)

            if coordinator_emails:
                body_coord = (f"<p>SPK untuk Toko <b>{nama_toko}</b> pada proyek <b>{jenis_toko} - {lingkup_pekerjaan}</b> ({row_data.get('Nomor Ulok')}) telah disetujui oleh Branch Manager.</p>"
                              f"<p>File PDF final terlampir.</p>")
                
                google_provider.send_email(to=coordinator_emails, subject=subject, html_body=body_coord, attachments=email_attachments)
                
                # Hapus dari other_recipients agar tidak dikirim double (jika coordinator juga initiator)
                for email in coordinator_emails:
                    other_recipients.discard(email)

            opname_recipients = set()
            opname_recipients.update(kontraktor_list)
            
            if opname_recipients:
                link_opname = f"<p>Silakan melakukan Opname melalui link berikut: <a href='https://sparta-alfamart.vercel.app' target='_blank' rel='noopener noreferrer'>Pengisian Opname</a></p>"
                body_opname = (f"<p>SPK untuk Toko <b>{nama_toko}</b> pada proyek <b>{jenis_toko} - {lingkup_pekerjaan}</b> ({row_data.get('Nomor Ulok')}) telah disetujui.</p>"
                               f"{link_opname}"
                               f"<p>File PDF final terlampir.</p>")
                google_provider.send_email(to=list(opname_recipients), subject=subject, html_body=body_opname, attachments=email_attachments)
                
                for email in opname_recipients:
                    other_recipients.discard(email)

            if other_recipients:
                body_default = (f"<p>SPK yang Anda ajukan untuk Toko <b>{nama_toko}</b> pada proyek <b>{jenis_toko} - {lingkup_pekerjaan}</b> ({row_data.get('Nomor Ulok')}) telah disetujui oleh Branch Manager.</p>"
                                f"<p>File PDF final terlampir.</p>")
                google_provider.send_email(to=list(other_recipients), subject=subject, html_body=body_default, attachments=email_attachments)
            
            return render_template('response_page.html', title='Persetujuan Berhasil', message='Terima kasih. Persetujuan Anda telah dicatat.', logo_url=logo_url)

        elif action == 'reject':
            new_status = config.STATUS.SPK_REJECTED
            google_provider.update_cell_by_sheet(spk_sheet, row_index, 'Status', new_status)
            
            google_provider.update_cell_by_sheet(spk_sheet, row_index, 'Alasan Penolakan', reason)

            nama_toko = row_data.get('Nama_Toko', row_data.get('nama_toko', 'N/A'))
            jenis_toko = row_data.get('Jenis_Toko', row_data.get('Proyek', 'N/A'))
            lingkup_pekerjaan = row_data.get('Lingkup Pekerjaan', row_data.get('Lingkup_Pekerjaan', row_data.get('lingkup_pekerjaan', 'N/A')))

            if initiator_email:
                kode_toko = row_data.get('Kode Toko', 'N/A')
                subject = f"[DITOLAK] SPK untuk Proyek {nama_toko} ({kode_toko}): {jenis_toko} - {lingkup_pekerjaan}"
                body = (f"<p>SPK yang Anda ajukan untuk Toko <b>{nama_toko}</b> pada proyek <b>{jenis_toko} - {lingkup_pekerjaan}</b> ({row_data.get('Nomor Ulok')}) telah ditolak oleh Branch Manager.</p>"
                        f"<p><b>Alasan Penolakan:</b></p>"
                        f"<p><i>{reason}</i></p>"
                        f"<p>Silakan ajukan revisi SPK Anda melalui link berikut:</p>"
                        f"<p><a href='https://sparta-alfamart.vercel.app' target='_blank' rel='noopener noreferrer'>Input Ulang SPK</a></p>")
                google_provider.send_email(to=initiator_email, subject=subject, html_body=body)

            return render_template('response_page.html', title='Permintaan Ditolak', message='Status permintaan telah diperbarui menjadi ditolak.', logo_url=logo_url)

    except Exception as e:
        traceback.print_exc()
        return render_template('response_page.html', title='Error Internal', message=f'Terjadi kesalahan: {str(e)}', logo_url=logo_url), 500

# --- ENDPOINTS UNTUK FORM PENGAWASAN ---
@app.route('/api/pengawasan/init_data', methods=['GET'])
def get_pengawasan_init_data():
    cabang = request.args.get('cabang')
    if not cabang:
        return jsonify({"status": "error", "message": "Parameter cabang dibutuhkan."}), 400
    try:
        pic_list, _, _ = google_provider.get_user_info_by_cabang(cabang)
        spk_list = google_provider.get_spk_data_by_cabang(cabang)
        
        return jsonify({
            "status": "success",
            "picList": pic_list,
            "spkList": spk_list
        }), 200
    except Exception as e:
        traceback.print_exc()
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route('/api/pengawasan/get_rab_url', methods=['GET'])
def get_rab_url():
    kode_ulok = request.args.get('kode_ulok')
    if not kode_ulok:
        return jsonify({"status": "error", "message": "Parameter kode_ulok dibutuhkan."}), 400
    try:
        rab_url = google_provider.get_rab_url_by_ulok(kode_ulok)
        if rab_url:
            return jsonify({"status": "success", "rabUrl": rab_url}), 200
        else:
            return jsonify({"status": "error", "message": "URL RAB tidak ditemukan."}), 404
    except Exception as e:
        traceback.print_exc()
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route('/api/pengawasan/get_spk_url', methods=['GET'])
def get_spk_url():
    kode_ulok = request.args.get('kode_ulok')
    if not kode_ulok:
        return jsonify({"status": "error", "message": "Parameter kode_ulok dibutuhkan."}), 400
    try:
        spk_url = google_provider.get_spk_url_by_ulok(kode_ulok)
        if spk_url:
            return jsonify({"status": "success", "spkUrl": spk_url}), 200
        else:
            return jsonify({"status": "error", "message": "URL SPK tidak ditemukan."}), 404
    except Exception as e:
        traceback.print_exc()
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route('/api/pengawasan/submit', methods=['POST'])
def submit_pengawasan():
    data = request.get_json()
    if not data:
        return jsonify({"status": "error", "message": "Data JSON tidak valid"}), 400

    try:
        form_type = data.get('form_type')
        WIB = timezone(timedelta(hours=7))
        timestamp = datetime.datetime.now(WIB)
        
        cabang = data.get('cabang', 'N/A')
        
        if form_type != 'input_pic':
            kode_ulok = data.get('kode_ulok')
            if kode_ulok:
                pic_email = google_provider.get_pic_email_by_ulok(kode_ulok)
                if pic_email:
                    data['pic_building_support'] = pic_email
                else:
                    return jsonify({"status": "error", "message": f"PIC tidak ditemukan untuk Kode Ulok {kode_ulok}. Pastikan proyek ini sudah diinisiasi."}), 404

        pic_list, koordinator_info, manager_info = google_provider.get_user_info_by_cabang(cabang)
        user_info = {
            'pic_list': pic_list,
            'koordinator_info': koordinator_info,
            'manager_info': manager_info
        }

        if form_type == 'input_pic':
            input_pic_data = {
                'Timestamp': timestamp.isoformat(),
                'Cabang': data.get('cabang'),
                'Kode_Ulok': data.get('kode_ulok'),
                'Kategori_Lokasi': data.get('kategori_lokasi'),
                'Tanggal_Mulai_SPK': data.get('tanggal_spk'),
                'PIC_Building_Support': data.get('pic_building_support'),
                'SPK_URL': data.get('spkUrl'),
                'RAB_URL': data.get('rabUrl')
            }
            google_provider.append_to_dynamic_sheet(
                config.PENGAWASAN_SPREADSHEET_ID, 
                config.INPUT_PIC_SHEET_NAME, 
                input_pic_data
            )

            penugasan_data = {
                'Email_BBS': data.get('pic_building_support'),
                'Kode_Ulok': data.get('kode_ulok'),
                'Cabang': data.get('cabang')
            }
            google_provider.append_to_dynamic_sheet(
                config.PENGAWASAN_SPREADSHEET_ID, 
                config.PENUGASAN_SHEET_NAME,
                penugasan_data
            )
            
            tanggal_spk_obj = datetime.datetime.fromisoformat(data.get('tanggal_spk'))
            tanggal_mengawas = get_tanggal_h(tanggal_spk_obj, 2)
            data['tanggal_mengawas'] = tanggal_mengawas.strftime('%d %B %Y')

        else:
            data_to_sheet = {}
            header_mapping = {
                "timestamp": "Timestamp", "kode_ulok": "Kode_Ulok", "status_lokasi": "Status_Lokasi",
                "status_progress1": "Status_Progress1", "catatan1": "Catatan1",
                "status_progress2": "Status_Progress2", "catatan2": "Catatan2",
                "status_progress3": "Status_Progress3", "catatan3": "Catatan3",
                "pengukuran_bowplank": "Pengukuran_Dan_Pemasangan_Bowplank",
                "pekerjaan_tanah": "Pekerjaan_Tanah",
                "berkas_pengawasan": "Berkas_Pengawasan"
            }
            
            if form_type == 'serah_terima':
                data_to_sheet = data
                data_to_sheet['Timestamp'] = timestamp.isoformat()
            else:
                 for key, value in data.items():
                    sheet_header = header_mapping.get(key, key.replace('_', ' ').title().replace(' ', '_'))
                    data_to_sheet[sheet_header] = value
                 data_to_sheet['Timestamp'] = timestamp.isoformat()

            sheet_map = {
                'h2': 'DataH2', 'h5': 'DataH5', 'h7': 'DataH7', 'h8': 'DataH8', 'h10': 'DataH10',
                'h12': 'DataH12', 'h14': 'DataH14', 'h16': 'DataH16', 'h17': 'DataH17',
                'h18': 'DataH18', 'h22': 'DataH22', 'h23': 'DataH23', 'h25': 'DataH25',
                'h28': 'DataH28', 'h32': 'DataH32', 'h33': 'DataH33', 'h41': 'DataH41',
                'serah_terima': 'SerahTerima'
            }
            target_sheet = sheet_map.get(form_type)
            if target_sheet:
                 google_provider.append_to_dynamic_sheet(
                    config.PENGAWASAN_SPREADSHEET_ID, target_sheet, data_to_sheet
                )

        email_details = get_email_details(form_type, data, user_info)
        
        if not email_details['recipients']:
            return jsonify({
                "status": "error", 
                "message": "Tidak ada penerima email yang valid. Pastikan PIC Building Support dipilih dan/atau Koordinator/Manajer terdaftar untuk cabang ini."
            }), 400

        base_url = "https://building-alfamart.vercel.app" 
        next_form_path = FORM_LINKS.get(form_type, {}).get(data.get('kategori_lokasi'), '#')
        
        next_url_with_redirect = f"{base_url.strip('/')}/?redirectTo={next_form_path}" if next_form_path != '#' else None

        email_html = render_template('pengawasan_email_template.html', 
                                     form_data=data,
                                     user_info=user_info,
                                     next_form_url=next_url_with_redirect,
                                     form_type=form_type
                                    )
        
        google_provider.send_email(
            to=email_details['recipients'],
            subject=email_details['subject'],
            html_body=email_html
        )
        
        if form_type == 'input_pic' and 'tanggal_mengawas' in data:
            event_date_obj = datetime.datetime.strptime(data['tanggal_mengawas'], '%d %B %Y')
            google_provider.create_calendar_event({
                'title': f"[REMINDER] Pengawasan H+2: {data.get('kode_ulok')}",
                'description': f"Ini adalah pengingat untuk melakukan pengawasan H+2 untuk toko {data.get('kode_ulok')}. Link untuk mengisi laporan akan dikirimkan melalui email terpisah.",
                'date': event_date_obj.strftime('%Y-%m-%d'),
                'guests': email_details['recipients']
            })

        return jsonify({"status": "success", "message": "Laporan berhasil dikirim."}), 200

    except Exception as e:
        traceback.print_exc()
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route('/api/pengawasan/active_projects', methods=['GET'])
def get_active_projects():
    email = request.args.get('email')
    if not email:
        return jsonify({"status": "error", "message": "Parameter email dibutuhkan."}), 400
    try:
        active_projects = google_provider.get_active_pengawasan_by_pic(email)
        return jsonify({"status": "success", "projects": active_projects}), 200
    except Exception as e:
        traceback.print_exc()
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route('/api/user_info_by_email', methods=['GET'])
def name_dan_cabang_by_email():
    email = request.args.get('email')
    if not email:
        return jsonify({"status": "error", "message": "Parameter email dibutuhkan."}), 400
    try:
        name, cabang = google_provider.get_nama_lengkap_dan_cabang_by_email(email)
        if name:
            return jsonify({"status": "success", "name": name, "cabang": cabang}), 200
        else:
            return jsonify({"status": "error", "message": "Nama tidak ditemukan untuk email tersebut."}), 404
    except Exception as e:
        traceback.print_exc()
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route('/health', methods=['GET'])
def health():
    return jsonify({"status": "ok", "message": "Service is alive"}), 200

# --- ENDPOINT PROCESS SUMMARY OPNAME ---
@app.route('/api/process_summary_opname', methods=['POST'])
def process_summary_opname():
    """
    Endpoint untuk memproses summary opname.
    
    Request body:
    {
        "no_ulok": "Z001-2512-TEST",
        "lingkup_pekerjaan": "SIPIL",
        "jenis_pekerjaan": "PEKERJAAN PERSIAPAN"
    }
    
    Logika:
    1. Cari total_harga_akhir di opname_final berdasarkan no_ulok, lingkup_pekerjaan, jenis_pekerjaan
    2. Cari row di summary_opname berdasarkan no_ulok dan lingkup_pekerjaan
    3. Jika total_harga_akhir positif -> tambahkan ke kerja_tambah
       Jika total_harga_akhir negatif -> tambahkan ke kerja_kurang
    """
    data = request.get_json()
    if not data:
        return jsonify({"status": "error", "message": "Request body tidak boleh kosong."}), 400
    
    no_ulok = data.get('no_ulok') or data.get('nomor_ulok')
    lingkup_pekerjaan = data.get('lingkup_pekerjaan')
    jenis_pekerjaan = data.get('jenis_pekerjaan')
    
    # Validasi input
    if not no_ulok:
        return jsonify({
            "status": "error",
            "message": "Parameter 'no_ulok' atau 'nomor_ulok' dibutuhkan."
        }), 400
    
    if not lingkup_pekerjaan:
        return jsonify({
            "status": "error",
            "message": "Parameter 'lingkup_pekerjaan' dibutuhkan."
        }), 400
    
    if not jenis_pekerjaan:
        return jsonify({
            "status": "error",
            "message": "Parameter 'jenis_pekerjaan' dibutuhkan."
        }), 400
    
    try:
        result = google_provider.process_summary_opname(no_ulok, lingkup_pekerjaan, jenis_pekerjaan)
        
        if result.get("status") == "success":
            return jsonify(result), 200
        else:
            return jsonify(result), 400
            
    except Exception as e:
        traceback.print_exc()
        return jsonify({
            "status": "error",
            "message": f"Terjadi kesalahan server: {str(e)}"
        }), 500


# --- ENDPOINT CHECK STATUS ITEM OPNAME ---
@app.route('/api/check_status_item_opname', methods=['GET'])
def check_status_item_opname():
    """
    Endpoint untuk mengecek status approval item opname.
    
    Query Parameters:
    - no_ulok: Nomor Ulok (wajib)
    - lingkup_pekerjaan: Lingkup Pekerjaan SIPIL/ME (wajib)
    
    Response jika semua approved:
    {
        "status": "approved",
        "message": "Semua item opname sudah ter-approved.",
        "tanggal_opname_final": "06/01/2026",
        ...
    }
    
    Response jika masih ada pending:
    {
        "status": "pending",
        "message": "Masih ada X item yang belum ter-approved.",
        "pending_items": [...],
        ...
    }
    """
    no_ulok = request.args.get('no_ulok') or request.args.get('nomor_ulok')
    lingkup_pekerjaan = request.args.get('lingkup_pekerjaan')
    
    # Validasi input
    if not no_ulok:
        return jsonify({
            "status": "error",
            "message": "Parameter 'no_ulok' atau 'nomor_ulok' dibutuhkan."
        }), 400
    
    if not lingkup_pekerjaan:
        return jsonify({
            "status": "error",
            "message": "Parameter 'lingkup_pekerjaan' dibutuhkan."
        }), 400
    
    try:
        result = google_provider.check_opname_approval_status(no_ulok, lingkup_pekerjaan)
        
        if result.get("status") == "error":
            return jsonify(result), 400
        else:
            return jsonify(result), 200
            
    except Exception as e:
        traceback.print_exc()
        return jsonify({
            "status": "error",
            "message": f"Terjadi kesalahan server: {str(e)}"
        }), 500


# --- ENDPOINT OPNAME LOCKED ---
@app.route('/api/opname_locked', methods=['POST'])
def opname_locked():
    """
    Endpoint untuk mengunci opname dan menyimpan tanggal hari ini ke tanggal_opname_final.
    
    Request body:
    {
        "status": "locked",
        "ulok": "Z001-2512-TEST",
        "lingkup_pekerjaan": "SIPIL"
    }
    
    Response:
    {
        "status": "success",
        "message": "Opname berhasil dikunci (locked) pada tanggal 06/01/2026",
        "tanggal_opname_final": "06/01/2026",
        ...
    }
    """
    data = request.get_json()
    if not data:
        return jsonify({"status": "error", "message": "Request body tidak boleh kosong."}), 400
    
    status_lock = data.get('status')
    no_ulok = data.get('ulok') or data.get('no_ulok') or data.get('nomor_ulok')
    lingkup_pekerjaan = data.get('lingkup_pekerjaan') or data.get('lingkup')
    
    # Validasi input
    if not status_lock or status_lock.lower() != 'locked':
        return jsonify({
            "status": "error",
            "message": "Parameter 'status' harus bernilai 'locked'."
        }), 400
    
    if not no_ulok:
        return jsonify({
            "status": "error",
            "message": "Parameter 'ulok', 'no_ulok', atau 'nomor_ulok' dibutuhkan."
        }), 400
    
    if not lingkup_pekerjaan:
        return jsonify({
            "status": "error",
            "message": "Parameter 'lingkup_pekerjaan' atau 'lingkup' dibutuhkan."
        }), 400
    
    try:
        result = google_provider.lock_opname(no_ulok, lingkup_pekerjaan)
        
        if result.get("status") == "success":
            return jsonify(result), 200
        else:
            return jsonify(result), 400
            
    except Exception as e:
        traceback.print_exc()
        return jsonify({
            "status": "error",
            "message": f"Terjadi kesalahan server: {str(e)}"
        }), 500


# --- ENDPOINT PROXY GAS (MIGRASI DARI BACKEND LAMA) ---
@app.route('/api/form', methods=['GET', 'POST'])
def proxy_form():
    # Ambil form ID dari query param atau body
    form_id = request.args.get('form') or (request.json.get('form') if request.is_json else None)
    
    if not form_id:
        return jsonify({"error": "Invalid or missing form ID"}), 400
        
    form_id = form_id.lower()
    gas_url = GAS_URLS.get(form_id)
    
    if not gas_url:
        return jsonify({"error": "Invalid or missing form ID mapping"}), 400

    try:
        if request.method == 'GET':
            # Forward semua query params kecuali 'form'
            params = {k: v for k, v in request.args.items() if k != 'form'}
            response = requests.get(gas_url, params=params)
            return jsonify(response.json()), response.status_code

        elif request.method == 'POST':
            # Forward json body kecuali 'form'
            payload = request.get_json() if request.is_json else {}
            if 'form' in payload:
                del payload['form']
                
            response = requests.post(gas_url, json=payload, headers={"Content-Type": "application/json"})
            return jsonify(response.json()), response.status_code

    except Exception as e:
        print(f"GAS Proxy error: {str(e)}")
        return jsonify({"error": "Gagal mengakses Google Apps Script", "details": str(e)}), 500

# --- ENDPOINTS APPROVAL PERPANJANGAN SPK (MIGRASI DARI BACKEND LAMA) ---

@app.route('/approval/approve', methods=['GET'])
def approve_perpanjangan():
    try:
        gas_url = request.args.get('gas_url')
        row = request.args.get('row')
        approver = request.args.get('approver')
        ulok = request.args.get('ulok')

        if not gas_url:
             return render_template("response_page.html", title="Error", message="URL GAS tidak ditemukan.")

        # 1. Panggil GAS untuk proses approval
        approval_response = requests.get(f"{gas_url}?action=processApproval&row={row}&approver={approver}")
        approval_data = approval_response.json()

        if approval_data.get('status') != 'success':
            return render_template("response_page.html", title="Error", message=approval_data.get('message', 'Gagal memproses persetujuan.'))

        # 2. Ambil Info Penerima Email dari GAS
        recipient_response = requests.get(f"{gas_url}?action=getRecipientInfo&row={row}")
        final_data = recipient_response.json()

        # [UPDATED] Kirim Email Notifikasi
        final_data['status_persetujuan'] = 'DISETUJUI' 
        
        # Gunakan helper untuk generate subject & body
        subject, body = generate_perpanjangan_email_body(final_data)
        
        # Ambil list email penerima
        recipients = final_data.get('recipients', [])
        
        # Kirim pakai google_provider
        google_provider.send_email(to=recipients, subject=subject, html_body=body)
        # 4. Render halaman sukses
        return render_template("response_page.html", title="Berhasil", message=f"Perpanjangan SPK untuk Ulok {ulok} berhasil DISETUJUI.")

    except Exception as e:
        print(f"Error processing approval: {str(e)}")
        return render_template("response_page.html", title="Error", message="Terjadi kesalahan server saat memproses persetujuan.")


@app.route('/approval/reject', methods=['GET'])
def reject_perpanjangan_page():
    gas_url = request.args.get('gas_url')
    row = request.args.get('row')
    approver = request.args.get('approver')
    ulok = request.args.get('ulok')
    
    # Menggunakan template rejection_form.html yang sudah ada di Sparta Backend
    # Kita perlu sesuaikan parameter yang dikirim agar cocok dengan template
    return render_template(
        "rejection_form.html", 
        form_action="/approval/submit-rejection", # Action mengarah ke endpoint python baru
        gas_url=gas_url, # Perlu passing hidden field ini di template jika belum ada
        row=row,
        approver=approver,
        ulok=ulok,
        item_type="Perpanjangan SPK",
        item_identifier=ulok,
        logo_url=url_for('static', filename='Alfamart-Emblem.png', _external=True)
    )

@app.route('/approval/submit-rejection', methods=['POST'])
def submit_rejection_perpanjangan():
    try:
        # Flask menggunakan request.form untuk data dari form HTML
        gas_url = request.form.get('gas_url') 
        # Cek: Karena rejection_form.html di Sparta mungkin tidak punya input hidden 'gas_url', 
        # Anda mungkin perlu menyesuaikan template atau logika ini. 
        # Asumsi: Jika gas_url tidak ada di form, kita fallback ke default perpanjangan_spk URL
        if not gas_url:
            gas_url = GAS_URLS['perpanjangan_spk']

        row = request.form.get('row')
        approver = request.form.get('approver')
        ulok = request.form.get('ulok') or request.form.get('item_identifier') # Adaptasi nama field
        reason = request.form.get('reason')

        # 1. Panggil GAS untuk proses rejection
        # Gunakan requests, encode params otomatis handle URL encoding
        payload = {
            'action': 'processRejection',
            'row': row,
            'approver': approver,
            'reason': reason
        }
        rejection_response = requests.get(gas_url, params=payload)
        rejection_data = rejection_response.json()

        if rejection_data.get('status') != 'success':
             return render_template("response_page.html", title="Error", message=rejection_data.get('message', 'Gagal memproses penolakan.'))

        # 2. Ambil Info Penerima Email
        recipient_response = requests.get(f"{gas_url}?action=getRecipientInfo&row={row}")
        final_data = recipient_response.json()

        # 3. Kirim Email Notifikasi
        # [UPDATED] Kirim Email Notifikasi
        final_data['status_persetujuan'] = 'DITOLAK'
        final_data['alasan_penolakan'] = reason
        
        # Gunakan helper
        subject, body = generate_perpanjangan_email_body(final_data)
        recipients = final_data.get('recipients', [])
        
        google_provider.send_email(to=recipients, subject=subject, html_body=body)
        return render_template("response_page.html", title="Berhasil", message=f"Perpanjangan SPK untuk Ulok {ulok} berhasil DITOLAK.")

    except Exception as e:
        print(f"Error processing rejection: {str(e)}")
        return render_template("response_page.html", title="Error", message="Gagal memproses penolakan.")

# --- ENDPOINT KIRIM EMAIL GENERAL (MIGRASI DARI BACKEND LAMA) ---
@app.route('/api/send-email', methods=['POST'])
def send_email_general():
    try:
        data = request.get_json()
        form_type = data.get('formType') # Di JS parameternya 'formType' (di body)
        
        # 1. Handle Materai Upload
        if form_type == 'materai_upload':
            # Kirim ke Manager
            manager_recipients = data.get('managerRecipients', [])
            if manager_recipients:
                emails = [r['email'] for r in manager_recipients if 'email' in r]
                subject, body = generate_materai_email_body(data, role='manager')
                google_provider.send_email(to=emails, subject=subject, html_body=body)
                
            # Kirim ke Koordinator/Lainnya
            other_recipients = data.get('otherRecipients', [])
            if other_recipients:
                emails = [r['email'] for r in other_recipients if 'email' in r]
                subject, body = generate_materai_email_body(data, role='coordinator')
                google_provider.send_email(to=emails, subject=subject, html_body=body)

            return jsonify({"status": "success", "message": "Email materai berhasil dikirim"}), 200

        # Jika formType tidak dikenali
        return jsonify({"status": "error", "message": f"Form type '{form_type}' tidak didukung di endpoint ini."}), 400

    except Exception as e:
        traceback.print_exc()
        return jsonify({"status": "error", "message": str(e)}), 500


if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5001))
    app.run(host='0.0.0.0', port=port)