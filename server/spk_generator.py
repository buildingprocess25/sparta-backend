import os
import locale
from weasyprint import HTML
from flask import render_template
from datetime import datetime, timedelta
from num2words import num2words
import config

try:
    locale.setlocale(locale.LC_TIME, 'id_ID.UTF-8')
except locale.Error:
    try:
        locale.setlocale(locale.LC_TIME, 'Indonesian_Indonesia.1252')
    except locale.Error:
        print("Peringatan: Locale Bahasa Indonesia tidak ditemukan.")

def get_nama_lengkap_by_email(google_provider, email, cabang=None, expected_jabatan=None):
    if not email:
        return ""

    normalized_email = str(email).strip().lower()
    normalized_cabang = str(cabang).strip().lower() if cabang else None
    normalized_jabatan = str(expected_jabatan).strip().upper() if expected_jabatan else None
    fallback_name = ""

    try:
        cabang_sheet = google_provider.sheet.worksheet(config.CABANG_SHEET_NAME)
        records = cabang_sheet.get_all_records()

        for record in records:
            record_email = str(record.get('EMAIL_SAT', '')).strip().lower()
            if record_email != normalized_email:
                continue

            record_name = str(record.get('NAMA LENGKAP', '')).strip()
            record_cabang = str(record.get('CABANG', '')).strip().lower()
            record_jabatan = str(record.get('JABATAN', '')).strip().upper()

            if not fallback_name:
                fallback_name = record_name

            if normalized_cabang and record_cabang != normalized_cabang:
                continue

            if normalized_jabatan and record_jabatan != normalized_jabatan:
                continue

            return record_name
    except Exception as e:
        print(f"Error getting name for email {email}: {e}")

    return fallback_name or email

def parse_flexible_timestamp(ts_string):
    if not ts_string or not isinstance(ts_string, str):
        return None
    try:
        return datetime.fromisoformat(ts_string)
    except (ValueError, TypeError):
        pass
    for fmt in ['%m/%d/%Y %H:%M:%S', '%Y-%m-%d %H:%M:%S']:
        try:
            return datetime.strptime(ts_string, fmt)
        except (ValueError, TypeError):
            continue
    return None

def create_approval_details_block(google_provider, approver_value, approval_time_str, cabang=None, expected_jabatan=None):
    if not approver_value:
        return ""

    approver_value_str = str(approver_value).strip()
    if "@" in approver_value_str:
        approver_name = get_nama_lengkap_by_email(
            google_provider,
            approver_value_str,
            cabang=cabang,
            expected_jabatan=expected_jabatan
        )
    else:
        approver_name = approver_value_str

    approval_dt = parse_flexible_timestamp(approval_time_str)
    formatted_time = approval_dt.strftime('%d %B %Y, %H:%M WIB') if approval_dt else "Waktu tidak tersedia"
    
    name_display = f"<strong>( {approver_name} )</strong><br>" if approver_name else f"<strong>( {approver_value_str} )</strong><br>"
    return f"""
    <div class="approval-details">
        {name_display}
        <span class="timestamp">Disetujui pada: {formatted_time}</span>
    </div>
    """

def create_spk_pdf(google_provider, spk_data):
    cabang = spk_data.get('Cabang')
    is_batam_branch = str(cabang or '').strip().upper() == 'BATAM'
    initiator_jabatan = config.JABATAN.KOORDINATOR if is_batam_branch else config.JABATAN.MANAGER
    initiator_role_title = 'Branch Building Coordinator' if is_batam_branch else 'Branch Building & Maintenance Manager'

    initiator_email = spk_data.get('Dibuat Oleh')
    initiator_name = (
        spk_data.get('Dibuat Oleh Nama') or
        spk_data.get('Nama Dibuat Oleh') or
        spk_data.get('Nama Pembuat') or
        initiator_email
    )
    initiator_timestamp = spk_data.get('Timestamp') 
    
    approver_email = spk_data.get('Disetujui Oleh')
    approver_name = (
        spk_data.get('Disetujui Oleh Nama') or
        spk_data.get('Nama Disetujui Oleh') or
        spk_data.get('Nama Branch Manager') or
        approver_email
    )
    approval_time = spk_data.get('Waktu Persetujuan') 

    initiator_details_html = create_approval_details_block(
        google_provider,
        initiator_name,
        initiator_timestamp,
        cabang=cabang,
        expected_jabatan=initiator_jabatan
    )
    
    approver_details_html = create_approval_details_block(
        google_provider,
        approver_name,
        approval_time,
        cabang=cabang,
        expected_jabatan=config.JABATAN.BRANCH_MANAGER
    )

    start_date_obj = datetime.fromisoformat(spk_data.get('Waktu Mulai'))
    duration = int(spk_data.get('Durasi'))
    end_date_obj = start_date_obj + timedelta(days=duration - 1)
    start_date_formatted = start_date_obj.strftime('%d %B %Y')
    end_date_formatted = end_date_obj.strftime('%d %B %Y')

    total_cost = float(spk_data.get('Grand Total', 0))
    total_cost_formatted = f"{total_cost:,.0f}".replace(",", ".")
    terbilang = num2words(int(total_cost), lang='id').title()

    template_context = {
        "logo_path": 'file:///' + os.path.abspath(os.path.join('static', 'ALFALOGO.png')),
        "spk_location": spk_data.get('Cabang'),
        "spk_date": datetime.now().strftime('%d %B %Y'),
        "spk_number": spk_data.get('Nomor SPK', '____/PROPNDEV-____/____/____'),
        "par_number": spk_data.get('PAR', '____/PROPNDEV-____-____-____'),
        "contractor_name": spk_data.get('Nama Kontraktor'),
        "lingkup_pekerjaan": spk_data.get('Lingkup Pekerjaan'),
        "proyek": spk_data.get('Proyek'),
        "project_address": spk_data.get('Alamat'),
        "nama_toko": spk_data.get('Nama_Toko'),
        "kode_toko": spk_data.get('Kode Toko', 'N/A'),
        "total_cost_formatted": total_cost_formatted,
        "terbilang": terbilang,
        "start_date": start_date_formatted,
        "end_date": end_date_formatted,
        "duration": duration,
        "initiator_details_html": initiator_details_html,
        "approver_details_html": approver_details_html,
        "initiator_role_title": initiator_role_title
    }
    
    html_string = render_template('spk_template.html', **template_context)
    return HTML(string=html_string).write_pdf()