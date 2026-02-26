import os
from dotenv import load_dotenv

load_dotenv()

# --- Google & Spreadsheet Configuration ----

# Data form sheet
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID", "1LA1TlhgltT2bqSN3H-LYasq9PtInVlqq98VPru8txoo")

# Data form sheet untuk RAB 2 / Instruksi Lapangan
SPREADSHEET_ID_RAB_2 = "1P_3cGMS1jOAA6nc398idLHHIhY3r4XANlSMzDdlE6cc"

# ID Folder Google Drive utama untuk menyimpan file PDF hasil generate dari form
PDF_STORAGE_FOLDER_ID = "1lvPxOwNILXHmagVfPGkVlNEtfv3U4Emj"

# ID Spreadsheet untuk data kontraktor
KONTRAKTOR_SHEET_ID = "1s95mAc0yXEyDwUDyyOzsDdIqIPEETZkA62_jQQBWXyw"

# ID Spreadsheet untuk semua data Pengawasan
PENGAWASAN_SPREADSHEET_ID = "1CcJQSk7l-A08R7vlOngG5Sh_k8XYZIr98xHSS-Ahcug"

# ID Folder Google Drive untuk upload file SPK dari form pengawasan
INPUT_PIC_DRIVE_FOLDER_ID = "1gkGZhOJYVo7zv7sZnUIOYAafL-NbNHf8"

# ID Spreadsheet untuk Opname
OPNAME_SHEET_ID = "1ssXGBJ-D4O8JVB1emOBuqcdmBKvdymciVsv7eMaBb64"

# ID Spreadsheet untuk Penyimpanan Dokumen
DOC_DRIVE_ROOT_ID = os.getenv("DRIVE_ROOT_ID", "14hjuP33ez1v1WDxkTi7A3k-XfKOZKVTc") # ID Folder Root Drive Penyimpanan
DOC_SHEET_NAME = os.getenv("SHEET_NAME", "penyimpanan_dokumen")

# awal dokumentasi bangunan
DOC_SHEET_ID = os.getenv("DOC_SHEET_ID", "1LA1TlhgltT2bqSN3H-LYasq9PtInVlqq98VPru8txoo")
DOC_FOLDER_ID = os.getenv("DOC_FOLDER_ID", "1ZTHC7vvvKOIejATqAYeluxfVDaGax1cd")
DOC_DEFAULT_PHOTO_ID = os.getenv("DOC_DEFAULT_PHOTO_ID", "14x-tF0TDAZA9t4lbk6GrHXL8RccxxnjQ")
DOC_GOOGLE_CLIENT_ID = os.getenv("DOC_GOOGLE_CLIENT_ID")
DOC_GOOGLE_CLIENT_SECRET = os.getenv("DOC_GOOGLE_CLIENT_SECRET")
DOC_GOOGLE_REFRESH_TOKEN = os.getenv("DOC_GOOGLE_REFRESH_TOKEN")
DOC_SERVICE_ACCOUNT_FILE = "service_account_doc.json"
DOC_SHEET_NAME_TEMP = "dokumentasi_temp"
DOC_SHEET_NAME_FINAL = "dokumentasi_bangunan"
DOC_SHEET_NAME_LOG = "login_dokumentasi"
# Akhir dokumentasi bangunan

# Nama-nama sheet

# Rab
DATA_ENTRY_SHEET_NAME = "Form2"
DATA_ENTRY_SHEET_NAME_RAB_2 = "Form2"
APPROVED_DATA_SHEET_NAME = "Form3"
APPROVED_DATA_SHEET_NAME_RAB_2 = "Form3"
# Akhir Rab

# Cabang untuk login dan data cabang
CABANG_SHEET_NAME = "Cabang"

# SPK
SPK_DATA_SHEET_NAME = "SPK_Data"

# Nama sheet untuk kontraktor
KONTRAKTOR_SHEET_NAME = "Monitoring Kontraktor"

# Nama sheet untuk Pengawasan
INPUT_PIC_SHEET_NAME = "InputPIC"
PENUGASAN_SHEET_NAME = "Penugasan"

# Nama sheet untuk Gantt Chart
GANTT_CHART_SHEET_NAME = "gantt_chart"
DAY_GANTT_CHART_SHEET_NAME = "day_gantt_chart"
DEPENDENCY_GANTT_SHEET_NAME = "dependency_gantt"

# Nama sheet untuk Opname
SUMMARY_OPNAME_SHEET_NAME = "summary_opname"
OPNAME_SHEET_NAME = "opname_final"
SUMMARY_DATA_SHEET_NAME = "summary"
SUMMARY_SERAH_TERIMA_SHEET_NAME = "summary_serahterima"
# Akhir sheet untuk Opname

# --- Nama Kolom ---
class COLUMN_NAMES:
    STATUS = "Status"
    TIMESTAMP = "Timestamp"
    EMAIL_PEMBUAT = "Email_Pembuat"
    LOKASI = "Nomor Ulok"
    PROYEK = "Proyek"
    CABANG = "Cabang"
    NAMA_PT = "Nama_PT"
    LINGKUP_PEKERJAAN = "Lingkup_Pekerjaan"
    DIREKTUR_APPROVER = "Pemberi Persetujuan Direktur"
    DIREKTUR_APPROVAL_TIME = "Waktu Persetujuan Direktur"
    KOORDINATOR_APPROVER = "Pemberi Persetujuan Koordinator"
    KOORDINATOR_APPROVAL_TIME = "Waktu Persetujuan Koordinator"
    MANAGER_APPROVER = "Pemberi Persetujuan Manager"
    MANAGER_APPROVAL_TIME = "Waktu Persetujuan Manager"
    LINK_PDF = "Link PDF"
    LINK_SURAT_PENAWARAN = "Link Surat Penawaran"
    LOGO = "Logo"
    LINK_PDF_NONSBO = "Link PDF Non-SBO"
    LINK_PDF_REKAP = "Link PDF Rekapitulasi"
    GRAND_TOTAL = "Grand Total"
    GRAND_TOTAL_NONSBO = "Grand Total Non-SBO"
    ALAMAT = "Alamat"
    ALASAN_PENOLAKAN_RAB = "Alasan Penolakan"
    ALASAN_PENOLAKAN_SPK = "Alasan Penolakan"
    GRAND_TOTAL_FINAL = "Grand Total Final"
    LINK_PDF_IL = "Link PDF IL"
    NAMA_TOKO = "nama_toko"
    KODE_TOKO = "kode_toko"
    AWAL_SPK = "awal_spk"
    AKHIR_SPK = "akhir_spk"
    TAMBAH_SPK = "tambah_spk"
    TANGGAL_SERAH_TERIMA = "tanggal_serah_terima"
    TANGGAL_OPNAME_FINAL = "tanggal_opname_final"
    HARI_AWAL = "h_awal"
    HARI_AKHIR = "h_akhir"
    DURASI_SPK = "Durasi"
    KATEGORI = "Kategori"
    KATEGORI_TERIKAT = "Kategori_Terikat"
    DURASI_PEKERJAAN = "Durasi_Pekerjaan"
    KATEGORI_LOKASI = "Kategori_Lokasi"


# --- Jabatan & Status ---
class JABATAN:
    SUPPORT = "BRANCH BUILDING SUPPORT"
    KOORDINATOR = "BRANCH BUILDING COORDINATOR"
    MANAGER = "BRANCH BUILDING & MAINTENANCE MANAGER"
    BRANCH_MANAGER = "BRANCH MANAGER"
    KONTRAKTOR = "KONTRAKTOR"
    DIREKTUR = "DIREKTUR"

class STATUS:
    # Status RAB
    WAITING_FOR_DIREKTUR_APPROVAL = "Menunggu Persetujuan Direktur"
    REJECTED_BY_DIREKTUR = "Ditolak oleh Direktur"
    WAITING_FOR_COORDINATOR = "Menunggu Persetujuan Koordinator"
    REJECTED_BY_COORDINATOR = "Ditolak oleh Koordinator"
    WAITING_FOR_MANAGER = "Menunggu Persetujuan Manajer"
    REJECTED_BY_MANAGER = "Ditolak oleh Manajer"
    APPROVED = "Disetujui"
    # Status SPK
    WAITING_FOR_BM_APPROVAL = "Menunggu Persetujuan Branch Manager"
    SPK_APPROVED = "SPK Disetujui"
    SPK_REJECTED = "SPK Ditolak"

# Data untuk kategori pekerjaan
KATEGORI_SIPIL = [
    "PEKERJAAN PERSIAPAN",
    "PEKERJAAN BOBOKAN / BONGKARAN",
    "PEKERJAAN TANAH",
    "PEKERJAAN PONDASI & BETON",
    "PEKERJAAN PASANGAN",
    "PEKERJAAN BESI",
    "PEKERJAAN KERAMIK",
    "PEKERJAAN PLUMBING",
    "PEKERJAAN SANITARY & ACECORIES",
    "PEKERJAAN JANITOR",
    "PEKERJAAN ATAP",
    "PEKERJAAN KUSEN, PINTU & KACA",
    "PEKERJAAN FINISHING",
    "PEKERJAAN BEANSPOT",
    "PEKERJAAN TAMBAHAN",
    "PEKERJAAN SBO",
    "PEKERJAAN AREA TERBUKA"
]

KATEGORI_ME = [
    "INSTALASI",
    "FIXTURE",
    "PEKERJAAN TAMBAHAN",
    "PEKERJAAN SBO"
]