import os
from dotenv import load_dotenv

load_dotenv()
# --- Google & Spreadsheet Configuration ---
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID", "1LA1TlhgltT2bqSN3H-LYasq9PtInVlqq98VPru8txoo")
SPREADSHEET_ID_RAB_2 = "1P_3cGMS1jOAA6nc398idLHHIhY3r4XANlSMzDdlE6cc"

PDF_STORAGE_FOLDER_ID = "1lvPxOwNILXHmagVfPGkVlNEtfv3U4Emj"
KONTRAKTOR_SHEET_ID = "1s95mAc0yXEyDwUDyyOzsDdIqIPEETZkA62_jQQBWXyw"
# ID Spreadsheet untuk semua data Pengawasan
PENGAWASAN_SPREADSHEET_ID = "1CcJQSk7l-A08R7vlOngG5Sh_k8XYZIr98xHSS-Ahcug"
# ID Folder Google Drive untuk upload file SPK dari form pengawasan
INPUT_PIC_DRIVE_FOLDER_ID = "1gkGZhOJYVo7zv7sZnUIOYAafL-NbNHf8"
OPNAME_SHEET_ID = "1ssXGBJ-D4O8JVB1emOBuqcdmBKvdymciVsv7eMaBb64"
DOC_SHEET_NAME = os.getenv("SHEET_NAME", "penyimpanan_dokumen")
DOC_DRIVE_ROOT_ID = os.getenv("DRIVE_ROOT_ID", "14hjuP33ez1v1WDxkTi7A3k-XfKOZKVTc") # ID Folder Root Drive Penyimpanan

# Nama-nama sheet
DATA_ENTRY_SHEET_NAME = "Form2"
DATA_ENTRY_SHEET_NAME_RAB_2 = "Form2"

APPROVED_DATA_SHEET_NAME = "Form3"
APPROVED_DATA_SHEET_NAME_RAB_2 = "Form3"

CABANG_SHEET_NAME = "Cabang"
SPK_DATA_SHEET_NAME = "SPK_Data"
KONTRAKTOR_SHEET_NAME = "Monitoring Kontraktor"

# Nama sheet untuk Pengawasan
INPUT_PIC_SHEET_NAME = "InputPIC"
PENUGASAN_SHEET_NAME = "Penugasan"
GANTT_CHART_SHEET_NAME = "gantt_chart"
DAY_GANTT_CHART_SHEET_NAME = "day_gantt_chart"

SUMMARY_OPNAME_SHEET_NAME = "summary_opname"
OPNAME_SHEET_NAME = "opname_final"
SUMMARY_DATA_SHEET_NAME = "summary"
SUMMARY_SERAH_TERIMA_SHEET_NAME = "summary_serahterima"

# Nama sheet dinamis akan ditangani di kode, contoh: "DataH2", "SerahTerima"

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
    KOORDINATOR_APPROVER = "Pemberi Persetujuan Koordinator"
    KOORDINATOR_APPROVAL_TIME = "Waktu Persetujuan Koordinator"
    MANAGER_APPROVER = "Pemberi Persetujuan Manager"
    MANAGER_APPROVAL_TIME = "Waktu Persetujuan Manager"
    LINK_PDF = "Link PDF"
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


# --- Jabatan & Status ---
class JABATAN:
    SUPPORT = "BRANCH BUILDING SUPPORT"
    KOORDINATOR = "BRANCH BUILDING COORDINATOR"
    MANAGER = "BRANCH BUILDING & MAINTENANCE MANAGER"
    BRANCH_MANAGER = "BRANCH MANAGER"
    KONTRAKTOR = "KONTRAKTOR"

class STATUS:
    # Status RAB
    WAITING_FOR_COORDINATOR = "Menunggu Persetujuan Koordinator"
    REJECTED_BY_COORDINATOR = "Ditolak oleh Koordinator"
    WAITING_FOR_MANAGER = "Menunggu Persetujuan Manajer"
    REJECTED_BY_MANAGER = "Ditolak oleh Manajer"
    APPROVED = "Disetujui"
    # Status SPK
    WAITING_FOR_BM_APPROVAL = "Menunggu Persetujuan Branch Manager"
    SPK_APPROVED = "SPK Disetujui"
    SPK_REJECTED = "SPK Ditolak"

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