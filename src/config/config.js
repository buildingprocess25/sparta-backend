// src/config/config.js
require('dotenv').config();

const config = {
    // --- Google & Spreadsheet Configuration ---
    // Menggunakan process.env agar fleksibel di Vercel, dengan fallback ke nilai default dari file lama
    SPREADSHEET_ID: process.env.SPREADSHEET_ID || "1LA1TlhgltT2bqSN3H-LYasq9PtInVlqq98VPru8txoo",
    SPREADSHEET_ID_RAB_2: process.env.SPREADSHEET_ID_RAB_2 || "1P_3cGMS1jOAA6nc398idLHHIhY3r4XANlSMzDdlE6cc",
    PDF_STORAGE_FOLDER_ID: process.env.PDF_STORAGE_FOLDER_ID || "1lvPxOwNILXHmagVfPGkVlNEtfv3U4Emj",
    KONTRAKTOR_SHEET_ID: process.env.KONTRAKTOR_SHEET_ID || "1s95mAc0yXEyDwUDyyOzsDdIqIPEETZkA62_jQQBWXyw",
    PENGAWASAN_SPREADSHEET_ID: process.env.PENGAWASAN_SPREADSHEET_ID || "1zy6BBKJwwmSSvFrMZSZG39pf0YgmjXockZNf10_OFLo",
    INPUT_PIC_DRIVE_FOLDER_ID: process.env.INPUT_PIC_DRIVE_FOLDER_ID || "1gkGZhOJYVo7zv7sZnUIOYAafL-NbNHf8",

    // --- Nama-nama Sheet ---
    DATA_ENTRY_SHEET_NAME: "Form2",
    DATA_ENTRY_SHEET_NAME_RAB_2: "Form2",
    APPROVED_DATA_SHEET_NAME: "Form3",
    APPROVED_DATA_SHEET_NAME_RAB_2: "Form3",
    CABANG_SHEET_NAME: "Cabang",
    SPK_DATA_SHEET_NAME: "SPK_Data",
    KONTRAKTOR_SHEET_NAME: "Monitoring Kontraktor",

    // Nama sheet untuk Pengawasan
    INPUT_PIC_SHEET_NAME: "InputPIC",
    PENUGASAN_SHEET_NAME: "Penugasan",

    // --- Nama Kolom (Konstanta) ---
    COLUMN_NAMES: {
        STATUS: "Status",
        TIMESTAMP: "Timestamp",
        EMAIL_PEMBUAT: "Email_Pembuat",
        LOKASI: "Nomor Ulok",
        PROYEK: "Proyek",
        CABANG: "Cabang",
        LINGKUP_PEKERJAAN: "Lingkup_Pekerjaan",
        KOORDINATOR_APPROVER: "Pemberi Persetujuan Koordinator",
        KOORDINATOR_APPROVAL_TIME: "Waktu Persetujuan Koordinator",
        MANAGER_APPROVER: "Pemberi Persetujuan Manager",
        MANAGER_APPROVAL_TIME: "Waktu Persetujuan Manager",
        LINK_PDF: "Link PDF",
        LINK_PDF_NONSBO: "Link PDF Non-SBO",
        LINK_PDF_REKAP: "Link PDF Rekapitulasi",
        GRAND_TOTAL: "Grand Total",
        GRAND_TOTAL_NONSBO: "Grand Total Non-SBO",
        ALAMAT: "Alamat",
        ALASAN_PENOLAKAN_RAB: "Alasan Penolakan",
        ALASAN_PENOLAKAN_SPK: "Alasan Penolakan",
        GRAND_TOTAL_FINAL: "Grand Total Final",
        LINK_PDF_IL: "Link PDF IL"
    },

    // --- Jabatan ---
    JABATAN: {
        SUPPORT: "BRANCH BUILDING SUPPORT",
        KOORDINATOR: "BRANCH BUILDING COORDINATOR",
        MANAGER: "BRANCH BUILDING & MAINTENANCE MANAGER",
        BRANCH_MANAGER: "BRANCH MANAGER",
        KONTRAKTOR: "KONTRAKTOR"
    },

    // --- Status Persetujuan ---
    STATUS: {
        // Status RAB
        WAITING_FOR_COORDINATOR: "Menunggu Persetujuan Koordinator",
        REJECTED_BY_COORDINATOR: "Ditolak oleh Koordinator",
        WAITING_FOR_MANAGER: "Menunggu Persetujuan Manajer",
        REJECTED_BY_MANAGER: "Ditolak oleh Manajer",
        APPROVED: "Disetujui",
        // Status SPK
        WAITING_FOR_BM_APPROVAL: "Menunggu Persetujuan Branch Manager",
        SPK_APPROVED: "SPK Disetujui",
        SPK_REJECTED: "SPK Ditolak"
    }
};

module.exports = config;