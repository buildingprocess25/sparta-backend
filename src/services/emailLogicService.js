// src/services/emailLogicService.js
const config = require('../config/config');

// --- Definisi Alur Laporan (Flows) ---
// Berdasarkan file HTML yang ada di folder Client/Pengawasan
const REPORT_FLOWS = {
    // Format: 'Kategori': ['urutan_form_1', 'urutan_form_2', ...]

    // Renovasi Cepat (10 - 14 Hari)
    '10_HARI': ['input_pic', 'h2', 'h5', 'h8', 'h10', 'serah_terima'],
    '14_HARI': ['input_pic', 'h2', 'h7', 'h10', 'h14', 'serah_terima'],

    // Renovasi Menengah (20 - 30 Hari)
    '20_HARI': ['input_pic', 'h2', 'h12', 'h16', 'serah_terima'], // Adjusted based on file availability
    '30_HARI': ['input_pic', 'h2', 'h7', 'h14', 'h18', 'h23', 'serah_terima'],

    // Renovasi Standar (35 - 40 Hari)
    '35_HARI': ['input_pic', 'h2', 'h7', 'h17', 'h22', 'h28', 'serah_terima'],
    '40_HARI': ['input_pic', 'h2', 'h7', 'h17', 'h25', 'h33', 'serah_terima'],

    // Renovasi Berat (48+ Hari / Sipil Berat)
    '48_HARI': ['input_pic', 'h2', 'h10', 'h25', 'h32', 'h41', 'serah_terima'], // Adjusted

    // Default fallback
    'DEFAULT': ['input_pic', 'h2', 'h7', 'h14', 'h21', 'serah_terima']
};

const emailLogicService = {

    /**
     * Menentukan URL/Path untuk formulir selanjutnya berdasarkan progres saat ini.
     * @param {string} currentForm - Tipe form saat ini (misal: 'h2', 'h7')
     * @param {string} kategoriLokasi - Kategori proyek (misal: 'RENOV_10_HARI')
     * @returns {string} - Path URL untuk redirect, atau '#' jika selesai.
     */
    getNextFormPath: (currentForm, kategoriLokasi) => {
        if (!kategoriLokasi) return '#';

        // Normalisasi kategori (hapus spasi, uppercase) agar cocok dengan key REPORT_FLOWS
        // Contoh input: "Renov 10 Hari" -> "10_HARI"
        let cleanKey = 'DEFAULT';
        const upperCat = kategoriLokasi.toUpperCase();

        if (upperCat.includes('10')) cleanKey = '10_HARI';
        else if (upperCat.includes('14')) cleanKey = '14_HARI';
        else if (upperCat.includes('20')) cleanKey = '20_HARI';
        else if (upperCat.includes('30')) cleanKey = '30_HARI';
        else if (upperCat.includes('35')) cleanKey = '35_HARI';
        else if (upperCat.includes('40')) cleanKey = '40_HARI';
        else if (upperCat.includes('48')) cleanKey = '48_HARI';

        const flow = REPORT_FLOWS[cleanKey];
        const currentIndex = flow.indexOf(currentForm);

        // Jika form ditemukan dan bukan yang terakhir
        if (currentIndex !== -1 && currentIndex < flow.length - 1) {
            const nextStep = flow[currentIndex + 1];

            // Logic khusus untuk mapping nama step ke nama file HTML
            // File HTML format: h{hari}_{total_hari}hr.html (contoh: h2_10hr.html)
            // Kecuali serah_terima.html

            if (nextStep === 'serah_terima') {
                return 'Pengawasan/serah_terima.html';
            }

            // Generate nama file dinamis
            // Ambil angka durasi dari cleanKey (misal 10_HARI -> 10)
            const duration = cleanKey.split('_')[0];
            return `Pengawasan/${nextStep}_${duration}hr.html`;
        }

        return '#'; // Tidak ada langkah selanjutnya (Selesai)
    },

    /**
     * Menentukan Subjek Email dan Daftar Penerima
     */
    getEmailDetails: (formType, data, userInfo) => {
        let subject = "";
        let recipients = [];

        // 1. Tentukan Penerima (Recipients)
        // Default: PIC (Pembuat) + Koordinator
        const emailPic = data.pic_building_support || data.Email_Pembuat; // Handle field name differences
        const emailKoordinator = userInfo.koordinator_info?.email;
        const emailManager = userInfo.manager_info?.email;

        if (emailPic) recipients.push(emailPic);
        if (emailKoordinator) recipients.push(emailKoordinator);

        // Khusus Serah Terima: Tambahkan Manager
        if (formType === 'serah_terima') {
            if (emailManager) recipients.push(emailManager);
        }

        // Hapus duplikat dan null/undefined
        recipients = [...new Set(recipients)].filter(e => e);

        // 2. Tentukan Subjek Email
        const namaToko = data.nama_toko || data.Nama_Toko || 'Toko Tanpa Nama';
        const ulok = data.kode_ulok || data.Nomor_Ulok || 'No Ulok';

        // Format Nama Laporan
        let reportName = formType.toUpperCase();
        if (formType.startsWith('h')) {
            reportName = `Laporan Progres ${formType.toUpperCase()}`;
        } else if (formType === 'input_pic') {
            reportName = "Inisiasi Pengawasan (Input PIC)";
        } else if (formType === 'serah_terima') {
            reportName = "Berita Acara Serah Terima (BAST)";
        }

        subject = `[${reportName}] ${namaToko} (${ulok})`;

        return {
            subject,
            recipients
        };
    }
};

module.exports = { emailLogicService };