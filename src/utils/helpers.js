// src/utils/helpers.js
const moment = require('moment');

const helpers = {
    // Hitung Tanggal Kerja (Skip Sabtu & Minggu)
    // Digunakan untuk menentukan target tanggal pengawasan (H+2, H+5, dst)
    getTanggalH: (startDateStr, days) => {
        let date = moment(startDateStr);
        let count = 0;

        // Loop sampai jumlah hari kerja tercapai
        while (count < days) {
            date.add(1, 'days');
            // isoWeekday: 1=Senin ... 5=Jumat, 6=Sabtu, 7=Minggu
            // Kita anggap hari kerja hanya Senin-Jumat
            if (date.isoWeekday() <= 5) {
                count++;
            }
        }
        return date;
    },

    // Parser Angka: Mengubah string "Rp 1.000.000" atau "1,000" menjadi float murni
    parseNumber: (str) => {
        if (!str) return 0;
        // Hapus karakter non-numeric kecuali titik dan minus
        // Sesuaikan regex ini dengan format input frontend Anda
        // Jika input "1.000.000", hapus titiknya dulu
        const cleanStr = String(str).replace(/[^0-9.-]+/g, "");
        return parseFloat(cleanStr) || 0;
    },

    // Helper untuk membuat Delay (berguna saat retry koneksi)
    sleep: (ms) => {
        return new Promise(resolve => setTimeout(resolve, ms));
    }
};

module.exports = helpers;