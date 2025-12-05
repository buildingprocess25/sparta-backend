// src/utils/formatters.js
const angkaTerbilang = require('angka-terbilang');

const formatters = {
    // Format Angka ke Rupiah (Contoh: 1000000 -> "Rp 1.000.000")
    formatRupiah: (number) => {
        return new Intl.NumberFormat('id-ID', {
            style: 'currency',
            currency: 'IDR',
            minimumFractionDigits: 0,
            maximumFractionDigits: 0
        }).format(number || 0).replace("Rp", "Rp "); // Tambah spasi setelah Rp
    },

    // Format Angka ke Terbilang dengan Title Case
    // Contoh: 1500 -> "Seribu Lima Ratus Rupiah"
    formatTerbilang: (number) => {
        if (!number) return "";

        // Konversi angka ke teks
        const text = angkaTerbilang(number);

        // Ubah menjadi Title Case (Huruf besar di setiap awal kata)
        const titleCase = text.replace(/\w\S*/g, (txt) => {
            return txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase();
        });

        return `${titleCase} Rupiah`;
    }
};

module.exports = formatters;