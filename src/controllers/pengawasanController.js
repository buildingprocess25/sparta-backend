// src/controllers/pengawasanController.js
const { googleService } = require('../services/googleService');
// Kita akan buat emailLogicService.js di langkah berikutnya untuk logika routing form
const { emailLogicService } = require('../services/emailLogicService');
const config = require('../config/config');
const moment = require('moment');

// Helper: Hitung Tanggal Kerja (skip Sabtu-Minggu)
// Menggantikan fungsi get_tanggal_h di app.py
const getTanggalH = (startDateStr, days) => {
    // days = jumlah hari kerja yang ingin ditambahkan
    let date = moment(startDateStr);
    let count = 0;

    // Loop sampai jumlah hari kerja tercapai
    while (count < days) {
        date.add(1, 'days');
        // isoWeekday: 1=Senin ... 6=Sabtu, 7=Minggu
        // Kita anggap hari kerja Senin-Jumat (<= 5)
        if (date.isoWeekday() <= 5) {
            count++;
        }
    }
    return date;
};

const pengawasanController = {
    // --- 1. Init Data (Load PIC & SPK List untuk Form Input) ---
    getInitData: async (req, res) => {
        try {
            const { cabang } = req.query;
            if (!cabang) {
                return res.status(400).json({ status: "error", message: "Parameter cabang dibutuhkan." });
            }

            // Ambil Info User (PIC, Koordinator, Manager) & Data SPK dari Spreadsheet
            const userInfo = await googleService.getUserInfoByCabang(cabang);
            const spkList = await googleService.getSpkDataByCabang(cabang);

            res.json({
                status: "success",
                picList: userInfo.picList,
                spkList: spkList
            });
        } catch (error) {
            console.error("Init Data Error:", error);
            res.status(500).json({ status: "error", message: error.message });
        }
    },

    // --- 2. Get RAB URL ---
    getRabUrl: async (req, res) => {
        try {
            const { kode_ulok } = req.query;
            if (!kode_ulok) return res.status(400).json({ status: "error", message: "Parameter kode_ulok dibutuhkan." });

            const rabUrl = await googleService.getRabUrlByUlok(kode_ulok);
            if (rabUrl) {
                res.json({ status: "success", rabUrl });
            } else {
                res.status(404).json({ status: "error", message: "URL RAB tidak ditemukan." });
            }
        } catch (error) {
            res.status(500).json({ status: "error", message: error.message });
        }
    },

    // --- 3. Get SPK URL ---
    getSpkUrl: async (req, res) => {
        try {
            const { kode_ulok } = req.query;
            if (!kode_ulok) return res.status(400).json({ status: "error", message: "Parameter kode_ulok dibutuhkan." });

            const spkUrl = await googleService.getSpkUrlByUlok(kode_ulok);
            if (spkUrl) {
                res.json({ status: "success", spkUrl });
            } else {
                res.status(404).json({ status: "error", message: "URL SPK tidak ditemukan." });
            }
        } catch (error) {
            res.status(500).json({ status: "error", message: error.message });
        }
    },

    // --- 4. Get Active Projects (Untuk Dropdown Form Harian) ---
    getActiveProjects: async (req, res) => {
        try {
            const { email } = req.query;
            if (!email) return res.status(400).json({ status: "error", message: "Parameter email dibutuhkan." });

            const projects = await googleService.getActivePengawasanByPic(email);
            res.json({ status: "success", projects });
        } catch (error) {
            res.status(500).json({ status: "error", message: error.message });
        }
    },

    // --- 5. Submit Laporan Pengawasan (INTI LOGIC) ---
    submitPengawasan: async (req, res) => {
        try {
            const data = req.body; // Data dari Form
            const formType = data.form_type; // 'input_pic', 'h2', 'h7', dll
            const timestamp = moment().format(); // ISO String

            const cabang = data.cabang || 'N/A';

            // -- Logic Cari Email PIC (jika bukan form input awal) --
            if (formType !== 'input_pic') {
                const kodeUlok = data.kode_ulok;
                if (kodeUlok) {
                    const picEmail = await googleService.getPicEmailByUlok(kodeUlok);
                    if (picEmail) {
                        data.pic_building_support = picEmail;
                    } else {
                        return res.status(404).json({
                            status: "error",
                            message: `PIC tidak ditemukan untuk Kode Ulok ${kodeUlok}. Pastikan proyek ini sudah diinisiasi.`
                        });
                    }
                }
            }

            // -- Siapkan Info User untuk Email --
            const userInfo = await googleService.getUserInfoByCabang(cabang);

            // -- LOGIKA INPUT PIC BARU --
            if (formType === 'input_pic') {
                // 1. Simpan ke Sheet 'InputPIC'
                const inputPicData = {
                    'Timestamp': timestamp,
                    'Cabang': data.cabang,
                    'Kode_Ulok': data.kode_ulok,
                    'Kategori_Lokasi': data.kategori_lokasi,
                    'Tanggal_Mulai_SPK': data.tanggal_spk,
                    'PIC_Building_Support': data.pic_building_support,
                    'SPK_URL': data.spkUrl,
                    'RAB_URL': data.rabUrl
                };
                await googleService.appendRow(
                    config.PENGAWASAN_SPREADSHEET_ID,
                    config.INPUT_PIC_SHEET_NAME,
                    inputPicData
                );

                // 2. Simpan ke Sheet 'Penugasan' (Agar muncul di dropdown form harian)
                const penugasanData = {
                    'Email_BBS': data.pic_building_support,
                    'Kode_Ulok': data.kode_ulok,
                    'Cabang': data.cabang
                };
                await googleService.appendRow(
                    config.PENGAWASAN_SPREADSHEET_ID,
                    config.PENUGASAN_SHEET_NAME,
                    penugasanData
                );

                // 3. Hitung Tanggal Mengawas (H+2 Hari Kerja dari SPK)
                const tanggalSpk = data.tanggal_spk;
                const tanggalMengawas = getTanggalH(tanggalSpk, 2);
                data.tanggal_mengawas = tanggalMengawas.format('D MMMM YYYY');
                data.raw_tanggal_mengawas = tanggalMengawas.format('YYYY-MM-DD');

            } else {
                // -- LOGIKA LAPORAN HARIAN (H2, H5, dst) --

                // Mapping field frontend ke header spreadsheet
                const headerMapping = {
                    "timestamp": "Timestamp", "kode_ulok": "Kode_Ulok", "status_lokasi": "Status_Lokasi",
                    "status_progress1": "Status_Progress1", "catatan1": "Catatan1",
                    "status_progress2": "Status_Progress2", "catatan2": "Catatan2",
                    "status_progress3": "Status_Progress3", "catatan3": "Catatan3",
                    "pengukuran_bowplank": "Pengukuran_Dan_Pemasangan_Bowplank",
                    "pekerjaan_tanah": "Pekerjaan_Tanah",
                    "berkas_pengawasan": "Berkas_Pengawasan"
                };

                const dataToSheet = {};

                // Khusus Serah Terima, pakai raw data
                if (formType === 'serah_terima') {
                    Object.assign(dataToSheet, data);
                } else {
                    // Mapping manual untuk form harian
                    for (const [key, value] of Object.entries(data)) {
                        // Jika ada di mapping gunakan, jika tidak ubah snake_case jadi Title_Case
                        const sheetHeader = headerMapping[key] || key.replace(/_/g, ' ').replace(/\w\S*/g, (txt) => txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase()).replace(/ /g, '_');
                        dataToSheet[sheetHeader] = value;
                    }
                }
                dataToSheet['Timestamp'] = timestamp;

                // Tentukan Nama Sheet Tujuan (misal: DataH2, DataH5)
                const sheetMap = {
                    'h2': 'DataH2', 'h5': 'DataH5', 'h7': 'DataH7', 'h8': 'DataH8', 'h10': 'DataH10',
                    'h12': 'DataH12', 'h14': 'DataH14', 'h16': 'DataH16', 'h17': 'DataH17',
                    'h18': 'DataH18', 'h22': 'DataH22', 'h23': 'DataH23', 'h25': 'DataH25',
                    'h28': 'DataH28', 'h32': 'DataH32', 'h33': 'DataH33', 'h41': 'DataH41',
                    'serah_terima': 'SerahTerima'
                };

                const targetSheetName = sheetMap[formType];
                if (targetSheetName) {
                    await googleService.appendRow(
                        config.PENGAWASAN_SPREADSHEET_ID,
                        targetSheetName,
                        dataToSheet
                    );
                }
            }

            // -- KIRIM EMAIL NOTIFIKASI --
            // Logic penerima & subjek ada di service terpisah
            const emailDetails = emailLogicService.getEmailDetails(formType, data, userInfo);

            if (!emailDetails.recipients || emailDetails.recipients.length === 0) {
                return res.status(400).json({
                    status: "error",
                    message: "Tidak ada penerima email valid. Cek data PIC/Koordinator."
                });
            }

            // Dapatkan URL form selanjutnya (Next Step)
            const nextFormPath = emailLogicService.getNextFormPath(formType, data.kategori_lokasi);
            // URL dasar frontend (Vercel)
            const baseUrl = "https://instruksi-lapangan.vercel.app";
            const nextFormUrl = nextFormPath !== '#' ? `${baseUrl}/?redirectTo=${nextFormPath}` : null;

            // Render Template Email (EJS)
            await googleService.sendPengawasanEmail(
                emailDetails.recipients,
                emailDetails.subject,
                {
                    form_type: formType,
                    form_data: data,
                    next_form_url: nextFormUrl
                }
            );

            // -- BUAT EVENT KALENDER (Khusus Input PIC) --
            if (formType === 'input_pic' && data.raw_tanggal_mengawas) {
                await googleService.createCalendarEvent({
                    title: `[REMINDER] Pengawasan H+2: ${data.kode_ulok}`,
                    description: `Pengingat pengawasan H+2 untuk toko ${data.kode_ulok}.`,
                    date: data.raw_tanggal_mengawas, // YYYY-MM-DD
                    guests: emailDetails.recipients
                });
            }

            res.json({ status: "success", message: "Laporan berhasil dikirim." });

        } catch (error) {
            console.error("Submit Pengawasan Error:", error);
            res.status(500).json({ status: "error", message: error.message });
        }
    }
};

module.exports = pengawasanController;