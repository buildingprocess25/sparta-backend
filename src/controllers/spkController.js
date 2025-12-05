// src/controllers/spkController.js
const { googleService } = require('../services/googleService');
const { pdfService } = require('../services/pdfService');
const config = require('../config/config');
const moment = require('moment');
const terbilang = require('angka-terbilang'); // Library konversi angka ke teks

const spkController = {
    // --- 1. Cek Status SPK ---
    // Digunakan saat akan membuat SPK baru untuk mencegah duplikasi
    getSpkStatus: async (req, res) => {
        try {
            const { ulok, lingkup } = req.query;

            if (!ulok || !lingkup) {
                return res.status(400).json({ status: "error", message: "Parameter ulok dan lingkup dibutuhkan." });
            }

            // Cari baris di Sheet SPK
            const spkData = await googleService.findSpkRow(
                config.SPK_DATA_SHEET_NAME,
                ulok,
                lingkup
            );

            // Jika tidak ada, return null (berarti belum ada SPK, boleh buat baru)
            if (!spkData) {
                return res.json(null);
            }

            return res.json({
                Status: spkData.Status,
                RowIndex: spkData.rowIndex,
                Data: spkData.data
            });

        } catch (error) {
            console.error("Get SPK Status Error:", error);
            res.status(500).json({ status: "error", message: error.message });
        }
    },

    // --- 2. Submit SPK Baru / Revisi ---
    submitSpk: async (req, res) => {
        let newRowIndex = null;
        try {
            const data = req.body;
            const isRevision = data.Revisi === "YES";
            const rowIndexForUpdate = data.RowIndex;

            // 1. Setup Timestamp & Status
            const now = moment();
            data[config.COLUMN_NAMES.TIMESTAMP] = now.format(); // ISO 8601
            data[config.COLUMN_NAMES.STATUS] = config.STATUS.WAITING_FOR_BM_APPROVAL;

            // 2. Hitung Tanggal Selesai
            const startDate = moment(data['Waktu Mulai']);
            const duration = parseInt(data['Durasi']);
            const endDate = startDate.clone().add(duration - 1, 'days');
            data['Waktu Selesai'] = endDate.format('YYYY-MM-DD');

            // 3. Konversi Biaya ke Terbilang
            const totalCost = parseFloat(data['Grand Total']) || 0;
            const terbilangText = terbilang(totalCost);
            // Format huruf kapital di awal kata (Title Case)
            const terbilangTitleCase = terbilangText.replace(/\w\S*/g, (txt) => txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase());

            data['Biaya'] = totalCost;
            data['Terbilang'] = `( ${terbilangTitleCase} Rupiah )`;

            // 4. Generate Nomor SPK (Hanya jika bukan revisi)
            // Format: XXX/PROPNDEV-CODE/ROM/YEAR
            if (!isRevision) {
                const spkSequence = await googleService.getNextSpkSequence(data.Cabang, now.year(), now.month() + 1);
                // Padding nomor urut menjadi 3 digit (contoh: 001)
                const sequenceStr = String(spkSequence).padStart(3, '0');

                // Ambil kode cabang dari peta (nanti ada di helper/service)
                const cabangCode = await googleService.getCabangCode(data.Cabang);

                data['Nomor SPK'] = `${sequenceStr}/PROPNDEV-${cabangCode}/${data.spk_manual_1}/${data.spk_manual_2}`;
            }
            // Jika revisi, gunakan Nomor SPK yang dikirim dari frontend (tidak berubah)

            // 5. Generate PDF SPK
            const pdfBuffer = await pdfService.generateSpkPDF(data);
            const pdfFilename = `SPK_${data.Proyek}_${data['Nomor Ulok']}.pdf`;

            // 6. Upload PDF ke Drive
            const linkPdf = await googleService.uploadFile(pdfBuffer, pdfFilename, 'application/pdf');
            data[config.COLUMN_NAMES.LINK_PDF] = linkPdf;

            // 7. Simpan ke Spreadsheet (Append atau Update)
            if (isRevision && rowIndexForUpdate) {
                // Update baris lama
                newRowIndex = rowIndexForUpdate;
                await googleService.updateRow(config.SPREADSHEET_ID, config.SPK_DATA_SHEET_NAME, newRowIndex, data);
            } else {
                // Tambah baris baru
                newRowIndex = await googleService.appendRow(config.SPREADSHEET_ID, config.SPK_DATA_SHEET_NAME, data);
            }

            // 8. Kirim Email ke Branch Manager
            const namaToko = data.Nama_Toko || data.nama_toko || 'N/A';
            const jenisToko = data.Jenis_Toko || data.Proyek || 'N/A';

            await googleService.sendApprovalEmail(
                data.Cabang,
                config.JABATAN.BRANCH_MANAGER,
                {
                    type: 'SPK',
                    data: data,
                    links: { pdf: linkPdf },
                    rowIndex: newRowIndex,
                    customSubject: `[PERLU PERSETUJUAN BM] SPK Proyek ${namaToko}: ${jenisToko}`
                }
            );

            res.json({
                status: "success",
                message: "SPK berhasil dikirim untuk persetujuan.",
                data: { linkPdf }
            });

        } catch (error) {
            console.error("Submit SPK Error:", error);
            // Rollback jika insert baru gagal (opsional)
            if (newRowIndex && !req.body.RowIndex) {
                // await googleService.deleteRow(...) 
            }
            res.status(500).json({ status: "error", message: error.message });
        }
    },

    // --- 3. Handle Approval SPK (Approve/Reject) ---
    handleApproval: async (req, res) => {
        try {
            const { action, row, approver, reason } = req.query;

            if (!action || !row || !approver) {
                return res.status(400).send("Parameter tidak lengkap.");
            }

            const sheetName = config.SPK_DATA_SHEET_NAME;
            const ssId = config.SPREADSHEET_ID;

            // Ambil data terkini dari baris tersebut
            const rowData = await googleService.getRowData(ssId, sheetName, row);

            if (!rowData) {
                return res.status(404).send("Data SPK tidak ditemukan.");
            }

            // Cek Status saat ini
            if (rowData.Status !== config.STATUS.WAITING_FOR_BM_APPROVAL) {
                return res.send(`Tindakan ini sudah diproses. Status saat ini: ${rowData.Status}`);
            }

            const now = moment().format();

            if (action === 'approve') {
                // 1. Update Status & Info Approver
                const updates = {
                    [config.COLUMN_NAMES.STATUS]: config.STATUS.SPK_APPROVED,
                    'Disetujui Oleh': approver,
                    'Waktu Persetujuan': now
                };

                // Update object lokal untuk PDF generation
                const updatedData = { ...rowData, ...updates };

                // 2. Generate PDF Final (dengan tanda tangan/nama approver)
                const finalPdfBuffer = await pdfService.generateSpkPDF(updatedData);
                const finalFilename = `SPK_DISETUJUI_${updatedData.Proyek}_${updatedData['Nomor Ulok']}.pdf`;

                // 3. Upload PDF Final
                const finalLinkPdf = await googleService.uploadFile(finalPdfBuffer, finalFilename, 'application/pdf');

                updates[config.COLUMN_NAMES.LINK_PDF] = finalLinkPdf;

                // 4. Simpan Update ke Sheet
                // Kita update cell satu per satu atau update row parsial (tergantung implementasi googleService)
                // Di sini diasumsikan updateRow menangani merge data
                await googleService.updateRow(ssId, sheetName, row, updates);

                // 5. Kirim Email Notifikasi ke Semua Pihak
                // (Branch Manager, Building Manager, Kontraktor, Support, Pembuat)
                await googleService.sendSpkFinalNotification(updatedData, finalLinkPdf, approver);

                res.send("<h1>SPK Disetujui</h1><p>SPK telah berhasil disetujui dan didistribusikan.</p>");

            } else if (action === 'reject') {
                // 1. Update Status Reject
                const updates = {
                    [config.COLUMN_NAMES.STATUS]: config.STATUS.SPK_REJECTED,
                    'Alasan Penolakan': reason || '-'
                };

                await googleService.updateRow(ssId, sheetName, row, updates);

                // 2. Kirim Notifikasi Penolakan ke Pembuat
                const emailPembuat = rowData['Dibuat Oleh']; // Atau 'Email_Pembuat' tergantung data
                if (emailPembuat) {
                    await googleService.sendRejectionEmail(rowData, reason, approver, 'SPK');
                }

                res.send("<h1>SPK Ditolak</h1><p>Notifikasi penolakan telah dikirim ke pembuat.</p>");
            }

        } catch (error) {
            console.error("SPK Approval Error:", error);
            res.status(500).send(`Error processing approval: ${error.message}`);
        }
    }
};

module.exports = spkController;