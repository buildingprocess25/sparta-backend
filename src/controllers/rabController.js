// src/controllers/rabController.js
const { googleService } = require('../services/googleService');
const { pdfService } = require('../services/pdfService');
const config = require('../config/config');
const moment = require('moment'); // Library untuk tanggal (npm install moment)

// Fungsi bantuan untuk menghitung total RAB
const calculateTotals = (data) => {
    let totalNonSbo = 0;
    let totalSemuaItem = 0;

    // Loop 200 item (sesuai form frontend)
    for (let i = 1; i <= 200; i++) {
        const kategori = data[`Kategori_Pekerjaan_${i}`];
        const totalItemStr = data[`Total_Harga_Item_${i}`];

        let totalItem = 0;
        if (totalItemStr) {
            // Bersihkan format string (misal "Rp 1.000") jadi angka
            totalItem = parseFloat(String(totalItemStr).replace(/[^0-9.-]+/g, "")) || 0;
        }

        if (totalItem > 0) {
            totalSemuaItem += totalItem;
            if (kategori && kategori !== 'PEKERJAAN SBO') {
                totalNonSbo += totalItem;
            }
        }
    }

    // Pembulatan ke bawah kelipatan 10.000
    const pembulatan = Math.floor(totalSemuaItem / 10000) * 10000;
    // PPN 11%
    const ppn = pembulatan * 0.11;
    const finalGrandTotal = pembulatan + ppn;

    return { totalNonSbo, totalSemuaItem, finalGrandTotal };
};

const rabController = {
    // --- 1. Submit RAB Reguler ---
    submitRab: async (req, res) => {
        let newRowIndex = null;
        try {
            const data = req.body;
            const nomorUlok = data[config.COLUMN_NAMES.LOKASI];
            const lingkup = data[config.COLUMN_NAMES.LINGKUP_PEKERJAAN];

            // 1. Validasi
            if (!nomorUlok) throw new Error("Nomor Ulok wajib diisi.");

            // 2. Cek Duplikasi (via Google Service)
            const exists = await googleService.checkUlokExists(
                nomorUlok,
                lingkup,
                config.SPREADSHEET_ID,
                config.DATA_ENTRY_SHEET_NAME
            );

            // Cek apakah ini revisi (ditolak sebelumnya)
            const isRevision = await googleService.isRevision(
                nomorUlok,
                data.Email_Pembuat,
                config.SPREADSHEET_ID,
                config.DATA_ENTRY_SHEET_NAME
            );

            if (exists && !isRevision) {
                return res.status(409).json({
                    status: "error",
                    message: `Nomor Ulok ${nomorUlok} (${lingkup}) sudah diajukan dan sedang diproses.`
                });
            }

            // 3. Set Status & Timestamp
            data[config.COLUMN_NAMES.STATUS] = config.STATUS.WAITING_FOR_COORDINATOR;
            data[config.COLUMN_NAMES.TIMESTAMP] = moment().format(); // ISO String

            // 4. Hitung Total
            const totals = calculateTotals(data);
            data[config.COLUMN_NAMES.GRAND_TOTAL_NONSBO] = totals.totalNonSbo;
            data[config.COLUMN_NAMES.GRAND_TOTAL] = totals.totalSemuaItem;
            data[config.COLUMN_NAMES.GRAND_TOTAL_FINAL] = totals.finalGrandTotal;

            // 5. Arsip Item Details ke JSON string (agar sheet tidak penuh kolomnya)
            const itemKeys = ['Kategori_', 'Jenis_', 'Satuan_', 'Volume_', 'Harga_', 'Total_'];
            const itemDetails = {};
            Object.keys(data).forEach(key => {
                if (itemKeys.some(prefix => key.startsWith(prefix))) {
                    itemDetails[key] = data[key];
                }
            });
            data['Item_Details_JSON'] = JSON.stringify(itemDetails);

            // 6. Generate PDF (Memanggil PDF Service)
            // Ini akan mereturn Buffer PDF
            const pdfNonSboBuffer = await pdfService.generateRABPDF(data, { excludeSbo: true });
            const pdfRecapBuffer = await pdfService.generateRecapPDF(data);

            // 7. Upload ke Google Drive
            const filenameBase = `RAB_${data.Proyek}_${nomorUlok.replace(/[^a-zA-Z0-9]/g, '-')}`;

            const linkNonSbo = await googleService.uploadFile(
                pdfNonSboBuffer,
                `NON-SBO_${filenameBase}.pdf`,
                'application/pdf'
            );

            const linkRecap = await googleService.uploadFile(
                pdfRecapBuffer,
                `RECAP_${filenameBase}.pdf`,
                'application/pdf'
            );

            data[config.COLUMN_NAMES.LINK_PDF_NONSBO] = linkNonSbo;
            data[config.COLUMN_NAMES.LINK_PDF_REKAP] = linkRecap;

            // 8. Simpan ke Spreadsheet
            newRowIndex = await googleService.appendRow(
                config.SPREADSHEET_ID,
                config.DATA_ENTRY_SHEET_NAME,
                data
            );

            // 9. Kirim Email Notifikasi
            await googleService.sendApprovalEmail(
                data.Cabang,
                config.JABATAN.KOORDINATOR,
                {
                    type: 'RAB',
                    data: data,
                    links: { nonSbo: linkNonSbo, recap: linkRecap },
                    rowIndex: newRowIndex
                }
            );

            res.json({
                status: "success",
                message: "RAB berhasil dikirim dan menunggu persetujuan.",
                data: { linkNonSbo, linkRecap }
            });

        } catch (error) {
            console.error("Submit RAB Error:", error);
            // Rollback (Hapus baris jika sudah terlanjur input tapi error di email)
            if (newRowIndex) {
                await googleService.deleteRow(config.SPREADSHEET_ID, config.DATA_ENTRY_SHEET_NAME, newRowIndex);
            }
            res.status(500).json({ status: "error", message: error.message });
        }
    },

    // --- 2. Submit RAB Kedua (Instruksi Lapangan / Revisi) ---
    // Endpoint ini menangani upload file manual (req.file dari multer)
    submitRabKedua: async (req, res) => {
        let newRowIndex = null;
        try {
            // Data dari form-data (req.body) dan file (req.file)
            const data = req.body;
            const fileManual = req.file; // File PDF lampiran tambahan

            // Validasi & Logika sama seperti submitRab (Totals, JSON Archive, PDF Gen)
            // ... (Kode dipersingkat agar fokus pada perbedaannya)

            // Hitung Total (Sama)
            const totals = calculateTotals(data);
            data[config.COLUMN_NAMES.GRAND_TOTAL_FINAL] = totals.finalGrandTotal;

            // Set Status & Timestamp
            data[config.COLUMN_NAMES.STATUS] = config.STATUS.WAITING_FOR_COORDINATOR;
            data[config.COLUMN_NAMES.TIMESTAMP] = moment().format();

            // Handle Upload File Manual (Jika ada)
            let linkPdfManual = "";
            if (fileManual) {
                linkPdfManual = await googleService.uploadFile(
                    fileManual.buffer,
                    `IL_MANUAL_${fileManual.originalname}`,
                    fileManual.mimetype
                );
                data[config.COLUMN_NAMES.LINK_PDF_IL] = linkPdfManual;
            }

            // Generate PDF System
            const pdfNonSboBuffer = await pdfService.generateRABPDF(data, { excludeSbo: true });
            const linkNonSbo = await googleService.uploadFile(pdfNonSboBuffer, `IL_SYS_${data.Nomor_Ulok}.pdf`, 'application/pdf');

            data[config.COLUMN_NAMES.LINK_PDF_NONSBO] = linkNonSbo;

            // Simpan ke Spreadsheet RAB 2 (ID berbeda)
            // Cek apakah ini revisi dari data yang ditolak sebelumnya (Overwrite)
            const existingRow = await googleService.findRejectedRow(
                config.SPREADSHEET_ID_RAB_2,
                config.DATA_ENTRY_SHEET_NAME_RAB_2,
                data[config.COLUMN_NAMES.LOKASI]
            );

            if (existingRow) {
                // Update baris lama
                newRowIndex = existingRow;
                await googleService.updateRow(
                    config.SPREADSHEET_ID_RAB_2,
                    config.DATA_ENTRY_SHEET_NAME_RAB_2,
                    newRowIndex,
                    data
                );
            } else {
                // Buat baris baru
                newRowIndex = await googleService.appendRow(
                    config.SPREADSHEET_ID_RAB_2,
                    config.DATA_ENTRY_SHEET_NAME_RAB_2,
                    data
                );
            }

            // Kirim Email
            await googleService.sendApprovalEmail(
                data.Cabang,
                config.JABATAN.KOORDINATOR,
                {
                    type: 'IL', // Instruksi Lapangan
                    data: data,
                    links: { nonSbo: linkNonSbo, manual: linkPdfManual },
                    rowIndex: newRowIndex,
                    isRab2: true // Flag untuk service email
                }
            );

            res.json({ status: "success", message: "Instruksi Lapangan berhasil dikirim." });

        } catch (error) {
            console.error("Submit RAB 2 Error:", error);
            res.status(500).json({ status: "error", message: error.message });
        }
    },

    // --- 3. Cek Status Pengajuan ---
    checkStatus: async (req, res) => {
        try {
            const { email, cabang, type } = req.query;

            // Tentukan sheet ID berdasarkan tipe (RAB 1 atau RAB 2)
            const spreadSheetId = (type === 'rab2') ? config.SPREADSHEET_ID_RAB_2 : config.SPREADSHEET_ID;
            const sheetName = (type === 'rab2') ? config.DATA_ENTRY_SHEET_NAME_RAB_2 : config.DATA_ENTRY_SHEET_NAME;

            const result = await googleService.checkUserSubmissions(
                email,
                cabang,
                spreadSheetId,
                sheetName
            );

            res.json(result);
        } catch (error) {
            res.status(500).json({ error: error.message });
        }
    },

    // --- 4. Handle Approval (URL yang diakses dari tombol di email) ---
    // Bisa digunakan untuk Approve/Reject baik oleh Koordinator maupun Manager
    handleApproval: async (req, res) => {
        try {
            const { action, row, level, approver, reason, isRab2 } = req.query; // atau req.body

            // Tentukan Config Sheet
            const ssId = isRab2 === 'true' ? config.SPREADSHEET_ID_RAB_2 : config.SPREADSHEET_ID;
            const sheetName = isRab2 === 'true' ? config.DATA_ENTRY_SHEET_NAME_RAB_2 : config.DATA_ENTRY_SHEET_NAME;

            if (action === 'approve') {
                if (level === 'coordinator') {
                    // Update status jadi Waiting for Manager
                    await googleService.updateCell(ssId, sheetName, row, config.COLUMN_NAMES.STATUS, config.STATUS.WAITING_FOR_MANAGER);
                    await googleService.updateCell(ssId, sheetName, row, config.COLUMN_NAMES.KOORDINATOR_APPROVER, approver);

                    // Trigger email ke Manager
                    // Ambil data row lengkap dulu untuk info email
                    const rowData = await googleService.getRowData(ssId, sheetName, row);
                    await googleService.sendApprovalEmail(
                        rowData.Cabang,
                        config.JABATAN.MANAGER,
                        { type: isRab2 ? 'IL' : 'RAB', data: rowData, rowIndex: row, isRab2: isRab2 === 'true' }
                    );

                } else if (level === 'manager') {
                    // Final Approval
                    await googleService.updateCell(ssId, sheetName, row, config.COLUMN_NAMES.STATUS, config.STATUS.APPROVED);
                    await googleService.updateCell(ssId, sheetName, row, config.COLUMN_NAMES.MANAGER_APPROVER, approver);

                    // Pindahkan ke Sheet Approved (Form3)
                    const rowData = await googleService.getRowData(ssId, sheetName, row);
                    const approvedSheetName = isRab2 === 'true' ? config.APPROVED_DATA_SHEET_NAME_RAB_2 : config.APPROVED_DATA_SHEET_NAME;

                    await googleService.appendRow(ssId, approvedSheetName, rowData);

                    // Kirim Email Final ke Semua (Pembuat, Koordinator, Manager)
                    await googleService.sendFinalNotification(rowData, isRab2 === 'true');
                }

                // Render halaman sukses HTML
                res.send("<h1>Persetujuan Berhasil Diproses</h1><p>Terima kasih.</p>");

            } else if (action === 'reject') {
                // Update status jadi Rejected
                const statusReject = level === 'coordinator' ? config.STATUS.REJECTED_BY_COORDINATOR : config.STATUS.REJECTED_BY_MANAGER;
                await googleService.updateCell(ssId, sheetName, row, config.COLUMN_NAMES.STATUS, statusReject);
                await googleService.updateCell(ssId, sheetName, row, 'Alasan Penolakan', reason || '-');

                // Notifikasi ke Pembuat
                const rowData = await googleService.getRowData(ssId, sheetName, row);
                await googleService.sendRejectionEmail(rowData, reason, approver);

                res.send("<h1>Pengajuan Telah Ditolak</h1><p>Notifikasi telah dikirim ke pembuat.</p>");
            }

        } catch (error) {
            console.error("Approval Error:", error);
            res.status(500).send(`Error processing approval: ${error.message}`);
        }
    }
};

module.exports = rabController;