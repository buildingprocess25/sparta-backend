const { GoogleSpreadsheet } = require('google-spreadsheet');
const { google } = require('googleapis');
const { JWT } = require('google-auth-library');
const config = require('../config/config');
const moment = require('moment');
const stream = require('stream');

// --- Setup Autentikasi (Google APIs Generic) ---
// Kita gunakan JWT manual untuk Drive, Gmail, & Calendar
const serviceAccountAuth = new JWT({
    email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
    key: process.env.GOOGLE_PRIVATE_KEY ? process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n') : "",
    scopes: [
        'https://www.googleapis.com/auth/drive',
        'https://www.googleapis.com/auth/gmail.send',
        'https://www.googleapis.com/auth/calendar'
    ]
});

const drive = google.drive({ version: 'v3', auth: serviceAccountAuth });
const gmail = google.gmail({ version: 'v1', auth: serviceAccountAuth });
const calendar = google.calendar({ version: 'v3', auth: serviceAccountAuth });

// --- Helper Functions ---

// Helper: Get Doc (Versi 3.3.0 Style)
const getDoc = async (spreadsheetId) => {
    const doc = new GoogleSpreadsheet(spreadsheetId);
    // V3: Auth dilakukan dengan method useServiceAccountAuth
    await doc.useServiceAccountAuth({
        client_email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
        private_key: process.env.GOOGLE_PRIVATE_KEY ? process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n') : "",
    });
    await doc.loadInfo();
    return doc;
};

// Helper: Membuat Raw Email String (MIME)
const createRawEmail = (to, subject, htmlBody, attachments = []) => {
    const boundary = "foo_bar_baz";
    const toStr = Array.isArray(to) ? to.join(',') : to;

    let email = [
        `MIME-Version: 1.0`,
        `To: ${toStr}`,
        `Subject: ${subject}`,
        `Content-Type: multipart/mixed; boundary="${boundary}"`,
        ``,
        `--${boundary}`,
        `Content-Type: text/html; charset="UTF-8"`,
        `Content-Transfer-Encoding: 7bit`,
        ``,
        htmlBody,
        ``
    ].join('\r\n');

    if (attachments && attachments.length > 0) {
        attachments.forEach(att => {
            const { filename, content, type } = att;
            email += [
                `--${boundary}`,
                `Content-Type: ${type}; name="${filename}"`,
                `Content-Disposition: attachment; filename="${filename}"`,
                `Content-Transfer-Encoding: base64`,
                ``,
                content,
                ``
            ].join('\r\n');
        });
    }

    email += `--${boundary}--`;
    return Buffer.from(email).toString('base64').replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
};

const googleService = {
    // --- 1. User & Auth Logic ---

    validateUser: async (email, cabang) => {
        try {
            const doc = await getDoc(config.SPREADSHEET_ID);
            const sheet = doc.sheetsByTitle[config.CABANG_SHEET_NAME];
            const rows = await sheet.getRows();

            // V3: Akses data kolom langsung sebagai properti object (case sensitive sesuai header di sheet)
            const user = rows.find(row =>
                String(row['EMAIL_SAT'] || "").trim().toLowerCase() === String(email).trim().toLowerCase() &&
                String(row['CABANG'] || "").trim().toLowerCase() === String(cabang).trim().toLowerCase()
            );

            return user ? user['JABATAN'] : null;
        } catch (error) {
            console.error("Validate User Error:", error);
            return null;
        }
    },

    getUserInfoByCabang: async (cabang) => {
        try {
            const doc = await getDoc(config.SPREADSHEET_ID);
            const sheet = doc.sheetsByTitle[config.CABANG_SHEET_NAME];
            const rows = await sheet.getRows();

            const picList = [];
            let koordinatorInfo = {};
            let managerInfo = {};

            rows.forEach(row => {
                const rowCabang = String(row['CABANG'] || "").trim().toLowerCase();
                if (rowCabang === String(cabang).trim().toLowerCase()) {
                    const jabatan = String(row['JABATAN'] || "").toUpperCase();
                    const email = row['EMAIL_SAT'];
                    const nama = row['NAMA LENGKAP'];

                    if (jabatan.includes("SUPPORT")) {
                        picList.push({ email, nama });
                    } else if (jabatan.includes("COORDINATOR")) {
                        koordinatorInfo = { email, nama };
                    } else if (jabatan.includes("MANAGER") && !jabatan.includes("BRANCH MANAGER")) {
                        managerInfo = { email, nama };
                    }
                }
            });

            return { picList, koordinator_info: koordinatorInfo, manager_info: managerInfo };
        } catch (error) {
            console.error("Get User Info Error:", error);
            return { picList: [], koordinator_info: {}, manager_info: {} };
        }
    },

    getEmailsByJabatan: async (cabang, jabatanTarget) => {
        try {
            const doc = await getDoc(config.SPREADSHEET_ID);
            const sheet = doc.sheetsByTitle[config.CABANG_SHEET_NAME];
            const rows = await sheet.getRows();

            const emails = [];
            rows.forEach(row => {
                const rowCabang = String(row['CABANG'] || "").trim().toLowerCase();
                const rowJabatan = String(row['JABATAN'] || "").trim().toUpperCase();

                if (rowCabang === String(cabang).trim().toLowerCase() &&
                    rowJabatan === String(jabatanTarget).trim().toUpperCase()) {
                    if (row['EMAIL_SAT']) emails.push(row['EMAIL_SAT']);
                }
            });
            return emails;
        } catch (error) {
            console.error("Get Emails By Jabatan Error:", error);
            return [];
        }
    },

    getEmailByJabatan: async (cabang, jabatanTarget) => {
        const emails = await googleService.getEmailsByJabatan(cabang, jabatanTarget);
        return emails.length > 0 ? emails[0] : null;
    },

    // --- 2. Spreadsheet Logic (RAB & General) ---

    checkUserSubmissions: async (email, cabang, spreadsheetId, sheetName) => {
        try {
            const doc = await getDoc(spreadsheetId);
            const sheet = doc.sheetsByTitle[sheetName];
            const rows = await sheet.getRows();

            const pending = [];
            const approved = [];
            const rejected = [];
            const processedLocations = new Set();
            const userCabang = String(cabang).trim().toLowerCase();

            // Loop dari bawah (terbaru) ke atas
            for (let i = rows.length - 1; i >= 0; i--) {
                const row = rows[i];
                // V3: Akses menggunakan nama kolom dari config
                const lokasi = String(row[config.COLUMN_NAMES.LOKASI] || "").trim().toUpperCase();

                if (!lokasi || processedLocations.has(lokasi)) continue;

                const status = String(row[config.COLUMN_NAMES.STATUS] || "").trim();
                const recordCabang = String(row[config.COLUMN_NAMES.CABANG] || "").trim().toLowerCase();

                if ([config.STATUS.WAITING_FOR_COORDINATOR, config.STATUS.WAITING_FOR_MANAGER].includes(status)) {
                    pending.push(lokasi);
                } else if (status === config.STATUS.APPROVED) {
                    approved.push(lokasi);
                } else if ([config.STATUS.REJECTED_BY_COORDINATOR, config.STATUS.REJECTED_BY_MANAGER].includes(status)) {
                    if (recordCabang === userCabang) {
                        // Di V3, row object bisa diakses propertinya, tapi untuk JSON field kita perlu parse
                        const rowData = {};
                        // Salin semua properti row ke rowData agar aman
                        sheet.headerValues.forEach(h => {
                            rowData[h] = row[h];
                        });

                        const jsonDetails = row['Item_Details_JSON'];
                        if (jsonDetails) {
                            try {
                                const parsed = JSON.parse(jsonDetails);
                                Object.assign(rowData, parsed);
                            } catch (e) { console.warn("JSON Parse Error"); }
                        }
                        rejected.push(rowData);
                    }
                }
                processedLocations.add(lokasi);
            }

            return {
                active_codes: { pending, approved },
                rejected_submissions: rejected
            };
        } catch (error) {
            throw error;
        }
    },

    checkUlokExists: async (ulok, lingkup, spreadsheetId, sheetName) => {
        const doc = await getDoc(spreadsheetId);
        const sheet = doc.sheetsByTitle[sheetName];
        const rows = await sheet.getRows();

        const normalizedUlok = String(ulok).replace(/-/g, '');

        return rows.some(row => {
            const rowUlok = String(row[config.COLUMN_NAMES.LOKASI] || "").replace(/-/g, '');
            const rowLingkup = String(row[config.COLUMN_NAMES.LINGKUP_PEKERJAAN] || "").trim();
            const status = row[config.COLUMN_NAMES.STATUS];

            return (rowUlok === normalizedUlok &&
                rowLingkup === lingkup &&
                [config.STATUS.WAITING_FOR_COORDINATOR, config.STATUS.WAITING_FOR_MANAGER, config.STATUS.APPROVED].includes(status));
        });
    },

    isRevision: async (ulok, email, spreadsheetId, sheetName) => {
        const doc = await getDoc(spreadsheetId);
        const sheet = doc.sheetsByTitle[sheetName];
        const rows = await sheet.getRows();

        const normalizedUlok = String(ulok).replace(/-/g, '');

        for (let i = rows.length - 1; i >= 0; i--) {
            const row = rows[i];
            const rowUlok = String(row[config.COLUMN_NAMES.LOKASI] || "").replace(/-/g, '');
            const rowEmail = row[config.COLUMN_NAMES.EMAIL_PEMBUAT];

            if (rowUlok === normalizedUlok && rowEmail === email) {
                const status = row[config.COLUMN_NAMES.STATUS];
                return [config.STATUS.REJECTED_BY_COORDINATOR, config.STATUS.REJECTED_BY_MANAGER].includes(status);
            }
        }
        return false;
    },

    findRejectedRow: async (spreadsheetId, sheetName, ulok) => {
        const doc = await getDoc(spreadsheetId);
        const sheet = doc.sheetsByTitle[sheetName];
        const rows = await sheet.getRows();
        const normalizedUlok = String(ulok).replace(/-/g, '').trim().toUpperCase();

        for (let i = 0; i < rows.length; i++) {
            const row = rows[i];
            const rowUlok = String(row[config.COLUMN_NAMES.LOKASI] || "").replace(/-/g, '').trim().toUpperCase();
            const status = row[config.COLUMN_NAMES.STATUS];

            if (rowUlok === normalizedUlok &&
                [config.STATUS.REJECTED_BY_COORDINATOR, config.STATUS.REJECTED_BY_MANAGER].includes(status)) {
                return row.rowIndex; // V3 menggunakan rowIndex (1-based index di Excel)
            }
        }
        return null;
    },

    appendRow: async (spreadsheetId, sheetName, data) => {
        try {
            const doc = await getDoc(spreadsheetId);
            const sheet = doc.sheetsByTitle[sheetName];

            // V3: addRow langsung menerima object { Header: Value }
            const addedRow = await sheet.addRow(data);
            return addedRow.rowIndex;
        } catch (error) {
            console.error("Append Row Error:", error);
            throw error;
        }
    },

    updateRow: async (spreadsheetId, sheetName, rowIndex, data) => {
        try {
            const doc = await getDoc(spreadsheetId);
            const sheet = doc.sheetsByTitle[sheetName];
            const rows = await sheet.getRows();

            // rowIndex dari findRejectedRow adalah 1-based (Excel row number).
            // sheet.getRows() mengembalikan array mulai dari baris data (biasanya baris 2).
            // Jadi row ke-0 di array = baris 2 di Excel.
            // Rumus: arrayIndex = rowIndex - 2
            const rowToUpdate = rows[rowIndex - 2];

            if (rowToUpdate) {
                // V3: Assign value ke properti object row
                Object.keys(data).forEach(key => {
                    rowToUpdate[key] = data[key];
                });
                await rowToUpdate.save(); // Simpan perubahan
                return true;
            }
            return false;
        } catch (error) {
            console.error("Update Row Error:", error);
            throw error;
        }
    },

    updateCell: async (spreadsheetId, sheetName, rowIndex, colName, value) => {
        try {
            const doc = await getDoc(spreadsheetId);
            const sheet = doc.sheetsByTitle[sheetName];
            const rows = await sheet.getRows();
            const row = rows[rowIndex - 2];

            if (row) {
                row[colName] = value; // Assign langsung
                await row.save();
                return true;
            }
            return false;
        } catch (error) {
            console.error("Update Cell Error:", error);
            return false;
        }
    },

    deleteRow: async (spreadsheetId, sheetName, rowIndex) => {
        try {
            const doc = await getDoc(spreadsheetId);
            const sheet = doc.sheetsByTitle[sheetName];
            const rows = await sheet.getRows();
            const row = rows[rowIndex - 2];
            if (row) {
                await row.delete();
            }
        } catch (error) {
            console.error("Delete Row Error:", error);
        }
    },

    getRowData: async (spreadsheetId, sheetName, rowIndex) => {
        try {
            const doc = await getDoc(spreadsheetId);
            const sheet = doc.sheetsByTitle[sheetName];
            const rows = await sheet.getRows();
            const row = rows[rowIndex - 2];

            if (!row) return null;

            // Convert row V3 object to plain object
            const rowData = {};
            sheet.headerValues.forEach(header => {
                rowData[header] = row[header];
            });
            return rowData;
        } catch (error) {
            return null;
        }
    },

    copyToApprovedSheet: async (rowData) => {
        return googleService.appendRow(config.SPREADSHEET_ID, config.APPROVED_DATA_SHEET_NAME, rowData);
    },

    copyToApprovedSheetKedua: async (rowData) => {
        return googleService.appendRow(config.SPREADSHEET_ID_RAB_2, config.APPROVED_DATA_SHEET_NAME_RAB_2, rowData);
    },

    // --- 3. Drive Logic ---

    uploadFile: async (fileBuffer, fileName, mimeType, folderId = config.PDF_STORAGE_FOLDER_ID) => {
        try {
            const bufferStream = new stream.PassThrough();
            bufferStream.end(fileBuffer);

            const fileMetadata = {
                name: fileName,
                parents: [folderId]
            };

            const media = {
                mimeType: mimeType,
                body: bufferStream
            };

            const file = await drive.files.create({
                resource: fileMetadata,
                media: media,
                fields: 'id, webViewLink'
            });

            await drive.permissions.create({
                fileId: file.data.id,
                requestBody: {
                    role: 'reader',
                    type: 'anyone'
                }
            });

            return file.data.webViewLink;
        } catch (error) {
            console.error("Upload Drive Error:", error);
            throw error;
        }
    },

    downloadFileFromLink: async (fileLink) => {
        try {
            if (!fileLink || !fileLink.includes("drive.google.com")) return [null, null, null];

            let fileId = null;
            if (fileLink.includes("/d/")) {
                fileId = fileLink.split("/d/")[1].split("/")[0];
            } else if (fileLink.includes("id=")) {
                fileId = fileLink.split("id=")[1].split("&")[0];
            }

            if (!fileId) return [null, null, null];

            const metadata = await drive.files.get({ fileId, fields: 'name, mimeType' });
            const name = metadata.data.name || "Lampiran.pdf";
            const mimeType = metadata.data.mimeType || "application/pdf";

            const response = await drive.files.get(
                { fileId, alt: 'media' },
                { responseType: 'arraybuffer' }
            );

            return [name, Buffer.from(response.data), mimeType];
        } catch (error) {
            console.error("Download Drive Error:", error);
            return [null, null, null];
        }
    },

    // --- 4. Gmail Logic ---

    sendEmail: async (to, subject, htmlBody, attachments = []) => {
        try {
            const processedAttachments = attachments.map(att => {
                if (Array.isArray(att)) {
                    return {
                        filename: att[0],
                        content: att[1].toString('base64'),
                        type: att[2]
                    };
                }
                return {
                    filename: att.filename,
                    content: att.content.toString('base64'),
                    type: att.type || 'application/pdf'
                };
            });

            const rawMessage = createRawEmail(to, subject, htmlBody, processedAttachments);

            await gmail.users.messages.send({
                userId: 'me',
                requestBody: { raw: rawMessage }
            });
            console.log(`Email sent to ${to}`);
        } catch (error) {
            console.error("Send Email Error:", error);
        }
    },

    // Ini placeholder, logic sesungguhnya ada di controller/pdfService untuk render template
    sendApprovalEmail: async (cabang, jabatanTarget, context) => {
        /* Logic rendering email ada di controller/service lain yang memanggil sendEmail ini */
    },

    // --- 5. SPK & General Helpers ---

    getSpkDataByCabang: async (cabang) => {
        try {
            const doc = await getDoc(config.SPREADSHEET_ID);
            const sheet = doc.sheetsByTitle[config.SPK_DATA_SHEET_NAME];
            const rows = await sheet.getRows();

            const spkList = [];
            rows.forEach(row => {
                if (String(row['Cabang']).toLowerCase() === String(cabang).toLowerCase() &&
                    row['Status'] === config.STATUS.SPK_APPROVED) {
                    spkList.push({
                        "Nomor Ulok": row["Nomor Ulok"],
                        "Lingkup Pekerjaan": row["Lingkup Pekerjaan"],
                        "Link PDF": row["Link PDF"]
                    });
                }
            });
            return spkList.reverse();
        } catch (error) {
            return [];
        }
    },

    getRabUrlByUlok: async (ulok) => {
        try {
            const doc = await getDoc(config.SPREADSHEET_ID);
            const sheet1 = doc.sheetsByTitle[config.APPROVED_DATA_SHEET_NAME];
            const rows1 = await sheet1.getRows();
            let found = rows1.reverse().find(r => r["Nomor Ulok"] === ulok);

            if (found) return found["Link PDF Non-SBO"] || found["Link PDF"];

            const doc2 = await getDoc(config.SPREADSHEET_ID_RAB_2);
            const sheet2 = doc2.sheetsByTitle[config.APPROVED_DATA_SHEET_NAME_RAB_2];
            const rows2 = await sheet2.getRows();
            found = rows2.reverse().find(r => r["Nomor Ulok"] === ulok);

            if (found) return found["Link PDF Non-SBO"] || found["Link PDF"];

            return null;
        } catch (e) { return null; }
    },

    getSpkUrlByUlok: async (ulok) => {
        try {
            const doc = await getDoc(config.SPREADSHEET_ID);
            const sheet = doc.sheetsByTitle[config.SPK_DATA_SHEET_NAME];
            const rows = await sheet.getRows();
            const found = rows.reverse().find(r =>
                r["Nomor Ulok"] === ulok &&
                r['Status'] === config.STATUS.SPK_APPROVED
            );
            return found ? found["Link PDF"] : null;
        } catch (e) { return null; }
    },

    getKontraktorByCabang: async (cabang) => {
        try {
            const doc = await getDoc(config.KONTRAKTOR_SHEET_ID);
            const sheet = doc.sheetsByTitle[config.KONTRAKTOR_SHEET_NAME];
            const rows = await sheet.getRows();

            const kontraktor = new Set();
            rows.forEach(row => {
                if (String(row['NAMA CABANG']).toLowerCase() === String(cabang).toLowerCase() &&
                    String(row['STATUS KONTRAKTOR']).toUpperCase() === 'AKTIF') {
                    kontraktor.add(row['NAMA KONTRAKTOR']);
                }
            });
            return Array.from(kontraktor).sort();
        } catch (e) { return []; }
    },

    getApprovedRabByCabang: async (cabang) => {
        try {
            const doc = await getDoc(config.SPREADSHEET_ID);
            const sheet = doc.sheetsByTitle[config.APPROVED_DATA_SHEET_NAME];
            const rows = await sheet.getRows();

            // Logic filter sederhana
            const relevantRows = rows.filter(r => String(r['Cabang']).toLowerCase().includes(String(cabang).toLowerCase().split(' ')[0]));

            // Convert to object
            return relevantRows.map(r => {
                const obj = {};
                sheet.headerValues.forEach(h => obj[h] = r[h]);
                return obj;
            });
        } catch (e) { return []; }
    },

    getApprovedRabByCabangKedua: async (cabang) => {
        try {
            const doc = await getDoc(config.SPREADSHEET_ID_RAB_2);
            const sheet = doc.sheetsByTitle[config.APPROVED_DATA_SHEET_NAME_RAB_2];
            const rows = await sheet.getRows();

            const relevantRows = rows.filter(r => String(r['Cabang']).toLowerCase().includes(String(cabang).toLowerCase().split(' ')[0]));

            return relevantRows.map(r => {
                const obj = {};
                sheet.headerValues.forEach(h => obj[h] = r[h]);
                return obj;
            });
        } catch (e) { return []; }
    },

    findSpkRow: async (sheetName, ulok, lingkup) => {
        try {
            const doc = await getDoc(config.SPREADSHEET_ID);
            const sheet = doc.sheetsByTitle[sheetName];
            const rows = await sheet.getRows();

            const foundRow = rows.find(r =>
                String(r["Nomor Ulok"]).trim() === String(ulok).trim() &&
                String(r["Lingkup Pekerjaan"]).trim().toLowerCase() === String(lingkup).trim().toLowerCase()
            );

            if (foundRow) {
                const rowData = {};
                sheet.headerValues.forEach(h => rowData[h] = foundRow[h]);

                return {
                    Status: foundRow['Status'],
                    rowIndex: foundRow.rowIndex,
                    data: rowData
                };
            }
            return null;
        } catch (e) { return null; }
    },

    getNextSpkSequence: async (cabang, year, month) => {
        const doc = await getDoc(config.SPREADSHEET_ID);
        const sheet = doc.sheetsByTitle[config.SPK_DATA_SHEET_NAME];
        const rows = await sheet.getRows();

        let count = 0;
        rows.forEach(row => {
            const ts = row['Timestamp'];
            if (ts && String(row['Cabang']).toLowerCase() === String(cabang).toLowerCase()) {
                const date = moment(ts);
                if (date.year() === year && date.month() + 1 === month) {
                    count++;
                }
            }
        });
        return count + 1;
    },

    getCabangCode: (cabangName) => {
        const map = {
            "WHC IMAM BONJOL": "7AZ1", "LUWU": "2VZ1", "KARAWANG": "1JZ1", "REMBANG": "2AZ1",
            "BANJARMASIN": "1GZ1", "PARUNG": "1MZ1", "TEGAL": "2PZ1", "GORONTALO": "2SZ1",
            "PONTIANAK": "1PZ1", "LOMBOK": "1SZ1", "KOTABUMI": "1VZ1", "SERANG": "2GZ1",
            "CIANJUR": "2JZ1", "BALARAJA": "TZ01", "SIDOARJO": "UZ01", "MEDAN": "WZ01",
            "BOGOR": "XZ01", "JEMBER": "YZ01", "BALI": "QZ01", "PALEMBANG": "PZ01",
            "KLATEN": "OZ01", "MAKASSAR": "RZ01", "PLUMBON": "VZ01", "PEKANBARU": "1AZ1",
            "JAMBI": "1DZ1", "HEAD OFFICE": "Z001", "BANDUNG 1": "BZ01", "BANDUNG 2": "NZ01",
            "BEKASI": "CZ01", "CILACAP": "IZ01", "CILEUNGSI2": "JZ01", "SEMARANG": "HZ01",
            "CIKOKOL": "KZ01", "LAMPUNG": "LZ01", "MALANG": "MZ01", "MANADO": "1YZ1",
            "BATAM": "2DZ1", "MADIUN": "2MZ1"
        };
        return map[String(cabangName).toUpperCase()] || cabangName;
    },

    getRabCreatorByUlok: async (ulok) => {
        try {
            const doc = await getDoc(config.SPREADSHEET_ID);
            const sheet = doc.sheetsByTitle[config.DATA_ENTRY_SHEET_NAME];
            const rows = await sheet.getRows();
            const found = rows.find(r => r['Nomor Ulok'] === ulok);
            return found ? found[config.COLUMN_NAMES.EMAIL_PEMBUAT] : null;
        } catch (e) { return null; }
    },

    getActivePengawasanByPic: async (email) => {
        try {
            const doc = await getDoc(config.PENGAWASAN_SPREADSHEET_ID);
            const sheet = doc.sheetsByTitle[config.PENUGASAN_SHEET_NAME];
            const rows = await sheet.getRows();
            const projects = [];
            rows.forEach(row => {
                if (String(row['Email_BBS']).trim().toLowerCase() === String(email).trim().toLowerCase()) {
                    projects.push({
                        kode_ulok: row['Kode_Ulok'],
                        cabang: row['Cabang']
                    });
                }
            });
            return projects;
        } catch (e) { return []; }
    },

    getPicEmailByUlok: async (ulok) => {
        try {
            const doc = await getDoc(config.PENGAWASAN_SPREADSHEET_ID);
            const sheet = doc.sheetsByTitle[config.PENUGASAN_SHEET_NAME];
            const rows = await sheet.getRows();
            const found = rows.reverse().find(r => r['Kode_Ulok'] === ulok);
            return found ? found['Email_BBS'] : null;
        } catch (e) { return null; }
    },

    // --- 6. Calendar Logic ---

    createCalendarEvent: async (eventData) => {
        try {
            const event = {
                'summary': eventData.title,
                'description': eventData.description,
                'start': { 'date': eventData.date },
                'end': { 'date': eventData.date },
                'attendees': eventData.guests.map(email => ({ email })),
            };

            await calendar.events.insert({
                calendarId: 'primary',
                resource: event,
                sendUpdates: 'all'
            });
            console.log("Calendar event created");
        } catch (error) {
            console.error("Calendar Error:", error);
        }
    },

    // --- 7. Email Helper Khusus ---
    sendSpkFinalNotification: async (rowData, pdfLink, approver) => {
        /* Akan diimplementasikan di controller masing-masing, atau placeholder */
    },
    sendRejectionEmail: async (rowData, reason, approver, type = 'RAB') => {
        const email = rowData[config.COLUMN_NAMES.EMAIL_PEMBUAT] || rowData['Dibuat Oleh'];
        if (!email) return;

        const subject = `[DITOLAK] Pengajuan ${type} Proyek ${rowData.Nama_Toko || rowData.nama_toko}`;
        const body = `
            <p>Pengajuan ${type} Anda telah ditolak oleh ${approver}.</p>
            <p><b>Alasan:</b> ${reason}</p>
            <p>Silakan revisi dan ajukan kembali.</p>
        `;

        await googleService.sendEmail(email, subject, body);
    },

    // Fungsi ini khusus untuk render email template dengan data
    sendPengawasanEmail: async (recipients, subject, data) => {
        const { emailLogicService } = require('./emailLogicService'); // Hindari circular dependency
        // Render HTML disini atau pass data ke controller
        // Untuk simplicity, kita anggap controller memanggil sendEmail langsung dengan HTML body
    }
};

module.exports = { googleService };