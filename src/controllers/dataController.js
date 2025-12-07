const { GoogleSpreadsheet } = require('google-spreadsheet');
// Hapus import JWT karena kita akan biarkan google-spreadsheet yang menanganinya
// const { JWT } = require('google-auth-library'); 
const config = require('../config/config');

// --- Konfigurasi ID Spreadsheet per Cabang ---
const SPREADSHEET_IDS = {
    "ACEH": { "ME": "1KZyh0VVn7dZRyvEm6q7RV4iA5jzg7u2oRTIvSHxeFu0", "SIPIL": "11b_oUEmsjqFkB8CX8uOg8SUjlUpfgjZQq6qN1BtVBm4" },
    "BALARAJA": { "ME": "1FVRlRK1Qop1Q7OlHKsIc14BhRSHn9XH2gLWarxWMON4", "SIPIL": "1nBPJjM17vwO1tTsC2m8VRnnhQQKC_bUEpfFebfib_g0" },
    "BALI": { "ME": "1ih2kuwqcCa7EiMTbJX6qiNsrv7vpmPE825a_9qAz2Ng", "SIPIL": "1gD_z66c4zPBXMXcKGxIFs8C2h2uiDXwO3J7r_e1BxJc" },
    "BANDUNG 1": { "ME": "13xo-gjJnXlCWXKNfW3ZugEYxK6qz1jvUq2ScSBJF1cw", "SIPIL": "11Weq3EIPCo-_bOPvTrSBQPcpXNPu5VTIoWHPA4ZJfHw" },
    "BANDUNG 2": { "ME": "15kOWwgAQTZKZ-Ofm_jrlaKHJbbU_wDkYTptefCVOTCY", "SIPIL": "1LsIBm2829RqC7Tu9fdNnWTgDKLT9gCz4zWCCHjMbmzI" },
    "BANGKA": { "ME": "1nhjwDMrG8fWUXu-4tTdcfgbLpEfOfEh90SSqOVhaaPc", "SIPIL": "1SmZJ5hNYAQba5byj-cduo4yRRNl74LEwHvayt_Wj27g" },
    "BANJARMASIN": { "ME": "1_1alrg4qaA2HeI_FpKqCczP-S8iasQ6P93jUDUISfgw", "SIPIL": "13PEkDV55bcV3SRU1gEvaU2EIDRUYfAREXLAkIYIX660" },
    "BATAM": { "ME": "1pUf7XBVN-eR7ptaeCqjHIokeRyu2xrE-_Y4b0KK0Uz8", "SIPIL": "1sl-1CfcEerMzpU-Boc-p7ePSZWbj53kq4DXjrC145R4" },
    "BEKASI": { "ME": "1TV8OqBBvHG93tuBe__aYm6Yo-tfRvlASoRe87qTD5WI", "SIPIL": "1rwdlpE0G_NP3zAsRECtALgXFh1g6Fp5-XC-bxdwvXhw" },
    "BELITUNG": { "ME": "1nhjwDMrG8fWUXu-4tTdcfgbLpEfOfEh90SSqOVhaaPc", "SIPIL": "1bLyVM6OEMHzR6LHyNlatDePmDOKE7UtHGXPfU-CqQsY" },
    "BENGKULU": { "ME": "1nhjwDMrG8fWUXu-4tTdcfgbLpEfOfEh90SSqOVhaaPc", "SIPIL": "1ZE69uEhHuG8hVTPqGge53wQBh4uy2j_UTXi0q8gHQNE" },
    "BOGOR": { "ME": "1ESn1O-gHsjJFoBYZGhzMxkI4IHO-L82aESTRWO_n26Y", "SIPIL": "1DMUZYn8ElI5j7yFn1Eeu_d7n4CIzYWbvLcdyjIsOrUI" },
    "CIANJUR": { "ME": "14MwcXFxSAkWAYf0CFSADNkBaONgcNjJBs9BsluMHAKU", "SIPIL": "1yrDQzOjYlg7YxI0E9clBZRG_Iz8KHoo9FA6abYxgDRs" },
    "CIKOKOL": { "ME": "1oEX2bmyi40u09LLfjsL-A6ec7Iw7hAXKVfk0SJDnnBw", "SIPIL": "1aSpjDMumbtEa0BrrK2T7qaGd2AIep09qSX14BwzI5gY" },
    "CILACAP": { "ME": "1nm9L7rNKnexO06dW0YkmfXO9hhlKN7Ai88iEWIC8Olw", "SIPIL": "1Lgjo4WB1cXNF9Howneisr-PdJLCjrJSE60C9JtxNIYI" },
    "CILEUNGSI": { "ME": "1JVtRGj9LxkYrBMCBfK1Tc4IMZLqiXyzhQBgQw0FBsj8", "SIPIL": "1yhAK38hk_pSzCII_rF329j8ebNOUbxDHju5EbG7L1t0" },
    "GORONTALO": { "ME": "1sYssoeMYr2iA0n-XLFLJTHo5Hr0ogOrAwIUnO8JKwro", "SIPIL": "1ZGsW7uzggEhvT77Kd8bqr-VokwPFlIjg6VOY_tntBN4" },
    "JAMBI": { "ME": "1SjIkwq9jLRKPeK3YXdBpEMJdCOlnIxaQ496mWG-aqWQ", "SIPIL": "1G_B08N-XSUqrVtcvizvUoZITPzST7pIiYbLouRFcfMc" },
    "JEMBER": { "ME": "1VjYiZ0C-h-ADIbhrJ4fAlr_oynqQabtASOTv5bc_JxY", "SIPIL": "1BBikMALwiAKam5Odz1e01Vb1aWtToCjRb2KNP4Zq-uE" },
    "KARAWANG": { "ME": "1pyhKAeNerRsOhok9VQ0Phb4h67W136ftaF7nJTXomXI", "SIPIL": "1W5nob_MXaCQCO90tdFO002irPBOTQZVNr_NMeoNjzf0" },
    "KLATEN": { "ME": "1qDUOB2xtNYEp3Kc_HbYaznc6cSfp4pa5vI9ZZZc3Hhk", "SIPIL": "1xA5qqKWlniFE8FCn0YTAnIpVfI7HKjPrFTOwUIv-7KI" },
    "KOTABUMI": { "ME": "14s6kC8qvAEjYW9E5-T1_cj_pXIkk82BZTrQd-UmPdmg", "SIPIL": "1TYCSAQ1Vu_K7KnfEfKdXTJnFuk-t6ZtKFktieXt199w" },
    "LAMPUNG": { "ME": "1Imw8469Od3dHiomBxPXgVQmn5_JOJkTT0iH-EJA6w0Y", "SIPIL": "1zY8SbebteKSb4OOGqQxDBFxJXlZHfCRz3AbmePY-4Mk" },
    "LOMBOK": { "ME": "1QeicLcNbNK7D0aYsVrM9klr273tiFZa2y9i1A3uK7Wc", "SIPIL": "17iy_6ONzglPDhOk4pm3KQpV4AIT53pZLgmERTMSrE5c" },
    "LUWU": { "ME": "1_2-ZXbpSoQN9uuHEkyFVVUjugx2n-QW3Qi0cfDeh_qM", "SIPIL": "1TDZCEUZC6TKeOWwX7I5VCDTInqy-nEVjNHYMr5ZiAbM" },
    "MADIUN": { "ME": "1UPv6sdqv5TtsL4UC3v0EOqY0MnvAB3Et-aTt2KMOrmI", "SIPIL": "1vJGFwQj92f9Mxs8dKT4dHDDCXer0DyFOPlUh__C956c" },
    "MAKASSAR": { "ME": "1XnFay1yEsQfUJ-fdccCCY8jfVuDks2sTUTgpuFrAj_k", "SIPIL": "1uUEIU7Ieqs7Nm68HBaRo4MHwob4RLTPUEtI_0LI6X28" },
    "MALANG": { "ME": "1FLZ-faMU1Cyp6OOVS5oQXdZNc4nDCvMs5rIzRvlEY0s", "SIPIL": "1EdKp2Yxz4GkEb7fyx8dLzIVNG7mSemZN_CnN3qQN_lI" },
    "MANADO": { "ME": "1kJBHqvxmvb8Kw4UvyxD_ecZsrjGGgtkCql2EAab5xfw", "SIPIL": "1sbt6Fu-OvH5qXf-LN0MpQaBnLF3tn1QcCUTNBgrVVIU" },
    "MANOKWARI": { "ME": "1O9tKmAojv42gRsNn6lryZ6Npo_kSt_J1LDOT9ZCNVMs", "SIPIL": "1DXvR6m9gZ4K_1m0hJwyJnSRAXtzy8y8uEyk68_yqziE" },
    "MEDAN": { "ME": "1TaEaBMTAXxXRS73VB8Ijnp53pZI6F6odweCBQV4vaFo", "SIPIL": "1ty-YLmkZqGAmVu7UGxjdtyBkmoYRhGibTk7ZkUt2dPs" },
    "NTT": { "ME": "1anZZF1ptwE-j_DroCLep6yfRC0A8jrFXp8v9cZ5gh7k", "SIPIL": "1dCfRyTOgMayV-m7y647X0Gqm_FPbBVl4rwJ8SPrav9o" },
    "PALEMBANG": { "ME": "1nhjwDMrG8fWUXu-4tTdcfgbLpEfOfEh90SSqOVhaaPc", "SIPIL": "1xfqITP8OOBTS7JdxBbc49U4HIlA0qbBsXZjuEuh_g4" },
    "PARUNG": { "ME": "1BTkXU0aVovA2p4zYUok3dPq57aqTiHwj_RboQUh6-54", "SIPIL": "1w7aBoRFAvt2kkuezHoPeKYxUx2FPgHCJ8eNWA7AvvCw" },
    "PEKANBARU": { "ME": "18x4TJ751pGnSeCtGXBWbbuKTJ68oK8MHYMQaCF-nQ_8", "SIPIL": "14vFw1rx4wDCCw2pKDIWKvnNqA4aAvnwrD0eRjNtguzo" },
    "PLUMBON": { "ME": "1mHnA9lQE5ymjUL4bc51MmdwTOkCO_okXhqdg3-XmuMY", "SIPIL": "1CPQcKkJrlvvyt4SSRbcu5_ee4tjFpS7upwxcVYRIZAQ" },
    "PONTIANAK": { "ME": "1o7RgfXKoS1FLok68SrSETBO0TQy1-yI0OtM0LtHriHo", "SIPIL": "1bzjdA0DJZqtTR3W0OdYU-J85Jipwk64WtYt_3TPwnAk" },
    "REMBANG": { "ME": "1KF5HvMQoNu9dk1lmaitc_I19nofgaiLPMzulHAXanJs", "SIPIL": "18eIGamo7yGhLKu16V5n4AceIIbMW1KxQ4PJfpZGl9D8" },
    "SEMARANG": { "ME": "1EAvelezXOEj06yiHeYxsf8R9H9wuUJImG0fLlLlfaA4", "SIPIL": "1a3aPYLepwr4u5lR0MmTO_0N-wb8XTmOld8flJCgRX30" },
    "SERANG": { "ME": "1hyjjNHkHZLJo3Z-n5c-Vik6SZCc1Sd2FbnjpnvrFz7A", "SIPIL": "12y3OZxSwPlrXFgiIA_6ediJ9DK6hJ_GKVznOsm1KBok" },
    "SIDOARJO": { "ME": "1BcUs28NrcsqPk9FA3oRrVmBBmlWyGjRspZ9dJ217kAQ", "SIPIL": "1k2eBX4vZaAHL30SZcAX3kXNastEuscirr-2RuiTKIRw" },
    "SIDOARJO BPN_SMD": { "ME": "1j1B7Yvgz5X02VcCuNSJC7btZv_tumFFmeMvxNEaS0ec", "SIPIL": "18BT0u4sHuNZA-NKxdcDeNTkxT35Mag4zOW9RjduRE8A" },
    "SORONG": { "ME": "1BfAWqaN-7fk5kZMWhK0SOKousSYM6pis6jVEplQLnYs", "SIPIL": "1NbG4HOt_Zl1auzR_kw-e8mG0ajrP9P-FFLfjNCD6Oyg" },
    "SUMBAWA": { "ME": "1WB6KA2sFD11Ol81IYbM3ugH34Sg1JU1Aw6Ny8zUgCpg", "SIPIL": "1Y3AqDtyXUyJhyrvT0slbQRlW14VUd9zA8P_1vPqqI8A" },
    "TEGAL": { "ME": "13qXfMnPrOYB-fJEf_7FGgL-UENA3KWBj5uTiY4PLoAQ", "SIPIL": "1i_2TLKswkCUoNm1Z6hDmc3j9k0qNcF38MZoC-z2-PNM" },
    "HEAD OFFICE": { "ME": "1oQfZkWSP-TWQmQMY-gM1qVcLP_i47REBmJj1IfDNzkg", "SIPIL": "1Jf_qTHOMpmyLWp9zR_5CiwjyzWWtD8cH99qt4kJvLOw" }
};

// ID SBO
const SBO_SPREADSHEET_ID = "11efRx3l6fXn5XLd_HZEQHmFWwQX5vRUIBVgFecCYMXQ";

// Mapping Nama Cabang ke Kode Ulok
const BRANCH_TO_ULOK_MAP = {
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

// --- Helper Functions ---

const processPriceValue = (rawValue) => {
    const valueStr = String(rawValue).trim().toLowerCase();
    if (valueStr === 'kondisional') return 'Kondisional';
    if (valueStr === 'sbo') return 'SBO';
    if (valueStr.includes('kontraktor')) return 0.0;

    // Hapus koma dan karakter non-numeric
    const cleanStr = String(rawValue).replace(/,/g, '').trim();
    const floatVal = parseFloat(cleanStr);
    return isNaN(floatVal) ? 0.0 : floatVal;
};

// Proses Sheet Utama (Non-SBO) dengan V3 Style
const processSheet = async (doc, lingkup) => {
    const sheet = doc.sheetsByIndex[0];

    // V3: Load semua cells agar bisa diakses dengan getCell
    await sheet.loadCells();

    const categorizedPrices = {};
    let currentCategory = "Uncategorized";

    // Kolom (0-based) sesuai kode Python lama:
    // B=1 (No), D=3 (Jenis Pekerjaan), E=4 (Satuan)
    const noColIndex = 1;
    const jenisPekerjaanColIndex = 3;
    const satColIndex = 4;

    // Tentukan baris header (0-based)
    // Sipil baris 16 (index 15), ME baris 13 (index 12)
    const targetHeaderRowIndex = (lingkup === "SIPIL") ? 15 : 12;

    // Cek apakah sheet cukup panjang
    if (sheet.rowCount <= targetHeaderRowIndex) {
        throw new Error(`Sheet tidak memiliki cukup baris untuk header ${lingkup}.`);
    }

    // Cari kolom Material dan Upah di baris header
    let materialColIndex = -1;
    let upahColIndex = -1;

    for (let c = 0; c < sheet.columnCount; c++) {
        const cellValue = String(sheet.getCell(targetHeaderRowIndex, c).value || "").toLowerCase();
        if (cellValue.includes("material")) materialColIndex = c;
        if (cellValue.includes("upah")) upahColIndex = c;
    }

    if (materialColIndex === -1 || upahColIndex === -1) {
        throw new Error("Header 'Material' atau 'Upah' tidak ditemukan.");
    }

    // Loop data mulai dari baris setelah header
    for (let r = targetHeaderRowIndex + 1; r < sheet.rowCount; r++) {
        // Ambil Jenis Pekerjaan
        const jenisPekerjaan = String(sheet.getCell(r, jenisPekerjaanColIndex).value || "").trim();

        // Skip jika kosong
        if (!jenisPekerjaan || jenisPekerjaan.toUpperCase() === "JENIS PEKERJAAN") continue;

        const noVal = String(sheet.getCell(r, noColIndex).value || "").trim();

        // Cek apakah ini baris Kategori (Romawi I, II, V, dll)
        if (/^[IVXLCDM]+$/.test(noVal)) {
            currentCategory = jenisPekerjaan;
            if (!categorizedPrices[currentCategory]) {
                categorizedPrices[currentCategory] = [];
            }
            continue;
        }

        const satuanVal = String(sheet.getCell(r, satColIndex).value || "").trim();
        if (!satuanVal) continue;

        const hargaMaterialRaw = sheet.getCell(r, materialColIndex).value || 0;
        const hargaUpahRaw = sheet.getCell(r, upahColIndex).value || 0;

        const itemData = {
            "Jenis Pekerjaan": jenisPekerjaan,
            "Satuan": satuanVal,
            "Harga Material": processPriceValue(hargaMaterialRaw),
            "Harga Upah": processPriceValue(hargaUpahRaw)
        };

        if (!categorizedPrices[currentCategory]) {
            categorizedPrices[currentCategory] = [];
        }
        categorizedPrices[currentCategory].push(itemData);
    }

    return categorizedPrices;
};

// Proses Sheet SBO (Menggunakan getRows di v3)
const processSboSheet = async (doc, cabangKode, lingkup) => {
    const sheet = doc.sheetsByIndex[0];
    const rows = await sheet.getRows(); // V3 mengembalikan array object row

    const sboItems = [];

    rows.forEach(row => {
        // V3: Akses data via properti object langsung (Header di baris 1)
        const rowLingkup = String(row['Lingkup_Pekerjaan'] || "").trim().toUpperCase();
        const rowKodeCabang = String(row['Kode Cabang'] || "");

        if (rowLingkup === lingkup && rowKodeCabang.includes(cabangKode)) {
            sboItems.push({
                "Jenis Pekerjaan": row['Item Pekerjaan'],
                "Satuan": row['Satuan'],
                "Harga Material": processPriceValue(row['Harga Material']),
                "Harga Upah": 0.0 // SBO Upah 0
            });
        }
    });

    return sboItems.length > 0 ? { "PEKERJAAN SBO": sboItems } : {};
};

const dataController = {
    getData: async (req, res) => {
        const { cabang, lingkup } = req.query;

        console.log(`\n--- MULAI REQUEST: ${cabang} - ${lingkup} ---`);

        if (!cabang || !lingkup) {
            return res.status(400).json({ error: "Missing 'cabang' or 'lingkup' parameter" });
        }

        const cabangKey = cabang.toUpperCase();
        const lingkupKey = lingkup.toUpperCase();

        if (!SPREADSHEET_IDS[cabangKey] || !SPREADSHEET_IDS[cabangKey][lingkupKey]) {
            return res.status(404).json({ error: "Invalid 'cabang' or 'lingkup' parameter" });
        }

        const spreadsheetId = SPREADSHEET_IDS[cabangKey][lingkupKey];
        const cabangKode = BRANCH_TO_ULOK_MAP[cabangKey];

        try {
            const envEmail = process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL;
            const envKey = process.env.GOOGLE_PRIVATE_KEY;

            // --- DEBUG LOGGING ---
            console.log("1. Cek Credentials:");
            console.log("   - Email di Env:", envEmail);
            console.log("   - Panjang Key:", envKey ? envKey.length : "KOSONG");

            if (!envEmail || !envKey) {
                throw new Error("Credentials Error: Env vars kosong.");
            }

            const privateKey = envKey.replace(/\\n/g, '\n');
            const creds = { client_email: envEmail, private_key: privateKey };

            // 1. Ambil Data Utama
            console.log(`2. Mencoba akses Sheet UTAMA (${lingkupKey}):`);
            console.log(`   - ID: ${spreadsheetId}`);

            const doc = new GoogleSpreadsheet(spreadsheetId);
            await doc.useServiceAccountAuth(creds);
            await doc.loadInfo();
            console.log("   ✅ Sukses akses Sheet Utama");

            const processedData = await processSheet(doc, lingkupKey);

            // 2. Ambil Data SBO
            if (cabangKode) {
                console.log("3. Mencoba akses Sheet SBO:");
                console.log(`   - ID: ${SBO_SPREADSHEET_ID}`);

                try {
                    const sboDoc = new GoogleSpreadsheet(SBO_SPREADSHEET_ID);
                    await sboDoc.useServiceAccountAuth(creds);
                    await sboDoc.loadInfo();
                    console.log("   ✅ Sukses akses Sheet SBO");

                    const sboData = await processSboSheet(sboDoc, cabangKode, lingkupKey);
                    if (sboData) Object.assign(processedData, sboData);
                } catch (e) {
                    console.error("   ❌ GAGAL akses Sheet SBO (Cek Permission file ini!)");
                    console.error("   - Error:", e.message);
                    // Jangan throw error supaya data utama tetap muncul, cuma warning aja
                }
            } else {
                console.log("3. Skip Sheet SBO (Tidak ada kode cabang)");
            }

            res.json(processedData);

        } catch (error) {
            console.error("!!! ERROR FATAL !!!");
            console.error(error);
            res.status(500).json({
                error: `Server Error: ${error.message}`,
                hint: "Cek log Vercel untuk detail ID spreadsheet yang gagal."
            });
        }
    }
};

module.exports = dataController;