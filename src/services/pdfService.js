// src/services/pdfService.js
const chromium = require('@sparticuz/chromium');
const puppeteer = require('puppeteer-core');
const ejs = require('ejs');
const path = require('path');
const moment = require('moment');

// --- Helper Functions ---

// Format Rupiah (Rp 1.000.000)
const formatRupiah = (number) => {
    return new Intl.NumberFormat('id-ID', {
        style: 'currency',
        currency: 'IDR',
        minimumFractionDigits: 0
    }).format(number || 0).replace("Rp", "Rp "); // Tambah spasi
};

// Parser: Mengubah data form flat (Item_1, Item_2...) menjadi Array Terstruktur
const parseItems = (data) => {
    const items = [];
    const itemsSbo = [];
    let totalNonSbo = 0;
    let totalSbo = 0;

    // Loop 200 baris sesuai limit form di frontend
    for (let i = 1; i <= 200; i++) {
        const kategori = data[`Kategori_Pekerjaan_${i}`];
        const jenis = data[`Jenis_Pekerjaan_${i}`];
        const satuan = data[`Satuan_${i}`];
        const volumeStr = data[`Volume_${i}`];
        const hargaStr = data[`Harga_Satuan_${i}`];
        const totalStr = data[`Total_Harga_Item_${i}`];

        // Skip jika baris kosong
        if (!kategori && !jenis) continue;

        // Bersihkan format angka (hapus Rp dan titik)
        const parseNum = (str) => parseFloat(String(str).replace(/[^0-9.-]+/g, "")) || 0;

        const volume = parseNum(volumeStr);
        const harga = parseNum(hargaStr);
        const total = parseNum(totalStr);

        const itemObj = {
            no: i, // Nomor urut bisa di-reset nanti saat rendering
            kategori,
            jenis,
            satuan,
            volume, // Akan diformat di template
            harga,
            total
        };

        if (kategori === 'PEKERJAAN SBO') {
            itemsSbo.push(itemObj);
            totalSbo += total;
        } else {
            items.push(itemObj);
            totalNonSbo += total;
        }
    }

    return { items, itemsSbo, totalNonSbo, totalSbo };
};

// Core: Generate PDF Buffer menggunakan Puppeteer
const createPdfBuffer = async (htmlContent) => {
    let browser = null;
    try {
        // Konfigurasi Browser untuk Vercel (Serverless)
        // Jika testing lokal, pastikan chrome path sesuai
        const executablePath = await chromium.executablePath() ||
            (process.platform === 'win32' ? 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe' : '/usr/bin/google-chrome');

        browser = await puppeteer.launch({
            args: chromium.args,
            defaultViewport: chromium.defaultViewport,
            executablePath: executablePath,
            headless: chromium.headless,
            ignoreHTTPSErrors: true,
        });

        const page = await browser.newPage();

        // Set Content
        await page.setContent(htmlContent, { waitUntil: 'networkidle0' });

        // Generate PDF
        const pdfBuffer = await page.pdf({
            format: 'A4',
            printBackground: true,
            margin: {
                top: '1cm',
                right: '1cm',
                bottom: '1cm',
                left: '1cm'
            }
        });

        return pdfBuffer;
    } catch (error) {
        console.error("Puppeteer Error:", error);
        throw error;
    } finally {
        if (browser) await browser.close();
    }
};

const pdfService = {
    // 1. Generate RAB PDF (Non-SBO atau Full)
    generateRABPDF: async (data, options = { excludeSbo: false }) => {
        try {
            // Parse data items
            const { items, itemsSbo, totalNonSbo } = parseItems(data);

            // Pilih items yang akan ditampilkan
            // Jika excludeSbo true (untuk PDF instruksi lapangan), hanya tampilkan items Non-SBO
            const displayItems = options.excludeSbo ? items : [...items, ...itemsSbo];

            // Hitung Grand Total untuk display
            // Perhatikan logic pembulatan (Round down ke puluhan ribu terdekat)
            const rawTotal = options.excludeSbo ? totalNonSbo : (totalNonSbo + itemsSbo.reduce((acc, curr) => acc + curr.total, 0));
            const pembulatan = Math.floor(rawTotal / 10000) * 10000;
            const ppn = pembulatan * 0.11;
            const grandTotal = pembulatan + ppn;

            // Render Template EJS
            // Lokasi: src/templates/pdf_report.ejs
            const templatePath = path.join(process.cwd(), 'src', 'templates', 'pdf_report.ejs');

            const html = await ejs.renderFile(templatePath, {
                data: data,
                items: displayItems,
                totals: {
                    subtotal: rawTotal,
                    pembulatan: pembulatan,
                    ppn: ppn,
                    grandTotal: grandTotal
                },
                helpers: { formatRupiah, moment }
            });

            return await createPdfBuffer(html);

        } catch (error) {
            throw new Error(`Gagal membuat PDF RAB: ${error.message}`);
        }
    },

    // 2. Generate Rekap PDF
    generateRecapPDF: async (data) => {
        try {
            const { items, itemsSbo, totalNonSbo, totalSbo } = parseItems(data);

            const pembulatan = Math.floor((totalNonSbo + totalSbo) / 10000) * 10000;
            const ppn = pembulatan * 0.11;
            const grandTotal = pembulatan + ppn;

            const templatePath = path.join(process.cwd(), 'src', 'templates', 'recap_report.ejs');

            const html = await ejs.renderFile(templatePath, {
                data: data,
                totalNonSbo: totalNonSbo,
                totalSbo: totalSbo,
                grandTotal: grandTotal, // Sebelum PPN & Pembulatan (Total Raw)
                finalTotal: grandTotal, // Setelah PPN logic (di Python recap logic agak beda, sesuaikan kebutuhan)
                helpers: { formatRupiah, moment }
            });

            return await createPdfBuffer(html);

        } catch (error) {
            throw new Error(`Gagal membuat PDF Rekap: ${error.message}`);
        }
    },

    // 3. Generate SPK PDF
    generateSpkPDF: async (data) => {
        try {
            const templatePath = path.join(process.cwd(), 'src', 'templates', 'spk_template.ejs');

            // Data 'Biaya' di SPK biasanya sudah final dari input user/database
            // Pastikan format tanggal benar

            const html = await ejs.renderFile(templatePath, {
                data: data,
                helpers: { formatRupiah, moment }
            });

            return await createPdfBuffer(html);

        } catch (error) {
            throw new Error(`Gagal membuat PDF SPK: ${error.message}`);
        }
    }
};

module.exports = { pdfService };