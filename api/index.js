// api/index.js
const express = require('express');
const cors = require('cors');

// Import Routes (Akan kita buat file-file ini di langkah selanjutnya dalam folder src/routes)
// Ini memecah logic yang menumpuk di app.py dan data_api.py
const authRoutes = require('../src/routes/authRoutes');
const rabRoutes = require('../src/routes/rabRoutes');
const spkRoutes = require('../src/routes/spkRoutes');
const pengawasanRoutes = require('../src/routes/pengawasanRoutes');
const dataRoutes = require('../src/routes/dataRoutes');

const app = express();

// --- 1. Konfigurasi Middleware ---
// Menggantikan setup CORS di app.py
app.use(cors({
    origin: [
        "http://127.0.0.1:5500",
        "http://localhost:5500",
        "http://168.110.201.69",
        "http://168.110.201.69:8082",
        "https://instruksi-lapangan.vercel.app",
        "https://instruksi-lapangan.vercel.app/"
    ],
    methods: ["GET", "POST", "OPTIONS", "PUT", "PATCH", "DELETE"],
    allowedHeaders: ["Content-Type", "Authorization"],
    credentials: true
}));

// Middleware untuk membaca JSON dan Form Data (pengganti request.get_json() di Flask)
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// --- 2. Route Dasar ---

// Root Endpoint (Pengganti route '/' di app.py)
app.get('/', (req, res) => {
    res.status(200).send("Backend server is running and healthy (Express on Vercel).");
});

// Health Check (Pengganti route '/health' di app.py)
app.get('/health', (req, res) => {
    res.status(200).json({ status: "ok", message: "Service is alive" });
});

// --- 3. Registrasi API Routes ---
// Kita mengelompokkan route berdasarkan fiturnya agar lebih rapi

// Route Otentikasi (Login) -> Menggantikan '/api/login'
app.use('/api', authRoutes);

// Route RAB (Submit, Approval, Check Status) 
// Menggantikan '/api/submit_rab', '/api/check_status', dll
app.use('/api', rabRoutes);

// Route SPK (Submit, Approval, Get Data)
// Menggantikan '/api/submit_spk', '/api/get_spk_status', dll
app.use('/api', spkRoutes);

// Route Pengawasan (Input PIC, Laporan Harian)
// Menggantikan group '/api/pengawasan/...'
app.use('/api/pengawasan', pengawasanRoutes);

// Route Data Umum (Get Data dari Spreadsheet)
// Menggantikan Blueprint 'data_api' di Flask
app.use('/', dataRoutes);

// --- 4. Export App untuk Vercel ---
// Tidak menggunakan app.listen() karena Vercel yang akan menjalankan fungsi ini
module.exports = app;