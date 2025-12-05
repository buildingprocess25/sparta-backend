const express = require('express');
const router = express.Router();
const pengawasanController = require('../controllers/pengawasanController');

// Route: GET /api/pengawasan/init_data
// Mengambil data awal (PIC list, SPK list) untuk form input
router.get('/init_data', pengawasanController.getInitData);

// Route: GET /api/pengawasan/get_rab_url
// Mengambil URL PDF RAB berdasarkan kode ulok
router.get('/get_rab_url', pengawasanController.getRabUrl);

// Route: GET /api/pengawasan/get_spk_url
// Mengambil URL PDF SPK berdasarkan kode ulok
router.get('/get_spk_url', pengawasanController.getSpkUrl);

// Route: GET /api/pengawasan/active_projects
// Mengambil daftar proyek aktif untuk dropdown form harian
router.get('/active_projects', pengawasanController.getActiveProjects);

// Route: POST /api/pengawasan/submit
// Menyimpan laporan (Input PIC atau Laporan Harian) dan kirim email
router.post('/submit', pengawasanController.submitPengawasan);

module.exports = router;