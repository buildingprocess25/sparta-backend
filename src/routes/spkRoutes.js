// src/routes/spkRoutes.js
const express = require('express');
const router = express.Router();
const spkController = require('../controllers/spkController');

// --- API Endpoints untuk SPK ---

// 1. Cek Status SPK (Mencegah duplikasi SPK untuk Ulok yang sama)
// Endpoint: GET /api/get_spk_status?ulok=...&lingkup=...
router.get('/get_spk_status', spkController.getSpkStatus);

// 2. Submit SPK (Baru atau Revisi)
// Endpoint: POST /api/submit_spk
router.post('/submit_spk', spkController.submitSpk);

// 3. Handle Approval SPK (Link dari Email untuk Approve/Reject)
// Endpoint: GET /api/handle_spk_approval?action=...&row=...
router.get('/handle_spk_approval', spkController.handleApproval);
router.post('/handle_spk_approval', spkController.handleApproval); // Support POST jika diperlukan

module.exports = router;