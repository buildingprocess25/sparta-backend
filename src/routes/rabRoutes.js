// src/routes/rabRoutes.js
const express = require('express');
const router = express.Router();
const rabController = require('../controllers/rabController');
const multer = require('multer');

// Konfigurasi Multer untuk upload file (disimpan di memori sementara agar bisa di-upload ke GDrive)
const upload = multer({
    storage: multer.memoryStorage(),
    limits: {
        fileSize: 10 * 1024 * 1024 // Batas ukuran file 10MB
    }
});

// --- API Endpoints ---

// 1. Submit RAB Reguler (Tanpa upload file manual)
router.post('/submit_rab', rabController.submitRab);

// 2. Submit RAB Kedua / Instruksi Lapangan (Wajib upload file PDF)
// Middleware 'upload.single' akan membaca file dari field bernama 'file_pdf'
router.post('/submit_rab_kedua', upload.single('file_pdf'), rabController.submitRabKedua);

// 3. Cek Status Pengajuan (Pending/Approved/Rejected)
// Endpoint ini menggantikan '/check_status' dan '/check_status_rab_2' di Python
// Frontend mengirim parameter ?type=rab1 atau ?type=rab2
router.get('/check_status', rabController.checkStatus);
router.get('/check_status_rab_2', rabController.checkStatus); // Alias untuk backward compatibility

// 4. Handle Approval (Link dari Email)
// Menangani Approve & Reject baik untuk RAB 1 maupun RAB 2
router.get('/handle_rab_approval', rabController.handleApproval);
router.post('/handle_rab_approval', rabController.handleApproval); // Support POST jika dari form

// Alias untuk RAB 2 agar link lama (jika ada) tetap jalan
router.get('/handle_rab_2_approval', rabController.handleApproval);
router.post('/handle_rab_2_approval', rabController.handleApproval);

module.exports = router;