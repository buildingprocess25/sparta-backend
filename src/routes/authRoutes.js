// src/routes/authRoutes.js
const express = require('express');
const router = express.Router();
const authController = require('../controllers/authController');

// Route: POST /api/login
// Menerima JSON body: { email, cabang }
router.post('/login', authController.login);

module.exports = router;