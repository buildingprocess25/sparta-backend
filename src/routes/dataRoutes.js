// src/routes/dataRoutes.js
const express = require('express');
const router = express.Router();
const dataController = require('../controllers/dataController');

// Route: GET /get-data
// Menggantikan endpoint '/get-data' di data_api.py Flask
router.get('/get-data', dataController.getData);

module.exports = router;