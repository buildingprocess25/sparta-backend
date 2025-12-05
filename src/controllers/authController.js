// src/controllers/authController.js
const { googleService } = require('../services/googleService');

const authController = {
    // Menangani request login
    // Menggantikan fungsi login() di app.py baris 57-70
    login: async (req, res) => {
        try {
            const { email, cabang } = req.body;
            // Debug: console.log("Login Request:", req.body);
            console.log("Login Request:", req.body);

            // 1. Validasi Input
            if (!email || !cabang) {
                return res.status(400).json({
                    status: "error",
                    message: "Email and cabang are required"
                });
            }

            // 2. Panggil Service untuk Cek Credentials di Google Sheets
            // Fungsi validateUser akan kita buat nanti di googleService.js
            // Return value: string role (jabatan) jika sukses, atau null jika gagal
            const role = await googleService.validateUser(email, cabang);

            // 3. Kirim Response
            if (role) {
                // Login Sukses
                return res.status(200).json({
                    status: "success",
                    message: "Login successful",
                    role: role
                });
            } else {
                // Login Gagal (Email/Cabang tidak cocok)
                return res.status(401).json({
                    status: "error",
                    message: "Invalid credentials"
                });
            }

        } catch (error) {
            console.error("Login Error:", error);
            // Error Handler Server
            return res.status(500).json({
                status: "error",
                message: "An internal server error occurred"
            });
        }
    }
};

module.exports = authController;