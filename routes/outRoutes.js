const express = require('express');
const router = express.Router();
const multer = require('multer');
const path = require('path');

const midController = require('../controllers/midController');
const desaController = require("../controllers/desaController");
const mainMidController = require("../controllers/mainMidController");

// Konfigurasi multer
const storage = multer.memoryStorage();
const upload = multer({
    storage: storage,
    fileFilter: (req, file, cb) => {
        console.log("📂 File diterima:", file); // Debugging

        // Ekstensi yang diperbolehkan
        const allowedExt = ['.xlsx', '.xls'];
        const fileExt = path.extname(file.originalname).toLowerCase();

        // MIME type yang valid (bisa dicek dengan console.log)
        const allowedMime = [
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "application/vnd.ms-excel"
        ];

        // Cek apakah format sesuai
        if (allowedExt.includes(fileExt) && allowedMime.includes(file.mimetype)) {
            return cb(null, true);
        } else {
            console.log("❌ Format tidak valid!", fileExt, file.mimetype);
            return cb(new Error('Only .xlsx and .xls files are allowed!'));
        }
    },
    limits: { fileSize: 5 * 1024 * 1024 } // Maksimum 5MB
});

// Middleware untuk menangani error multer
const uploadMiddleware = (req, res, next) => {
    upload.single('file')(req, res, (err) => {
        if (err instanceof multer.MulterError) {
            return res.status(400).json({ message: `Multer error: ${err.message}` });
        } else if (err) {
            return res.status(400).json({ message: err.message });
        }
        next();
    });
};

// Endpoint untuk upload Master MID
router.post('/upload-master-mid', uploadMiddleware, midController.uploadMasterMID);

// Endpoint untuk upload MID
router.post('/upload-mid', uploadMiddleware, mainMidController.uploadMid);

// Endpoint untuk upload Desa
router.post("/upload-desa", uploadMiddleware, desaController.uploadDesa);

module.exports = router;
