const express = require('express');
const router = express.Router();
const multer = require('multer');
const midController = require('../controllers/midController');
const desaController = require("../controllers/desaController");
const mainMidController = require("../controllers/mainMidController");

// Setup multer dengan memoryStorage + limit size + file filter
const storage = multer.memoryStorage();
const upload = multer({
    storage: storage,
    limits: { fileSize: 5 * 1024 * 1024 }, // Maksimum 5MB
    fileFilter: (req, file, cb) => {
        if (!file.originalname.match(/\.(xlsx)$/)) {
            return cb(new Error('Hanya file Excel (.xlsx) yang diperbolehkan!'), false);
        }
        cb(null, true);
    }
});

// Middleware untuk menangani error saat upload file
const uploadErrorHandler = (err, req, res, next) => {
    if (err instanceof multer.MulterError) {
        return res.status(400).json({ message: `Error upload file: ${err.message}` });
    } else if (err) {
        return res.status(400).json({ message: err.message });
    }
    next();
};

// 🟢 Rute untuk upload file Master MID
router.post('/upload-master-mid', upload.single('file'), uploadErrorHandler, midController.uploadMasterMID);

// 🟢 Rute untuk upload file MID
router.post('/upload-mid', upload.single('file'), uploadErrorHandler, mainMidController.uploadMid);

// 🟢 Rute untuk upload file Desa
router.post('/upload-desa', upload.single('file'), uploadErrorHandler, desaController.uploadDesa);

module.exports = router;
