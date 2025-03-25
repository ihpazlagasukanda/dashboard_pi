const express = require('express');
const router = express.Router();
const multer = require('multer');
const fileController = require('../controllers/fileController');  // Controller untuk file utama

const storage = multer.memoryStorage(); // Simpan file dalam memori sebagai buffer
// Setup multer middleware untuk upload
const upload = multer({ storage: storage });

// Route untuk meng-upload file utama
router.post('/upload', upload.array('files', 14), fileController.uploadFile);

module.exports = router;
