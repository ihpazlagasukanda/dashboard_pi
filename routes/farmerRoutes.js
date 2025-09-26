const express = require('express');
const router = express.Router();
const farmerController = require('../controllers/farmerController');

// Refresh data dari database (bisa dipanggil periodik via cron)
router.post('/refresh', farmerController.refreshData);

// Get data dengan pagination dan filter
router.get('/farmers', farmerController.getFarmers);

// Get semua data by kabupaten (tanpa pagination)
router.get('/by-kabupaten', farmerController.getAllFarmersByKabupaten);

// Get daftar kabupaten
router.get('/kabupaten-list', farmerController.getKabupatenList);

// Get statistik data
router.get('/stats', farmerController.getStats);

// Export data ke CSV
router.get('/export', farmerController.exportData);

// Endpoint baru untuk mendapatkan daftar tahun yang tersedia
router.get('/available-years', farmerController.getAvailableYears);

// Endpoint untuk mengubah tahun aktif
router.post('/set-active-year', farmerController.setActiveYear);

module.exports = router;