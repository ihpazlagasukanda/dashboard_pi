const express = require("express");
const router = express.Router();
const db = require("../config/db"); // Sesuaikan dengan koneksi MySQL kamu
const { generateAllReports } = require('../services/reportGenerator');

router.get('/generate-excel-2403', async (req, res) => {
    generateAllReports();
    res.send('Excel file generated manually.');
});

module.exports = router;
