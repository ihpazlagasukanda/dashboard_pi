const express = require("express");
const router = express.Router();
const farmerController = require("../controllers/farmerController");

// Generate JSON manual
router.get("/generate-farmers", farmerController.generateFarmersJSON);

// Ambil data JSON
router.get("/farmers", farmerController.getFarmers);

// routes.js - Tambahkan route baru
router.get('/all-farmers', farmerController.getAllFarmersByKabupaten);

router.get('/kabupaten-list', farmerController.getKabupatenList);

module.exports = router;