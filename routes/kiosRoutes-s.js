const express = require("express");
const router = express.Router();
const { downloadRealisasiKios } = require("../controllers/kiosController-s");

router.get("/realisasi-kios", downloadRealisasiKios);

module.exports = router;
