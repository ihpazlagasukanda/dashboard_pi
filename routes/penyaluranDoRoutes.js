const express = require("express");
const multer = require("multer");
const { uploadPenyaluranDo } = require("../controllers/doController");
const { uploadPoktan } = require("../controllers/poktanController");

const router = express.Router();

// Konfigurasi Multer untuk upload file
const storage = multer.memoryStorage();
const upload = multer({ storage });

router.post("/upload-penyalurando", upload.single("file"), uploadPenyaluranDo);

router.post("/upload-poktan", upload.single("file"), uploadPoktan);

module.exports = router;
