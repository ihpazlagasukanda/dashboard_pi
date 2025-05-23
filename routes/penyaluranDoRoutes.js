const express = require("express");
const multer = require("multer");
const { uploadPenyaluranDo } = require("../controllers/doController");

const router = express.Router();

// Konfigurasi Multer untuk upload file
const storage = multer.memoryStorage();
const upload = multer({ storage });

router.post("/upload-penyalurando", upload.single("file"), uploadPenyaluranDo);

module.exports = router;
