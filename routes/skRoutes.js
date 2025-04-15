const express = require("express");
const multer = require("multer");
const { uploadSkBupati } = require("../controllers/skbupatiController");

const router = express.Router();

// Konfigurasi Multer untuk upload file
const storage = multer.memoryStorage();
const upload = multer({ storage });

router.post("/upload-skbupati", upload.single("file"), uploadSkBupati);

module.exports = router;
