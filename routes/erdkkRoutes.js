const express = require("express");
const multer = require("multer");
const { uploadErdkk } = require("../controllers/erdkkController");

const router = express.Router();

// Konfigurasi Multer untuk upload file
const storage = multer.memoryStorage();
const upload = multer({ storage });

router.post("/upload-erdkk", upload.single("file"), uploadErdkk);

module.exports = router;
