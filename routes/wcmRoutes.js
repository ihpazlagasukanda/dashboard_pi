const express = require("express");
const multer = require("multer");
const { uploadWcm } = require("../controllers/wcmController");

const router = express.Router();

// Konfigurasi Multer untuk upload file
const storage = multer.memoryStorage();
const upload = multer({ storage });

router.post("/upload-wcm", upload.single("file"), uploadWcm);

module.exports = router;
