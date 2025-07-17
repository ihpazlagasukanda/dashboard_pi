const express = require("express");
const multer = require("multer");
const { uploadWcm } = require("../controllers/wcmController");
const { uploadF5 } = require("../controllers/uploadF5Controller");

const router = express.Router();

// Konfigurasi Multer untuk upload file
const storage = multer.memoryStorage();
const upload = multer({ storage });

router.post("/upload-wcm", upload.single("file"), uploadWcm);
router.post("/upload-f5", upload.single("file"), uploadF5);

module.exports = router;
