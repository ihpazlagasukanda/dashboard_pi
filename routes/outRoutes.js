const express = require('express');
const router = express.Router();
const multer = require('multer');
const midController = require('../controllers/midController');  // Controller untuk Master MID
const desaController = require("../controllers/desaController");
const mainMidController = require("../controllers/mainMidController");

const storage = multer.memoryStorage(); // Simpan file dalam memori sebagai buffer
// Setup multer middleware untuk upload
const upload = multer({ storage: storage });

router.post('/upload-master-mid', upload.single('file'), midController.uploadMasterMID);

router.post('/upload-mid', upload.single('file'), mainMidController.uploadMid);

router.post("/upload-desa", upload.single("file"), desaController.uploadDesa);

