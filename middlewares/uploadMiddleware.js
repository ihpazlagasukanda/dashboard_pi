const multer = require("multer");
const path = require("path");

// Konfigurasi penyimpanan file
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, "uploads/"); // Simpan ke folder uploads/
  },
  filename: (req, file, cb) => {
    cb(null, `file_${Date.now()}${path.extname(file.originalname)}`);
  },
});

// Filter hanya menerima file Excel
const fileFilter = (req, file, cb) => {
  if (file.mimetype.includes("spreadsheetml")) {
    cb(null, true);
  } else {
    cb(new Error("Hanya file Excel (.xlsx) yang diperbolehkan"), false);
  }
};

const upload = multer({ storage, fileFilter });

module.exports = upload;
