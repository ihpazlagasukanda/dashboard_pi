const pool = require("../config/db");
const fs = require("fs");
const path = require("path");

const jsonFilePath = path.join(__dirname, "../data/farmers.json");

// Generate file JSON dari query sisa
exports.generateFarmersJSON = async (req, res) => {
  try {
    const query = `
      SELECT 
    e.kabupaten,
    e.kecamatan,
    e.desa,
    e.poktan,
    e.nik,
    e.nama_petani,
    COALESCE(SUM(e.urea),0) - COALESCE(SUM(v.urea),0) AS sisa_urea,
    COALESCE(SUM(e.npk),0) - COALESCE(SUM(v.npk),0) AS sisa_npk,
    COALESCE(SUM(e.npk_formula),0) - COALESCE(SUM(v.npk_formula),0) AS sisa_npk_formula,
    COALESCE(SUM(e.organik),0) - COALESCE(SUM(v.organik),0) AS sisa_organik,
    (
        (COALESCE(SUM(e.urea),0) - COALESCE(SUM(v.urea),0)) +
        (COALESCE(SUM(e.npk),0) - COALESCE(SUM(v.npk),0)) +
        (COALESCE(SUM(e.npk_formula),0) - COALESCE(SUM(v.npk_formula),0)) +
        (COALESCE(SUM(e.organik),0) - COALESCE(SUM(v.organik),0))
    ) AS total_sisa
FROM erdkk e
LEFT JOIN verval v 
    ON e.nik = v.nik
   AND YEAR(v.tanggal_tebus) = 2025
WHERE e.tahun = 2025 
  AND e.kabupaten IN 
      ('BOYOLALI', 'KLATEN', 'SUKOHARJO', 'KARANGANYAR', 'WONOGIRI', 'SRAGEN', 
       'KOTA SURAKARTA', 'SLEMAN', 'BANTUL', 'GUNUNG KIDUL', 'KULON PROGO', 'KOTA YOGYAKARTA')
GROUP BY 
    e.kabupaten,
    e.kecamatan,
    e.desa,
    e.poktan,
    e.nik,
    e.nama_petani
ORDER BY 
    e.kabupaten, e.kecamatan, e.desa, e.poktan, e.nama_petani;
    `;

    const [rows] = await pool.query(query);

    fs.writeFileSync(jsonFilePath, JSON.stringify(rows, null, 2));

    res.json({
      message: "File JSON berhasil dibuat",
      count: rows.length,
    });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: "Gagal membuat file JSON" });
  }
};

// API untuk ambil data JSON
exports.getFarmers = (req, res) => {
  try {
    const { page = 1, limit = 20, kabupaten, nik, nama } = req.query;

    // Baca data dari file JSON
    const rawData = fs.readFileSync(jsonFilePath, "utf8");
    let data = JSON.parse(rawData);

    // Filter data
    let filtered = data;

    if (kabupaten) {
      filtered = filtered.filter(f => f.kabupaten && f.kabupaten.toLowerCase() === kabupaten.toLowerCase());
    }

    if (nik) {
      filtered = filtered.filter(f => f.nik && f.nik.toString().includes(nik));
    }

    if (nama) {
      const namaLower = nama.toLowerCase();
      filtered = filtered.filter(f => f.nama_petani && f.nama_petani.toLowerCase().includes(namaLower));
    }

    // Pagination
    const startIndex = (page - 1) * limit;
    const endIndex = startIndex + parseInt(limit);
    const paginatedData = filtered.slice(startIndex, endIndex);

    res.json({
      total: filtered.length,
      page: parseInt(page),
      limit: parseInt(limit),
      data: paginatedData
    });
  } catch (err) {
    console.error("Error in getFarmers:", err);
    res.status(500).json({ error: "Gagal membaca data JSON" });
  }
};

exports.getAllFarmersByKabupaten = (req, res) => {
  try {
    const { kabupaten } = req.query;

    const rawData = fs.readFileSync(jsonFilePath, "utf8");
    let data = JSON.parse(rawData);

    // Filter by kabupaten only
    if (kabupaten) {
      data = data.filter(f => f.kabupaten && f.kabupaten.toLowerCase() === kabupaten.toLowerCase());
    }

    res.json(data);
  } catch (err) {
    console.error("Error in getAllFarmersByKabupaten:", err);
    res.status(500).json({ error: "Gagal membaca data JSON" });
  }
};

exports.getKabupatenList = (req, res) => {
  try {
    const rawData = fs.readFileSync(jsonFilePath, "utf8");
    const data = JSON.parse(rawData);
    
    // Ekstrak daftar kabupaten unik
    const kabupatenList = [...new Set(data.map(item => item.kabupaten).filter(Boolean))].sort();
    
    res.json(kabupatenList);
  } catch (err) {
    console.error("Error in getKabupatenList:", err);
    res.status(500).json({ error: "Gagal membaca data JSON" });
  }
};
