const pool = require("../config/db");
const fs = require("fs");
const path = require("path");

const jsonFilePath = path.join(__dirname, "../data/farmers.json");

// Generate file JSON dari query sisa
exports.generateFarmersJSON = async (req, res) => {
  try {
    const query = `
    SELECT 
        COALESCE(e.kabupaten, v.kabupaten) AS kabupaten, 
        COALESCE(e.kecamatan, v.kecamatan) AS kecamatan,
        COALESCE(e.nik, v.nik) AS nik, 
        COALESCE(e.nama_petani, '') AS nama_petani, 
        COALESCE(e.kode_kios, v.kode_kios) AS kode_kios,
        COALESCE(e.desa, '') AS desa,
        COALESCE(e.poktan, '') AS poktan,
        COALESCE(e.tahun, v.tahun) AS tahun, 

        (COALESCE(e.urea,0) - COALESCE(v.tebus_urea,0)) AS sisa_urea,
        (COALESCE(e.npk,0) - COALESCE(v.tebus_npk,0)) AS sisa_npk,
        (COALESCE(e.npk_formula,0) - COALESCE(v.tebus_npk_formula,0)) AS sisa_npk_formula,
        (COALESCE(e.organik,0) - COALESCE(v.tebus_organik,0)) AS sisa_organik,
        (
            (COALESCE(e.urea,0) - COALESCE(v.tebus_urea,0)) +
            (COALESCE(e.npk,0) - COALESCE(v.tebus_npk,0)) +
            (COALESCE(e.npk_formula,0) - COALESCE(v.tebus_npk_formula,0)) +
            (COALESCE(e.organik,0) - COALESCE(v.tebus_organik,0))
        ) AS total_sisa

    FROM (
        -- Data dari ERDKK
        SELECT 
            e.kabupaten, e.kecamatan, e.nik, e.nama_petani, e.kode_kios,
            e.tahun, e.desa, e.poktan,
            SUM(e.urea) AS urea, 
            SUM(e.npk) AS npk, 
            SUM(e.npk_formula) AS npk_formula, 
            SUM(e.organik) AS organik
        FROM erdkk e
        WHERE e.tahun = 2025
          AND e.kabupaten IN ('BOYOLALI', 'KLATEN', 'SUKOHARJO', 'KARANGANYAR', 'WONOGIRI', 'SRAGEN', 
       'KOTA SURAKARTA', 'SLEMAN', 'BANTUL', 'GUNUNG KIDUL', 'KULON PROGO', 'KOTA YOGYAKARTA')
        GROUP BY e.kabupaten, e.kecamatan, e.nik, e.nama_petani, e.kode_kios, e.tahun, e.desa, e.poktan
        
        UNION ALL
        
        -- Data dari VERVAL yang tidak ada di ERDKK (alokasi = 0)
        SELECT 
            v.kabupaten, v.kecamatan, v.nik, '' AS nama_petani, v.kode_kios,
            v.tahun, '' AS desa, '' AS poktan,
            0 AS urea, 0 AS npk, 0 AS npk_formula, 0 AS organik
        FROM verval_summary v
        LEFT JOIN erdkk e 
            ON e.nik = v.nik
           AND e.kabupaten = v.kabupaten
           AND e.tahun = v.tahun
           AND e.kode_kios = v.kode_kios
        WHERE v.tahun = 2025
          AND v.kabupaten IN ('BOYOLALI', 'KLATEN', 'SUKOHARJO', 'KARANGANYAR', 'WONOGIRI', 'SRAGEN', 
       'KOTA SURAKARTA', 'SLEMAN', 'BANTUL', 'GUNUNG KIDUL', 'KULON PROGO', 'KOTA YOGYAKARTA')
          AND e.nik IS NULL
     ) AS combined

    LEFT JOIN verval_summary v 
        ON combined.nik = v.nik
       AND combined.kabupaten = v.kabupaten
       AND combined.tahun = v.tahun
       AND combined.kode_kios = v.kode_kios

    LEFT JOIN erdkk e 
        ON combined.nik = e.nik
       AND combined.kabupaten = e.kabupaten
       AND combined.tahun = e.tahun
       AND combined.kode_kios = e.kode_kios

    WHERE combined.tahun = 2025
      AND combined.kabupaten IN ('BOYOLALI', 'KLATEN', 'SUKOHARJO', 'KARANGANYAR', 'WONOGIRI', 'SRAGEN', 
       'KOTA SURAKARTA', 'SLEMAN', 'BANTUL', 'GUNUNG KIDUL', 'KULON PROGO', 'KOTA YOGYAKARTA');
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
