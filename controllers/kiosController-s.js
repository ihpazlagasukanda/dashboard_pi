// controllers/realisasiKiosController.js
const ExcelJS = require("exceljs");
const db = require("../config/db"); // asumsi pakai mysql2/promise pool

exports.downloadRealisasiKios = async (req, res) => {
  try {
    // Query data realisasi kios 2025
    const [rows] = await db.query(`
      SELECT 
          v.kabupaten,
          v.kecamatan,
          v.kode_kios,
          COALESCE(SUM(v.tebus_urea), 0) AS total_urea,
          COALESCE(SUM(v.tebus_npk), 0) AS total_npk,
          COALESCE(SUM(v.tebus_organik), 0) AS total_organik,
          COALESCE(SUM(v.tebus_npk_formula), 0) AS total_npk_formula
      FROM verval_summary v
      WHERE v.tahun = 2025 AND v.kabupaten IN
      ('BOYOLALI', 'KLATEN', 'SUKOHARJO', 'KARANGANYAR', 'WONOGIRI', 'SRAGEN', 
       'KOTA SURAKARTA', 'SLEMAN', 'BANTUL', 'GUNUNG KIDUL', 'KULON PROGO', 'KOTA YOGYAKARTA')
      GROUP BY v.kabupaten, v.kecamatan, v.kode_kios
      ORDER BY v.kabupaten, v.kecamatan;
    `);

    // Buat workbook Excel
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Realisasi Kios 2025");

    // Tambah header
    worksheet.columns = [
      { header: "Kabupaten", key: "kabupaten", width: 15 },
      { header: "Kecamatan", key: "kecamatan", width: 15 },
      { header: "Kode Kios", key: "kode_kios", width: 15 },
      { header: "Total Urea", key: "total_urea", width: 15 },
      { header: "Total NPK", key: "total_npk", width: 15 },
      { header: "Total Organik", key: "total_organik", width: 18 },
      { header: "Total NPK Formula", key: "total_npk_formula", width: 20 },
    ];

    // Tambah data
    rows.forEach((row) => {
      worksheet.addRow(row);
    });

    // Styling header biar rapi
    worksheet.getRow(1).eachCell((cell) => {
      cell.font = { bold: true };
      cell.alignment = { vertical: "middle", horizontal: "center" };
    });

    // Kirim file sebagai response download
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader(
      "Content-Disposition",
      'attachment; filename="Realisasi_Kios_2025.xlsx"'
    );

    await workbook.xlsx.write(res);
    res.end();
  } catch (err) {
    console.error("Error generate Excel:", err);
    res.status(500).json({ message: "Gagal generate laporan" });
  }
};
