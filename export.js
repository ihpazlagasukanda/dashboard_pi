const fs = require('fs');
const path = require('path');
const Excel = require('exceljs');
const pool = require("./config/db");

async function exportExcel() {
    const connection = await pool.getConnection();

    const query = `
    SELECT
  e.nik,
  e.nama_petani,

  -- UREA per bulan
  SUM(CASE WHEN MONTH(v.tanggal_tebus) = 1 THEN COALESCE(v.urea, 0) ELSE 0 END) AS urea_jan,
  SUM(CASE WHEN MONTH(v.tanggal_tebus) = 2 THEN COALESCE(v.urea, 0) ELSE 0 END) AS urea_feb,
  SUM(CASE WHEN MONTH(v.tanggal_tebus) = 3 THEN COALESCE(v.urea, 0) ELSE 0 END) AS urea_mar,
  SUM(CASE WHEN MONTH(v.tanggal_tebus) = 4 THEN COALESCE(v.urea, 0) ELSE 0 END) AS urea_apr,
  SUM(CASE WHEN MONTH(v.tanggal_tebus) = 5 THEN COALESCE(v.urea, 0) ELSE 0 END) AS urea_mei,
  SUM(CASE WHEN MONTH(v.tanggal_tebus) = 6 THEN COALESCE(v.urea, 0) ELSE 0 END) AS urea_jun,
  SUM(CASE WHEN MONTH(v.tanggal_tebus) = 7 THEN COALESCE(v.urea, 0) ELSE 0 END) AS urea_jul,
  SUM(CASE WHEN MONTH(v.tanggal_tebus) = 8 THEN COALESCE(v.urea, 0) ELSE 0 END) AS urea_agu,
  SUM(CASE WHEN MONTH(v.tanggal_tebus) = 9 THEN COALESCE(v.urea, 0) ELSE 0 END) AS urea_sep,
  SUM(CASE WHEN MONTH(v.tanggal_tebus) = 10 THEN COALESCE(v.urea, 0) ELSE 0 END) AS urea_okt,
  SUM(CASE WHEN MONTH(v.tanggal_tebus) = 11 THEN COALESCE(v.urea, 0) ELSE 0 END) AS urea_nov,
  SUM(CASE WHEN MONTH(v.tanggal_tebus) = 12 THEN COALESCE(v.urea, 0) ELSE 0 END) AS urea_des,
  SUM(COALESCE(v.urea, 0)) AS total_urea,

  -- NPK per bulan
  SUM(CASE WHEN MONTH(v.tanggal_tebus) = 1 THEN COALESCE(v.npk, 0) ELSE 0 END) AS npk_jan,
  SUM(CASE WHEN MONTH(v.tanggal_tebus) = 2 THEN COALESCE(v.npk, 0) ELSE 0 END) AS npk_feb,
  SUM(CASE WHEN MONTH(v.tanggal_tebus) = 3 THEN COALESCE(v.npk, 0) ELSE 0 END) AS npk_mar,
  SUM(CASE WHEN MONTH(v.tanggal_tebus) = 4 THEN COALESCE(v.npk, 0) ELSE 0 END) AS npk_apr,
  SUM(CASE WHEN MONTH(v.tanggal_tebus) = 5 THEN COALESCE(v.npk, 0) ELSE 0 END) AS npk_mei,
  SUM(CASE WHEN MONTH(v.tanggal_tebus) = 6 THEN COALESCE(v.npk, 0) ELSE 0 END) AS npk_jun,
  SUM(CASE WHEN MONTH(v.tanggal_tebus) = 7 THEN COALESCE(v.npk, 0) ELSE 0 END) AS npk_jul,
  SUM(CASE WHEN MONTH(v.tanggal_tebus) = 8 THEN COALESCE(v.npk, 0) ELSE 0 END) AS npk_agu,
  SUM(CASE WHEN MONTH(v.tanggal_tebus) = 9 THEN COALESCE(v.npk, 0) ELSE 0 END) AS npk_sep,
  SUM(CASE WHEN MONTH(v.tanggal_tebus) = 10 THEN COALESCE(v.npk, 0) ELSE 0 END) AS npk_okt,
  SUM(CASE WHEN MONTH(v.tanggal_tebus) = 11 THEN COALESCE(v.npk, 0) ELSE 0 END) AS npk_nov,
  SUM(CASE WHEN MONTH(v.tanggal_tebus) = 12 THEN COALESCE(v.npk, 0) ELSE 0 END) AS npk_des,
  SUM(COALESCE(v.npk, 0)) AS total_npk,

  -- NPK per bulan
  SUM(CASE WHEN MONTH(v.tanggal_tebus) = 1 THEN COALESCE(v.npk_formula, 0) ELSE 0 END) AS npk_formula_jan,
  SUM(CASE WHEN MONTH(v.tanggal_tebus) = 2 THEN COALESCE(v.npk_formula, 0) ELSE 0 END) AS npk_formula_feb,
  SUM(CASE WHEN MONTH(v.tanggal_tebus) = 3 THEN COALESCE(v.npk_formula, 0) ELSE 0 END) AS npk_formula_mar,
  SUM(CASE WHEN MONTH(v.tanggal_tebus) = 4 THEN COALESCE(v.npk_formula, 0) ELSE 0 END) AS npk_formula_apr,
  SUM(CASE WHEN MONTH(v.tanggal_tebus) = 5 THEN COALESCE(v.npk_formula, 0) ELSE 0 END) AS npk_formula_mei,
  SUM(CASE WHEN MONTH(v.tanggal_tebus) = 6 THEN COALESCE(v.npk_formula, 0) ELSE 0 END) AS npk_formula_jun,
  SUM(CASE WHEN MONTH(v.tanggal_tebus) = 7 THEN COALESCE(v.npk_formula, 0) ELSE 0 END) AS npk_formula_jul,
  SUM(CASE WHEN MONTH(v.tanggal_tebus) = 8 THEN COALESCE(v.npk_formula, 0) ELSE 0 END) AS npk_formula_agu,
  SUM(CASE WHEN MONTH(v.tanggal_tebus) = 9 THEN COALESCE(v.npk_formula, 0) ELSE 0 END) AS npk_formula_sep,
  SUM(CASE WHEN MONTH(v.tanggal_tebus) = 10 THEN COALESCE(v.npk_formula, 0) ELSE 0 END) AS npk_formula_okt,
  SUM(CASE WHEN MONTH(v.tanggal_tebus) = 11 THEN COALESCE(v.npk_formula, 0) ELSE 0 END) AS npk_formula_nov,
  SUM(CASE WHEN MONTH(v.tanggal_tebus) = 12 THEN COALESCE(v.npk_formula, 0) ELSE 0 END) AS npk_formula_des,
  SUM(COALESCE(v.npk_formula, 0)) AS total_npk_formula,


  -- oRGANIK per bulan
  SUM(CASE WHEN MONTH(v.tanggal_tebus) = 1 THEN COALESCE(v.organik, 0) ELSE 0 END) AS organik_jan,
  SUM(CASE WHEN MONTH(v.tanggal_tebus) = 2 THEN COALESCE(v.organik, 0) ELSE 0 END) AS organik_feb,
  SUM(CASE WHEN MONTH(v.tanggal_tebus) = 3 THEN COALESCE(v.organik, 0) ELSE 0 END) AS organik_mar,
  SUM(CASE WHEN MONTH(v.tanggal_tebus) = 4 THEN COALESCE(v.organik, 0) ELSE 0 END) AS organik_apr,
  SUM(CASE WHEN MONTH(v.tanggal_tebus) = 5 THEN COALESCE(v.organik, 0) ELSE 0 END) AS organik_mei,
  SUM(CASE WHEN MONTH(v.tanggal_tebus) = 6 THEN COALESCE(v.organik, 0) ELSE 0 END) AS organik_jun,
  SUM(CASE WHEN MONTH(v.tanggal_tebus) = 7 THEN COALESCE(v.organik, 0) ELSE 0 END) AS organik_jul,
  SUM(CASE WHEN MONTH(v.tanggal_tebus) = 8 THEN COALESCE(v.organik, 0) ELSE 0 END) AS organik_agu,
  SUM(CASE WHEN MONTH(v.tanggal_tebus) = 9 THEN COALESCE(v.organik, 0) ELSE 0 END) AS organik_sep,
  SUM(CASE WHEN MONTH(v.tanggal_tebus) = 10 THEN COALESCE(v.organik, 0) ELSE 0 END) AS organik_okt,
  SUM(CASE WHEN MONTH(v.tanggal_tebus) = 11 THEN COALESCE(v.organik, 0) ELSE 0 END) AS organik_nov,
  SUM(CASE WHEN MONTH(v.tanggal_tebus) = 12 THEN COALESCE(v.organik, 0) ELSE 0 END) AS organik_des,
  SUM(COALESCE(v.organik, 0)) AS total_organik
FROM erdkk e
LEFT JOIN verval v
  ON e.nik = v.nik
AND YEAR(v.tanggal_tebus) = 2025
  AND v.kabupaten = 'KOTA MAGELANG'
  AND v.nama_kios = 'TANI MAKMUR KOTA'

WHERE e.tahun = 2025
  AND e.kabupaten = 'KOTA MAGELANG'
  AND e.nama_kios = 'TANI MAKMUR KOTA'

GROUP BY e.nik, e.nama_petani
ORDER BY e.nama_petani;

  `;

    try {
        const [rows] = await connection.query(query);

        const workbook = new Excel.stream.xlsx.WorkbookWriter({
            filename: path.join(__dirname, 'exports', `laporan_tebusan_${Date.now()}.xlsx`)
        });

        const sheet = workbook.addWorksheet('Laporan');

        if (rows.length > 0) {
            // Tambahkan header otomatis dari field nama
            const headers = Object.keys(rows[0]).map(key => ({ header: key, key }));
            sheet.columns = headers;

            for (const row of rows) {
                sheet.addRow(row).commit();
            }
        }

        await workbook.commit();
        console.log('✅ Export selesai.');
    } catch (err) {
        console.error('❌ Gagal export:', err.message);
    } finally {
        connection.release();
    }
}

exportExcel();
