const express = require('express');
const router = express.Router();
const pool = require('../config/db'); // koneksi mysql
const ExcelJS = require('exceljs');

router.get('/export-belum-tebus', async (req, res) => {
  try {
    // Set headers first
    res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    );
    res.setHeader(
      'Content-Disposition',
      'attachment; filename=petani_belum_tebus.xlsx'
    );

    // Create workbook with streaming
    const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({
      stream: res,
      useStyles: false,
      useSharedStrings: false
    });

    const worksheet = workbook.addWorksheet('Petani Belum Tebus');

    // Define columns
    worksheet.columns = [
      { header: 'Kabupaten', key: 'kabupaten', width: 20 },
      { header: 'Kecamatan', key: 'kecamatan', width: 20 },
      { header: 'Desa', key: 'desa', width: 20 },
      { header: 'Poktan', key: 'poktan', width: 20 },
      { header: 'NIK', key: 'nik', width: 20 },
      { header: 'Nama Petani', key: 'nama_petani', width: 25 },
      { header: 'Tidak Tebus 2023', key: 'tidak_tebus_2023', width: 15 },
      { header: 'Tidak Tebus 2024', key: 'tidak_tebus_2024', width: 15 },
      { header: 'Tidak Tebus 2025', key: 'tidak_tebus_2025', width: 15 },
    ];

    // Process data in chunks
    const chunkSize = 1000;
    let offset = 0;
    let hasMoreData = true;

    while (hasMoreData) {
      const [rows] = await pool.query(`
        SELECT 
          kabupaten,
          kecamatan,
          desa,
          poktan,
          nik,
          nama_petani,
          tidak_tebus_2023,
          tidak_tebus_2024,
          tidak_tebus_2025
        FROM petani_belum_tebus WHERE kabupaten IN 
        ('BOYOLALI','KLATEN','KARANGANYAR','SUKOHARJO','SRAGEN','WONOGIRI','KOTA SURAKARTA')
        ORDER BY kabupaten, kecamatan, nik
        LIMIT ? OFFSET ?
      `, [chunkSize, offset]);

      if (rows.length === 0) {
        hasMoreData = false;
        break;
      }

      // Add rows to worksheet
      rows.forEach(row => {
        worksheet.addRow(row).commit();
      });

      offset += chunkSize;
    }

    // Finalize the workbook
    await workbook.commit();
    res.end();

  } catch (err) {
    console.error('Export error:', err);
    if (!res.headersSent) {
      res.status(500).send('Gagal export Excel');
    }
  }
});

module.exports = router;
