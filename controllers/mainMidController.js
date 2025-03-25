const ExcelJS = require('exceljs');
const db = require('../config/db'); // Pastikan path ke koneksi database benar

exports.uploadMid = async (req, res) => {
    try {
        const file = req.file;
        if (!file) {
            return res.status(400).json({ message: 'No file uploaded' });
        }

        // 📌 Gunakan buffer untuk membaca file dari memori
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(file.buffer);
        const worksheet = workbook.worksheets[0];

        if (!worksheet) {
            return res.status(400).json({ message: "Sheet tidak ditemukan dalam file" });
        }

        const data = [];

        // 🔹 Loop mulai dari baris ke-2 (skip header)
        worksheet.eachRow((row, rowNumber) => {
            if (rowNumber === 1) return; // Skip header

            let mid = row.getCell(1).text.trim(); // Ambil MID dari kolom pertama
            const kodeKios = row.getCell(2).text.trim(); // Ambil kode_kios dari kolom kedua

            // Jika MID kosong, set NULL
            if (!mid) mid = null;

            // Pastikan kode_kios tidak kosong
            if (kodeKios) {
                data.push([mid, kodeKios]);
            } else {
                console.log(`❌ Data tidak valid: ${mid}, ${kodeKios}`);
            }
        });

        if (data.length > 0) {
            const query = `
                INSERT INTO master_mid (mid, kode_kios) 
                VALUES ?
            `;
            await db.query(query, [data]);
            console.log(`✅ ${data.length} baris data berhasil dimasukkan ke database.`);
        } else {
            console.log('⚠ Tidak ada data yang valid untuk dimasukkan.');
        }

        res.status(200).json({ message: 'Master MID data uploaded successfully' });
    } catch (error) {
        console.error('❌ Error uploading file:', error);
        res.status(500).json({ message: 'Error uploading file' });
    }
};
