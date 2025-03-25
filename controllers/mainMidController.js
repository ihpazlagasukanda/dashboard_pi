const ExcelJS = require('exceljs');
const fs = require('fs');
const db = require('../config/db'); // Pastikan path benar
const util = require('util');

const query = util.promisify(db.query).bind(db); // Konversi ke async/await

exports.uploadMid = async (req, res) => {
    try {
        const file = req.file;
        if (!file) {
            return res.status(400).send({ message: 'No file uploaded' });
        }

        const workbook = new ExcelJS.Workbook();

        // 🔥 Cek apakah pakai memoryStorage atau diskStorage
        if (file.buffer) {
            await workbook.xlsx.load(file.buffer); // Pakai buffer jika pakai memoryStorage
        } else {
            await workbook.xlsx.readFile(file.path); // Pakai file path jika pakai diskStorage
        }

        const worksheet = workbook.worksheets[0];

        let data = [];
        worksheet.eachRow((row, rowNumber) => {
            if (rowNumber > 1) { // Skip header
                const kodeKios = row.getCell(2).text.trim(); // Kode Kios dari kolom ke-2
                let mid = row.getCell(1).text.trim(); // MID dari kolom ke-1

                if (!mid) mid = null;

                if (kodeKios) {
                    data.push([mid, kodeKios]);
                } else {
                    console.log(`❌ Data tidak valid: ${kodeKios}, ${mid}`);
                }
            }
        });

        if (data.length > 0) {
            const queryText = `INSERT INTO master_mid (mid, kode_kios) VALUES ?`;

            await query(queryText, [data]); // Pakai `await` agar query dieksekusi
            console.log(`✅ ${data.length} baris data berhasil dimasukkan ke database.`);
        } else {
            console.log('⚠ Tidak ada data yang valid untuk dimasukkan.');
        }

        // Hapus file jika pakai diskStorage
        if (file.path) fs.unlinkSync(file.path);

        res.status(200).send({ message: 'Master MID data uploaded successfully' });
    } catch (error) {
        console.error('❌ Error uploading file: ', error);
        res.status(500).send({ message: 'Error uploading file', error });
    }
};
