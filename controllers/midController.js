const ExcelJS = require('exceljs');
const fs = require('fs');
const db = require('../config/db'); // Pastikan path benar
const util = require('util');

const query = util.promisify(db.query).bind(db); // Konversi ke async/await

exports.uploadMasterMID = async (req, res) => {
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
                const kabupaten = row.getCell(1).text.trim();
                const kecamatan = row.getCell(2).text.trim();
                const kodeKios = row.getCell(3).text.trim();
                const namaKios = row.getCell(4).text.trim();
                let mid = row.getCell(5).text.trim();

                if (!mid) mid = null;

                if (kabupaten && kecamatan && kodeKios && namaKios) {
                    data.push([kabupaten, kecamatan, kodeKios, namaKios, mid]);
                } else {
                    console.log(`❌ Data tidak valid: ${kabupaten}, ${kecamatan}, ${kodeKios}, ${namaKios}, ${mid}`);
                }
            }
        });

        if (data.length > 0) {
            const queryText = `
                INSERT INTO main_mid (kabupaten, kecamatan, kode_kios, nama_kios, mid) 
                VALUES ? 
                ON DUPLICATE KEY UPDATE 
                kode_kios = VALUES(kode_kios), 
                nama_kios = VALUES(nama_kios), 
                kabupaten = VALUES(kabupaten), 
                kecamatan = VALUES(kecamatan),
                mid = VALUES(mid)
            `;

            await query(queryText, [data]); // Pakai `await` agar query dieksekusi
            console.log(`✅ ${data.length} baris data berhasil dimasukkan ke database.`);
        } else {
            console.log('⚠ Tidak ada data yang valid untuk dimasukkan.');
        }

        // Hapus file jika pakai diskStorage
        if (file.path) fs.unlinkSync(file.path);

        res.status(200).send({ message: 'Main MID data uploaded successfully' });
    } catch (error) {
        console.error('❌ Error uploading file: ', error);
        res.status(500).send({ message: 'Error uploading file', error });
    }
};
