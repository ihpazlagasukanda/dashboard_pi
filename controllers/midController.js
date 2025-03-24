const ExcelJS = require('exceljs');
const fs = require('fs');
const db = require('../config/db'); // Pastikan path ke koneksi database benar

exports.uploadMasterMID = async (req, res) => {
    try {
        const file = req.file;
        if (!file) {
            return res.status(400).send({ message: 'No file uploaded' });
        }

        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(file.path);
        const worksheet = workbook.worksheets[0];

        // Loop mulai dari baris ke-2 untuk menghindari header
        const rows = worksheet.getRows(2, worksheet.rowCount);
        const data = [];

        for (let row of rows) {
            const kabupaten = row.getCell(1).text.trim();
            const kecamatan = row.getCell(2).text.trim();
            const kodeKios = row.getCell(3).text.trim();
            const namaKios = row.getCell(4).text.trim();
            let mid = row.getCell(5).text.trim();

            // Jika MID kosong, set NULL
            if (!mid) mid = null;

            // Pastikan kabupaten, kode_kios, dan nama_kios tidak kosong
            if (kabupaten && kecamatan && kodeKios && namaKios) {
                data.push([kabupaten, kecamatan, kodeKios, namaKios, mid]);
            } else {
                console.log(`Data tidak valid: ${kabupaten}, ${kecamatan}, ${kodeKios}, ${namaKios}, ${mid}`);
            }
        }

        if (data.length > 0) {
            const query = `
                INSERT INTO main_mid (kabupaten, kecamatan, kode_kios, nama_kios, mid) 
                VALUES ? 
                ON DUPLICATE KEY UPDATE 
                kode_kios = VALUES(kode_kios), 
                nama_kios = VALUES(nama_kios), 
                kabupaten = VALUES(kabupaten), 
                kecamatan = VALUES(kecamatan),
                mid = VALUES(mid)
            `;
            await db.query(query, [data]);
            console.log(`${data.length} baris data berhasil dimasukkan ke database.`);
        } else {
            console.log('Tidak ada data yang valid untuk dimasukkan.');
        }

        fs.unlinkSync(file.path);
        res.status(200).send({ message: 'Main MID data uploaded successfully' });
    } catch (error) {
        console.error('Error uploading file: ', error);
        res.status(500).send({ message: 'Error uploading file' });
    }
};
