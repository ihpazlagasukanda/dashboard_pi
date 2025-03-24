const ExcelJS = require('exceljs');
const fs = require('fs');
const db = require('../config/db'); // Pastikan path ke koneksi database benar

exports.uploadMID = async (req, res) => {
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
            const kodeKios = row.getCell(2).text.trim(); // Ambil kode_kios dari kolom pertama
            let mid = row.getCell(1).text.trim(); // Ambil MID dari kolom kedua

            // Jika MID kosong, set NULL
            if (!mid) mid = null;

            // Pastikan kode_kios tidak kosong
            if (kodeKios) {
                data.push([kodeKios, mid]);
            } else {
                console.log(`Data tidak valid: ${kodeKios}, ${mid}`);
            }
        }

        if (data.length > 0) {
            const query = `
                INSERT INTO master_mid (mid, kode_kios) 
                VALUES ? 
            `;
            await db.query(query, [data]);
            console.log(`${data.length} baris data berhasil dimasukkan ke database.`);
        } else {
            console.log('Tidak ada data yang valid untuk dimasukkan.');
        }

        fs.unlinkSync(file.path);
        res.status(200).send({ message: 'Master MID data uploaded successfully' });
    } catch (error) {
        console.error('Error uploading file: ', error);
        res.status(500).send({ message: 'Error uploading file' });
    }
};
