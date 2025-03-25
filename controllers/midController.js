const ExcelJS = require('exceljs');
const db = require('../config/db'); // Pastikan path ke koneksi database benar

exports.uploadMasterMID = async (req, res) => {
    try {
        const file = req.file;
        if (!file) {
            return res.status(400).send({ message: 'No file uploaded' });
        }

        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(file.buffer); // Baca dari buffer, bukan file path
        const worksheet = workbook.worksheets[0];

        if (!worksheet) {
            return res.status(400).send({ message: 'Invalid Excel file, no worksheet found' });
        }

        let data = [];

        worksheet.eachRow((row, rowNumber) => {
            if (rowNumber === 1) return; // Lewati header

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
        });

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
            console.log(`✅ ${data.length} baris data berhasil dimasukkan ke database.`);
            res.status(200).send({ message: 'Main MID data uploaded successfully' });
        } else {
            console.log('⚠ Tidak ada data yang valid untuk dimasukkan.');
            res.status(400).send({ message: 'No valid data found in the file' });
        }
    } catch (error) {
        console.error('❌ Error uploading file:', error);
        res.status(500).send({ message: 'Error uploading file', error: error.message });
    }
};
