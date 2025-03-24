const ExcelJS = require("exceljs");
const db = require("../config/db");

exports.uploadDesa = async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ message: "File tidak ditemukan" });
        }

        // 🔹 1. Baca file Excel menggunakan ExcelJS
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(req.file.buffer);
        const sheet = workbook.worksheets[0];

        if (!sheet) {
            return res.status(400).json({ message: "Sheet tidak ditemukan dalam file" });
        }

        let desaData = [];

        // 🔹 2. Loop dari baris ke-2 (baris pertama biasanya header)
        sheet.eachRow((row, rowNumber) => {
            if (rowNumber === 1) return; // Skip header

            desaData.push([
                row.getCell(1).value, // Provinsi
                row.getCell(2).value, // Kabupaten
                row.getCell(3).value, // Kecamatan
                row.getCell(4).value, // Kode Desa
                row.getCell(5).value, // Kelurahan/Nama Desa
                row.getCell(6).value, // ID Poktan
                row.getCell(7).value, // Nama Poktan
                row.getCell(8).value, // Nama Kios
                row.getCell(9).value  // PIHC Kode
            ]);
        });

        if (desaData.length === 0) {
            return res.status(400).json({ message: "Data tidak ditemukan dalam file" });
        }

        // 🔹 3. Simpan ke database
        const sql = `
            INSERT INTO desa (provinsi, kabupaten, kecamatan, kode_desa, kelurahan, id_poktan, nama_poktan, nama_kios, pihc_kode)
            VALUES ?
            ON DUPLICATE KEY UPDATE 
            kelurahan = VALUES(kelurahan),
            id_poktan = VALUES(id_poktan),
            nama_poktan = VALUES(nama_poktan),
            nama_kios = VALUES(nama_kios),
            pihc_kode = VALUES(pihc_kode)
        `;

        await db.query(sql, [desaData]);

        res.json({ message: "Data Desa berhasil diupload", data: desaData });
    } catch (error) {
        console.error(error);
        res.status(500).json({ message: "Terjadi kesalahan saat upload desa" });
    }
};
