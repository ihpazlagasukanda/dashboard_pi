const ExcelJS = require("exceljs");
const db = require("../config/db");
const util = require("util");

const query = util.promisify(db.query).bind(db); // Agar bisa pakai `await`

exports.uploadDesa = async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ message: "File tidak ditemukan" });
        }

        const workbook = new ExcelJS.Workbook();

        // 🔥 Cek apakah pakai `multer.memoryStorage()` atau `diskStorage`
        if (req.file.buffer) {
            await workbook.xlsx.load(req.file.buffer); // Pakai buffer jika `memoryStorage`
        } else {
            await workbook.xlsx.readFile(req.file.path); // Pakai file path jika `diskStorage`
        }

        const sheet = workbook.worksheets[0];
        if (!sheet) {
            return res.status(400).json({ message: "Sheet tidak ditemukan dalam file" });
        }

        let desaData = [];

        // 🔹 Loop dari baris ke-2 (skip header)
        sheet.eachRow((row, rowNumber) => {
            if (rowNumber === 1) return;

            const provinsi = row.getCell(1).text?.trim() || null;
            const kabupaten = row.getCell(2).text?.trim() || null;
            const kecamatan = row.getCell(3).text?.trim() || null;
            const kodeDesa = row.getCell(4).text?.trim() || null;
            const kelurahan = row.getCell(5).text?.trim() || null;
            const idPoktan = row.getCell(6).text?.trim() || null;
            const namaPoktan = row.getCell(7).text?.trim() || null;
            const namaKios = row.getCell(8).text?.trim() || null;
            const pihcKode = row.getCell(9).text?.trim() || null;

            if (kodeDesa) { // Pastikan `kode_desa` tidak kosong
                desaData.push([
                    provinsi, kabupaten, kecamatan, kodeDesa,
                    kelurahan, idPoktan, namaPoktan, namaKios, pihcKode
                ]);
            }
        });

        if (desaData.length === 0) {
            return res.status(400).json({ message: "Data tidak ditemukan dalam file" });
        }

        // 🔹 Simpan ke database
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

        await query(sql, [desaData]);

        res.json({ message: "✅ Data Desa berhasil diupload", data: desaData });

    } catch (error) {
        console.error("❌ Error upload desa: ", error);
        res.status(500).json({ message: "Terjadi kesalahan saat upload desa", error });
    }
};
