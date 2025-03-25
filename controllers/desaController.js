const ExcelJS = require("exceljs");
const db = require("../config/db");

exports.uploadDesa = async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ message: "File tidak ditemukan" });
        }

        // 🔹 1. Baca file Excel dari buffer
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

            const provinsi = row.getCell(1).text.trim();
            const kabupaten = row.getCell(2).text.trim();
            const kecamatan = row.getCell(3).text.trim();
            const kodeDesa = row.getCell(4).text.trim();
            const kelurahan = row.getCell(5).text.trim();
            const idPoktan = row.getCell(6).text.trim() || null; // Bisa NULL
            const namaPoktan = row.getCell(7).text.trim();
            const namaKios = row.getCell(8).text.trim();
            const pihcKode = row.getCell(9).text.trim() || null; // Bisa NULL

            // 🔹 Validasi: Jika data utama kosong, skip baris ini
            if (!provinsi || !kabupaten || !kecamatan || !kodeDesa || !kelurahan) {
                console.log(`❌ Data tidak valid di baris ${rowNumber}, dilewati.`);
                return;
            }

            desaData.push([provinsi, kabupaten, kecamatan, kodeDesa, kelurahan, idPoktan, namaPoktan, namaKios, pihcKode]);
        });

        if (desaData.length === 0) {
            return res.status(400).json({ message: "Data tidak ditemukan atau tidak valid dalam file" });
        }

        // 🔹 3. Simpan ke database dengan ON DUPLICATE KEY UPDATE
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

        console.log(`✅ ${desaData.length} baris berhasil disimpan ke database.`);
        res.json({ message: "Data Desa berhasil diupload", total: desaData.length });
    } catch (error) {
        console.error("❌ Error saat upload desa:", error);
        res.status(500).json({ message: "Terjadi kesalahan saat upload desa", error: error.message });
    }
};
