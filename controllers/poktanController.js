const ExcelJS = require("exceljs");
const db = require("../config/db");

const cekDataPoktan = async (kabupaten) => {
    try {
        const [rows] = await db.execute(
            `SELECT COUNT(*) AS count 
             FROM poktan 
             WHERE kabupaten = ?`,
            [kabupaten]
        );
        return (rows[0]?.count || 0) > 0;
    } catch (error) {
        console.error("❌ Error saat mengecek data Penyaluran DO:", error);
        throw error;
    }
};

exports.uploadPoktan = async (req, res) => {

    const { kabupaten, forceUpload } = req.body;
    if (!kabupaten) {
        return res.status(400).json({
            success: false,
            message: "Tahun dan bulan wajib diisi."
        });
    }

    let insertedCount = 0, skippedCount = 0, duplicateCount = 0;
    let connection;
    try {
        // Cek apakah file ada
        if (!req.file) {
            return res.status(400).json({ success: false, message: "File tidak ditemukan" });
        }

        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(req.file.buffer);
        const sheet = workbook.worksheets[0];

        if (!sheet) {
            return res.status(400).json({ success: false, message: "Sheet tidak ditemukan dalam file" });
        }

        const expectedColumnsPoktan = [
            'KABUPATEN', 'KECAMATAN', 'KODE DESA', 'NAMA DESA', 'KODE KIOS', 'NAMA KIOS', 'NAMA POKTAN', 'NAMA KETUA POKTAN', 'NO HP POKTAN'
        ];

        const rowValues = sheet.getRow(1).values;
        if (rowValues[0] === undefined) rowValues.shift();
        const actualColumns = rowValues.map(cell => (cell ? cell.toString().trim() : ""));

        if (actualColumns.length !== expectedColumnsPoktan.length) {
            return res.status(400).json({ success: false, message: "Jumlah kolom tidak sesuai dengan struktur yang diharapkan" });
        }

        for (let i = 0; i < expectedColumnsPoktan.length; i++) {
            if (expectedColumnsPoktan[i] !== actualColumns[i]) {
                return res.status(400).json({ success: false, message: `Kolom ke-${i + 1} tidak sesuai. Harus: "${expectedColumnsPoktan[i]}", ditemukan: "${actualColumns[i]}"` });
            }
        }

        const PoktanDataMap = new Map();
        // Iterasi melalui setiap baris data
        sheet.eachRow((row, rowNumber) => {
            if (rowNumber === 1) return; // Lewati baris pertama (header)

            const data = {
                kabupaten: row.getCell(1).value || "",
                kecamatan: row.getCell(2).value || "",
                kode_desa: row.getCell(3).value || "",
                desa: row.getCell(4).value || "",
                kode_kios: row.getCell(5).value || "",
                nama_kios: row.getCell(6).value || "",
                poktan: (row.getCell(7).value || "").toString().toUpperCase(),
                ketua_poktan: row.getCell(8).value || "",
                nomor_hp: row.getCell(9).result?.toString().trim() || row.getCell(9).text?.toString().trim() || "",

            };

            const key = `${data.kabupaten}-${data.kecamatan}-${data.kode_desa}-${data.desa}-${data.kode_kios}-${data.nama_kios}-${data.poktan}-${data.ketua_poktan}-${data.nomor_hp}`;
            // Jika sudah ada data dengan key yang sama, jumlahkan qty-nya
            if (!PoktanDataMap.has(key)) {
                PoktanDataMap.set(key, data);
            }

        });

        // Cek apakah data Penyaluran DO untuk tahun dan bulan sudah ada
        const isDataExist = await cekDataPoktan(kabupaten);
        if (isDataExist && !forceUpload) {
            return res.status(200).json({
                success: false,
                message: `Data poktan ${kabupaten} sudah ada`,
                confirmRequired: true
            });
        }

        // Konversi Map ke array
        let PoktanData = Array.from(PoktanDataMap.values());

        connection = await db.getConnection();
        await connection.beginTransaction();

        const batchSize = 2000;
        insertedCount = 0;
        duplicateCount = 0;

        for (let i = 0; i < PoktanData.length; i += batchSize) {
            const batch = PoktanData.slice(i, i + batchSize);
            const result = await simpanBatchPoktan(connection, batch);
            insertedCount += result.insertedCount;
            duplicateCount += result.duplicateCount;
        }

        await connection.commit();

        res.status(200).json({
            success: true,
            message: "Data Poktan berhasil diupload & update",
            insertedCount,
            skippedCount,
            duplicateCount,
        });

    } catch (error) {
        if (connection) await connection.rollback();
        console.error("❌ Error saat upload poktan:", error);
        res.status(500).json({ success: false, message: "Terjadi kesalahan saat upload poktan", error: error.message });
    } finally {
        if (connection) connection.release();
    }
};

// Fungsi untuk menyimpan data Penyaluran DO dalam batch
const simpanBatchPoktan = async (connection, data) => {
    try {
        const values = data.map(row => [
            row.kabupaten,
            row.kecamatan,
            row.kode_desa,
            row.desa,
            row.kode_kios,
            row.nama_kios,
            row.poktan,
            row.ketua_poktan,
            row.nomor_hp
        ]);

        const [result] = await connection.query(
            `INSERT INTO poktan (
                kabupaten, kecamatan, kode_desa, desa, kode_kios, nama_kios,
                poktan, ketua_poktan, nomor_hp
            ) VALUES ? 
            ON DUPLICATE KEY UPDATE 
                kabupaten = VALUES(kabupaten), 
                kecamatan = VALUES(kecamatan), 
                kode_desa = VALUES(kode_desa), 
                desa = VALUES(desa), 
                kode_kios = VALUES(kode_kios),
                nama_kios = VALUES(nama_kios),
                poktan = VALUES(poktan),
                ketua_poktan = VALUES(ketua_poktan),
                nomor_hp = VALUES(nomor_hp)`,
            [values]
        );

        return {
            insertedCount: result.affectedRows,
            duplicateCount: result.warningStatus === 1 ? data.length - result.affectedRows : 0
        };
    } catch (error) {
        console.error("❌ Error saat menyimpan batch:", error);
        throw error;
    }
};