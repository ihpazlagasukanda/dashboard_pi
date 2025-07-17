const ExcelJS = require("exceljs");
const db = require("../config/db");

// Fungsi untuk mengecek apakah semua nilai stok 0
const isAllStockZero = (data) => {
    return data.stok_awal === 0 &&
        data.penebusan === 0 &&
        data.penyaluran === 0 &&
        data.stok_akhir === 0;
};

// Fungsi untuk mengecek apakah data F5 sudah ada
const cekDataF5 = async (tahun, bulan) => {
    try {
        const [rows] = await db.execute(
            "SELECT COUNT(*) AS count FROM f5 WHERE tahun = ? AND bulan = ?",
            [tahun, bulan]
        );
        return (rows[0]?.count || 0) > 0;
    } catch (error) {
        console.error("❌ Error saat mengecek data F5:", error);
        throw error;
    }
};

exports.uploadF5 = async (req, res) => {
    let insertedCount = 0, skippedCount = 0, duplicateCount = 0;
    let connection;
    try {
        // Cek apakah file ada
        if (!req.file) {
            return res.status(400).json({ success: false, message: "File tidak ditemukan" });
        }

        // Cek apakah tahun dan bulan diisi
        const { tahun, bulan, forceUpload } = req.body;
        if (!tahun || !bulan) {
            return res.status(400).json({ success: false, message: "Tahun dan bulan harus diisi" });
        }

        // Cek apakah data F5 untuk tahun dan bulan sudah ada
        const isDataExist = await cekDataF5(tahun, bulan);
        if (isDataExist && !forceUpload) {
            return res.status(200).json({
                success: false,
                message: `Data F5 Tahun ${tahun} bulan ${bulan} sudah ada. Yakin ingin mengupload ulang?`,
                confirmRequired: true
            });
        }

        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(req.file.buffer);
        const sheet = workbook.worksheets[0];

        if (!sheet) {
            return res.status(400).json({ success: false, message: "Sheet tidak ditemukan dalam file" });
        }

        const expectedColumnsF5 = [
  'Kode Produsen',
  'Produsen',
  'No F5',
  'Kode Distributor',
  'Nama Distributor',
  'Tahun',
  'Bulan',
  'Kode Provinsi',
  'Provinsi',
  'Kode Kabupaten',
  'Kabupaten',
  'Status',
  'Kode Produk',
  'Produk',
  'Stok Awal',
  'Penebusan',
  'Penyaluran',
  'Stok Akhir'
];


        const rowValues = sheet.getRow(1).values;
        if (rowValues[0] === undefined) rowValues.shift();
        const actualColumns = rowValues.map(cell => (cell ? cell.toString().trim() : ""));

        if (actualColumns.length !== expectedColumnsF5.length) {
            return res.status(400).json({ success: false, message: "Jumlah kolom tidak sesuai dengan struktur yang diharapkan" });
        }

        for (let i = 0; i < expectedColumnsF5.length; i++) {
            if (expectedColumnsF5[i] !== actualColumns[i]) {
                return res.status(400).json({ success: false, message: `Kolom ke-${i + 1} tidak sesuai. Harus: "${expectedColumnsF5[i]}", ditemukan: "${actualColumns[i]}"` });
            }
        }

        const F5DataMap = new Map();
        const produkDiizinkan = ["UREA", "NPK", "ORGANIK", "NPK KAKAO", "ZA", "ORGANIK CAIR", "SP-36"];

        // Iterasi melalui setiap baris data
        sheet.eachRow((row, rowNumber) => {
            if (rowNumber === 1) return; // Lewati baris pertama (header)

            const rawKabupaten = row.getCell(11).value || "";
            const processedKabupaten = rawKabupaten
                .replace(/^Kab\.\s*/i, "")
                .toUpperCase();

            const data = {
                kode_produsen: row.getCell(1).value || "",
                produsen: row.getCell(2).value || "",
                no_f5: row.getCell(3).value || "",
                kode_distributor: row.getCell(4).value || "",
                nama_distributor: row.getCell(5).value || "",
                tahun: parseInt(row.getCell(6).value) || 0,
                bulan: parseInt(row.getCell(7).value) || 0,
                kode_provinsi: row.getCell(8).value || "",
                provinsi: row.getCell(9).value || "",
                kode_kabupaten: row.getCell(10).value || "",
                kabupaten: processedKabupaten,
                status: row.getCell(12).value || "",
                kode_produk: row.getCell(13).value || "",
                produk: row.getCell(14).value || "",
                stok_awal: parseFloat(row.getCell(15).value) || 0,
                penebusan: parseFloat(row.getCell(16).value) || 0,
                penyaluran: parseFloat(row.getCell(17).value) || 0,
                stok_akhir: parseFloat(row.getCell(18).value) || 0,
            };

            // Skip jika produk tidak diizinkan
            if (!produkDiizinkan.includes(data.produk.toUpperCase())) {
                skippedCount++;
                return;
            }

            // Skip jika semua stok 0
            if (isAllStockZero(data)) {
                skippedCount++;
                return;
            }

            const key = `${data.produsen}-${data.nomor}-${data.kode_distributor}-${data.kode_provinsi}-${data.kode_kabupaten}-${data.kode_kecamatan}-${data.produk}-${data.kode_kios}-${data.tahun}-${data.bulan}`;
            F5DataMap.set(key, data);
        });

        // Konversi Map ke array
        let F5Data = Array.from(F5DataMap.values());

        connection = await db.getConnection();
        await connection.beginTransaction();

        const batchSize = 2000;
        let insertedCount = 0;
        let duplicateCount = 0;

        for (let i = 0; i < F5Data.length; i += batchSize) {
            const batch = F5Data.slice(i, i + batchSize);
            const result = await simpanBatchF5(connection, batch);
            insertedCount += result.affectedRows;
            duplicateCount += result.duplicateCount;
        }

        await connection.commit();

        res.status(200).json({
            success: true,
            message: "Data F5 berhasil diupload & update",
            insertedCount,
            skippedCount,
            duplicateCount
        });

    } catch (error) {
        if (connection) await connection.rollback();
        res.status(500).json({ success: false, message: "Terjadi kesalahan saat upload F5", error: error.message });
    } finally {
        if (connection) connection.release();
    }
};

// Fungsi untuk menyimpan data F5 dalam batch
const simpanBatchF5 = async (connection, data) => {
    try {
        const values = data.map(row => [
            row.kode_produsen,
            row.produsen,
            row.no_f5,
            row.kode_distributor,
            row.nama_distributor,
            row.tahun,
            row.bulan,
            row.kode_provinsi,
            row.provinsi,
            row.kode_kabupaten,
            row.kabupaten,
            row.status,
            row.kode_produk,
            row.produk,
            row.stok_awal,
            row.penebusan,
            row.penyaluran,
            row.stok_akhir
        ]);


        const [result] = await connection.query(
            `INSERT INTO f5 (
                kode_produsen, produsen, no_f5, kode_distributor, distributor,
                tahun, bulan, kode_provinsi, provinsi, kode_kabupaten, kabupaten,
                status, kode_produk, produk, stok_awal, penebusan, penyaluran, stok_akhir
            )
            VALUES ? 
            ON DUPLICATE KEY UPDATE 
                kode_produsen = VALUES(kode_produsen),
                produsen = VALUES(produsen),
                no_f5 = VALUES(no_f5),
                kode_distributor = VALUES(kode_distributor),
                distributor = VALUES(distributor),
                tahun = VALUES(tahun),
                bulan = VALUES(bulan),
                kode_provinsi = VALUES(kode_provinsi),
                provinsi = VALUES(provinsi),
                kode_kabupaten = VALUES(kode_kabupaten),
                kabupaten = VALUES(kabupaten),
                status = VALUES(status),
                kode_produk = VALUES(kode_produk),
                produk = VALUES(produk),
                stok_awal = VALUES(stok_awal),
                penebusan = VALUES(penebusan),
                penyaluran = VALUES(penyaluran),
                stok_akhir = VALUES(stok_akhir)`,
            [values]
        );


        return {
            affectedRows: result.affectedRows,
            duplicateCount: 0
        };
    } catch (error) {
        console.error("❌ Error saat menyimpan batch F5:", error);
        throw error;
    }
};