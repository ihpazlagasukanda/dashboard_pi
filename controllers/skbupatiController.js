const ExcelJS = require("exceljs");
const db = require("../config/db");

// Fungsi untuk mengecek apakah data SK Bupati sudah ada
const cekDataSkBupati = async (tahun, produk) => {
    try {
        const [rows] = await db.execute(
            "SELECT COUNT(*) AS count FROM sk_bupati WHERE tahun = ? AND produk = ?",
            [tahun, produk]
        );
        return (rows[0]?.count || 0) > 0;
    } catch (error) {
        console.error("❌ Error saat mengecek data SK Bupati:", error);
        throw error;
    }
};

exports.uploadSkBupati = async (req, res) => {
    let insertedCount = 0, skippedCount = 0, duplicateCount = 0;
    let connection;

    try {
        // Validasi file
        if (!req.file) {
            return res.status(400).json({ success: false, message: "File tidak ditemukan" });
        }

        const { tahun, produk, forceUpload } = req.body;
        if (!tahun || !produk) {
            return res.status(400).json({ success: false, message: "Tahun dan Produk harus diisi" });
        }

        // Cek data existing
        const isDataExist = await cekDataSkBupati(tahun, produk);
        if (isDataExist && !forceUpload) {
            return res.status(200).json({
                success: false,
                message: `SK Bupati Tahun ${tahun} ${produk} sudah ada. Yakin ingin mengupload ulang?`,
                confirmRequired: true
            });
        }

        // Load Excel
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(req.file.buffer);
        const sheet = workbook.worksheets[0];

        if (!sheet) {
            return res.status(400).json({ success: false, message: "Sheet tidak ditemukan dalam file" });
        }

        // Validasi header
        const expectedColumns = [
            'KABUPATEN', 'ID DISTRIBUTOR', 'DISTRIBUTOR', 'KODE KEC',
            'KECAMATAN', 'TOTAL'
        ];

        const rowValues = sheet.getRow(1).values;
        if (rowValues[0] === undefined) rowValues.shift();
        const actualColumns = rowValues.map(cell => (cell ? cell.toString().trim() : ""));

        if (actualColumns.length !== expectedColumns.length) {
            return res.status(400).json({ success: false, message: "Jumlah kolom tidak sesuai dengan struktur yang diharapkan" });
        }

        for (let i = 0; i < expectedColumns.length; i++) {
            if (expectedColumns[i] !== actualColumns[i]) {
                return res.status(400).json({
                    success: false,
                    message: `Kolom ke-${i + 1} tidak sesuai. Harus: "${expectedColumns[i]}", ditemukan: "${actualColumns[i]}"`
                });
            }
        }

        // Mulai transaksi database
        connection = await db.getConnection();
        await connection.beginTransaction();

        const skBupatiMap = new Map();

        // Loop data mulai dari baris ke-2
        sheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
            if (rowNumber === 1) return; // Skip header

            const kabupaten = row.getCell(1).value?.toString().trim() || "";
            const kodeDistributor = row.getCell(2).value?.toString().trim() || "";
            const distributor = row.getCell(3).value?.toString().trim() || "";
            const kodeKecamatan = row.getCell(4).value?.toString().trim() || "";
            const kecamatan = row.getCell(5).value?.toString().trim() || "";
            const alokasi = (parseFloat(row.getCell(6).value) || 0) * 1000; // Konversi ton ke kg


            // Skip jika data tidak lengkap
            if (!kabupaten || !kodeDistributor || !kodeKecamatan) {
                skippedCount++;
                return;
            }

            const key = `${kabupaten}|${kodeDistributor}|${kodeKecamatan}`;

            if (!skBupatiMap.has(key)) {
                skBupatiMap.set(key, {
                    tahun,
                    kabupaten,
                    kodeDistributor,
                    distributor,
                    kodeKecamatan,
                    kecamatan,
                    produk,
                    alokasi
                });
            } else {
                // Jika duplikat, jumlahkan alokasinya
                const existing = skBupatiMap.get(key);
                existing.alokasi += alokasi;
            }
        });

        // Simpan ke database
        for (const data of skBupatiMap.values()) {
            await connection.execute(
                `INSERT INTO sk_bupati 
                (tahun, kabupaten, kode_distributor, distributor, kode_kecamatan, kecamatan, produk, alokasi) 
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)`,
                [
                    data.tahun,
                    data.kabupaten,
                    data.kodeDistributor,
                    data.distributor,
                    data.kodeKecamatan,
                    data.kecamatan,
                    data.produk,
                    data.alokasi
                ]
            );
            insertedCount++;
        }

        await connection.commit();

        return res.status(200).json({
            success: true,
            message: `Berhasil upload SK Bupati Tahun ${tahun} ${produk}.`,
            inserted: insertedCount,
            skipped: skippedCount,
            duplicateMerged: skBupatiMap.size - insertedCount
        });

    } catch (error) {
        console.error("❌ Error upload SK Bupati:", error);
        if (connection) await connection.rollback();
        return res.status(500).json({ success: false, message: "Terjadi kesalahan saat upload data", error: error.message });
    } finally {
        if (connection) connection.release();
    }
};
