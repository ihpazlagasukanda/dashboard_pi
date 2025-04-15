const ExcelJS = require("exceljs");
const db = require("../config/db");

// Fungsi untuk mengecek apakah semua nilai stok 0
const isAllStockZero = (data) => {
    return data.stok_awal === 0 &&
        data.penebusan === 0 &&
        data.penyaluran === 0 &&
        data.stok_akhir === 0;
};

// Fungsi untuk mengecek apakah data WCM sudah ada
const cekDataWcm = async (tahun, bulan) => {
    try {
        const [rows] = await db.execute(
            "SELECT COUNT(*) AS count FROM wcm WHERE tahun = ? AND bulan = ?",
            [tahun, bulan]
        );
        return (rows[0]?.count || 0) > 0;
    } catch (error) {
        console.error("❌ Error saat mengecek data WCM:", error);
        throw error;
    }
};

exports.uploadWcm = async (req, res) => {
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

        // Cek apakah data WCM untuk tahun dan bulan sudah ada
        const isDataExist = await cekDataWcm(tahun, bulan);
        if (isDataExist && !forceUpload) {
            return res.status(200).json({
                success: false,
                message: `Data WCM Tahun ${tahun} bulan ${bulan} sudah ada. Yakin ingin mengupload ulang?`,
                confirmRequired: true
            });
        }

        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(req.file.buffer);
        const sheet = workbook.worksheets[0];

        if (!sheet) {
            return res.status(400).json({ success: false, message: "Sheet tidak ditemukan dalam file" });
        }

        const expectedColumnsWcm = [
            'Produsen', 'Nomor', 'Kode Distributor', 'Nama Distributor', 'Kode Provinsi',
            'Nama Provinsi', 'Kode Kota/Kabupaten', 'Nama Kota/Kab', 'Kode Kecamatan',
            'Nama Kecamatan', 'Bulan', 'Tahun', 'Kode Pengecer', 'Nama Pengecer',
            'Produk', 'Stok Awal', 'Penebusan', 'Penyaluran', 'Stok Akhir', 'Status'
        ];

        const rowValues = sheet.getRow(1).values;
        if (rowValues[0] === undefined) rowValues.shift();
        const actualColumns = rowValues.map(cell => (cell ? cell.toString().trim() : ""));

        if (actualColumns.length !== expectedColumnsWcm.length) {
            return res.status(400).json({ success: false, message: "Jumlah kolom tidak sesuai dengan struktur yang diharapkan" });
        }

        for (let i = 0; i < expectedColumnsWcm.length; i++) {
            if (expectedColumnsWcm[i] !== actualColumns[i]) {
                return res.status(400).json({ success: false, message: `Kolom ke-${i + 1} tidak sesuai. Harus: "${expectedColumnsWcm[i]}", ditemukan: "${actualColumns[i]}"` });
            }
        }

        const WcmDataMap = new Map();
        const produkDiizinkan = ["UREA", "NPK", "ORGANIK", "NPK KAKAO"];

        // Iterasi melalui setiap baris data
        sheet.eachRow((row, rowNumber) => {
            if (rowNumber === 1) return; // Lewati baris pertama (header)

            const rawKabupaten = row.getCell(8).value || "";
            const processedKabupaten = rawKabupaten
                .replace(/^Kab\.\s*/i, "")
                .toUpperCase();

            const data = {
                produsen: row.getCell(1).value || "",
                nomor: row.getCell(2).value || "",
                kode_distributor: row.getCell(3).value || "",
                distributor: row.getCell(4).value || "",
                kode_provinsi: row.getCell(5).value || "",
                provinsi: row.getCell(6).value || "",
                kode_kabupaten: row.getCell(7).value || "",
                kabupaten: processedKabupaten,
                kode_kecamatan: row.getCell(9).value || "",
                kecamatan: row.getCell(10).value || "",
                bulan: row.getCell(11).value || 0,
                tahun: row.getCell(12).value || 0,
                kode_kios: row.getCell(13).value || "",
                nama_kios: row.getCell(14).value || "",
                produk: row.getCell(15).value || "",
                stok_awal: parseFloat(row.getCell(16).value) || 0,
                penebusan: parseFloat(row.getCell(17).value) || 0,
                penyaluran: parseFloat(row.getCell(18).value) || 0,
                stok_akhir: parseFloat(row.getCell(19).value) || 0,
                status: row.getCell(20).value || "",
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
            WcmDataMap.set(key, data);
        });

        // Konversi Map ke array
        let WcmData = Array.from(WcmDataMap.values());

        connection = await db.getConnection();
        await connection.beginTransaction();

        const batchSize = 2000;
        let insertedCount = 0;
        let duplicateCount = 0;

        for (let i = 0; i < WcmData.length; i += batchSize) {
            const batch = WcmData.slice(i, i + batchSize);
            const result = await simpanBatchWcm(connection, batch);
            insertedCount += result.affectedRows;
            duplicateCount += result.duplicateCount;
        }

        await connection.commit();

        res.status(200).json({
            success: true,
            message: "Data WCM berhasil diupload & update",
            insertedCount,
            skippedCount,
            duplicateCount
        });

    } catch (error) {
        if (connection) await connection.rollback();
        res.status(500).json({ success: false, message: "Terjadi kesalahan saat upload WCM", error: error.message });
    } finally {
        if (connection) connection.release();
    }
};

// Fungsi untuk menyimpan data WCM dalam batch
const simpanBatchWcm = async (connection, data) => {
    try {
        const values = data.map(row => [
            row.produsen, row.nomor, row.kode_distributor, row.distributor, row.kode_provinsi,
            row.provinsi, row.kode_kabupaten, row.kabupaten, row.kode_kecamatan, row.kecamatan,
            row.bulan, row.tahun, row.kode_kios, row.nama_kios, row.produk, row.stok_awal,
            row.penebusan, row.penyaluran, row.stok_akhir, row.status
        ]);

        const [result] = await connection.query(
            `INSERT INTO wcm (produsen, nomor, kode_distributor, distributor, kode_provinsi, provinsi, kode_kabupaten, kabupaten, kode_kecamatan, kecamatan, bulan, tahun, kode_kios, nama_kios, produk, stok_awal, penebusan, penyaluran, stok_akhir, status)
            VALUES ? 
            ON DUPLICATE KEY UPDATE 
            produsen = VALUES(produsen), 
            nomor = VALUES(nomor), 
            kode_distributor = VALUES(kode_distributor), 
            distributor = VALUES(distributor), 
            kode_provinsi = VALUES(kode_provinsi), 
            provinsi = VALUES(provinsi), 
            kode_kabupaten = VALUES(kode_kabupaten), 
            kabupaten = VALUES(kabupaten), 
            kode_kecamatan = VALUES(kode_kecamatan), 
            kecamatan = VALUES(kecamatan), 
            bulan = VALUES(bulan), 
            tahun = VALUES(tahun), 
            kode_kios = VALUES(kode_kios), 
            nama_kios = VALUES(nama_kios), 
            produk = VALUES(produk), 
            stok_awal = VALUES(stok_awal), 
            penebusan = VALUES(penebusan), 
            penyaluran = VALUES(penyaluran), 
            stok_akhir = VALUES(stok_akhir), 
            status = VALUES(status)`,
            [values]
        );

        return {
            affectedRows: result.affectedRows,
            duplicateCount: 0
        };
    } catch (error) {
        console.error("❌ Error saat menyimpan batch WCM:", error);
        throw error;
    }
};