const ExcelJS = require("exceljs");
const db = require("../config/db");

// Fungsi untuk mengecek apakah data Penyaluran DO sudah ada
const cekDataPenyaluranDo = async (tahun, bulan) => {
    try {
        const [rows] = await db.execute(
            `SELECT COUNT(*) AS count 
             FROM penyaluran_do 
             WHERE MONTH(tanggal_penyaluran) = ? AND YEAR(tanggal_penyaluran) = ?`,
            [tahun, bulan]
        );
        return (rows[0]?.count || 0) > 0;
    } catch (error) {
        console.error("❌ Error saat mengecek data Penyaluran DO:", error);
        throw error;
    }
};

exports.uploadPenyaluranDo = async (req, res) => {

    const { tahun, bulan, forceUpload } = req.body;
    if (!tahun || !bulan) {
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

        const expectedColumnsPenyaluranDo = [
            'Nama Produsen', 'Nomor F5/F6', 'Kode Distributor', 'Nama Distributor', 'Tipe Penyaluran', 'Nomor Order', 'Nomor SO', 'Nomor Penyaluran', 'Kode Provinsi',
            'Nama Provinsi', 'Kode Kabupaten', 'Nama Kabupaten', 'Kode Kecamatan',
            'Nama Kecamatan', 'Kode Retail', 'Nama Retail', 'Tanggal Penyaluran', 'Produk', 'QTY', 'status F5/F6', 'Status'
        ];

        const rowValues = sheet.getRow(1).values;
        if (rowValues[0] === undefined) rowValues.shift();
        const actualColumns = rowValues.map(cell => (cell ? cell.toString().trim() : ""));

        if (actualColumns.length !== expectedColumnsPenyaluranDo.length) {
            return res.status(400).json({ success: false, message: "Jumlah kolom tidak sesuai dengan struktur yang diharapkan" });
        }

        for (let i = 0; i < expectedColumnsPenyaluranDo.length; i++) {
            if (expectedColumnsPenyaluranDo[i] !== actualColumns[i]) {
                return res.status(400).json({ success: false, message: `Kolom ke-${i + 1} tidak sesuai. Harus: "${expectedColumnsPenyaluranDo[i]}", ditemukan: "${actualColumns[i]}"` });
            }
        }

        const PenyaluranDoDataMap = new Map();
        const produkDiizinkan = ["UREA", "NPK", "ORGANIK", "NPK KAKAO"];
        const tipePenyaluranDiizinkan = ["order"];

        // Iterasi melalui setiap baris data
        sheet.eachRow((row, rowNumber) => {
            if (rowNumber === 1) return; // Lewati baris pertama (header)

            const rawKabupaten = row.getCell(12).value || "";
            const processedKabupaten = rawKabupaten
                .replace(/^Kab\.\s*/i, "")
                .toUpperCase();

            // Handle date format (ExcelJS returns date as a number or Date object)
            let tanggalPenyaluran = row.getCell(17).value;
            let tanggalDate;

            if (tanggalPenyaluran instanceof Date) {
                tanggalDate = tanggalPenyaluran;
                tanggalPenyaluran = tanggalPenyaluran.toISOString().split('T')[0];
            } else if (typeof tanggalPenyaluran === 'number') {
                // Convert Excel date number to JS Date
                tanggalDate = new Date((tanggalPenyaluran - (25567 + 1)) * 86400 * 1000);
                tanggalPenyaluran = tanggalDate.toISOString().split('T')[0];
            } else if (typeof tanggalPenyaluran === 'string') {
                // Parse string date in format dd-mm-yyyy
                const parts = tanggalPenyaluran.split('-');
                if (parts.length === 3) {
                    tanggalDate = new Date(`${parts[2]}-${parts[1]}-${parts[0]}`);
                    tanggalPenyaluran = `${parts[2]}-${parts[1]}-${parts[0]}`;
                } else {
                    // Try to parse other date formats
                    tanggalDate = new Date(tanggalPenyaluran);
                    if (isNaN(tanggalDate.getTime())) {
                        tanggalDate = null;
                    } else {
                        tanggalPenyaluran = tanggalDate.toISOString().split('T')[0];
                    }
                }
            }

            const data = {
                produsen: row.getCell(1).value || "",
                nomor_ff: row.getCell(2).value || "",
                kode_distributor: row.getCell(3).value || "",
                distributor: row.getCell(4).value || "",
                tipe_penyaluran: (row.getCell(5).value || "").toLowerCase(),
                nama_order: row.getCell(6).value || "",
                nomor_so: row.getCell(7).value || "",
                nomor_penyaluran: row.getCell(8).value || "",
                kode_provinsi: row.getCell(9).value || "",
                provinsi: row.getCell(10).value || "",
                kode_kabupaten: row.getCell(11).value || "",
                kabupaten: processedKabupaten,
                kode_kecamatan: row.getCell(13).value || "",
                kecamatan: row.getCell(14).value || "",
                kode_kios: row.getCell(15).value || "",
                nama_kios: row.getCell(16).value || "",
                tanggal_penyaluran: tanggalPenyaluran,
                produk: (row.getCell(18).value || "").toUpperCase(),
                qty: parseFloat(row.getCell(19).value) || 0,
                status_ff: row.getCell(20).value || "",
                status: row.getCell(21).value || "",
            };

            // Skip jika produk tidak diizinkan
            if (!produkDiizinkan.includes(data.produk)) {
                skippedCount++;
                return;
            }

            if (!tipePenyaluranDiizinkan.includes(data.tipe_penyaluran)) {
                skippedCount++;
                return;
            }

            const key = `${data.produsen}-${data.nomor_ff}-${data.kode_distributor}-${data.kode_provinsi}-${data.kode_kabupaten}-${data.kode_kecamatan}-${data.produk}-${data.kode_kios}-${data.tanggal_penyaluran}`;
            // Jika sudah ada data dengan key yang sama, jumlahkan qty-nya
            if (PenyaluranDoDataMap.has(key)) {
                const existingData = PenyaluranDoDataMap.get(key);
                existingData.qty += data.qty; // Jumlahkan qty
                PenyaluranDoDataMap.set(key, existingData);
            } else {
                PenyaluranDoDataMap.set(key, data);
            }
        });

        // Cek apakah data Penyaluran DO untuk tahun dan bulan sudah ada
        const isDataExist = await cekDataPenyaluranDo(tahun, bulan);
        if (isDataExist && !forceUpload) {
            return res.status(200).json({
                success: false,
                message: `Data Penyaluran DO Tahun ${tahun} bulan ${bulan} sudah ada. Yakin ingin mengupload ulang?`,
                confirmRequired: true
            });
        }

        // Konversi Map ke array
        let PenyaluranDoData = Array.from(PenyaluranDoDataMap.values());

        connection = await db.getConnection();
        await connection.beginTransaction();

        const batchSize = 2000;
        insertedCount = 0;
        duplicateCount = 0;

        for (let i = 0; i < PenyaluranDoData.length; i += batchSize) {
            const batch = PenyaluranDoData.slice(i, i + batchSize);
            const result = await simpanBatchPenyaluranDo(connection, batch);
            insertedCount += result.insertedCount;
            duplicateCount += result.duplicateCount;
        }

        await connection.commit();

        res.status(200).json({
            success: true,
            message: "Data Penyaluran DO berhasil diupload & update",
            insertedCount,
            skippedCount,
            duplicateCount,
        });

    } catch (error) {
        if (connection) await connection.rollback();
        console.error("❌ Error saat upload Penyaluran DO:", error);
        res.status(500).json({ success: false, message: "Terjadi kesalahan saat upload Penyaluran DO", error: error.message });
    } finally {
        if (connection) connection.release();
    }
};

// Fungsi untuk menyimpan data Penyaluran DO dalam batch
const simpanBatchPenyaluranDo = async (connection, data) => {
    try {
        const values = data.map(row => [
            row.produsen,
            row.nomor_ff,
            row.kode_distributor,
            row.distributor,
            row.tipe_penyaluran,
            row.nama_order,
            row.nomor_so,
            row.nomor_penyaluran,
            row.kode_provinsi,
            row.provinsi,
            row.kode_kabupaten,
            row.kabupaten,
            row.kode_kecamatan,
            row.kecamatan,
            row.kode_kios,
            row.nama_kios,
            row.tanggal_penyaluran,
            row.produk,
            row.qty,
            row.status_ff,
            row.status
        ]);

        const [result] = await connection.query(
            `INSERT INTO penyaluran_do (
                produsen, nomor_ff, kode_distributor, distributor, tipe_penyaluran, nama_order,
                nomor_so, nomor_penyaluran, kode_provinsi, provinsi, kode_kabupaten, kabupaten,
                kode_kecamatan, kecamatan, kode_kios, nama_kios, tanggal_penyaluran, produk,
                qty, status_ff, status
            ) VALUES ? 
            ON DUPLICATE KEY UPDATE 
                produsen = VALUES(produsen), 
                nomor_ff = VALUES(nomor_ff), 
                kode_distributor = VALUES(kode_distributor), 
                distributor = VALUES(distributor), 
                tipe_penyaluran = VALUES(tipe_penyaluran),
                nama_order = VALUES(nama_order),
                nomor_so = VALUES(nomor_so),
                nomor_penyaluran = VALUES(nomor_penyaluran),
                kode_provinsi = VALUES(kode_provinsi),
                provinsi = VALUES(provinsi),
                kode_kabupaten = VALUES(kode_kabupaten),
                kabupaten = VALUES(kabupaten),
                kode_kecamatan = VALUES(kode_kecamatan),
                kecamatan = VALUES(kecamatan),
                kode_kios = VALUES(kode_kios),
                nama_kios = VALUES(nama_kios),
                tanggal_penyaluran = VALUES(tanggal_penyaluran),
                produk = VALUES(produk),
                qty = VALUES(qty),
                status_ff = VALUES(status_ff),
                status = VALUES(status)`,
            [values]
        );

        return {
            insertedCount: result.affectedRows,
            duplicateCount: result.warningStatus === 1 ? data.length - result.affectedRows : 0
        };
    } catch (error) {
        console.error("❌ Error saat menyimpan batch Penyaluran DO:", error);
        throw error;
    }
};