const ExcelJS = require("exceljs");
const db = require("../config/db");

// Fungsi untuk mengecek apakah data SK Bupati sudah ada
const cekDataSkBupati = async (tahun) => {
    try {
        const [rows] = await db.execute(
            "SELECT COUNT(*) AS count FROM sk_bupati WHERE tahun = ?",
            [tahun]
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

        const { tahun, forceUpload } = req.body;
        if (!tahun) {
            return res.status(400).json({ success: false, message: "Tahun harus diisi" });
        }

        // Cek data existing
        const isDataExist = await cekDataSkBupati(tahun);
        if (isDataExist && !forceUpload) {
            return res.status(200).json({
                success: false,
                message: `SK Bupati Tahun ${tahun} sudah ada. Yakin ingin mengupload ulang?`,
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
            'Provinsi', 'Kabupaten', 'Kecamatan', 'Tahun',
            'Alokasi Urea', 'Alokasi Npk', 'Alokasi Organik',
            'Alokasi Npk Formula', 'Alokasi Npk Kakao'
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

            const provinsi = row.getCell(1).value?.toString().trim() || "";
            const kabupaten = row.getCell(2).value?.toString().trim() || "";
            const kecamatan = row.getCell(3).value?.toString().trim() || "";
            const tahunCell = row.getCell(4).value;
            const tahunData = tahunCell?.result || tahunCell?.toString().trim() || "";

            const urea = parseFloat(row.getCell(5).value) || 0;
            const npk = parseFloat(row.getCell(6).value) || 0;
            const organik = parseFloat(row.getCell(7).value) || 0;
            const npkFormula = parseFloat(row.getCell(8).value) || 0;
            const npkKakao = parseFloat(row.getCell(9).value) || 0;

            // Skip jika data tidak lengkap
            if (!provinsi || !kabupaten || !tahunData) {
                skippedCount++;
                return;
            }

            const key = `${provinsi}|${kabupaten}|${tahunData}`;

            if (!skBupatiMap.has(key)) {
                skBupatiMap.set(key, {
                    provinsi,
                    kabupaten,
                    kecamatan,
                    tahun: tahunData,
                    urea,
                    npk,
                    organik,
                    npkFormula,
                    npkKakao
                });
            } else {
                // Jika duplikat, jumlahkan alokasinya
                const existing = skBupatiMap.get(key);
                existing.urea += urea;
                existing.npk += npk;
                existing.organik += organik;
                existing.npkFormula += npkFormula;
                existing.npkKakao += npkKakao;
            }
        });

        // Simpan ke database
        for (const data of skBupatiMap.values()) {
            await connection.execute(
                `INSERT INTO sk_bupati 
                (provinsi, kabupaten, kecamatan, tahun, alokasi_urea, alokasi_npk, alokasi_organik, alokasi_npk_formula, alokasi_npk_kakao) 
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`,
                [
                    data.provinsi,
                    data.kabupaten,
                    data.kecamatan,
                    data.tahun,
                    data.urea,
                    data.npk,
                    data.organik,
                    data.npkFormula,
                    data.npkKakao
                ]
            );
            insertedCount++;
        }

        await connection.commit();

        return res.status(200).json({
            success: true,
            message: `Berhasil upload SK Bupati Tahun ${tahun}.`,
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
