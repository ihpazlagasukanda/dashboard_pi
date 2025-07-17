const ExcelJS = require("exceljs");
const db = require("../config/db");

// Fungsi untuk mengecek apakah data ERDKK sudah ada
const cekDataErdkk = async (tahun, kabupaten) => {
    try {
        const [rows] = await db.execute(
            "SELECT COUNT(*) AS count FROM erdkk WHERE tahun = ? AND kabupaten = ?",
            [tahun, kabupaten]
        );
        return (rows[0]?.count || 0) > 0; // Pastikan nilai aman
    } catch (error) {
        console.error("❌ Error saat mengecek data ERDKK:", error);
        throw error;
    }
};

exports.uploadErdkk = async (req, res) => {
    let insertedCount = 0, skippedCount = 0, duplicateCount = 0;
    let connection;
    try {
        // Cek apakah file ada
        if (!req.file) {
            return res.status(400).json({ success: false, message: "File tidak ditemukan" });
        }

        // Cek apakah tahun dan kabupaten diisi
        const { tahun, kabupaten, forceUpload } = req.body;
        if (!tahun || !kabupaten) {
            return res.status(400).json({ success: false, message: "Tahun dan kabupaten harus diisi" });
        }

        // Cek apakah data ERDKK untuk tahun dan kabupaten sudah ada
        const isDataExist = await cekDataErdkk(tahun, kabupaten);
        if (isDataExist && !forceUpload) {
            return res.status(200).json({
                success: false,
                message: `Data ERDKK Tahun ${tahun} Kabupaten ${kabupaten} sudah ada. Yakin ingin mengupload ulang?`,
                confirmRequired: true
            });
        }
        // Lanjutkan proses upload jika data belum ada atau client mengonfirmasi upload ulang
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(req.file.buffer);
        const sheet = workbook.worksheets[0];

        // Cek apakah sheet ada
        if (!sheet) {
            return res.status(400).json({ success: false, message: "Sheet tidak ditemukan dalam file" });
        }

        const expectedColumnsNonDIY = [
            'Nama Penyuluh', 'Kode Desa', 'Kode Kios Pengecer', 'Nama Kios Pengecer', 'Gapoktan', 'Nama Poktan', 'Nama Petani',
            'KTP', 'Tempat Lahir', 'Tanggal Lahir', 'Nama Ibu Kandung', 'Alamat', 'Subsektor',
            'Komoditas MT1', 'Luas Lahan (Ha) MT1', 'Pupuk Urea (Kg) MT1', 'Pupuk NPK (Kg) MT1',
            'Pupuk NPK Formula (Kg) MT1', 'Pupuk Organik (Kg) MT1', 'Komoditas MT2',
            'Luas Lahan (Ha) MT2', 'Pupuk Urea (Kg) MT2', 'Pupuk NPK (Kg) MT2',
            'Pupuk NPK Formula (Kg) MT2', 'Pupuk Organik (Kg) MT2', 'Komoditas MT3',
            'Luas Lahan (Ha) MT3', 'Pupuk Urea (Kg) MT3', 'Pupuk NPK (Kg) MT3',
            'Pupuk NPK Formula (Kg) MT3', 'Pupuk Organik (Kg) MT3'
        ];

        // Daftar kolom yang diharapkan untuk DIY dan non-DIY
        const expectedColumnsDIY = [
            'Nama Penyuluh', 'Kode Desa', 'Kode Kios Pengecer', 'Nama Kios Pengecer', 'Gapoktan', 'Nama Poktan', 'Nama Petani',
            'KTP', 'Tempat Lahir', 'Tanggal Lahir', 'Nama Ibu Kandung', 'Alamat', 'Subsektor',
            'Komoditas MT1', 'Luas Lahan (Ha) MT1', 'Pupuk Urea (Kg) MT1', 'Pupuk NPK (Kg) MT1',
            'Pupuk NPK Formula (Kg) MT1', 'Pupuk Organik (Kg) MT1', 'Komoditas MT2',
            'Luas Lahan (Ha) MT2', 'Pupuk Urea (Kg) MT2', 'Pupuk NPK (Kg) MT2',
            'Pupuk NPK Formula (Kg) MT2', 'Pupuk Organik (Kg) MT2', 'Komoditas MT3',
            'Luas Lahan (Ha) MT3', 'Pupuk Urea (Kg) MT3', 'Pupuk NPK (Kg) MT3',
            'Pupuk NPK Formula (Kg) MT3', 'Pupuk Organik (Kg) MT3'
        ];

        // Tentukan apakah kabupaten termasuk DIY atau bukan
        const isDIY = ["SLEMAN", "BANTUL", "GUNUNG KIDUL", "KULON PROGO", "KOTA YOGYAKARTA"].includes(kabupaten.toUpperCase());
        const expectedColumns = isDIY ? expectedColumnsDIY : expectedColumnsNonDIY;

        // Ambil kolom dari baris pertama
        const rowValues = sheet.getRow(1).values;
        if (rowValues[0] === undefined) rowValues.shift(); // Hapus elemen pertama jika kosong
        const actualColumns = rowValues.map(cell => {
    if (!cell) return "";
    if (typeof cell === "string") return cell.trim();
    if (typeof cell === "number") return cell.toString().trim();
    if (typeof cell === "object") {
        // Tangani beberapa kemungkinan format Excel
        if (cell.text) return cell.text.trim();
        if (cell.result) return cell.result.toString().trim();
        if (cell.richText) return cell.richText.map(r => r.text).join('').trim();
        if (cell.formula && cell.result) return cell.result.toString().trim();
        return "[Invalid Object]";
    }
    return cell.toString().trim(); // fallback
});


        // Cek apakah struktur file sesuai
        if (actualColumns.length !== expectedColumns.length) {
            return res.status(400).json({ success: false, message: "Jumlah kolom tidak sesuai dengan struktur yang diharapkan" });
        }

        for (let i = 0; i < expectedColumns.length; i++) {
            if (expectedColumns[i] !== actualColumns[i]) {
                return res.status(400).json({ success: false, message: `Kolom ke-${i + 1} tidak sesuai. Harus: "${expectedColumns[i]}", ditemukan: "${actualColumns[i]}"` });
            }
        }

        const [kecRows] = await db.execute("SELECT kode_desa, kecamatan FROM desa");
        const kecMap = Object.fromEntries(kecRows.map(row => [row.kode_desa, row.kecamatan]));

        const [desaRows] = await db.execute("SELECT kode_desa, desa FROM desa");
        const desaMap = Object.fromEntries(desaRows.map(row => [row.kode_desa, row.desa]));

        // Proses data Excel
        let erdkkDataMap = new Map(); // Untuk menggabungkan baris duplikat
        const shift = 0; // Shift kolom untuk DIY dan non-DIY

        function cleanCell(cell) {
            const val = cell?.value;
            if (val == null) return "";
            if (typeof val === "object") {
                return val.text || val.result || (val.richText?.map(r => r.text).join('')) || "[Invalid Object]";
            }
            return val.toString().trim();
        }


        for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber++) {
           const row = sheet.getRow(rowNumber);
            const kodeDesa = cleanCell(row.getCell(2));
            const namaDesa = desaMap[kodeDesa] || "Desa Tidak Ditemukan";
            const kecamatan = kecMap[kodeDesa] || "Kecamatan Tidak Ditemukan";
            const kodeKios = cleanCell(row.getCell(3));
            const namaKios = cleanCell(row.getCell(4));
            const poktan = cleanCell(row.getCell(6)).toUpperCase();
            const nik = cleanCell(row.getCell(8)).replace(/`/g, "'");
            const namaPetani = cleanCell(row.getCell(7));

            if (!kecamatan || !nik || !namaPetani) {
                skippedCount++;
                continue;
            }

            console.log('Isi nama_petani:', namaPetani);
            console.log('Panjang nama_petani:', String(namaPetani).length);


            const parseExcelNumber = (value) => {
                if (value === null || value === undefined) return 0;

                // Jika sudah number, langsung kembalikan
                if (typeof value === 'number') return value;

                // Handle string
                if (typeof value === 'string') {
                    // Bersihkan spasi dan karakter non-numeric
                    let cleaned = value.trim().replace(/[^\d,.-]/g, '');

                    // Ganti koma dengan titik untuk decimal
                    cleaned = cleaned.replace(',', '.');

                    // Parse ke float
                    const parsed = parseFloat(cleaned);

                    return isNaN(parsed) ? 0 : parsed;
                }

                return 0;
            };

            // Gunakan dalam perhitungan:
            const getCellValue = (cell) => {
                return parseExcelNumber(row.getCell(cell).value);
            };


            // Hitung total pupuk
            const totalUrea = getCellValue(16 + shift) + getCellValue(22 + shift) + getCellValue(28 + shift);
            const totalNpk = getCellValue(17 + shift) + getCellValue(23 + shift) + getCellValue(29 + shift);
            const totalNpkFormula = getCellValue(18 + shift) + getCellValue(24 + shift) + getCellValue(30 + shift);
            const totalOrganik = getCellValue(19 + shift) + getCellValue(25 + shift) + getCellValue(31 + shift);


            // Gabungkan baris duplikat di Excel
            const key = `${kabupaten}-${nik}-${kodeKios}-${tahun}`;

            if (erdkkDataMap.has(key)) {
                const existingData = erdkkDataMap.get(key);

                existingData[8] += totalUrea;
                existingData[9] += totalNpk;
                existingData[10] += totalNpkFormula;
                existingData[11] += totalOrganik;
            } else {
                erdkkDataMap.set(key, [
                    tahun, kabupaten, kecamatan, namaDesa, poktan, kodeKios, namaKios, nik, namaPetani,
                    totalUrea, totalNpk, totalNpkFormula, totalOrganik
                ]);
            }
        }

        // Konversi Map ke array
        let erdkkData = Array.from(erdkkDataMap.values());

        // Ambil koneksi dari pool dan mulai transaksi
        connection = await db.getConnection();
        await connection.beginTransaction();

        const batchSize = 2000; // Sesuaikan kapasitas server
        let insertedCount = 0;
        let duplicateCount = 0;

        for (let i = 0; i < erdkkData.length; i += batchSize) {
            const batch = erdkkData.slice(i, i + batchSize);
            const result = await simpanBatch(connection, batch);
            insertedCount += result.affectedRows;
            duplicateCount += result.duplicateCount;
        }

        await connection.commit(); // Simpan transaksi ke database

        res.status(200).json({
            success: true,
            message: "Data ERDKK berhasil diupload & update",
            insertedCount,
            duplicateCount
        });

    } catch (error) {
        if (connection) await connection.rollback(); // Rollback jika ada error
        res.status(500).json({ success: false, message: "Terjadi kesalahan saat upload ERDKK", error: error.message });
    } finally {
        if (connection) connection.release(); // Lepaskan koneksi ke pool
    }
};

// Fungsi untuk menyimpan data dalam batch
const simpanBatch = async (connection, data) => {
    try {
        // Siapkan nilai untuk query INSERT
        const values = data.map(row => [
            row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7],
            row[8], row[9], row[10], row[11], row[12]
        ]);

        // Masukkan data baru dengan ON DUPLICATE KEY UPDATE
        const [result] = await connection.query(
            `INSERT INTO erdkk (tahun, kabupaten, kecamatan, desa, poktan, kode_kios, nama_kios, nik, nama_petani, urea, npk, npk_formula, organik)
            VALUES ? 
            ON DUPLICATE KEY UPDATE 
            urea = VALUES(urea), 
            npk = VALUES(npk), 
            npk_formula = VALUES(npk_formula), 
            organik = VALUES(organik)`,
            [values]
        );

        return {
            affectedRows: result.affectedRows,
            duplicateCount: 0 // Duplikat sudah ditangani ON DUPLICATE KEY UPDATE
        };
    } catch (error) {
        console.error("❌ Error saat menyimpan batch:", error);
        throw error;
    }
};
