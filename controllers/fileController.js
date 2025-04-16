const ExcelJS = require('exceljs');
const fs = require('fs');
const db = require('../config/db');
const { v4: uuidv4 } = require('uuid');

// Fungsi untuk konversi nilai menjadi angka
const convertToNumber = (value) => {
    if (typeof value === 'string') {
        value = value.replace(',', '.'); // Ubah koma ke titik untuk desimal
    }
    const numberValue = parseFloat(value);
    return isNaN(numberValue) ? 0 : numberValue;
};

// Fungsi untuk eksekusi query dengan retry mechanism
async function executeWithRetry(fn, maxRetries = 3, baseDelay = 1000) {
    let attempt = 1;
    while (attempt <= maxRetries) {
        try {
            return await fn();
        } catch (err) {
            if (err.code === 'ER_LOCK_WAIT_TIMEOUT' && attempt < maxRetries) {
                const delay = baseDelay * Math.pow(2, attempt - 1);
                console.warn(`⚠️ Lock timeout detected, retrying in ${delay}ms (attempt ${attempt}/${maxRetries})`);
                await new Promise(resolve => setTimeout(resolve, delay));
                attempt++;
                continue;
            }
            throw err;
        }
    }
}

// Fungsi untuk mendapatkan koneksi dari pool dengan timeout
async function getConnectionWithTimeout(timeout = 10000) {
    return new Promise((resolve, reject) => {
        let timer = setTimeout(() => {
            reject(new Error('Connection timeout'));
        }, timeout);

        db.getConnection((err, connection) => {
            clearTimeout(timer);
            if (err) {
                reject(err);
            } else {
                resolve(connection);
            }
        });
    });
}

// Fungsi untuk membersihkan teks
const cleanText = (text) => {
    return text ? text.toLowerCase().trim().replace(/[^a-z0-9\s]/g, '') : '';
};

// Fungsi untuk membersihkan nama kios
const cleanNamaKios = (text) => {
    if (!text) return "";

    let cleaned = text.toUpperCase();
    cleaned = cleaned.replace(/([A-Z])([A-Z][a-z])/g, '$1 $2');
    cleaned = cleaned.replace(/[.,\"-\/]/g, "");

    const wordsToRemove = ["KPL", "KIOS", "UD", "KUD", "TOKO", "TK", "GAPOKTAN", "CV"];
    let regex = new RegExp(`\\b(${wordsToRemove.join("|")})\\b`, "g");
    cleaned = cleaned.replace(regex, "").trim();

    return cleaned;
};

// Fungsi untuk memproses data Excel
async function processExcelFile(file, metodePenebusan) {
    console.log('📝 Memproses file:', file.originalname);

    const workbook = new ExcelJS.Workbook();
    let buffer;

    try {
        buffer = file.buffer || fs.readFileSync(file.path);
        await workbook.xlsx.load(buffer);
        console.log('✅ File berhasil dibaca!');
    } catch (err) {
        console.error('❌ Gagal membaca file Excel:', err);
        throw new Error('File Excel tidak valid');
    }

    const worksheet = workbook.worksheets[0];
    console.log(`📜 Sheet: ${worksheet.name}`);

    // Validasi kolom
    const expectedColumnsKartan = ['NO', 'KABUPATEN', 'KECAMATAN', 'KODE KIOS', 'NAMA KIOS', 'NIK', 'NAMA PETANI', 'UREA', 'NPK', 'SP36', 'ZA', 'NPK FORMULA', 'ORGANIK', 'ORGANIK CAIR', 'TGL TEBUS', 'TGL INPUT', 'STATUS'];
    const expectedColumnsIpubers = ['No', 'Kabupaten', 'Kecamatan', 'Kode Kios', 'Nama Kios', 'Kode TRX', 'No Transaksi', 'NIK', 'Nama Petani', 'Urea', 'NPK', 'SP36', 'ZA', 'NPK Formula', 'Organik', 'Organik Cair', 'Keterangan', 'Tanggal Tebus', 'Tanggal Entri', 'Tanggal Update', 'Tipe Tebus', 'NIK Perwakilan', 'Url Bukti', 'Status'];

    const headerRow = 1;
    const rowValues = worksheet.getRow(headerRow).values;

    if (rowValues[0] === undefined) {
        rowValues.shift();
    }

    const actualColumns = rowValues.map(cell => cell ? cell.toString().trim() : "");
    const expectedColumns = metodePenebusan === 'ipubers' ? expectedColumnsIpubers : expectedColumnsKartan;

    if (!expectedColumns.every((col, index) => col === actualColumns[index])) {
        console.error("❌ Struktur file tidak sesuai. Ditemukan:", actualColumns);
        throw new Error('Struktur file Excel tidak sesuai');
    }

    // Proses data
    const startRow = 2;
    let rows = worksheet.getRows(startRow, worksheet.rowCount - startRow);
    rows = rows.filter(row => row.getCell(1).value && row.getCell(2).value);

    console.log(`✅ Data siap diproses, total baris: ${rows.length}`);

    return rows.map((row) => {
        let tanggalTebus, tanggalExcel;
        const kodeTransaksi = metodePenebusan === 'ipubers' ? row.getCell(7).text : uuidv4();

        if (metodePenebusan === 'ipubers') {
            tanggalExcel = row.getCell(18).text;
        } else if (metodePenebusan === 'kartan') {
            tanggalExcel = row.getCell(15).text;
        }

        // Format tanggal
        if (tanggalExcel && (tanggalExcel instanceof Date || !isNaN(Date.parse(tanggalExcel)))) {
            const dateObj = new Date(tanggalExcel);
            const year = dateObj.getFullYear();
            const month = (dateObj.getMonth() + 1).toString().padStart(2, '0');
            const day = dateObj.getDate().toString().padStart(2, '0');
            tanggalTebus = `${year}/${month}/${day}`;
        } else if (typeof tanggalExcel === 'string' && tanggalExcel.includes('/')) {
            const parts = tanggalExcel.split('/');
            if (parts.length === 3) {
                tanggalTebus = `${parts[2]}/${parts[1]}/${parts[0]}`;
            } else {
                tanggalTebus = null;
            }
        } else {
            tanggalTebus = null;
        }

        if (metodePenebusan === 'ipubers') {
            return {
                kabupaten: (row.getCell(2).text || '').replace(/^KAB\.\s*/i, '').trim(),
                kecamatan: row.getCell(3).text || '',
                kodeTransaksi: kodeTransaksi,
                poktan: row.getCell(25).text || '',
                kodeKios: row.getCell(4).text || '',
                namaKios: row.getCell(5).text || '',
                nik: row.getCell(8).text || '',
                namaPetani: row.getCell(9).text || '',
                metodePenebusan: 'ipubers',
                tanggalTebus: tanggalTebus,
                urea: convertToNumber(row.getCell(10).text || '0'),
                npk: convertToNumber(row.getCell(11).text || '0'),
                sp36: convertToNumber(row.getCell(12).text || '0'),
                za: convertToNumber(row.getCell(13).text || '0'),
                npkFormula: convertToNumber(row.getCell(14).text || '0'),
                organik: convertToNumber(row.getCell(15).text || '0'),
                organikCair: convertToNumber(row.getCell(16).text || '0'),
                kakao: convertToNumber(row.getCell(26).text || '0'),
                status: row.getCell(24).text || 'Tidak Diketahui'
            };
        } else {
            const mid = row.getCell(4).text || '';
            const namaKios = row.getCell(5).text || '';

            return {
                kabupaten: (row.getCell(2).text || '').replace(/^KAB\.\s*/i, '').trim(),
                kecamatan: row.getCell(3).text || '',
                kodeTransaksi: kodeTransaksi,
                poktan: '',
                kodeKios: mid || '-',
                nik: row.getCell(6).text || '',
                namaPetani: row.getCell(7).text || '',
                metodePenebusan: 'kartan',
                tanggalTebus: tanggalTebus,
                urea: convertToNumber(row.getCell(8).text || '0'),
                npk: convertToNumber(row.getCell(9).text || '0'),
                sp36: convertToNumber(row.getCell(10).text || '0'),
                za: convertToNumber(row.getCell(11).text || '0'),
                npkFormula: convertToNumber(row.getCell(12).text || '0'),
                organik: convertToNumber(row.getCell(13).text || '0'),
                organikCair: convertToNumber(row.getCell(14).text || '0'),
                kakao: convertToNumber(row.getCell(19).text || '0'),
                mid: mid,
                namaKios: namaKios,
                status: row.getCell(17).text || 'kartan'
            };
        }
    });
}

exports.uploadFile = async (req, res) => {
    const sessionId = uuidv4(); // ID unik untuk tracking session upload
    let connection;
    let transactionActive = false;

    try {
        const files = req.files;
        const metodePenebusan = req.body.metodePenebusan;

        console.log(`[${sessionId}] 📤 Memulai upload, metode: ${metodePenebusan}, file: ${files.length}`);

        if (!files || files.length === 0) {
            return res.status(400).send({ message: 'Tidak ada file yang diunggah' });
        }

        // Dapatkan koneksi dengan timeout
        connection = await executeWithRetry(
            () => getConnectionWithTimeout(5000),
            3,
            1000
        );

        // Mulai transaksi
        await executeWithRetry(async () => {
            await connection.query('START TRANSACTION');
            transactionActive = true;
        });

        // 1. Load semua data referensi sekaligus
        console.log(`[${sessionId}] 🔍 Memuat data referensi...`);

        const [existingTransactions, allMainMID, allMasterMID] = await Promise.all([
            executeWithRetry(() => connection.query('SELECT kode_transaksi FROM verval')),
            executeWithRetry(() => connection.query('SELECT mid, kabupaten, kecamatan, nama_kios, kode_kios FROM main_mid')),
            executeWithRetry(() => connection.query('SELECT mid, kode_kios FROM master_mid'))
        ]);

        const existingTransactionsSet = new Set(existingTransactions.map(row => row.kode_transaksi));

        // 2. Bangun mapping untuk pencarian kode kios
        const midMap = new Map();
        const kabKecNamaMap = new Map();
        const masterMidMap = new Map();

        allMainMID.forEach(({ mid, kabupaten, kecamatan, nama_kios, kode_kios }) => {
            const normKabupaten = cleanText(kabupaten);
            const normKecamatan = cleanText(kecamatan);
            const normNamaKios = cleanNamaKios(nama_kios);

            if (mid) {
                midMap.set(`${cleanText(mid)}-${normKabupaten}`, kode_kios);
            }
            if (nama_kios) {
                kabKecNamaMap.set(`${normKabupaten}-${normKecamatan}-${normNamaKios}`, kode_kios);
            }
        });

        allMasterMID.forEach(({ mid, kode_kios }) => {
            masterMidMap.set(cleanText(mid), kode_kios);
        });

        let totalInserted = 0;
        let totalDuplicates = 0;

        // 3. Proses setiap file secara sequential
        for (const file of files) {
            try {
                if (!file.originalname.match(/\.(xlsx|xls)$/)) {
                    throw new Error('Hanya file Excel yang diperbolehkan');
                }

                const data = await processExcelFile(file, metodePenebusan);
                const batchSize = 50; // Jumlah row per batch insert
                let batchValues = [];

                for (const rowData of data) {
                    if (existingTransactionsSet.has(rowData.kodeTransaksi)) {
                        totalDuplicates++;
                        continue;
                    }

                    let kodeKios = rowData.kodeKios;

                    // Handle kartan kode kios lookup
                    if (rowData.metodePenebusan === 'kartan') {
                        const cleanMid = cleanText(rowData.mid);
                        const cleanKabupaten = cleanText(rowData.kabupaten);
                        const cleanKecamatan = cleanText(rowData.kecamatan);
                        const cleanedNamaKios = cleanNamaKios(rowData.namaKios);

                        if (midMap.has(`${cleanMid}-${cleanKabupaten}`) && midMap.get(`${cleanMid}-${cleanKabupaten}`) !== "-") {
                            kodeKios = midMap.get(`${cleanMid}-${cleanKabupaten}`);
                        } else if (kabKecNamaMap.has(`${cleanKabupaten}-${cleanKecamatan}-${cleanedNamaKios}`)) {
                            kodeKios = kabKecNamaMap.get(`${cleanKabupaten}-${cleanKecamatan}-${cleanedNamaKios}`);
                        } else if (masterMidMap.has(cleanMid)) {
                            kodeKios = masterMidMap.get(cleanMid);
                        } else {
                            kodeKios = rowData.mid ? `MID-${rowData.mid}` : '-';
                            console.log(`[${sessionId}] ℹ️ Menggunakan MID sebagai fallback kode kios`);
                        }
                    }

                    batchValues.push([
                        rowData.kabupaten, rowData.kecamatan, rowData.poktan, kodeKios,
                        rowData.namaKios, rowData.nik, rowData.namaPetani,
                        rowData.metodePenebusan, rowData.tanggalTebus, rowData.urea,
                        rowData.npk, rowData.sp36, rowData.za, rowData.npkFormula,
                        rowData.organik, rowData.organikCair, rowData.kakao,
                        rowData.kodeTransaksi, rowData.status
                    ]);

                    // Jika batch sudah penuh, insert ke database
                    if (batchValues.length >= batchSize) {
                        await executeWithRetry(async () => {
                            await connection.query(
                                `INSERT INTO verval 
                                (kabupaten, kecamatan, poktan, kode_kios, nama_kios, nik, nama_petani, 
                                 metode_penebusan, tanggal_tebus, urea, npk, sp36, za, npk_formula, 
                                 organik, organik_cair, kakao, kode_transaksi, status) 
                                VALUES ?`,
                                [batchValues]
                            );
                            totalInserted += batchValues.length;
                            batchValues = [];
                        });
                    }
                }

                // Insert sisa data yang belum terproses
                if (batchValues.length > 0) {
                    await executeWithRetry(async () => {
                        await connection.query(
                            `INSERT INTO verval 
                            (kabupaten, kecamatan, poktan, kode_kios, nama_kios, nik, nama_petani, 
                             metode_penebusan, tanggal_tebus, urea, npk, sp36, za, npk_formula, 
                             organik, organik_cair, kakao, kode_transaksi, status) 
                            VALUES ?`,
                            [batchValues]
                        );
                        totalInserted += batchValues.length;
                    });
                }

            } catch (error) {
                console.error(`[${sessionId}] ❌ Error memproses file ${file.originalname}:`, error);
                throw error;
            }
        }

        // Commit transaksi jika semua berhasil
        await executeWithRetry(async () => {
            await connection.query('COMMIT');
            transactionActive = false;
        });

        console.log(`[${sessionId}] ✅ Upload selesai. Total: ${totalInserted}, Duplikat: ${totalDuplicates}`);

        return res.status(200).send({
            message: `✅ Proses selesai. ${totalInserted} data ditambahkan, ${totalDuplicates} data duplikat.`,
            inserted: totalInserted,
            duplicates: totalDuplicates
        });

    } catch (error) {
        console.error(`[${sessionId}] ❌ Error utama:`, error);

        // Rollback transaksi jika masih aktif
        if (connection && transactionActive) {
            try {
                await connection.query('ROLLBACK');
            } catch (rollbackError) {
                console.error(`[${sessionId}] ❌ Gagal rollback:`, rollbackError);
            }
        }

        return res.status(500).send({
            message: '❌ Terjadi kesalahan dalam proses upload.',
            error: error.message,
            sessionId: sessionId
        });

    } finally {
        // Selalu lepaskan koneksi
        if (connection) {
            try {
                connection.release();
            } catch (releaseError) {
                console.error(`[${sessionId}] ❌ Gagal melepas koneksi:`, releaseError);
            }
        }
    }
};