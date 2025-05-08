const ExcelJS = require('exceljs');
const fs = require('fs');
const db = require('../config/db');
const { v4: uuidv4 } = require('uuid');

// Konfigurasi
const DB_LOCK_TIMEOUT = 60000; // 60 detik
const MAX_RETRIES = 5;
const RETRY_DELAY_BASE = 1000; // 1 detik dasar
const BATCH_SIZE = 100; // Jumlah data per batch insert

// Helper functions
const delay = (ms) => new Promise(resolve => setTimeout(resolve, ms));

const convertToNumber = (value) => {
    if (typeof value === 'string') {
        value = value.replace(',', '.').replace(/[^0-9.-]/g, '');
    }
    const numberValue = parseFloat(value);
    return isNaN(numberValue) ? 0 : numberValue;
};

const cleanText = (text) => {
    return text ? text.toLowerCase().trim().replace(/[^a-z0-9\s]/g, '') : '';
};

const cleanNamaKios = (text) => {
    if (!text) return "";
    let cleaned = text.toUpperCase()
        .replace(/([A-Z])([A-Z][a-z])/g, '$1 $2')
        .replace(/[.,"\-/]/g, "")
        .replace(/\b(KPL|KIOS|UD|KUD|TOKO|TK|GAPOKTAN|CV)\b/gi, "")
        .trim();
    return cleaned;
};

const generateKodeTransaksi = () => {
    return 'TX-' + Date.now().toString(36) + '-' + Math.random().toString(36).substr(2, 7).toUpperCase();
};

// Fungsi untuk memproses file Excel (dipertahankan sama seperti aslinya)
const processExcelFile = async (file, metodePenebusan) => {
    const workbook = new ExcelJS.Workbook();
    let buffer;

    try {
        buffer = file.buffer || fs.readFileSync(file.path);
        await workbook.xlsx.load(buffer);
    } catch (err) {
        throw new Error('File Excel tidak valid');
    }

    const worksheet = workbook.worksheets[0];
    const expectedColumnsKartan = ['NO', 'KABUPATEN', 'KECAMATAN', 'KODE KIOS', 'NAMA KIOS', 'NIK', 'NAMA PETANI', 'UREA', 'NPK', 'SP36', 'ZA', 'NPK FORMULA', 'ORGANIK', 'ORGANIK CAIR', 'TGL TEBUS', 'TGL INPUT', 'STATUS'];
//    const expectedColumnsIpubers = ['NO', 'KABUPATEN', 'KECAMATAN', 'NO TRANSAKSI', 'KODE KIOS', 'NAMA KIOS', 'POKTAN', 'NIK', 'NAMA PETANI', 'KOMODITAS','UREA', 'NPK','NPK FORMULA', 'ORGANIK', 'ORGANIK CAIR', 'TGL TEBUS', 'TGL INPUT', 'STATUS'];
const expectedColumnsIpubers = ['No', 'Kabupaten',	'Kecamatan', 'Kode Kios',	'Nama Kios',	'Kode TRX', 'No Transaksi',	'NIK',	'Nama Petani',	'Urea', 'NPK',	'SP36',	'ZA',	'NPK Formula',	'Organik',	'Organik Cair',	'Keterangan',	'Tanggal Tebus',	'Tanggal Entri',	'Tanggal Update',	'Tipe Tebus',	'NIK Perwakilan',	'Url Bukti',	'Status'];

    // Validasi header
    const headerRow = metodePenebusan === 'ipubers' ? 1 : 2;
    const rowValues = worksheet.getRow(headerRow).values;
    if (rowValues[0] === undefined) rowValues.shift();
    const actualColumns = rowValues.map(cell => cell ? cell.toString().trim() : "");
    const expectedColumns = metodePenebusan === 'ipubers' ? expectedColumnsIpubers : expectedColumnsKartan;

    if (!expectedColumns.every((col, index) => col === actualColumns[index])) {
        throw new Error('Struktur file Excel tidak sesuai');
    }

    // Proses data
    const startRow = headerRow + 1;
    let rows = worksheet.getRows(startRow, worksheet.rowCount - startRow);
    rows = rows.filter(row => row.getCell(1).value && row.getCell(2).value);

    return rows.map((row) => {
        let tanggalTebus, tanggalExcel;
        const kodeTransaksi = metodePenebusan === 'ipubers' ? row.getCell(7).text : generateKodeTransaksi();

        if (metodePenebusan === 'ipubers') {
            tanggalExcel = row.getCell(18).text;
        } else if (metodePenebusan === 'kartan') {
            tanggalExcel = row.getCell(15).text;
        }

        // Format tanggal (dipertahankan sama seperti aslinya)
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
            }
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
            const mid = `'${row.getCell(4).text || ''}`;
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
                namaKios: row.getCell(5).text || '',
                status: row.getCell(17).text || 'kartan'
            };
        }
    });
};

// Fungsi untuk mendapatkan mapping kode kios (dipertahankan sama seperti aslinya)
const getKodeKiosMappings = async () => {
    const [allMainMID] = await db.query('SELECT mid, kabupaten, kecamatan, nama_kios, kode_kios FROM main_mid');
    const [allMasterMID] = await db.query('SELECT mid, kode_kios FROM master_mid');

    const mappings = {
        midMap: new Map(),
        kabKecNamaMap: new Map(),
        masterMidMap: new Map()
    };

    allMainMID.forEach(({ mid, kabupaten, kecamatan, nama_kios, kode_kios }) => {
        const normKabupaten = cleanText(kabupaten);
        const normKecamatan = cleanText(kecamatan);
        const normNamaKios = cleanNamaKios(nama_kios);

        if (mid) {
            mappings.midMap.set(`${cleanText(mid)}-${normKabupaten}`, kode_kios);
        }
        if (nama_kios) {
            mappings.kabKecNamaMap.set(`${normKabupaten}-${normKecamatan}-${normNamaKios}`, kode_kios);
        }
    });

    allMasterMID.forEach(({ mid, kode_kios }) => {
        mappings.masterMidMap.set(cleanText(mid), kode_kios);
    });

    return mappings;
};

// Fungsi untuk menentukan kode kios (dipertahankan sama seperti aslinya)
const determineKodeKios = (rowData, mappings) => {
    if (rowData.metodePenebusan !== 'kartan') {
        return rowData.kodeKios || '-';
    }

    const { midMap, kabKecNamaMap, masterMidMap } = mappings;
    const cleanMid = cleanText(rowData.mid);
    const cleanKab = cleanText(rowData.kabupaten);
    const cleanKec = cleanText(rowData.kecamatan);
    const cleanNama = cleanNamaKios(rowData.namaKios);

    // 1. Coba match dengan MID + Kabupaten
    const midKabKey = `${cleanMid}-${cleanKab}`;
    if (midMap.has(midKabKey)) {
        const kode = midMap.get(midKabKey);
        if (kode && kode !== '-') return kode;
    }

    // 2. Coba match dengan Kab+Kec+Nama
    const kabKecNamaKey = `${cleanKab}-${cleanKec}-${cleanNama}`;
    if (kabKecNamaMap.has(kabKecNamaKey)) {
        return kabKecNamaMap.get(kabKecNamaKey);
    }

    // 3. Coba match dengan MID saja (master_mid)
    if (masterMidMap.has(cleanMid)) {
        return masterMidMap.get(cleanMid);
    }

    return rowData.kodeKios || '-';
};

// Fungsi baru untuk menyimpan data dengan retry mechanism
const saveWithRetry = async (values, retryCount = 0) => {
    const connection = await db.getConnection();
    try {
        await connection.beginTransaction();

        // Set lock timeout yang lebih besar
        await connection.query(`SET SESSION innodb_lock_wait_timeout = ${DB_LOCK_TIMEOUT / 1000}`);

        // Split into smaller batches if needed
        for (let i = 0; i < values.length; i += BATCH_SIZE) {
            const batch = values.slice(i, i + BATCH_SIZE);
            await connection.query(
                `INSERT INTO verval 
                (kabupaten, kecamatan, poktan, kode_kios, nama_kios, nik, nama_petani, metode_penebusan, tanggal_tebus, 
                urea, npk, sp36, za, npk_formula, organik, organik_cair, kakao, kode_transaksi, status) 
                VALUES ?`, [batch]
            );
        }

        await connection.commit();
        return true;
    } catch (error) {
        await connection.rollback();

        if ((error.code === 'ER_LOCK_WAIT_TIMEOUT' || error.code === 'ER_LOCK_DEADLOCK') && retryCount < MAX_RETRIES) {
            const delayTime = RETRY_DELAY_BASE * Math.pow(2, retryCount); // Exponential backoff
            console.warn(`[Retry ${retryCount + 1}] Lock timeout detected, retrying in ${delayTime}ms...`);
            await delay(delayTime);
            return saveWithRetry(values, retryCount + 1);
        }

        console.error('Failed to save after retries:', error);
        throw error;
    } finally {
        connection.release();
    }
};

// Controller utama dengan penambahan fitur untuk handle concurrent upload
exports.uploadFile = async (req, res) => {
    const requestId = uuidv4();
    const log = (message) => console.log(`[${requestId}] ${message}`);
    const logError = (message, error) => console.error(`[${requestId}] ${message}`, error);

    try {
        const { files, body } = req;
        const { metodePenebusan } = body;

        if (!files || files.length === 0) {
            return res.status(400).json({ error: 'Tidak ada file yang diunggah' });
        }

        log(`Memulai proses upload ${files.length} file dengan metode ${metodePenebusan}`);

        // 1. Load semua data referensi sekaligus di awal
        log('Memuat data referensi...');
        const [existingTransactions] = await db.query('SELECT kode_transaksi FROM verval');
        const existingTxSet = new Set(existingTransactions.map(tx => tx.kode_transaksi));
        const kodeKiosMappings = await getKodeKiosMappings();
        log('Data referensi berhasil dimuat');

        let totalInserted = 0;
        let totalDuplicates = 0;

        // 2. Proses setiap file
        for (const file of files) {
            log(`Memproses file: ${file.originalname}`);

            if (!file.originalname.match(/\.(xlsx|xls)$/i)) {
                return res.status(400).json({ error: 'Hanya file Excel (.xlsx, .xls) yang diperbolehkan' });
            }

            try {
                const fileData = await processExcelFile(file, metodePenebusan);
                const batches = [];
                let currentBatch = [];

                // 3. Siapkan data untuk insert
                for (const row of fileData) {
                    if (existingTxSet.has(row.kodeTransaksi)) {
                        totalDuplicates++;
                        continue;
                    }

                    const kodeKios = determineKodeKios(row, kodeKiosMappings);
                    currentBatch.push([
                        row.kabupaten, row.kecamatan, row.poktan || '', kodeKios,
                        row.namaKios || '', row.nik || '', row.namaPetani || '',
                        metodePenebusan, row.tanggalTebus || null,
                        row.urea || 0, row.npk || 0, row.sp36 || 0, row.za || 0,
                        row.npkFormula || 0, row.organik || 0, row.organikCair || 0,
                        row.kakao || 0, row.kodeTransaksi, row.status || 'pending'
                    ]);

                    if (currentBatch.length >= BATCH_SIZE) {
                        batches.push([...currentBatch]);
                        currentBatch = [];
                    }
                }

                if (currentBatch.length > 0) {
                    batches.push(currentBatch);
                }

                // 4. Simpan data per batch dengan retry mechanism
                for (const batch of batches) {
                    try {
                        await saveWithRetry(batch);
                        totalInserted += batch.length;
                        log(`Berhasil menyimpan batch ${batch.length} data`);
                    } catch (err) {
                        logError('Gagal menyimpan batch data', err);
                        throw err;
                    }
                }

            } catch (err) {
                logError('Gagal memproses file', err);
                return res.status(400).json({
                    error: `Gagal memproses file ${file.originalname}`,
                    details: err.message
                });
            }
        }

        log(`Proses selesai. Total: ${totalInserted} data baru, ${totalDuplicates} duplikat`);
        return res.json({
            success: true,
            inserted: totalInserted,
            duplicates: totalDuplicates,
            message: `Berhasil memproses ${files.length} file`
        });

    } catch (error) {
        logError('Error utama dalam uploadFile', error);
        return res.status(500).json({
            error: 'Terjadi kesalahan sistem',
            details: process.env.NODE_ENV === 'development' ? error.message : undefined
        });
    }
};
