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

// Fungsi untuk validasi struktur file Excel
const validateExcelStructure = (worksheet, metodePenebusan) => {
    const expectedColumnsKartan = ['NO', 'KABUPATEN', 'KECAMATAN', 'KODE KIOS', 'NAMA KIOS', 'NIK', 'NAMA PETANI', 'UREA', 'NPK', 'SP36', 'ZA', 'NPK FORMULA', 'ORGANIK', 'ORGANIK CAIR', 'TGL TEBUS', 'TGL INPUT', 'STATUS'];
    const expectedColumnsIpubers = ['NO', 'KABUPATEN', 'KECAMATAN', 'NO TRANSAKSI', 'KODE KIOS', 'NAMA KIOS', 'POKTAN', 'NIK', 'NAMA PETANI', 'KOMODITAS', 'UREA', 'NPK', 'SP36', 'ZA', 'NPK FORMULA', 'ORGANIK', 'ORGANIK CAIR', 'TGL TEBUS', 'TGL INPUT', 'STATUS'];

    const headerRow = 1;
    let rowValues;
    
    try {
        rowValues = worksheet.getRow(headerRow).values;
        if (rowValues[0] === undefined) rowValues.shift();
        const actualColumns = rowValues.map(cell => cell ? cell.toString().trim().toUpperCase() : "");
        const expectedColumns = metodePenebusan === 'ipubers' ? expectedColumnsIpubers : expectedColumnsKartan;

        console.log('Expected Columns:', expectedColumns);
        console.log('Actual Columns:', actualColumns);

        // Validasi kolom
        if (actualColumns.length < expectedColumns.length) {
            throw new Error(`Jumlah kolom tidak sesuai. Diharapkan: ${expectedColumns.length}, Aktual: ${actualColumns.length}`);
        }

        // Validasi setiap kolom
        const missingColumns = [];
        for (let i = 0; i < expectedColumns.length; i++) {
            if (actualColumns[i] !== expectedColumns[i]) {
                missingColumns.push({
                    expected: expectedColumns[i],
                    actual: actualColumns[i],
                    position: i + 1
                });
            }
        }

        if (missingColumns.length > 0) {
            throw new Error(`Struktur kolom tidak sesuai. Kolom yang bermasalah: ${JSON.stringify(missingColumns)}`);
        }

        return true;
    } catch (error) {
        throw new Error(`Validasi struktur file gagal: ${error.message}`);
    }
};

// Fungsi untuk memproses file Excel - DIPERBAIKI
const processExcelFile = async (file, metodePenebusan) => {
    const workbook = new ExcelJS.Workbook();
    let buffer;

    try {
        if (file.buffer) {
            buffer = file.buffer;
        } else if (file.path && fs.existsSync(file.path)) {
            buffer = fs.readFileSync(file.path);
        } else {
            throw new Error('File tidak ditemukan atau buffer tidak tersedia');
        }

        await workbook.xlsx.load(buffer);
        console.log(`File Excel berhasil dimuat: ${file.originalname}`);
    } catch (err) {
        throw new Error(`File Excel tidak valid atau korup: ${err.message}`);
    }

    const worksheet = workbook.worksheets[0];
    if (!worksheet) {
        throw new Error('Tidak ada worksheet yang ditemukan dalam file Excel');
    }

    console.log(`Worksheet: ${worksheet.name}, Total Rows: ${worksheet.rowCount}`);

    // Validasi struktur file
    validateExcelStructure(worksheet, metodePenebusan);

    // Proses data - DIPERBAIKI: Filter yang lebih baik
    const startRow = 2; // Row 1 untuk header, data mulai row 2
    
    let validRows = 0;
    let invalidRows = 0;
    const rows = [];

    for (let i = startRow; i <= worksheet.rowCount; i++) {
        try {
            const row = worksheet.getRow(i);
            
            // Periksa apakah baris kosong
            let isEmpty = true;
            for (let j = 1; j <= row.cellCount; j++) {
                const cell = row.getCell(j);
                if (cell.value !== null && cell.value !== undefined && cell.value.toString().trim() !== '') {
                    isEmpty = false;
                    break;
                }
            }

            if (isEmpty) {
                invalidRows++;
                continue;
            }

            // Validasi data minimal (harus ada NIK atau Nama Petani)
            const nik = row.getCell(metodePenebusan === 'ipubers' ? 8 : 6)?.text || '';
            const namaPetani = row.getCell(metodePenebusan === 'ipubers' ? 9 : 7)?.text || '';
            
            if (!nik && !namaPetani) {
                console.warn(`Row ${i} diabaikan: Tidak ada NIK dan Nama Petani`);
                invalidRows++;
                continue;
            }

            rows.push(row);
            validRows++;

        } catch (error) {
            console.warn(`Error membaca row ${i}:`, error.message);
            invalidRows++;
            continue;
        }
    }

    console.log(`File ${file.originalname}: ${validRows} baris valid, ${invalidRows} baris tidak valid/tidak diproses`);

    if (validRows === 0) {
        throw new Error('Tidak ada data valid yang ditemukan dalam file');
    }

    return rows;
};

// Fungsi konversi Excel date
function convertExcelDate(excelSerial) {
    try {
        const baseDate = new Date(1900, 0, 1);
        const date = new Date(baseDate.getTime() + (excelSerial - 1) * 24 * 60 * 60 * 1000);
        
        if (excelSerial > 59) {
            date.setTime(date.getTime() - 24 * 60 * 60 * 1000);
        }
        
        const year = date.getFullYear();
        const month = (date.getMonth() + 1).toString().padStart(2, '0');
        const day = date.getDate().toString().padStart(2, '0');
        
        console.log(`📊 Excel date: ${excelSerial} -> ${year}/${month}/${day}`);
        return `${year}/${month}/${day}`;
    } catch (error) {
        console.log(`❌ Error convert Excel date: ${error.message}`);
        return null;
    }
}

// Fungsi untuk parse berbagai format tanggal
// FUNGSI FIXED - TANGGAL PASTI BENAR
function tryMultipleDateFormats(dateStr) {
    if (!dateStr) return null;
    
    console.log(`🔍 START PARSING: "${dateStr}"`);

    // Bersihkan dan normalisasi
    const cleaned = dateStr.toString().trim().replace(/\s+/g, '');
    console.log(`🧹 Setelah cleaning: "${cleaned}"`);

    // SEMUA FORMAT DIASUMSIKAN SEBAGAI DD/MM/YYYY (Format Indonesia)
    const formats = [
        // Format: DD/MM/YYYY atau D/M/YYYY
        { 
            regex: /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/,
            handler: (match) => {
                const d = match[1];  // Day
                const m = match[2];  // Month  
                const y = match[3];  // Year
                return { d, m, y, format: 'DD/MM/YYYY' };
            }
        },
        // Format: DD/MM/YY atau D/M/YY
        { 
            regex: /^(\d{1,2})\/(\d{1,2})\/(\d{2})$/,
            handler: (match) => {
                const d = match[1];  // Day
                const m = match[2];  // Month
                let y = match[3];    // Year 2-digit
                y = parseInt(y) > 50 ? '19' + y : '20' + y;
                return { d, m, y, format: 'DD/MM/YY' };
            }
        },
        // Format: DD-MM-YYYY atau D-M-YYYY - INI YANG DIPERBAIKI
        { 
            regex: /^(\d{1,2})-(\d{1,2})-(\d{4})$/,
            handler: (match) => {
                const d = match[1];  // Day
                const m = match[2];  // Month
                const y = match[3];  // Year
                return { d, m, y, format: 'DD-MM-YYYY' };
            }
        },
        // Format: DD-MM-YY atau D-M-YY - INI YANG DIPERBAIKI
        { 
            regex: /^(\d{1,2})-(\d{1,2})-(\d{2})$/,
            handler: (match) => {
                const d = match[1];  // Day
                const m = match[2];  // Month
                let y = match[3];    // Year 2-digit
                y = parseInt(y) > 50 ? '19' + y : '20' + y;
                return { d, m, y, format: 'DD-MM-YY' };
            }
        },
        // Format: DD.MM.YYYY atau D.M.YYYY
        { 
            regex: /^(\d{1,2})\.(\d{1,2})\.(\d{4})$/,
            handler: (match) => {
                const d = match[1];  // Day
                const m = match[2];  // Month
                const y = match[3];  // Year
                return { d, m, y, format: 'DD.MM.YYYY' };
            }
        }
    ];

    for (const format of formats) {
        const match = cleaned.match(format.regex);
        if (match) {
            try {
                const { d, m, y, format: detectedFormat } = format.handler(match);
                
                const day = parseInt(d);
                const month = parseInt(m);
                const year = parseInt(y);
                
                console.log(`📅 Format terdeteksi: ${detectedFormat}`);
                console.log(`   Input: ${dateStr} -> Day: ${day}, Month: ${month}, Year: ${year}`);

                // Validasi bulan
                if (month < 1 || month > 12) {
                    console.log(`❌ Bulan tidak valid: ${month}`);
                    continue;
                }
                
                // Validasi hari
                if (day < 1 || day > 31) {
                    console.log(`❌ Hari tidak valid: ${day}`);
                    continue;
                }

                // Validasi hari per bulan
                const daysInMonth = new Date(year, month, 0).getDate();
                if (day > daysInMonth) {
                    console.log(`❌ Hari ${day} > ${daysInMonth} (max hari bulan ${month})`);
                    continue;
                }

                const result = `${year}/${month.toString().padStart(2, '0')}/${day.toString().padStart(2, '0')}`;
                console.log(`✅ BERHASIL: ${dateStr} -> ${result}`);
                return result;

            } catch (error) {
                console.log(`❌ Error: ${error.message}`);
                continue;
            }
        }
    }
    
    console.log(`❌ Gagal parse semua format: "${dateStr}"`);
    return null;
}

// FUNGSI FIXED - Parser utama
function parseTanggalAcak(tanggal) {
    if (!tanggal) {
        console.log('❌ Tanggal null/undefined');
        return null;
    }
    
    console.log(`🔄 parseTanggalAcak input: ${tanggal} (tipe: ${typeof tanggal})`);

    // Jika sudah Date object
    if (tanggal instanceof Date) {
        const year = tanggal.getFullYear();
        const month = (tanggal.getMonth() + 1).toString().padStart(2, '0');
        const day = tanggal.getDate().toString().padStart(2, '0');
        const result = `${year}/${month}/${day}`;
        console.log(`📅 Dari Date object: ${result}`);
        return result;
    }
    
    // Jika number (Excel serial date)
    if (typeof tanggal === 'number' && tanggal > 0) {
        const result = convertExcelDate(tanggal);
        console.log(`📅 Dari Excel serial: ${tanggal} -> ${result}`);
        return result;
    }
    
    // Jika string
    if (typeof tanggal === 'string') {
        const cleaned = tanggal.trim();
        
        // Coba parse sebagai Excel date string (hanya angka)
        if (/^\d+$/.test(cleaned)) {
            const result = convertExcelDate(parseInt(cleaned));
            console.log(`📅 Dari string angka: ${cleaned} -> ${result}`);
            return result;
        }
        
        // Gunakan parser custom untuk format teks
        console.log(`🔍 Parsing string: "${cleaned}"`);
        return tryMultipleDateFormats(cleaned);
    }
    
    console.log(`❌ Tipe data tidak didukung: ${typeof tanggal}`);
    return null;
}

// Fungsi validasi tanggal
function isValidDate(dateStr) {
    if (!dateStr || typeof dateStr !== 'string') return false;
    
    const match = dateStr.match(/^(\d{4})\/(\d{2})\/(\d{2})$/);
    if (!match) return false;
    
    const [_, year, month, day] = match;
    const y = parseInt(year), m = parseInt(month), d = parseInt(day);
    
    if (m < 1 || m > 12 || d < 1 || d > 31 || y < 1900 || y > 2100) {
        return false;
    }
    
    const daysInMonth = new Date(y, m, 0).getDate();
    return d <= daysInMonth;
}

// Fungsi normalisasi nama wilayah
function normalisasiNamaWilayah(nama) {
    if (!nama || typeof nama !== 'string') return nama || '';
    
    let normalized = nama.trim().toUpperCase();
    
    const specialCases = {
        'GUNUNGKIDUL': 'GUNUNG KIDUL',
        'GUNUNG KIDUL': 'GUNUNG KIDUL',
        'BAMBANGLIPURO': 'BAMBANG LIPURO',
        'BAMBANG LIPURO': 'BAMBANG LIPURO'
    };
    
    if (specialCases[normalized]) {
        return specialCases[normalized];
    }
    
    return normalized;
}

// Fungsi untuk memproses data row - VERSI FIXED
const processRowData = (row, metodePenebusan) => {
    let tanggalTebus, tanggalExcel;
    const kodeTransaksi = metodePenebusan === 'ipubers' ? (row.getCell(4).text || generateKodeTransaksi()) : generateKodeTransaksi();

    // Ambil tanggal dari cell yang benar
    if (metodePenebusan === 'ipubers') {
        tanggalExcel = row.getCell(18).value; // Gunakan .value bukan .text
    } else if (metodePenebusan === 'kartan') {
        tanggalExcel = row.getCell(15).value; // Gunakan .value bukan .text
    }

    console.log(`🎯 PROSES ROW - Tanggal excel: ${tanggalExcel} (tipe: ${typeof tanggalExcel})`);

    // PARSING TANGGAL - VERSI SIMPLE & PASTI BENAR
    if (tanggalExcel instanceof Date) {
        // Jika sudah Date object dari Excel
        const dateObj = tanggalExcel;
        const year = dateObj.getFullYear();
        const month = (dateObj.getMonth() + 1).toString().padStart(2, '0');
        const day = dateObj.getDate().toString().padStart(2, '0');
        tanggalTebus = `${year}/${month}/${day}`;
        console.log(`✅ Dari Date object Excel: ${tanggalTebus}`);
        
    } else {
        // Gunakan parser custom kita
        const parsed = parseTanggalAcak(tanggalExcel);
        if (parsed) {
            tanggalTebus = parsed;
            console.log(`✅ Berhasil parsing: ${tanggalExcel} -> ${tanggalTebus}`);
        } else {
            // Fallback ke tanggal hari ini
            console.warn(`⚠️ Gagal parsing: ${tanggalExcel}, gunakan tanggal hari ini`);
            tanggalTebus = new Date().toISOString().split('T')[0].replace(/-/g, '/');
        }
    }

    // Validasi akhir
    if (!tanggalTebus || !isValidDate(tanggalTebus)) {
        console.warn(`❌ Tanggal akhir tidak valid: ${tanggalTebus}, gunakan tanggal hari ini`);
        tanggalTebus = new Date().toISOString().split('T')[0].replace(/-/g, '/');
    }

    console.log(`🎯 FINAL TANGGAL: ${tanggalTebus}`);

    // Process data berdasarkan metode (sisa code tetap sama)
    if (metodePenebusan === 'ipubers') {
        return {
            kabupaten: normalisasiNamaWilayah((row.getCell(2).text || '').replace(/^KAB\.\s*/i, '').trim()),
            kecamatan: normalisasiNamaWilayah(row.getCell(3).text || ''),
            kodeTransaksi: kodeTransaksi,
            poktan: row.getCell(7)?.text || '',
            kodeKios: row.getCell(5).text || '',
            namaKios: row.getCell(6).text || '',
            nik: row.getCell(8).text || '',
            namaPetani: row.getCell(9).text || '',
            metodePenebusan: 'ipubers',
            tanggalTebus: tanggalTebus,
            urea: convertToNumber(row.getCell(11).text || '0'),
            npk: convertToNumber(row.getCell(12).text || '0'),
            sp36: convertToNumber(row.getCell(13).text || '0'),
            za: convertToNumber(row.getCell(14).text || '0'),
            npkFormula: convertToNumber(row.getCell(15).text || '0'),
            organik: convertToNumber(row.getCell(16).text || '0'),
            organikCair: convertToNumber(row.getCell(17).text || '0'),
            kakao: convertToNumber(row.getCell(26)?.text || '0'),
            status: row.getCell(20).text || 'Tidak Diketahui'
        };
    } else {
        const mid = `'${row.getCell(4).text || ''}`;
        return {
            kabupaten: normalisasiNamaWilayah((row.getCell(2).text || '').replace(/^KAB\.\s*/i, '').trim()),
            kecamatan: normalisasiNamaWilayah(row.getCell(3).text || ''),
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
            kakao: convertToNumber(row.getCell(19)?.text || '0'),
            mid: mid,
            namaKios: row.getCell(5).text || '',
            status: row.getCell(17).text || 'kartan'
        };
    }
};

// Fungsi untuk mendapatkan mapping kode kios
const getKodeKiosMappings = async () => {
    try {
        const [allMainMID] = await db.query('SELECT mid, kabupaten, kecamatan, nama_kios, kode_kios FROM main_mid');
        const [allMasterMID] = await db.query('SELECT mid, kode_kios FROM master_mid');
        const [masterMID2025] = await db.query('SELECT mid, kode_kios FROM master_mid_2025');

        const mappings = {
            midMap: new Map(),
            kabKecNamaMap: new Map(),
            masterMidMap: new Map(),
            mid2025Map: new Map()
        };

        allMasterMID.forEach(({ mid, kode_kios }) => {
            mappings.masterMidMap.set(cleanText(mid), kode_kios);
        });

        masterMID2025.forEach(({ mid, kode_kios }) => {
            mappings.mid2025Map.set(cleanText(mid), kode_kios);
        });

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

        console.log(`Mapping loaded: ${mappings.midMap.size} midMap, ${mappings.kabKecNamaMap.size} kabKecNamaMap, ${mappings.masterMidMap.size} masterMidMap, ${mappings.mid2025Map.size} mid2025Map`);
        return mappings;
    } catch (error) {
        throw new Error(`Gagal memuat data mapping kode kios: ${error.message}`);
    }
};

// Fungsi untuk menentukan kode kios
const determineKodeKios = (rowData, mappings) => {
    if (rowData.metodePenebusan !== 'kartan') {
        return rowData.kodeKios || '-';
    }

    const { midMap, kabKecNamaMap, masterMidMap, mid2025Map } = mappings;
    const cleanMid = cleanText(rowData.mid);
    const cleanKab = cleanText(rowData.kabupaten);
    const cleanKec = cleanText(rowData.kecamatan);
    const cleanNama = cleanNamaKios(rowData.namaKios);

    if (mid2025Map.has(cleanMid)) {
        return mid2025Map.get(cleanMid);
    }

    const midKabKey = `${cleanMid}-${cleanKab}`;
    if (midMap.has(midKabKey)) {
        const kode = midMap.get(midKabKey);
        if (kode && kode !== '-') return kode;
    }

    if (masterMidMap.has(cleanMid)) {
        return masterMidMap.get(cleanMid);
    }

    const kabKecNamaKey = `${cleanKab}-${cleanKec}-${cleanNama}`;
    if (kabKecNamaMap.has(kabKecNamaKey)) {
        return kabKecNamaMap.get(kabKecNamaKey);
    }

    return rowData.kodeKios || '-';
};

// Fungsi untuk menyimpan data dengan retry mechanism
const saveWithRetry = async (values, retryCount = 0) => {
    const connection = await db.getConnection();
    try {
        await connection.beginTransaction();
        await connection.query(`SET SESSION innodb_lock_wait_timeout = ${DB_LOCK_TIMEOUT / 1000}`);

        for (let i = 0; i < values.length; i += BATCH_SIZE) {
            const batch = values.slice(i, i + BATCH_SIZE);
            if (batch.length > 0) {
                await connection.query(
                    `INSERT INTO verval_test 
                    (kabupaten, kecamatan, poktan, kode_kios, nama_kios, nik, nama_petani, metode_penebusan, tanggal_tebus, 
                    urea, npk, sp36, za, npk_formula, organik, organik_cair, kakao, kode_transaksi, status) 
                    VALUES ?`, [batch]
                );
            }
        }

        await connection.commit();
        return true;
    } catch (error) {
        await connection.rollback();

        if ((error.code === 'ER_LOCK_WAIT_TIMEOUT' || error.code === 'ER_LOCK_DEADLOCK') && retryCount < MAX_RETRIES) {
            const delayTime = RETRY_DELAY_BASE * Math.pow(2, retryCount);
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
    const startTime = Date.now();
    
    const log = (message) => console.log(`[${requestId}] ${new Date().toISOString()} - ${message}`);
    const logError = (message, error) => console.error(`[${requestId}] ${new Date().toISOString()} - ${message}`, error);

    let totalInserted = 0;
    let totalDuplicates = 0;
    let totalFilesProcessed = 0;
    let totalRowsProcessed = 0;
    let totalInvalidRows = 0;

    try {
        const { files, body } = req;
        const { metodePenebusan } = body;

        if (!files || files.length === 0) {
            return res.status(400).json({ 
                success: false,
                error: 'Tidak ada file yang diunggah',
                details: 'Silakan pilih file Excel untuk diupload'
            });
        }

        if (!metodePenebusan || !['ipubers', 'kartan'].includes(metodePenebusan)) {
            return res.status(400).json({
                success: false,
                error: 'Metode penebusan tidak valid',
                details: 'Metode penebusan harus ipubers atau kartan'
            });
        }

        log(`Memulai proses upload ${files.length} file dengan metode ${metodePenebusan}`);

        // 1. Load semua data referensi sekaligus di awal
        log('Memuat data referensi...');
        const [existingTransactions] = await db.query('SELECT kode_transaksi FROM verval_test');
        const existingTxSet = new Set(existingTransactions.map(tx => tx.kode_transaksi));
        const kodeKiosMappings = await getKodeKiosMappings();
        log('Data referensi berhasil dimuat');

        // 2. Proses setiap file
        const fileResults = [];

        for (const file of files) {
            const fileStartTime = Date.now();
            const fileResult = {
                filename: file.originalname,
                success: false,
                processed: 0,
                inserted: 0,
                duplicates: 0,
                invalid: 0,
                error: null
            };

            log(`Memproses file: ${file.originalname}`);

            // Validasi file type
            if (!file.originalname.match(/\.(xlsx|xls)$/i)) {
                fileResult.error = 'Format file tidak didukung. Hanya file Excel (.xlsx, .xls) yang diperbolehkan';
                fileResults.push(fileResult);
                continue;
            }

            try {
                const fileRows = await processExcelFile(file, metodePenebusan);
                log(`Data valid dari ${file.originalname}: ${fileRows.length} baris`);

                const batches = [];
                let currentBatch = [];
                let fileInserted = 0;
                let fileDuplicates = 0;

                // 3. Process each row
                for (const row of fileRows) {
                    totalRowsProcessed++;
                    fileResult.processed++;

                    try {
                        const rowData = processRowData(row, metodePenebusan);

                        if (existingTxSet.has(rowData.kodeTransaksi)) {
                            totalDuplicates++;
                            fileDuplicates++;
                            fileResult.duplicates++;
                            continue;
                        }

                        const kodeKios = determineKodeKios(rowData, kodeKiosMappings);
                        currentBatch.push([
                            rowData.kabupaten, rowData.kecamatan, rowData.poktan || '', kodeKios,
                            rowData.namaKios || '', rowData.nik || '', rowData.namaPetani || '',
                            metodePenebusan, rowData.tanggalTebus || null,
                            rowData.urea || 0, rowData.npk || 0, rowData.sp36 || 0, rowData.za || 0,
                            rowData.npkFormula || 0, rowData.organik || 0, rowData.organikCair || 0,
                            rowData.kakao || 0, rowData.kodeTransaksi, rowData.status || 'pending'
                        ]);

                        if (currentBatch.length >= BATCH_SIZE) {
                            batches.push([...currentBatch]);
                            currentBatch = [];
                        }
                    } catch (rowError) {
                        logError(`Error processing row in ${file.originalname}`, rowError);
                        totalInvalidRows++;
                        fileResult.invalid++;
                        continue;
                    }
                }

                // Add remaining batch
                if (currentBatch.length > 0) {
                    batches.push(currentBatch);
                }

                log(`Menyiapkan ${batches.length} batch untuk disimpan dari ${file.originalname}`);

                // 4. Save batches
                for (const batch of batches) {
                    try {
                        if (batch.length > 0) {
                            await saveWithRetry(batch);
                            totalInserted += batch.length;
                            fileInserted += batch.length;
                            fileResult.inserted += batch.length;
                        }
                    } catch (batchError) {
                        logError(`Gagal menyimpan batch dari ${file.originalname}`, batchError);
                        throw batchError;
                    }
                }

                fileResult.success = true;
                totalFilesProcessed++;
                const fileProcessingTime = Date.now() - fileStartTime;
                
                log(`Selesai memproses file ${file.originalname} - ${fileInserted} inserted, ${fileDuplicates} duplicates, ${fileResult.invalid} invalid - ${fileProcessingTime}ms`);

            } catch (fileError) {
                fileResult.error = fileError.message;
                logError(`Gagal memproses file ${file.originalname}`, fileError);
            }

            fileResults.push(fileResult);
        }

        const totalProcessingTime = Date.now() - startTime;

        // 5. Prepare final response
        const successFiles = fileResults.filter(f => f.success).length;
        const failedFiles = fileResults.filter(f => !f.success).length;

        const response = {
            success: true,
            summary: {
                totalFiles: files.length,
                filesProcessed: totalFilesProcessed,
                filesSuccess: successFiles,
                filesFailed: failedFiles,
                totalRowsProcessed: totalRowsProcessed,
                totalInserted: totalInserted,
                totalDuplicates: totalDuplicates,
                totalInvalidRows: totalInvalidRows,
                processingTime: `${totalProcessingTime}ms`
            },
            details: fileResults,
            message: `Berhasil memproses ${successFiles} dari ${files.length} file. ${totalInserted} data baru ditambahkan, ${totalDuplicates} data duplikat ditemukan.`
        };

        log(`Proses selesai. Total waktu: ${totalProcessingTime}ms`);
        return res.json(response);

    } catch (error) {
        const totalProcessingTime = Date.now() - startTime;
        logError('Error utama dalam uploadFile', error);
        
        return res.status(500).json({
            success: false,
            error: 'Terjadi kesalahan sistem selama proses upload',
            details: process.env.NODE_ENV === 'development' ? error.message : 'Silakan coba lagi atau hubungi administrator',
            summary: {
                totalFiles: files?.length || 0,
                filesProcessed: totalFilesProcessed,
                totalInserted: totalInserted,
                totalDuplicates: totalDuplicates,
                totalInvalidRows: totalInvalidRows,
                processingTime: `${totalProcessingTime}ms`
            }
        });
    }
};