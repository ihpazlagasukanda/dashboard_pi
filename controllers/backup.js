const ExcelJS = require('exceljs');
const fs = require('fs');
const db = require('../config/db');  // Pastikan path database benar

// Fungsi untuk konversi nilai menjadi angka
const convertToNumber = (value) => {
    if (typeof value === 'string') {
        value = value.replace(',', '.'); // Ubah koma ke titik untuk desimal
    }
    const numberValue = parseFloat(value);
    return isNaN(numberValue) ? 0 : numberValue;
};

exports.uploadFile = async (req, res) => {
    try {
        const files = req.files; // Ambil semua file dari request
        const metodePenebusan = req.body.metodePenebusan;

        console.log('📤 Jumlah File Diterima:', files.length);
        console.log('🛠️ Metode Penebusan:', metodePenebusan);

        if (!files || files.length === 0) {
            return res.status(400).send({ message: 'Tidak ada file yang diunggah' });
        }

        let insertedCount = 0;
        let duplicateCount = 0;

        for (const file of files) {
            console.log('📝 File yang diproses:', file.originalname);
            console.log('📄 MIME Type:', file.mimetype);
            console.log('📂 File Path:', file.path);

            // Validasi format file
            if (!file.originalname.match(/\.(xlsx|xls)$/)) {
                return res.status(400).send({ message: 'Hanya file Excel yang diperbolehkan' });
            }

            const workbook = new ExcelJS.Workbook();
            let buffer;

            try {
                if (file.buffer) {
                    buffer = file.buffer; // Jika pakai memory storage
                } else {
                    buffer = fs.readFileSync(file.path); // Jika pakai disk storage
                }

                await workbook.xlsx.load(buffer);
                console.log('✅ File berhasil dibaca!');
            } catch (err) {
                console.error('❌ Gagal membaca file Excel:', err);
                return res.status(400).send({ message: 'File Excel tidak valid' });
            }

            const worksheet = workbook.worksheets[0]; // Ambil sheet pertama
            console.log(`📜 Sheet: ${worksheet.name}`);

            // Validasi kolom sesuai metode penebusan
            const expectedColumnsKartan = ['NO', 'KABUPATEN', 'KECAMATAN', 'KODE KIOS', 'NAMA KIOS', 'NIK', 'NAMA PETANI', 'UREA', 'NPK', 'SP36', 'ZA', 'NPK FORMULA', 'ORGANIK', 'ORGANIK CAIR', 'TGL TEBUS', 'TGL INPUT', 'STATUS'];
            const expectedColumnsIpubers = ['No', 'Kabupaten', 'Kecamatan', 'Kode Kios', 'Nama Kios', 'Kode TRX', 'No Transaksi', 'NIK', 'Nama Petani', 'Urea', 'NPK', 'SP36', 'ZA', 'NPK Formula', 'Organik', 'Organik Cair', 'Keterangan', 'Tanggal Tebus', 'Tanggal Entri', 'Tanggal Update', 'Tipe Tebus', 'NIK Perwakilan', 'Url Bukti', 'Status'];

            const headerRow = metodePenebusan === 'ipubers' ? 1 : 1; // Baris header
            const rowValues = worksheet.getRow(headerRow).values;

            // **Hapus elemen pertama jika undefined (karena values[0] sering kosong)**
            if (rowValues[0] === undefined) {
                rowValues.shift();
            }

            // **Ambil teks dari setiap cell dengan aman**
            const actualColumns = rowValues.map(cell => cell ? cell.toString().trim() : "");

            console.log('🔍 Header dari file Excel:', actualColumns);

            // **Validasi apakah header sesuai dengan yang diharapkan**
            const expectedColumns = metodePenebusan === 'ipubers' ? expectedColumnsIpubers : expectedColumnsKartan;

            if (!expectedColumns.every((col, index) => col === actualColumns[index])) {
                console.error("❌ Struktur file tidak sesuai. Ditemukan:", actualColumns);
                return res.status(400).send({ message: 'Struktur file Excel tidak sesuai' });
            }

            // **Tentukan dari baris mana data harus dibaca**
            const startRow = metodePenebusan === 'ipubers' ? 2 : 2;
            let rows = worksheet.getRows(startRow, worksheet.rowCount - startRow);

            // **Hapus baris kosong**
            rows = rows.filter(row => row.getCell(1).value && row.getCell(2).value);

            console.log(`✅ Data siap diproses, total baris: ${rows.length}`);


            const data = rows.map((row) => {
                let tanggalTebus, tanggalExcel;
                const kodeTransaksi = metodePenebusan === 'ipubers' ? row.getCell(7).text : row.getCell(4).text;

                if (metodePenebusan === 'ipubers') {
                    tanggalExcel = row.getCell(18).text;

                    // if (tanggalExcel && typeof tanggalExcel === 'string') {
                    //     if (/^\d{1,2}-\d{1,2}-\d{4}$/.test(tanggalExcel)) {
                    //         const parts = tanggalExcel.split('-');
                    //         if (parts.length === 3) {
                    //             parts[0] = parts[0].padStart(2, '0');
                    //             parts[1] = parts[1].padStart(2, '0');
                    //             tanggalExcel = parts.join('-');
                    //         }
                    //         const [dd, mm, yyyy] = tanggalExcel.split('-');
                    //         tanggalTebus = `${yyyy}/${mm}/${dd}`;
                    //     } else {
                    //         tanggalTebus = null;
                    //     }
                    // } else {
                    //     tanggalTebus = null;
                    // }

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

                } else if (metodePenebusan === 'kartan') {
                    tanggalExcel = row.getCell(15).text;

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
                }

                console.log('Formatted Tanggal Tebus:', tanggalTebus);

                function generateKodeTransaksi() {
                    return 'TX-' + Date.now().toString(36) + '-' + Math.random().toString(36).substring(2, 7).toUpperCase();
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
                } else if (metodePenebusan === 'kartan') {
                    const mid = `'${row.getCell(4).text || ''}`;
                    const namaKios = row.getCell(5).text || '';
                    const kodeTransaksi = generateKodeTransaksi(); // otomatis buat kode unik

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

            const cleanText = (text) => {
                return text ? text.toLowerCase().trim().replace(/[^a-z0-9\s]/g, '') : '';
            };

            const cleanNamaKios = (text) => {
                if (!text) return "";

                // Ubah ke huruf besar
                let cleaned = text.toUpperCase();

                // Tambahkan spasi sebelum huruf besar jika diikuti huruf kecil (memisahkan kata yang nyambung)
                cleaned = cleaned.replace(/([A-Z])([A-Z][a-z])/g, '$1 $2');

                // Hapus karakter khusus: . , " - /
                cleaned = cleaned.replace(/[.,\"-\/]/g, "");

                // Hapus kata tidak penting
                const wordsToRemove = ["KPL", "KIOS", "UD", "KUD", "TOKO", "TK", "GAPOKTAN", "CV"];
                let regex = new RegExp(`\\b(${wordsToRemove.join("|")})\\b`, "g");
                cleaned = cleaned.replace(regex, "").trim();

                return cleaned;
            };

            try {
                console.log("🔍 Memulai proses upload...");

                // 1️⃣ Ambil kode transaksi yang sudah ada di database untuk menghindari duplikasi
                const existingTransactionsSet = new Set();
                const [existingData] = await db.query('SELECT kode_transaksi FROM verval');
                existingData.forEach(row => existingTransactionsSet.add(row.kode_transaksi));

                console.log(`✅ ${existingTransactionsSet.size} kode transaksi sudah ada di database.`);

                // 2️⃣ Ambil data dari main_mid untuk mencocokkan kode kios
                const [allMainMID] = await db.query('SELECT mid, kabupaten, kecamatan, nama_kios, kode_kios FROM main_mid');
                console.log(`✅ ${allMainMID.length} data dari main_mid dimuat.`);

                // 3️⃣ Ambil data dari master_mid jika tidak ditemukan di main_mid
                const [allMasterMID] = await db.query('SELECT mid, kode_kios FROM master_mid');
                console.log(`✅ ${allMasterMID.length} data dari master_mid dimuat.`);

                // 4️⃣ Buat struktur pencarian kode kios
                const midMap = new Map(); // Simpan kode_kios berdasarkan MID + Kabupaten
                const kabKecNamaMap = new Map(); // Simpan kode_kios berdasarkan Kabupaten + Kecamatan + Nama Kios
                const masterMidMap = new Map(); // Simpan kode_kios berdasarkan MID saja (dari master_mid)

                // Isi data dari main_mid
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

                // Isi data dari master_mid
                allMasterMID.forEach(({ mid, kode_kios }) => {
                    masterMidMap.set(cleanText(mid), kode_kios);
                });

                let values = [];

                // 5️⃣ Looping data untuk diproses
                for (let rowData of data) {
                    let kodeKios = rowData.kodeKios;

                    if (existingTransactionsSet.has(rowData.kodeTransaksi)) {
                        duplicateCount++;
                        continue;
                    }


                    if (rowData.metodePenebusan === 'kartan') {
                        const cleanMid = cleanText(rowData.mid);
                        const cleanKabupaten = cleanText(rowData.kabupaten);
                        const cleanKecamatan = cleanText(rowData.kecamatan);
                        const cleanedNamaKios = cleanNamaKios(rowData.namaKios); // Variabel baru

                        // 1️⃣ Cek di main_mid berdasarkan MID + Kabupaten
                        if (midMap.has(`${cleanMid}-${cleanKabupaten}`) && midMap.get(`${cleanMid}-${cleanKabupaten}`) !== "-") {
                            kodeKios = midMap.get(`${cleanMid}-${cleanKabupaten}`);
                        }
                        // 2️⃣ Jika masih tidak ditemukan, cek di main_mid berdasarkan Kabupaten + Kecamatan + Nama Kios
                        else if (kabKecNamaMap.has(`${cleanKabupaten}-${cleanKecamatan}-${cleanedNamaKios}`)) {
                            kodeKios = kabKecNamaMap.get(`${cleanKabupaten}-${cleanKecamatan}-${cleanedNamaKios}`);
                        }
                        // 3️⃣ Jika masih tidak ditemukan, cek di master_mid hanya dengan MID
                        else if (masterMidMap.has(cleanMid)) {
                            kodeKios = masterMidMap.get(cleanMid);
                        }
                        // 4️⃣ Jika semua pencarian gagal, tetap gunakan "-"
                        else {
                            console.log(`❌ Kode kios tidak ditemukan.`);
                        }
                    }

                    // 6️⃣ Masukkan data ke dalam array untuk di-insert
                    values.push([
                        rowData.kabupaten, rowData.kecamatan, rowData.poktan, kodeKios, rowData.namaKios, rowData.nik, rowData.namaPetani,
                        rowData.metodePenebusan, rowData.tanggalTebus, rowData.urea, rowData.npk, rowData.sp36,
                        rowData.za, rowData.npkFormula, rowData.organik, rowData.organikCair, rowData.kakao,
                        rowData.kodeTransaksi, rowData.status
                    ]);

                    insertedCount++;
                }

                // 7️⃣ Simpan ke database jika ada data yang valid
                if (values.length > 0) {
                    try {
                        console.log("🚀 Menyimpan data ke database...");
                        await db.query('START TRANSACTION');
                        await db.query(
                            `INSERT INTO verval 
                            (kabupaten, kecamatan, poktan, kode_kios, nama_kios, nik, nama_petani, metode_penebusan, tanggal_tebus, 
                            urea, npk, sp36, za, npk_formula, organik, organik_cair, kakao, kode_transaksi, status) 
                            VALUES ?`, [values]
                        );
                        await db.query('COMMIT');
                        console.log(`✅ ${insertedCount} data berhasil disimpan.`);
                    } catch (sqlError) {
                        await db.query('ROLLBACK');
                        console.error("❌ Error saat menyimpan data:", sqlError);
                        throw sqlError;
                    }
                }

            } catch (err) {
                console.error('❌ Error:', err);
                res.status(500).send({
                    message: '❌ Terjadi kesalahan dalam proses upload.',
                    error: err.message
                });
            }
        }

        res.status(200).send({
            message: `✅ Proses selesai. ${insertedCount} data ditambahkan, ${duplicateCount} data sudah ada.`
        });

    } catch (err) {
        console.error('❌ Error:', err);
        await db.query('ROLLBACK');
        res.status(500).send({
            message: '❌ Terjadi kesalahan dalam proses upload.',
            error: err.message
        });
    }
};

query = `SELECT COUNT(*) AS total_baris_boyolali
FROM (
    SELECT 
        COALESCE(e.kabupaten, v.kabupaten) AS kabupaten, 
        COALESCE(e.kecamatan, v.kecamatan) AS kecamatan,
        COALESCE(e.nik, v.nik) AS nik, 
        COALESCE(e.nama_petani, '') AS nama_petani, 
        COALESCE(e.kode_kios, v.kode_kios) AS kode_kios,
        COALESCE(e.desa, '') AS desa,
        COALESCE(e.poktan, '') AS poktan,
        COALESCE(e.tahun, v.tahun) AS tahun, 

        (COALESCE(e.urea,0) - COALESCE(v.tebus_urea,0)) AS sisa_urea,
        (COALESCE(e.npk,0) - COALESCE(v.tebus_npk,0)) AS sisa_npk,
        (COALESCE(e.npk_formula,0) - COALESCE(v.tebus_npk_formula,0)) AS sisa_npk_formula,
        (COALESCE(e.organik,0) - COALESCE(v.tebus_organik,0)) AS sisa_organik,
        (
            (COALESCE(e.urea,0) - COALESCE(v.tebus_urea,0)) +
            (COALESCE(e.npk,0) - COALESCE(v.tebus_npk,0)) +
            (COALESCE(e.npk_formula,0) - COALESCE(v.tebus_npk_formula,0)) +
            (COALESCE(e.organik,0) - COALESCE(v.tebus_organik,0))
        ) AS total_sisa

    FROM (
        -- Data dari ERDKK
        SELECT 
            e.kabupaten, e.kecamatan, e.nik, e.nama_petani, e.kode_kios,
            e.tahun, e.desa, e.poktan,
            SUM(e.urea) AS urea, 
            SUM(e.npk) AS npk, 
            SUM(e.npk_formula) AS npk_formula, 
            SUM(e.organik) AS organik
        FROM erdkk e
        WHERE e.tahun = 2025
          AND e.kabupaten IN ('BOYOLALI', 'KLATEN', 'SUKOHARJO', 'KARANGANYAR', 'WONOGIRI', 'SRAGEN', 
       'KOTA SURAKARTA', 'SLEMAN', 'BANTUL', 'GUNUNG KIDUL', 'KULON PROGO', 'KOTA YOGYAKARTA')
        GROUP BY e.kabupaten, e.kecamatan, e.nik, e.nama_petani, e.kode_kios, e.tahun, e.desa, e.poktan
        
        UNION ALL
        
        -- Data dari VERVAL yang tidak ada di ERDKK (alokasi = 0)
        SELECT 
            v.kabupaten, v.kecamatan, v.nik, '' AS nama_petani, v.kode_kios,
            v.tahun, '' AS desa, '' AS poktan,
            0 AS urea, 0 AS npk, 0 AS npk_formula, 0 AS organik
        FROM verval_summary v
        LEFT JOIN erdkk e 
            ON e.nik = v.nik
           AND e.kabupaten = v.kabupaten
           AND e.tahun = v.tahun
           AND e.kode_kios = v.kode_kios
        WHERE v.tahun = 2025
          AND v.kabupaten IN ('BOYOLALI', 'KLATEN', 'SUKOHARJO', 'KARANGANYAR', 'WONOGIRI', 'SRAGEN', 
       'KOTA SURAKARTA', 'SLEMAN', 'BANTUL', 'GUNUNG KIDUL', 'KULON PROGO', 'KOTA YOGYAKARTA')
          AND e.nik IS NULL
     ) AS combined

    LEFT JOIN verval_summary v 
        ON combined.nik = v.nik
       AND combined.kabupaten = v.kabupaten
       AND combined.tahun = v.tahun
       AND combined.kode_kios = v.kode_kios

    LEFT JOIN erdkk e 
        ON combined.nik = e.nik
       AND combined.kabupaten = e.kabupaten
       AND combined.tahun = e.tahun
       AND combined.kode_kios = e.kode_kios

    WHERE combined.tahun = 2025
      AND combined.kabupaten IN ('BOYOLALI', 'KLATEN', 'SUKOHARJO', 'KARANGANYAR', 'WONOGIRI', 'SRAGEN', 
       'KOTA SURAKARTA', 'SLEMAN', 'BANTUL', 'GUNUNG KIDUL', 'KULON PROGO', 'KOTA YOGYAKARTA')
) AS hasil_final; `


// export excel wcmvsverval
exports.exportExcelWcmVsVerval = async (req, res) => {
    const startTime = Date.now();
    let cacheStatus = 'MISS';
    let cacheKey = '';

    try {
        const { tahun, bulan, produk, kabupaten, status = 'ALL' } = req.query;

        // Validasi parameter wajib
        if (!tahun) {
            return res.status(400).json({ message: 'Parameter tahun wajib diisi' });
        }

        cacheKey = generateExcelCacheKey({ tahun, bulan, produk, kabupaten, status });
        console.log('📊 EXCEL EXPORT REQUEST:', new Date().toLocaleString('id-ID'));
        console.log('🔑 Cache Key:', cacheKey);

        try {
            const cachedExcel = await client.getBuffer(cacheKey);
            if (cachedExcel) {
                cacheStatus = 'HIT';
                console.log('✅ ✅ ✅ EXCEL DIAMBIL DARI CACHE ✅ ✅ ✅');
                
                res.set({
                    'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    'Content-Disposition': `attachment; filename=wcm_vs_verval_${tahun}${bulan ? '_' + bulan : ''}.xlsx`,
                    'X-Cache-Status': 'HIT',
                    'X-Cache-Key': cacheKey,
                    'X-Response-Time': (Date.now() - startTime) + 'ms'
                });

                return res.send(cachedExcel);
            }
        } catch (cacheError) {
            console.log('⚠️  Cache miss untuk excel:', cacheError.message);
        }

        const produkFilter = produk && produk !== 'ALL' ? produk.toUpperCase() : 'ALL';
        const statusFilter = status.toUpperCase();

        const params = [];
        let whereClauses = ['g.tahun = ?'];
        params.push(tahun);

        if (bulan && bulan !== 'ALL') {
            whereClauses.push('g.bulan = ?');
            params.push(bulan);
        }

        if (produk && produk !== 'ALL') {
            whereClauses.push('g.produk = ?');
            params.push(produkFilter);
        }

        if (kabupaten && kabupaten !== 'ALL') {
            whereClauses.push('g.kabupaten = ?');
            params.push(kabupaten);
        }

        const whereSQL = whereClauses.length > 0 ? ' WHERE ' + whereClauses.join(' AND ') : '';

        const baseQuery = `
    WITH
penebusan AS (
  SELECT 
    pd.kode_kios,
    pd.kecamatan,
    pd.kabupaten,
    pd.produk,
    MONTH(pd.tanggal_penyaluran) AS bulan,
    YEAR(pd.tanggal_penyaluran) AS tahun,
    pd.kode_distributor,
    pd.distributor,
    pd.nama_kios,
    pd.provinsi,
    SUM(pd.qty) AS penebusan
  FROM penyaluran_do pd
  GROUP BY 
    pd.kode_kios,
    pd.kecamatan,
    pd.kabupaten,
    pd.produk,
    MONTH(pd.tanggal_penyaluran),
    YEAR(pd.tanggal_penyaluran),
    pd.kode_distributor,
    pd.distributor,
    pd.nama_kios,
    pd.provinsi
),

stok_akhir_prev AS (
  SELECT 
    w.kode_kios,
    w.kecamatan,
    w.kabupaten,
    w.produk,
    CASE WHEN w.bulan = 12 THEN 1 ELSE w.bulan + 1 END AS bulan,
    CASE WHEN w.bulan = 12 THEN w.tahun + 1 ELSE w.tahun END AS tahun,
    w.kode_distributor,
    w.nama_kios,
    w.provinsi,
    w.distributor,
    SUM(w.stok_akhir) AS stok_awal
  FROM wcm w
  WHERE w.stok_akhir > 0
  GROUP BY 
    w.kode_kios,
    w.kecamatan,
    w.kabupaten,
    w.produk,
    CASE WHEN w.bulan = 12 THEN 1 ELSE w.bulan + 1 END,
    CASE WHEN w.bulan = 12 THEN w.tahun + 1 ELSE w.tahun END,
    w.kode_distributor,
    w.nama_kios,
    w.provinsi,
    w.distributor
),

verval AS (
  SELECT 
    vf.kode_kios,
    vf.kecamatan,
    vf.kabupaten,
    vf.produk,
    vf.tahun,
    vf.bulan,
    SUM(vf.penyaluran) AS total_penyaluran
  FROM verval_f6 vf
  GROUP BY 
    vf.kode_kios,
    vf.kecamatan,
    vf.kabupaten,
    vf.produk,
    vf.tahun,
    vf.bulan
),

gabungan AS (
  SELECT 
    p.kode_kios,
    p.kecamatan,
    p.kabupaten,
    p.produk,
    p.tahun,
    p.bulan,
    p.kode_distributor,
    p.nama_kios,
    p.provinsi,
    p.distributor
  FROM penebusan p

  UNION

  SELECT 
    s.kode_kios,
    s.kecamatan,
    s.kabupaten,
    s.produk,
    s.tahun,
    s.bulan,
    s.kode_distributor,
    s.nama_kios,
    s.provinsi,
    s.distributor
  FROM stok_akhir_prev s
)

SELECT 
  g.provinsi,
  g.kabupaten,
  g.kecamatan,
  g.kode_kios,
  g.nama_kios,
  g.kode_distributor,
  g.distributor,
  g.tahun,
  g.bulan,
  g.produk,

  ROUND(COALESCE(sa.stok_awal, 0) * 1000, 0) AS stok_awal_wcm,
  ROUND(COALESCE(p.penebusan, 0) * 1000, 0) AS penebusan_wcm,
  ROUND(COALESCE(wcm.penyaluran, 0) * 1000, 0) AS penyaluran_wcm,
  ROUND(COALESCE(v.total_penyaluran, 0), 0) AS penyaluran_verval,

  ROUND(
    (COALESCE(sa.stok_awal, 0) * 1000 + COALESCE(p.penebusan, 0) * 1000 - COALESCE(v.total_penyaluran, 0)),
    0
  ) AS stok_akhir_wcm,

  ROUND(COALESCE(wcm.penyaluran, 0) * 1000, 0) - ROUND(COALESCE(v.total_penyaluran, 0), 0) AS status_penyaluran

FROM gabungan g
LEFT JOIN penebusan p
  ON g.kode_kios = p.kode_kios
  AND g.kecamatan = p.kecamatan
  AND g.kabupaten = p.kabupaten
  AND g.produk = p.produk
  AND g.tahun = p.tahun
  AND g.bulan = p.bulan
  AND g.kode_distributor = p.kode_distributor

LEFT JOIN stok_akhir_prev sa
  ON g.kode_kios = sa.kode_kios
  AND g.kecamatan = sa.kecamatan
  AND g.kabupaten = sa.kabupaten
  AND g.produk = sa.produk
  AND g.tahun = sa.tahun
  AND g.bulan = sa.bulan
  AND g.kode_distributor = sa.kode_distributor

LEFT JOIN wcm
  ON wcm.kode_kios = g.kode_kios
  AND wcm.kecamatan = g.kecamatan
  AND wcm.kabupaten = g.kabupaten
  AND wcm.produk = g.produk
  AND wcm.tahun = g.tahun
  AND wcm.bulan = g.bulan
  AND wcm.kode_distributor = g.kode_distributor

LEFT JOIN verval v
  ON v.kode_kios = g.kode_kios
  AND v.kecamatan = g.kecamatan
  AND v.kabupaten = g.kabupaten
  AND v.produk = g.produk
  AND v.tahun = g.tahun
  AND v.bulan = g.bulan

${whereSQL}

ORDER BY g.kode_kios
`;
        // Filter status jika bukan ALL
        let finalQuery = baseQuery;
        if (statusFilter !== 'ALL') {
            const operator = statusFilter === 'SESUAI' ? '=' : '!=';
            finalQuery = `
        SELECT * FROM (${baseQuery}) AS filtered_table
        WHERE status_penyaluran ${operator} ?
    `;
            params.push(0);
        }


        const [data] = await db.query(finalQuery, params);

        // Create Excel workbook
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('WCM vs Verval');

        // Styling untuk header
        const headerStyle = {
            font: { bold: true, color: { argb: 'FFFFFFFF' } },
            alignment: { vertical: 'middle', horizontal: 'center', wrapText: true },
            border: {
                top: { style: 'thin', color: { argb: 'FF000000' } },
                left: { style: 'thin', color: { argb: 'FF000000' } },
                bottom: { style: 'thin', color: { argb: 'FF000000' } },
                right: { style: 'thin', color: { argb: 'FF000000' } }
            }
        };

        // Header row 1
        worksheet.mergeCells('A1:A2');
        worksheet.getCell('A1').value = 'Provinsi';
        worksheet.getCell('A1').style = { ...headerStyle, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF7F7F7F' } } };

        worksheet.mergeCells('B1:B2');
        worksheet.getCell('B1').value = 'Kabupaten';
        worksheet.getCell('B1').style = { ...headerStyle, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF7F7F7F' } } };

        worksheet.mergeCells('C1:C2');
        worksheet.getCell('C1').value = 'Kode Distributor';
        worksheet.getCell('C1').style = { ...headerStyle, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF7F7F7F' } } };

        worksheet.mergeCells('D1:D2');
        worksheet.getCell('D1').value = 'Distributor';
        worksheet.getCell('D1').style = { ...headerStyle, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF7F7F7F' } } };

        worksheet.mergeCells('E1:E2');
        worksheet.getCell('E1').value = 'Kecamatan';
        worksheet.getCell('E1').style = { ...headerStyle, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF7F7F7F' } } };

        worksheet.mergeCells('F1:F2');
        worksheet.getCell('F1').value = 'Kode Kios';
        worksheet.getCell('F1').style = { ...headerStyle, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF7F7F7F' } } };

        worksheet.mergeCells('G1:G2');
        worksheet.getCell('G1').value = 'Nama Kios';
        worksheet.getCell('G1').style = { ...headerStyle, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF7F7F7F' } } };

        worksheet.mergeCells('H1:H2');
        worksheet.getCell('H1').value = 'Produk';
        worksheet.getCell('H1').style = { ...headerStyle, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF7F7F7F' } } };

        worksheet.mergeCells('I1:I2');
        worksheet.getCell('I1').value = 'Bulan';
        worksheet.getCell('I1').style = { ...headerStyle, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF7F7F7F' } } };

        worksheet.mergeCells('J1:L1');
        worksheet.getCell('J1').value = 'WCM';
        worksheet.getCell('J1').style = { ...headerStyle, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' } } };

        worksheet.getCell('M1').value = 'VERVAL';
        worksheet.getCell('M1').style = { ...headerStyle, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFEB9C' } }, font: { ...headerStyle.font, color: { argb: 'FF000000' } } };
        worksheet.getCell('M2').value = 'Penyaluran';
        worksheet.getCell('M2').style = { ...headerStyle, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFEB9C' } }, font: { ...headerStyle.font, color: { argb: 'FF000000' } } };

        worksheet.getCell('N1').value = 'WCM';
        worksheet.getCell('N1').style = { ...headerStyle, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' } } };
        worksheet.getCell('N2').value = 'Stok Akhir';
        worksheet.getCell('N2').style = { ...headerStyle, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' } } };

        worksheet.mergeCells('O1:O2');
        worksheet.getCell('O1').value = 'Selisih Penyaluran';
        worksheet.getCell('O1').style = { ...headerStyle, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF7F7F7F' } } };

        // Header row 2 (hanya untuk kolom WCM)
        worksheet.getCell('J2').value = 'Stok Awal';
        worksheet.getCell('J2').style = { ...headerStyle, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' } } };

        worksheet.getCell('K2').value = 'Penebusan';
        worksheet.getCell('K2').style = { ...headerStyle, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' } } };

        worksheet.getCell('L2').value = 'Penyaluran';
        worksheet.getCell('L2').style = { ...headerStyle, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' } } };

        // Add data rows
        data.forEach((row, index) => {
            const dataRow = worksheet.addRow([
                row.provinsi,
                row.kabupaten,
                row.kode_distributor,
                row.distributor,
                row.kecamatan,
                row.kode_kios,
                row.nama_kios,
                row.produk,
                row.bulan,
                row.stok_awal_wcm,
                row.penebusan_wcm,
                row.penyaluran_wcm,
                row.penyaluran_verval,
                row.stok_akhir_wcm,
                row.status_penyaluran
            ]);

            // Format angka dengan pemisah ribuan
            [9, 10, 11, 12, 13].forEach(col => {
                dataRow.getCell(col).numFmt = '#,##0';
            });

            // Warna status
            if (row.status_penyaluran === 'Sesuai') {
                dataRow.getCell(15).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC6EFCE' } };
            } else {
                dataRow.getCell(15).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFC7CE' } };
            }
        });

        // Set column widths
        worksheet.columns = [
            { width: 15 }, // Provinsi
            { width: 15 }, // Kabupaten
            { width: 15 }, // Kode Distributor
            { width: 20 }, // Distributor
            { width: 15 }, // Kecamatan
            { width: 15 }, // Kode Kios
            { width: 25 }, // Nama Kios
            { width: 15 }, // Produk
            { width: 10 }, // Bulan
            { width: 12 }, // Stok Awal WCM
            { width: 12 }, // Penebusan WCM
            { width: 12 }, // Penyaluran WCM
            { width: 12 }, // Penyaluran Verval
            { width: 12 }, // Stok Akhir WCM
            { width: 12 }  // Status
        ];

        // Freeze header rows
        worksheet.views = [{ state: 'frozen', xSplit: 0, ySplit: 2 }];

        // Set response headers
        res.setHeader(
            'Content-Type',
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        );
        res.setHeader(
            'Content-Disposition',
            `attachment; filename=wcm_vs_verval_${tahun}${bulan ? '_' + bulan : ''}.xlsx`
        );
        
        const excelBuffer = await workbook.xlsx.writeBuffer();

        // Simpan ke cache (1 jam = 3600 detik)
        try {
            await client.setEx(cacheKey, 3600, excelBuffer.toString('base64'));
            console.log('💾 Excel disimpan ke cache untuk 1 jam');
        } catch (cacheError) {
            console.log('❌ Gagal menyimpan excel ke cache:', cacheError.message);
        }

        // Kirim response
        res.set({
            'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'Content-Disposition': `attachment; filename=wcm_vs_verval_${tahun}${bulan ? '_' + bulan : ''}.xlsx`,
            'X-Cache-Status': cacheStatus,
            'X-Cache-Key': cacheKey,
            'X-Response-Time': (Date.now() - startTime) + 'ms'
        });

        res.send(excelBuffer);

    } catch (error) {
        console.error('Error in exportExcelWcmVsVerval:', error);
        res.status(500).json({
            message: 'Terjadi kesalahan saat mengekspor data.',
            error: error.message
        });
    }
};
