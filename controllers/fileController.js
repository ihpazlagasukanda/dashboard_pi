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
            const expectedColumnsKartan = ['NO', 'Kabupaten', 'Kecamatan', 'Kodetrx', 'Nik', 'Nama petani', 'Urea', 'NPK', 'SP36', 'ZA', 'NPK Formula', 'Organik', 'Organik Cair', 'Alasan Tolak', 'Tanggal Tebus', 'MID', 'Nama Kios'];
            const expectedColumnsIpubers = ['Kabupaten', 'Kecamatan', 'No Transaksi', 'Poktan', 'Kode Kios', 'Nama Kios', 'NIK', 'Nama Petani', 'Urea', 'NPK', 'SP36', 'ZA', 'NPK Formula', 'Organik', 'Organik Cair', 'Tanggal Tebus (tanggal-bulan-tahun)', 'tanggal Input', 'Tipe Tebus', 'Nik Perwakilan', 'Bukti trx', 'Status'];

            const headerRow = metodePenebusan === 'ipubers' ? 1 : 3; // Baris header
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
            const startRow = metodePenebusan === 'ipubers' ? 2 : 4;
            let rows = worksheet.getRows(startRow, worksheet.rowCount - startRow);

            // **Hapus baris kosong**
            rows = rows.filter(row => row.getCell(1).value && row.getCell(2).value);

            console.log(`✅ Data siap diproses, total baris: ${rows.length}`);


            const data = rows.map((row) => {
                let tanggalTebus, tanggalExcel;
                const kodeTransaksi = metodePenebusan === 'ipubers' ? row.getCell(3).text : row.getCell(4).text;

                if (metodePenebusan === 'ipubers') {
                    tanggalExcel = row.getCell(16).text;

                    if (tanggalExcel && typeof tanggalExcel === 'string') {
                        if (/^\d{1,2}-\d{1,2}-\d{4}$/.test(tanggalExcel)) {
                            const parts = tanggalExcel.split('-');
                            if (parts.length === 3) {
                                parts[0] = parts[0].padStart(2, '0');
                                parts[1] = parts[1].padStart(2, '0');
                                tanggalExcel = parts.join('-');
                            }
                            const [dd, mm, yyyy] = tanggalExcel.split('-');
                            tanggalTebus = `${yyyy}/${mm}/${dd}`;
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

                if (metodePenebusan === 'ipubers') {
                    return {
                        kabupaten: row.getCell(1).text || '',
                        kecamatan: row.getCell(2).text || '',
                        kodeTransaksi: kodeTransaksi,
                        poktan: row.getCell(4).text || '',
                        kodeKios: row.getCell(5).text || '',
                        namaKios: row.getCell(6).text || '',
                        nik: row.getCell(7).text || '',
                        namaPetani: row.getCell(8).text || '',
                        metodePenebusan: 'ipubers',
                        tanggalTebus: tanggalTebus,
                        urea: convertToNumber(row.getCell(9).text || '0'),
                        npk: convertToNumber(row.getCell(10).text || '0'),
                        sp36: convertToNumber(row.getCell(11).text || '0'),
                        za: convertToNumber(row.getCell(12).text || '0'),
                        npkFormula: convertToNumber(row.getCell(13).text || '0'),
                        organik: convertToNumber(row.getCell(14).text || '0'),
                        organikCair: convertToNumber(row.getCell(15).text || '0'),
                        kakao: convertToNumber(row.getCell(22).text || '0'),
                        status: row.getCell(21).text || 'Tidak Diketahui'
                    };
                } else if (metodePenebusan === 'kartan') {
                    const mid = row.getCell(16).text;
                    const namaKios = row.getCell(17).text || '';

                    return {
                        kabupaten: row.getCell(2).text || '',
                        kecamatan: row.getCell(3).text || '',
                        kodeTransaksi: kodeTransaksi,
                        poktan: '',
                        kodeKios: '-',
                        nik: row.getCell(5).text || '',
                        namaPetani: row.getCell(6).text || '',
                        metodePenebusan: 'kartan',
                        tanggalTebus: tanggalTebus,
                        urea: convertToNumber(row.getCell(7).text || '0'),
                        npk: convertToNumber(row.getCell(8).text || '0'),
                        sp36: convertToNumber(row.getCell(9).text || '0'),
                        za: convertToNumber(row.getCell(10).text || '0'),
                        npkFormula: convertToNumber(row.getCell(11).text || '0'),
                        organik: convertToNumber(row.getCell(12).text || '0'),
                        organikCair: convertToNumber(row.getCell(13).text || '0'),
                        kakao: convertToNumber(row.getCell(14).text || '0'),
                        mid: mid,
                        namaKios: namaKios,
                        status: 'kartan'
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