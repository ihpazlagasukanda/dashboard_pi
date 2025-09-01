const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');
const db = require('../config/db');

// Fungsi untuk generate semua laporan yang diperlukan
// async function generateAllReports() {
//     try {
//         const baseDir = path.join(__dirname, '../temp_exports');
//         if (!fs.existsSync(baseDir)) {
//             fs.mkdirSync(baseDir, { recursive: true });
//         }

//         const kabupatenList = [
//             'KOTA YOGYAKARTA',
//             'SLEMAN',
//             'BANTUL',
//             'GUNUNG KIDUL',
//             'KULON PROGO',
//              'BANJARNEGARA',
//     'BANYUMAS',
//     'BATANG',
//     'BLORA',
//     'BOYOLALI',
//     'BREBES',
//     'CILACAP',
//     'DEMAK',
//     'GROBOGAN',
//     'JEPARA',
//     'KARANGANYAR',
//     'KEBUMEN',
//     'KENDAL',
//     'KLATEN',
//     'KOTA MAGELANG',
//     'KOTA PEKALONGAN',
//     'KOTA SALATIGA',
//     'KOTA SEMARANG',
//     'KOTA SURAKARTA',
//     'KOTA TEGAL',
//     'KUDUS',
//     'MAGELANG',
//     'PATI',
//     'PEKALONGAN',
//     'PEMALANG',
//     'PURBALINGGA',
//     'PURWOREJO',
//     'REMBANG',
//     'SALATIGA',
//     'SEMARANG',
//     'SRAGEN',
//     'SUKOHARJO',
//     'TEGAL',
//     'TEMANGGUNG',
//     'WONOGIRI',
//     'WONOSOBO'
//         ];

//         const tahunList = [2025];

//         for (const kabupaten of kabupatenList) {
//             for (const tahun of tahunList) {
//                 console.log(`[INFO] Generating report for ${kabupaten} - ${tahun}`);
//                 await generateReport(kabupaten, tahun, baseDir);
//             }
//         }

//         // Generate laporan ALL (tanpa filter)
//         console.log('[INFO] Generating full report (ALL)...');
//         await generateReport(null, null, baseDir);

//         console.log('[DONE] Semua laporan telah digenerate');
//     } catch (error) {
//         console.error('[ERROR] Gagal generate reports:', error);
//     }
// }

const kabupatenList = [
  'KOTA YOGYAKARTA', 'SLEMAN', 'BANTUL', 'GUNUNG KIDUL', 'KULON PROGO',
  'BANJARNEGARA', 'BANYUMAS', 'BATANG', 'BLORA', 'BOYOLALI',
  'BREBES', 'CILACAP', 'DEMAK', 'GROBOGAN', 'JEPARA',
  'KARANGANYAR', 'KEBUMEN', 'KENDAL', 'KLATEN', 'KOTA MAGELANG',
  'KOTA PEKALONGAN', 'KOTA SALATIGA', 'KOTA SEMARANG', 'KOTA SURAKARTA', 'KOTA TEGAL',
  'KUDUS', 'MAGELANG', 'PATI', 'PEKALONGAN', 'PEMALANG',
  'PURBALINGGA', 'PURWOREJO', 'REMBANG', 'SALATIGA', 'SEMARANG',
  'SRAGEN', 'SUKOHARJO', 'TEGAL', 'TEMANGGUNG', 'WONOGIRI', 'WONOSOBO'
];

const tahunList = [2025];

async function generateAllReports(group, baseDir) {
  try {
    if (!fs.existsSync(baseDir)) {
      fs.mkdirSync(baseDir, { recursive: true });
    }

    let kabListGroup = [];

    if (group === 'all') {
      kabListGroup = kabupatenList;
    } else {
      const groupSize = Math.ceil(kabupatenList.length / 4);
      const start = group * groupSize;
      const end = start + groupSize;
      kabListGroup = kabupatenList.slice(start, end);
    }

    for (const kabupaten of kabListGroup) {
      for (const tahun of tahunList) {
        console.log(`[INFO] Generating report for ${kabupaten} - ${tahun}`);
        await generateReport(kabupaten, tahun, baseDir);
      }
    }

    if (group === 'all') {
      console.log('[INFO] Generating full report (ALL tanpa filter)...');
      await generateReport(null, null, baseDir);
    }

    console.log(`[DONE] Grup ${group} selesai.`);
  } catch (error) {
    console.error(`[ERROR] generateAllReports Grup ${group}:`, error);
  }
}

async function generateReport(kabupaten, tahun, baseDir) {
    try {
        // ===================== 1. SETUP QUERY =====================
        let query = `
        SELECT 
    COALESCE(e.kabupaten, v.kabupaten) AS kabupaten, 
    COALESCE(e.kecamatan, v.kecamatan) AS kecamatan,
    COALESCE(e.nik, v.nik) AS nik, 
    COALESCE(e.nama_petani, v.nama_petani) AS nama_petani, 
    COALESCE(e.kode_kios, v.kode_kios) AS kode_kios,
    COALESCE(e.desa, '') AS desa,
    COALESCE(e.poktan, v.poktan) AS poktan,
    COALESCE(e.tahun, v.tahun) AS tahun, 
    COALESCE(e.urea, 0) AS urea, 
    COALESCE(e.npk, 0) AS npk, 
    COALESCE(e.npk_formula, 0) AS npk_formula, 
    COALESCE(e.organik, 0) AS organik,

    COALESCE(v.tebus_urea, 0) AS tebus_urea,
    COALESCE(v.tebus_npk, 0) AS tebus_npk,
    COALESCE(v.tebus_npk_formula, 0) AS tebus_npk_formula,
    COALESCE(v.tebus_organik, 0) AS tebus_organik,

    (COALESCE(e.urea, 0) - COALESCE(v.tebus_urea, 0)) AS sisa_urea,
    (COALESCE(e.npk, 0) - COALESCE(v.tebus_npk, 0)) AS sisa_npk,
    (COALESCE(e.npk_formula, 0) - COALESCE(v.tebus_npk_formula, 0)) AS sisa_npk_formula,
    (COALESCE(e.organik, 0) - COALESCE(v.tebus_organik, 0)) AS sisa_organik,

    -- distribusi bulanan
    COALESCE(t.jan_urea, 0) AS jan_urea,
    COALESCE(t.feb_urea, 0) AS feb_urea,
    COALESCE(t.mar_urea, 0) AS mar_urea,
    COALESCE(t.apr_urea, 0) AS apr_urea,
    COALESCE(t.mei_urea, 0) AS mei_urea,
    COALESCE(t.jun_urea, 0) AS jun_urea,
    COALESCE(t.jul_urea, 0) AS jul_urea,
    COALESCE(t.agu_urea, 0) AS agu_urea,
    COALESCE(t.sep_urea, 0) AS sep_urea,
    COALESCE(t.okt_urea, 0) AS okt_urea,
    COALESCE(t.nov_urea, 0) AS nov_urea,
    COALESCE(t.des_urea, 0) AS des_urea,

    COALESCE(t.jan_npk, 0) AS jan_npk,
    COALESCE(t.feb_npk, 0) AS feb_npk,
    COALESCE(t.mar_npk, 0) AS mar_npk,
    COALESCE(t.apr_npk, 0) AS apr_npk,
    COALESCE(t.mei_npk, 0) AS mei_npk,
    COALESCE(t.jun_npk, 0) AS jun_npk,
    COALESCE(t.jul_npk, 0) AS jul_npk,
    COALESCE(t.agu_npk, 0) AS agu_npk,
    COALESCE(t.sep_npk, 0) AS sep_npk,
    COALESCE(t.okt_npk, 0) AS okt_npk,
    COALESCE(t.nov_npk, 0) AS nov_npk,
    COALESCE(t.des_npk, 0) AS des_npk,

    COALESCE(t.jan_npk_formula, 0) AS jan_npk_formula,
    COALESCE(t.feb_npk_formula, 0) AS feb_npk_formula,
    COALESCE(t.mar_npk_formula, 0) AS mar_npk_formula,
    COALESCE(t.apr_npk_formula, 0) AS apr_npk_formula,
    COALESCE(t.mei_npk_formula, 0) AS mei_npk_formula,
    COALESCE(t.jun_npk_formula, 0) AS jun_npk_formula,
    COALESCE(t.jul_npk_formula, 0) AS jul_npk_formula,
    COALESCE(t.agu_npk_formula, 0) AS agu_npk_formula,
    COALESCE(t.sep_npk_formula, 0) AS sep_npk_formula,
    COALESCE(t.okt_npk_formula, 0) AS okt_npk_formula,
    COALESCE(t.nov_npk_formula, 0) AS nov_npk_formula,
    COALESCE(t.des_npk_formula, 0) AS des_npk_formula,

    COALESCE(t.jan_organik, 0) AS jan_organik,
    COALESCE(t.feb_organik, 0) AS feb_organik,
    COALESCE(t.mar_organik, 0) AS mar_organik,
    COALESCE(t.apr_organik, 0) AS apr_organik,
    COALESCE(t.mei_organik, 0) AS mei_organik,
    COALESCE(t.jun_organik, 0) AS jun_organik,
    COALESCE(t.jul_organik, 0) AS jul_organik,
    COALESCE(t.agu_organik, 0) AS agu_organik,
    COALESCE(t.sep_organik, 0) AS sep_organik,
    COALESCE(t.okt_organik, 0) AS okt_organik,
    COALESCE(t.nov_organik, 0) AS nov_organik,
    COALESCE(t.des_organik, 0) AS des_organik

FROM (
    -- Data dari ERDKK
    SELECT 
        e.kabupaten, e.kecamatan, e.nik, e.nama_petani, e.kode_kios, e.nama_kios,
        e.tahun, e.desa, e.poktan,
        e.urea, e.npk, e.npk_formula, e.organik
    FROM erdkk e
    
    UNION ALL
    
    -- Data dari VERVAL yang tidak ada di ERDKK
    SELECT 
        v.kabupaten, v.kecamatan, v.nik, v.nama_petani, v.kode_kios, v.nama_kios,
        v.tahun, '' AS desa, v.poktan,   -- desa diisi kosong
        0 AS urea, 0 AS npk, 0 AS npk_formula, 0 AS organik
    FROM verval_summary v
    LEFT JOIN erdkk e 
        ON e.nik = v.nik
       AND e.kabupaten = v.kabupaten
       AND e.tahun = v.tahun
       AND e.kecamatan = v.kecamatan
       AND e.kode_kios = v.kode_kios
    WHERE e.nik IS NULL
) AS combined

LEFT JOIN erdkk e 
    ON combined.nik = e.nik
   AND combined.kabupaten = e.kabupaten
   AND combined.tahun = e.tahun
   AND combined.kecamatan = e.kecamatan
   AND combined.kode_kios = e.kode_kios

LEFT JOIN verval_summary v 
    ON combined.nik = v.nik
   AND combined.kabupaten = v.kabupaten
   AND combined.tahun = v.tahun
   AND combined.kecamatan = v.kecamatan
   AND combined.kode_kios = v.kode_kios

LEFT JOIN tebusan_per_bulan t 
    ON combined.nik = t.nik
   AND combined.kabupaten = t.kabupaten
   AND combined.tahun = t.tahun
   AND combined.kecamatan = t.kecamatan
   AND combined.kode_kios = t.kode_kios

WHERE 1=1
        `;

        let params = [];

        if (tahun) {
            query += " AND combined.tahun = ?";
            params.push(tahun);
        }

        if (kabupaten) {
            query += " AND combined.kabupaten = ?";
            params.push(kabupaten);
        }

        // ... (kode selanjutnya tetap sama seperti yang Anda berikan)
        let kab = kabupaten ? kabupaten.replace(/\s+/g, '_').toUpperCase() : 'ALL';
        let thn = tahun || '';
        let fileName = `data_summary_${kab}${thn ? `_${thn}` : ''}.xlsx`;

        // ===================== 2. BUAT WORKBOOK =====================
        const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({
            filename: path.join(baseDir, fileName),
            useStyles: true,
            useSharedStrings: true
        });
        const worksheet = workbook.addWorksheet("Summary");

        // ... (kode untuk setup header, styling, dan proses data tetap sama)
        // ===================== 3. SETUP HEADER =====================
        // Style border
        const borderStyle = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };

        // Merge cells untuk header utama
        worksheet.mergeCells('A1:A2'); // Kabupaten
        worksheet.mergeCells('B1:B2'); // Kecamatan
        worksheet.mergeCells('C1:C2'); // Desa
        worksheet.mergeCells('D1:D2'); // NIK
        worksheet.mergeCells('E1:E2'); // Nama Petani
        worksheet.mergeCells('F1:F2'); // Kode Kios
        worksheet.mergeCells('G1:G2'); // Poktan
        worksheet.mergeCells('H1:K1'); // Alokasi
        worksheet.mergeCells('L1:O1'); // Sisa
        worksheet.mergeCells('P1:S1'); // Tebusan

        // Merge cells untuk header bulan
        const bulanHeaders = ['T1:W1', 'X1:AA1', 'AB1:AE1', 'AF1:AI1', 'AJ1:AM1', 'AN1:AQ1',
            'AR1:AU1', 'AV1:AY1', 'AZ1:BC1', 'BD1:BG1', 'BH1:BK1', 'BL1:BO1'];
        const bulanNames = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
            'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];

        bulanHeaders.forEach(range => worksheet.mergeCells(range));

        // Set value dan style untuk header utama
        worksheet.getCell("H1").value = "Alokasi";
        worksheet.getCell("L1").value = "Sisa";
        worksheet.getCell("P1").value = "Tebusan";

        // Style untuk header utama
        ["H1", "L1", "P1"].forEach(cell => {
            worksheet.getCell(cell).alignment = { horizontal: "center", vertical: "middle" };
            worksheet.getCell(cell).font = { bold: true, size: 12 };
            worksheet.getCell(cell).fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFD9D9D9' }
            };
        });

        // Set value dan style untuk header bulan
        bulanHeaders.forEach((range, index) => {
            const cell = worksheet.getCell(range.split(':')[0]);
            cell.value = bulanNames[index];
            cell.alignment = { horizontal: "center", vertical: "middle" };
            cell.font = { bold: true, size: 12, color: { argb: 'FFFFFFFF' } };
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFFF0000' }
            };
        });

        // Header kolom
        const subHeaders = [
            'Kabupaten', 'Kecamatan', 'Desa', 'Nama Poktan', 'NIK', 'Nama Petani', 'Kode Kios',
            'Urea', 'NPK', 'NPK Formula', 'Organik', // Alokasi
            'Urea', 'NPK', 'NPK Formula', 'Organik', // Sisa
            'Urea', 'NPK', 'NPK Formula', 'Organik', // Tebusan
            'Urea', 'NPK', 'NPK Formula', 'Organik', // Januari
            'Urea', 'NPK', 'NPK Formula', 'Organik', // Februari
            'Urea', 'NPK', 'NPK Formula', 'Organik', // Maret
            'Urea', 'NPK', 'NPK Formula', 'Organik', // April
            'Urea', 'NPK', 'NPK Formula', 'Organik', // Mei
            'Urea', 'NPK', 'NPK Formula', 'Organik', // Juni
            'Urea', 'NPK', 'NPK Formula', 'Organik', // Juli
            'Urea', 'NPK', 'NPK Formula', 'Organik', // Agustus
            'Urea', 'NPK', 'NPK Formula', 'Organik', // September
            'Urea', 'NPK', 'NPK Formula', 'Organik', // Oktober
            'Urea', 'NPK', 'NPK Formula', 'Organik', // November
            'Urea', 'NPK', 'NPK Formula', 'Organik'  // Desember
        ];

        // Tambahkan header row
        const headerRow = worksheet.addRow(subHeaders);
        headerRow.font = { bold: true, size: 12 };
        headerRow.alignment = { horizontal: 'center' };
        headerRow.eachCell((cell) => {
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFE6E6E6' }
            };
            cell.border = borderStyle;
        });

        // ===================== 4. TAMBAHKAN BARIS TOTAL DI BAWAH HEADER =====================
        // Inisialisasi variabel total
        const totals = {
            urea: 0, npk: 0, npk_formula: 0, organik: 0,
            sisa_urea: 0, sisa_npk: 0, sisa_npk_formula: 0, sisa_organik: 0,
            tebus_urea: 0, tebus_npk: 0, tebus_npk_formula: 0, tebus_organik: 0,
            jan_urea: 0, jan_npk: 0, jan_npk_formula: 0, jan_organik: 0,
            feb_urea: 0, feb_npk: 0, feb_npk_formula: 0, feb_organik: 0,
            mar_urea: 0, mar_npk: 0, mar_npk_formula: 0, mar_organik: 0,
            apr_urea: 0, apr_npk: 0, apr_npk_formula: 0, apr_organik: 0,
            mei_urea: 0, mei_npk: 0, mei_npk_formula: 0, mei_organik: 0,
            jun_urea: 0, jun_npk: 0, jun_npk_formula: 0, jun_organik: 0,
            jul_urea: 0, jul_npk: 0, jul_npk_formula: 0, jul_organik: 0,
            agu_urea: 0, agu_npk: 0, agu_npk_formula: 0, agu_organik: 0,
            sep_urea: 0, sep_npk: 0, sep_npk_formula: 0, sep_organik: 0,
            okt_urea: 0, okt_npk: 0, okt_npk_formula: 0, okt_organik: 0,
            nov_urea: 0, nov_npk: 0, nov_npk_formula: 0, nov_organik: 0,
            des_urea: 0, des_npk: 0, des_npk_formula: 0, des_organik: 0
        };

        // Tambahkan baris total kosong terlebih dahulu
        const totalRow = worksheet.addRow([
            'TOTAL', '', '', '', '', '', '',
            totals.urea.toLocaleString(),
            totals.npk.toLocaleString(),
            totals.npk_formula.toLocaleString(),
            totals.organik.toLocaleString(),
            totals.sisa_urea.toLocaleString(),
            totals.sisa_npk.toLocaleString(),
            totals.sisa_npk_formula.toLocaleString(),
            totals.sisa_organik.toLocaleString(),
            totals.tebus_urea.toLocaleString(),
            totals.tebus_npk.toLocaleString(),
            totals.tebus_npk_formula.toLocaleString(),
            totals.tebus_organik.toLocaleString(),
            totals.jan_urea.toLocaleString(),
            totals.jan_npk.toLocaleString(),
            totals.jan_npk_formula.toLocaleString(),
            totals.jan_organik.toLocaleString(),
            totals.feb_urea.toLocaleString(),
            totals.feb_npk.toLocaleString(),
            totals.feb_npk_formula.toLocaleString(),
            totals.feb_organik.toLocaleString(),
            totals.mar_urea.toLocaleString(),
            totals.mar_npk.toLocaleString(),
            totals.mar_npk_formula.toLocaleString(),
            totals.mar_organik.toLocaleString(),
            totals.apr_urea.toLocaleString(),
            totals.apr_npk.toLocaleString(),
            totals.apr_npk_formula.toLocaleString(),
            totals.apr_organik.toLocaleString(),
            totals.mei_urea.toLocaleString(),
            totals.mei_npk.toLocaleString(),
            totals.mei_npk_formula.toLocaleString(),
            totals.mei_organik.toLocaleString(),
            totals.jun_urea.toLocaleString(),
            totals.jun_npk.toLocaleString(),
            totals.jun_npk_formula.toLocaleString(),
            totals.jun_organik.toLocaleString(),
            totals.jul_urea.toLocaleString(),
            totals.jul_npk.toLocaleString(),
            totals.jul_npk_formula.toLocaleString(),
            totals.jul_organik.toLocaleString(),
            totals.agu_urea.toLocaleString(),
            totals.agu_npk.toLocaleString(),
            totals.agu_npk_formula.toLocaleString(),
            totals.agu_organik.toLocaleString(),
            totals.sep_urea.toLocaleString(),
            totals.sep_npk.toLocaleString(),
            totals.sep_npk_formula.toLocaleString(),
            totals.sep_organik.toLocaleString(),
            totals.okt_urea.toLocaleString(),
            totals.okt_npk.toLocaleString(),
            totals.okt_npk_formula.toLocaleString(),
            totals.okt_organik.toLocaleString(),
            totals.nov_urea.toLocaleString(),
            totals.nov_npk.toLocaleString(),
            totals.nov_npk_formula.toLocaleString(),
            totals.nov_organik.toLocaleString(),
            totals.des_urea.toLocaleString(),
            totals.des_npk.toLocaleString(),
            totals.des_npk_formula.toLocaleString(),
            totals.des_organik.toLocaleString()
        ]);

        // Style untuk baris total
        totalRow.font = { bold: true };
        totalRow.alignment = { horizontal: 'center' };
        totalRow.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFC0C0C0' }
        };
        totalRow.eachCell(cell => {
            cell.border = borderStyle;
        });
        totalRow.commit();

        // ===================== 5. PROSES DATA BATCH =====================
        const batchSize = 900;
        let offset = 0;
        let hasMoreData = true;

        while (hasMoreData) {
            const batchQuery = `${query} LIMIT ${batchSize} OFFSET ${offset}`;
            const [rows] = await db.query(batchQuery, params);

            if (rows.length === 0) {
                hasMoreData = false;
                break;
            }

            // Proses setiap baris dalam batch
            for (const row of rows) {
                // Update totals
                Object.keys(totals).forEach(key => {
                    totals[key] += parseFloat(row[key] || 0);
                });

                // Tambahkan baris data
                const dataRow = worksheet.addRow([
                    row.kabupaten,
                    row.kecamatan,
                    row.desa || '',
                    row.poktan || '',
                    row.nik,
                    row.nama_petani,
                    row.kode_kios,
                    row.urea, row.npk, row.npk_formula, row.organik,
                    row.sisa_urea, row.sisa_npk, row.sisa_npk_formula, row.sisa_organik,
                    row.tebus_urea, row.tebus_npk, row.tebus_npk_formula, row.tebus_organik,
                    row.jan_urea, row.jan_npk, row.jan_npk_formula, row.jan_organik,
                    row.feb_urea, row.feb_npk, row.feb_npk_formula, row.feb_organik,
                    row.mar_urea, row.mar_npk, row.mar_npk_formula, row.mar_organik,
                    row.apr_urea, row.apr_npk, row.apr_npk_formula, row.apr_organik,
                    row.mei_urea, row.mei_npk, row.mei_npk_formula, row.mei_organik,
                    row.jun_urea, row.jun_npk, row.jun_npk_formula, row.jun_organik,
                    row.jul_urea, row.jul_npk, row.jul_npk_formula, row.jul_organik,
                    row.agu_urea, row.agu_npk, row.agu_npk_formula, row.agu_organik,
                    row.sep_urea, row.sep_npk, row.sep_npk_formula, row.sep_organik,
                    row.okt_urea, row.okt_npk, row.okt_npk_formula, row.okt_organik,
                    row.nov_urea, row.nov_npk, row.nov_npk_formula, row.nov_organik,
                    row.des_urea, row.des_npk, row.des_npk_formula, row.des_organik
                ]);

                dataRow.commit();
                dataRow.eachCell(cell => {
                    cell.border = borderStyle;
                });
            }

            offset += batchSize;
            console.log(`Memproses batch ${offset / batchSize}...`);
        }

       // ===================== 6. UPDATE BARIS TOTAL DENGAN DATA AKTUAL =====================

        // Tambahkan kembali baris total dengan data yang sudah dihitung
        const updatedTotalRow = worksheet.addRow([
            'TOTAL', '', '', '', '', '', '',
            totals.urea.toLocaleString(),
            totals.npk.toLocaleString(),
            totals.npk_formula.toLocaleString(),
            totals.organik.toLocaleString(),
            totals.sisa_urea.toLocaleString(),
            totals.sisa_npk.toLocaleString(),
            totals.sisa_npk_formula.toLocaleString(),
            totals.sisa_organik.toLocaleString(),
            totals.tebus_urea.toLocaleString(),
            totals.tebus_npk.toLocaleString(),
            totals.tebus_npk_formula.toLocaleString(),
            totals.tebus_organik.toLocaleString(),
            totals.jan_urea.toLocaleString(),
            totals.jan_npk.toLocaleString(),
            totals.jan_npk_formula.toLocaleString(),
            totals.jan_organik.toLocaleString(),
            totals.feb_urea.toLocaleString(),
            totals.feb_npk.toLocaleString(),
            totals.feb_npk_formula.toLocaleString(),
            totals.feb_organik.toLocaleString(),
            totals.mar_urea.toLocaleString(),
            totals.mar_npk.toLocaleString(),
            totals.mar_npk_formula.toLocaleString(),
            totals.mar_organik.toLocaleString(),
            totals.apr_urea.toLocaleString(),
            totals.apr_npk.toLocaleString(),
            totals.apr_npk_formula.toLocaleString(),
            totals.apr_organik.toLocaleString(),
            totals.mei_urea.toLocaleString(),
            totals.mei_npk.toLocaleString(),
            totals.mei_npk_formula.toLocaleString(),
            totals.mei_organik.toLocaleString(),
            totals.jun_urea.toLocaleString(),
            totals.jun_npk.toLocaleString(),
            totals.jun_npk_formula.toLocaleString(),
            totals.jun_organik.toLocaleString(),
            totals.jul_urea.toLocaleString(),
            totals.jul_npk.toLocaleString(),
            totals.jul_npk_formula.toLocaleString(),
            totals.jul_organik.toLocaleString(),
            totals.agu_urea.toLocaleString(),
            totals.agu_npk.toLocaleString(),
            totals.agu_npk_formula.toLocaleString(),
            totals.agu_organik.toLocaleString(),
            totals.sep_urea.toLocaleString(),
            totals.sep_npk.toLocaleString(),
            totals.sep_npk_formula.toLocaleString(),
            totals.sep_organik.toLocaleString(),
            totals.okt_urea.toLocaleString(),
            totals.okt_npk.toLocaleString(),
            totals.okt_npk_formula.toLocaleString(),
            totals.okt_organik.toLocaleString(),
            totals.nov_urea.toLocaleString(),
            totals.nov_npk.toLocaleString(),
            totals.nov_npk_formula.toLocaleString(),
            totals.nov_organik.toLocaleString(),
            totals.des_urea.toLocaleString(),
            totals.des_npk.toLocaleString(),
            totals.des_npk_formula.toLocaleString(),
            totals.des_organik.toLocaleString()
        ], 3); // Masukkan di baris ke-3 (setelah header)

        // Style untuk baris total yang diupdate
        updatedTotalRow.font = { bold: true };
        updatedTotalRow.alignment = { horizontal: 'center' };
        updatedTotalRow.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFC0C0C0' }
        };
        updatedTotalRow.eachCell(cell => {
            cell.border = borderStyle;
        });
        updatedTotalRow.commit();

        // ===================== 7. SIMPAN FILE =====================
        const filePath = path.join(baseDir, fileName);

        if (!fs.existsSync(baseDir)) {
            fs.mkdirSync(baseDir, { recursive: true });
        }

        worksheet.commit();
        await workbook.commit();

        console.log(`File ${fileName} berhasil digenerate di ${filePath}`);
        return filePath;

    } catch (error) {
        console.error(`Gagal generate report:`, error);
        throw error;
    }
}

module.exports = {
    generateReport,
    generateAllReports,
};