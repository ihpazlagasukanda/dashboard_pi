const { workerData, parentPort } = require('worker_threads');
const ExcelJS = require('exceljs');
const db = require('../config/db');
const path = require('path');

async function generateExcel() {
    const { kabupaten, tahun, filePath } = workerData;
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Data Petani');

    try {
        // 1. Buat query dinamis
        let query = `
      SELECT 
                e.kabupaten, 
                e.kecamatan,
                e.nik, 
                e.nama_petani, 
                e.kode_kios,
                e.tahun, 
                e.urea, 
                e.npk, 
                e.npk_formula, 
                e.organik,

                -- Data tebusan total
                COALESCE(v.tebus_urea, 0) AS tebus_urea,
                COALESCE(v.tebus_npk, 0) AS tebus_npk,
                COALESCE(v.tebus_npk_formula, 0) AS tebus_npk_formula,
                COALESCE(v.tebus_organik, 0) AS tebus_organik,

                -- Perhitungan sisa
                (e.urea - COALESCE(v.tebus_urea, 0)) AS sisa_urea,
                (e.npk - COALESCE(v.tebus_npk, 0)) AS sisa_npk,
                (e.npk_formula - COALESCE(v.tebus_npk_formula, 0)) AS sisa_npk_formula,
                (e.organik - COALESCE(v.tebus_organik, 0)) AS sisa_organik,

                -- Data tebusan per bulan
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
            FROM erdkk e FORCE INDEX (idx_erdkk_nik_kabupaten_tahun)
            LEFT JOIN verval_summary v 
                ON e.nik = v.nik
                AND e.kabupaten = v.kabupaten
                AND e.tahun = v.tahun
                AND e.kecamatan = v.kecamatan
                AND e.kode_kios = v.kode_kios
            LEFT JOIN tebusan_per_bulan t 
                ON e.nik = t.nik
                AND e.kabupaten = t.kabupaten
                AND e.tahun = t.tahun
                AND e.kecamatan = t.kecamatan
                AND e.kode_kios = t.kode_kios
      WHERE 1=1
    `;

        const params = [];
        if (kabupaten) {
            query += " AND e.kabupaten = ?";
            params.push(kabupaten);
        }
        if (tahun) {
            query += " AND e.tahun = ?";
            params.push(tahun);
        }

        // 2. Stream data dari database
        const [rows] = await db.query(query, params);

        // 3. Setup worksheet (mirip dengan kode awal Anda)
        const borderStyle = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };

        // 🔥 **Setup Header dengan Merge Cells**
        // Updated header merges to include new columns
        worksheet.mergeCells('A1:A2'); // Kabupaten
        worksheet.mergeCells('B1:B2'); // Kecamatan
        worksheet.mergeCells('C1:C2'); // NIK
        worksheet.mergeCells('D1:D2'); // Nama Petani
        worksheet.mergeCells('E1:E2'); // Kode Kios
        worksheet.mergeCells('F1:I1'); // Alokasi (shifted right by 2 columns)
        worksheet.mergeCells('J1:M1'); // Sisa
        worksheet.mergeCells('N1:Q1'); // Tebusan

        // Merge header bulanan, 4 kolom per bulan (shifted right by 2 columns)
        worksheet.mergeCells('R1:U1'); // Januari
        worksheet.mergeCells('V1:Y1'); // Februari
        worksheet.mergeCells('Z1:AC1'); // Maret
        worksheet.mergeCells('AD1:AG1'); // April
        worksheet.mergeCells('AH1:AK1'); // Mei
        worksheet.mergeCells('AL1:AO1'); // Juni
        worksheet.mergeCells('AP1:AS1'); // Juli
        worksheet.mergeCells('AT1:AW1'); // Agustus
        worksheet.mergeCells('AX1:AZ1'); // September
        worksheet.mergeCells('BA1:BD1'); // Oktober
        worksheet.mergeCells('BE1:BH1'); // November
        worksheet.mergeCells('BI1:BL1'); // Desember

        // Set Header Utama
        worksheet.getCell("F1").value = "Alokasi";
        worksheet.getCell("J1").value = "Sisa";
        worksheet.getCell("N1").value = "Tebusan";

        worksheet.getCell("R1").value = "Januari";
        worksheet.getCell("V1").value = "Februari";
        worksheet.getCell("Z1").value = "Maret";
        worksheet.getCell("AD1").value = "April";
        worksheet.getCell("AH1").value = "Mei";
        worksheet.getCell("AL1").value = "Juni";
        worksheet.getCell("AP1").value = "Juli";
        worksheet.getCell("AT1").value = "Agustus";
        worksheet.getCell("AX1").value = "September";
        worksheet.getCell("BA1").value = "Oktober";
        worksheet.getCell("BE1").value = "November";
        worksheet.getCell("BI1").value = "Desember";

        // Styling Header Utama
        ["F1", "J1", "N1"].forEach(cell => {
            worksheet.getCell(cell).alignment = { horizontal: "center", vertical: "middle" };
            worksheet.getCell(cell).font = { bold: true, size: 12 };
            worksheet.getCell(cell).fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFD9D9D9' } // Warna abu-abu muda
            };
        });

        // Styling Header Bulanan (Januari - Desember) -> Merah
        ["R1", "V1", "Z1", "AD1", "AH1", "AL1", "AP1", "AT1", "AX1", "BA1", "BE1", "BI1"].forEach(cell => {
            worksheet.getCell(cell).alignment = { horizontal: "center", vertical: "middle" };
            worksheet.getCell(cell).font = { bold: true, size: 12, color: { argb: 'FFFFFFFF' } }; // Warna teks putih
            worksheet.getCell(cell).fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFFF0000' } // Warna merah
            };
        });

        // ✅ **Tambahkan header kosong untuk memastikan Kabupaten, Kecamatan, NIK, Nama Petani, Kode Kios tetap muncul**
        worksheet.getRow(2).values = [
            'Kabupaten', 'Kecamatan', 'NIK', 'Nama Petani', 'Kode Kios', // New columns added
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

        // Styling Header Sub (Jenis Pupuk)
        worksheet.getRow(2).font = { bold: true, size: 12 };
        worksheet.getRow(2).alignment = { horizontal: 'center' };
        worksheet.getRow(2).eachCell((cell) => {
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFE6E6E6' } // Warna abu-abu lebih muda
            };
        });

        // 🔥 **Tambahkan Baris Sum Total**
        let totalRow = worksheet.addRow([
            'TOTAL', '', '', '', '', // Empty for new columns
            data.reduce((sum, r) => sum + parseFloat(r.urea || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.npk || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.npk_formula || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.organik || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.sisa_urea || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.sisa_npk || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.sisa_npk_formula || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.sisa_organik || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.tebus_urea || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.tebus_npk || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.tebus_npk_formula || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.tebus_organik || 0), 0).toLocaleString(),

            data.reduce((sum, r) => sum + parseFloat(r.jan_urea || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.jan_npk || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.jan_npk_formula || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.jan_organik || 0), 0).toLocaleString(),

            data.reduce((sum, r) => sum + parseFloat(r.feb_urea || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.feb_npk || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.feb_npk_formula || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.feb_organik || 0), 0).toLocaleString(),

            data.reduce((sum, r) => sum + parseFloat(r.mar_urea || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.mar_npk || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.mar_npk_formula || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.mar_organik || 0), 0).toLocaleString(),

            data.reduce((sum, r) => sum + parseFloat(r.apr_urea || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.apr_npk || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.apr_npk_formula || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.apr_organik || 0), 0).toLocaleString(),

            data.reduce((sum, r) => sum + parseFloat(r.mei_urea || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.mei_npk || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.mei_npk_formula || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.mei_organik || 0), 0).toLocaleString(),

            data.reduce((sum, r) => sum + parseFloat(r.jun_urea || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.jun_npk || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.jun_npk_formula || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.jun_organik || 0), 0).toLocaleString(),

            data.reduce((sum, r) => sum + parseFloat(r.jul_urea || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.jul_npk || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.jul_npk_formula || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.jul_organik || 0), 0).toLocaleString(),

            data.reduce((sum, r) => sum + parseFloat(r.agu_urea || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.agu_npk || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.agu_npk_formula || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.agu_organik || 0), 0).toLocaleString(),

            data.reduce((sum, r) => sum + parseFloat(r.sep_urea || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.sep_npk || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.sep_npk_formula || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.sep_organik || 0), 0).toLocaleString(),

            data.reduce((sum, r) => sum + parseFloat(r.okt_urea || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.okt_npk || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.okt_npk_formula || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.okt_organik || 0), 0).toLocaleString(),

            data.reduce((sum, r) => sum + parseFloat(r.nov_urea || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.nov_npk || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.nov_npk_formula || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.nov_organik || 0), 0).toLocaleString(),

            data.reduce((sum, r) => sum + parseFloat(r.des_urea || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.des_npk || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.des_npk_formula || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.des_organik || 0), 0).toLocaleString()
        ]);

        totalRow.font = { bold: true };
        totalRow.alignment = { horizontal: 'center' };
        totalRow.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFC0C0C0' } // Warna abu-abu untuk baris total
        };

        totalRow.eachCell((cell) => {
            cell.border = borderStyle;
        });

        // 4. Tambahkan data per chunk (1000 baris)
        const chunkSize = 1000;
        for (let i = 0; i < rows.length; i += chunkSize) {
            const chunk = rows.slice(i, i + chunkSize);
            chunk.forEach(row => {
                worksheet.addRow([
                    row.kabupaten,
                    row.kecamatan, // Added kecamatan
                    row.nik,
                    row.nama_petani,
                    row.kode_kios, // Added kode_kios
                    row.urea, row.npk, row.npk_formula, row.organik, // Alokasi
                    row.sisa_urea, row.sisa_npk, row.sisa_npk_formula, row.sisa_organik, // Sisa
                    row.tebus_urea, row.tebus_npk, row.tebus_npk_formula, row.tebus_organik, // Tebusan
                    row.jan_urea, row.jan_npk, row.jan_npk_formula, row.jan_organik, // Januari
                    row.feb_urea, row.feb_npk, row.feb_npk_formula, row.feb_organik, // Februari
                    row.mar_urea, row.mar_npk, row.mar_npk_formula, row.mar_organik, // Maret
                    row.apr_urea, row.apr_npk, row.apr_npk_formula, row.apr_organik, // April
                    row.mei_urea, row.mei_npk, row.mei_npk_formula, row.mei_organik, // Mei
                    row.jun_urea, row.jun_npk, row.jun_npk_formula, row.jun_organik, // Juni
                    row.jul_urea, row.jul_npk, row.jul_npk_formula, row.jul_organik, // Juli
                    row.agu_urea, row.agu_npk, row.agu_npk_formula, row.agu_organik, // Agustus
                    row.sep_urea, row.sep_npk, row.sep_npk_formula, row.sep_organik, // September
                    row.okt_urea, row.okt_npk, row.okt_npk_formula, row.okt_organik, // Oktober
                    row.nov_urea, row.nov_npk, row.nov_npk_formula, row.nov_organik, // November
                    row.des_urea, row.des_npk, row.des_npk_formula, row.des_organik  // Desember
                ]);
            });

            // Beri tahu progress
            parentPort.postMessage({
                progress: Math.round((i / rows.length) * 100)
            });
        }

        // 5. Simpan file
        await workbook.xlsx.writeFile(filePath);
        parentPort.postMessage({ success: true });

    } catch (error) {
        parentPort.postMessage({
            error: error.message,
            stack: error.stack
        });
    }
}

generateExcel();