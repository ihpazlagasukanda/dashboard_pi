const ExcelJS = require('exceljs');
const { buildPetaniQuery } = require('./queryBuilder');
const db = require('../config/db');

const generatePetaniReport = async (filters, outputPath) => {
    // 1. Setup Workbook Streaming
    const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({
        filename: outputPath,
        useStyles: true
    });

    const worksheet = workbook.addWorksheet('Data Petani');

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

    // Isi data
    data.forEach((row) => {
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

    worksheet.eachRow((row) => {
        row.eachCell((cell) => {
            cell.border = borderStyle;
        });
    });

    // 3. Stream Data dari Database
    const { query, params } = buildPetaniQuery(filters);
    const dbStream = await db.query(query, params).stream();

    dbStream.on('data', (row) => {
        worksheet.addRow(row).commit(); // Langsung commit per baris
    });

    return new Promise((resolve, reject) => {
        dbStream.on('end', () => {
            worksheet.commit();
            workbook.commit()
                .then(() => resolve(outputPath))
                .catch(reject);
        });
        dbStream.on('error', reject);
    });
};

module.exports = { generatePetaniReport };