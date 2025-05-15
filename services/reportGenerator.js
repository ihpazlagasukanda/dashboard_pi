const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');
const db = require('../config/db');

// Fungsi untuk generate semua laporan yang diperlukan
async function generateAllReports() {
    try {
        const baseDir = path.join(__dirname, '../public/reports');
        if (!fs.existsSync(baseDir)) {
            fs.mkdirSync(baseDir, { recursive: true });
        }

        // Daftar kabupaten yang perlu digenerate
        const kabupatenList = await db.query('SELECT DISTINCT kabupaten FROM petani_summary_cache_partitioned');
        const tahunList = await db.query('SELECT DISTINCT tahun FROM petani_summary_cache_partitioned ORDER BY tahun DESC');

        // Generate laporan untuk setiap kombinasi kabupaten dan tahun
        for (const kabupaten of kabupatenList[0]) {
            for (const tahun of tahunList[0]) {
                await generateReport(kabupaten.kabupaten, tahun.tahun, baseDir);
            }
        }

        // Generate laporan ALL (tanpa filter)
        await generateReport(null, null, baseDir);

        console.log('Semua laporan telah digenerate');
    } catch (error) {
        console.error('Error generate reports:', error);
    }
}

// Fungsi untuk generate single report
async function generateReport(kabupaten, tahun, baseDir) {
    try {
        let query = `SELECT * FROM petani_summary_cache_partitioned WHERE 1=1`;
        let params = [];
        let fileName = 'petani_summary_ALL.xlsx';

        if (kabupaten) {
            query += " AND kabupaten = ?";
            params.push(kabupaten);
            fileName = `petani_summary_${kabupaten.replace(/\s+/g, '_')}`;
        }

        if (tahun) {
            query += " AND tahun = ?";
            params.push(tahun);
            fileName += `_${tahun}`;
        }

        fileName += '.xlsx';

        const [data] = await db.query(query, params);

        // Proses pembuatan Excel (sama seperti kode Anda sebelumnya)
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet("Summary");
        // ... (kode pembuatan excel yang sudah ada)
        const borderStyle = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };
        // 🔥 **Setup Header dengan Merge Cells**
        worksheet.mergeCells('A1:A3'); // Kabupaten
        worksheet.mergeCells('B1:B3'); // Kecamatan
        worksheet.mergeCells('C1:C3'); // NIK
        worksheet.mergeCells('D1:D3'); // Nama Petani
        worksheet.mergeCells('E1:E3'); // Kode_kios
        worksheet.mergeCells('F1:I1'); // Alokasi
        worksheet.mergeCells('F2:I2'); // Sub Header Alokasi
        worksheet.mergeCells('J1:M1'); // Sisa
        worksheet.mergeCells('J2:M2'); // Sub Header Sisa
        worksheet.mergeCells('N1:Q1'); // Tebusan
        worksheet.mergeCells('N2:Q2'); // Sub Header Tebusan

        // Merge header bulanan, 8 kolom per bulan (4 Kartan + 4 Ipubers)
        worksheet.mergeCells('R1:Y1'); // Januari
        worksheet.mergeCells('R2:U2'); // Sub Header Januari - Kartan
        worksheet.mergeCells('V2:Y2'); // Sub Header Januari - Ipubers

        worksheet.mergeCells('Z1:AG1'); // Februari
        worksheet.mergeCells('Z2:AC2'); // Sub Header Februari - Kartan
        worksheet.mergeCells('AD2:AG2'); // Sub Header Februari - Ipubers

        worksheet.mergeCells('AH1:AO1'); // Maret
        worksheet.mergeCells('AH2:AK2'); // Sub Header Maret - Kartan
        worksheet.mergeCells('AL2:AO2'); // Sub Header Maret - Ipubers

        worksheet.mergeCells('AP1:AW1'); // April
        worksheet.mergeCells('AP2:AS2'); // Sub Header April - Kartan
        worksheet.mergeCells('AT2:AW2'); // Sub Header April - Ipubers

        worksheet.mergeCells('AX1:BE1'); // Mei
        worksheet.mergeCells('AX2:BA2'); // Sub Header Mei - Kartan
        worksheet.mergeCells('BB2:BE2'); // Sub Header Mei - Ipubers

        worksheet.mergeCells('BF1:BM1'); // Juni
        worksheet.mergeCells('BF2:BI2'); // Sub Header Juni - Kartan
        worksheet.mergeCells('BJ2:BM2'); // Sub Header Juni - Ipubers

        worksheet.mergeCells('BN1:BU1'); // Juli
        worksheet.mergeCells('BN2:BQ2'); // Sub Header Juli - Kartan
        worksheet.mergeCells('BR2:BU2'); // Sub Header Juli - Ipubers

        worksheet.mergeCells('BV1:CC1'); // Agustus
        worksheet.mergeCells('BV2:BY2'); // Sub Header Agustus - Kartan
        worksheet.mergeCells('BZ2:CC2'); // Sub Header Agustus - Ipubers

        worksheet.mergeCells('CD1:CK1'); // September
        worksheet.mergeCells('CD2:CG2'); // Sub Header September - Kartan
        worksheet.mergeCells('CH2:CK2'); // Sub Header September - Ipubers

        worksheet.mergeCells('CL1:CS1'); // Oktober
        worksheet.mergeCells('CL2:CO2'); // Sub Header Oktober - Kartan
        worksheet.mergeCells('CP2:CS2'); // Sub Header Oktober - Ipubers

        worksheet.mergeCells('CT1:DA1'); // November
        worksheet.mergeCells('CT2:CW2'); // Sub Header November - Kartan
        worksheet.mergeCells('CX2:DA2'); // Sub Header November - Ipubers

        worksheet.mergeCells('DB1:DI1'); // Desember
        worksheet.mergeCells('DB2:DE2'); // Sub Header Desember - Kartan
        worksheet.mergeCells('DF2:DI2'); // Sub Header Desember - Ipubers


        // Set Header Utama
        worksheet.getCell("F1").value = "Alokasi";
        worksheet.getCell("J1").value = "Sisa";
        worksheet.getCell("N1").value = "Tebusan";

        worksheet.getCell("R1").value = "Januari";
        worksheet.getCell("R2").value = "Kartan";
        worksheet.getCell("V2").value = "Ipubers";

        worksheet.getCell("Z1").value = "Februari";
        worksheet.getCell("Z2").value = "Kartan";
        worksheet.getCell("AD2").value = "Ipubers";

        worksheet.getCell("AH1").value = "Maret";
        worksheet.getCell("AH2").value = "Kartan";
        worksheet.getCell("AL2").value = "Ipubers";

        worksheet.getCell("AP1").value = "April";
        worksheet.getCell("AP2").value = "Kartan";
        worksheet.getCell("AT2").value = "Ipubers";

        worksheet.getCell("AX1").value = "Mei";
        worksheet.getCell("AX2").value = "Kartan";
        worksheet.getCell("BB2").value = "Ipubers";

        worksheet.getCell("BF1").value = "Juni";
        worksheet.getCell("BF2").value = "Kartan";
        worksheet.getCell("BJ2").value = "Ipubers";

        worksheet.getCell("BN1").value = "Juli";
        worksheet.getCell("BN2").value = "Kartan";
        worksheet.getCell("BR2").value = "Ipubers";

        worksheet.getCell("BV1").value = "Agustus";
        worksheet.getCell("BV2").value = "Kartan";
        worksheet.getCell("BZ2").value = "Ipubers";

        worksheet.getCell("CD1").value = "September";
        worksheet.getCell("CD2").value = "Kartan";
        worksheet.getCell("CH2").value = "Ipubers";

        worksheet.getCell("CL1").value = "Oktober";
        worksheet.getCell("CL2").value = "Kartan";
        worksheet.getCell("CP2").value = "Ipubers";

        worksheet.getCell("CT1").value = "November";
        worksheet.getCell("CT2").value = "Kartan";
        worksheet.getCell("CX2").value = "Ipubers";

        worksheet.getCell("DB1").value = "Desember";
        worksheet.getCell("DB2").value = "Kartan";
        worksheet.getCell("DF2").value = "Ipubers";



        // Styling Header Utama (geser 1 kolom dari sebelumnya)
        ["F1", "J1", "N1"].forEach(cell => {
            worksheet.getCell(cell).alignment = { horizontal: "center", vertical: "middle" };
            worksheet.getCell(cell).font = { bold: true, size: 12 };
            worksheet.getCell(cell).fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFD9D9D9' } // Warna abu-abu muda
            };
        });

        // Styling Header Bulanan (Januari - Desember) -> Merah (Baris 1, geser 1 kolom)
        ["R1", "Z1", "AH1", "AP1", "AX1", "BF1", "BN1", "BV1", "CD1", "CL1", "CT1", "DB1"].forEach(cell => {
            worksheet.getCell(cell).alignment = { horizontal: "center", vertical: "middle" };
            worksheet.getCell(cell).font = { bold: true, size: 12, color: { argb: 'FFFFFFFF' } }; // Warna teks putih
            worksheet.getCell(cell).fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFFF0000' } // Warna merah
            };
        });

        // Styling Header Bulanan (Baris 2, geser 1 kolom)
        ["R2", "Z2", "AH2", "AP2", "AX2", "BF2", "BN2", "BV2", "CD2", "CL2", "CT2", "DB2"].forEach(cell => {
            worksheet.getCell(cell).alignment = { horizontal: "center", vertical: "middle" };
            worksheet.getCell(cell).font = { bold: true, size: 12, color: { argb: 'FFFFFFFF' } }; // Warna teks putih
            worksheet.getCell(cell).fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFFF0000' } // Warna merah
            };
        });

        // Styling Header Bulanan tambahan (Baris 2, geser 1 kolom)
        ["V2", "AD2", "AL2", "AT2", "BB2", "BJ2", "BR2", "BZ2", "CH2", "CP2", "CX2", "DF2"].forEach(cell => {
            worksheet.getCell(cell).alignment = { horizontal: "center", vertical: "middle" };
            worksheet.getCell(cell).font = { bold: true, size: 12, color: { argb: 'FFFFFFFF' } }; // Warna teks putih
            worksheet.getCell(cell).fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFFF0000' } // Warna merah
            };
        });


        // ✅ **Tambahkan header kosong untuk memastikan Kabupaten, NIK, Nama Petani tetap muncul**
        worksheet.getRow(3).values = [
            'Kabupaten', 'Kecamatan', 'NIK', 'Nama Petani', 'Kode Kios', // Tambahkan kolom awal agar tidak hilang
            'Urea', 'NPK', 'NPK Formula', 'Organik', // Alokasi
            'Urea', 'NPK', 'NPK Formula', 'Organik', // Sisa
            'Urea', 'NPK', 'NPK Formula', 'Organik', // Tebusan
            'Urea', 'NPK', 'NPK Formula', 'Organik', // Januari Kartan
            'Urea', 'NPK', 'NPK Formula', 'Organik', // Januari Ipubers
            'Urea', 'NPK', 'NPK Formula', 'Organik',
            'Urea', 'NPK', 'NPK Formula', 'Organik',
            'Urea', 'NPK', 'NPK Formula', 'Organik',
            'Urea', 'NPK', 'NPK Formula', 'Organik',
            'Urea', 'NPK', 'NPK Formula', 'Organik',
            'Urea', 'NPK', 'NPK Formula', 'Organik',
            'Urea', 'NPK', 'NPK Formula', 'Organik',
            'Urea', 'NPK', 'NPK Formula', 'Organik',
            'Urea', 'NPK', 'NPK Formula', 'Organik',
            'Urea', 'NPK', 'NPK Formula', 'Organik',
            'Urea', 'NPK', 'NPK Formula', 'Organik',
            'Urea', 'NPK', 'NPK Formula', 'Organik',
            'Urea', 'NPK', 'NPK Formula', 'Organik',
            'Urea', 'NPK', 'NPK Formula', 'Organik',
            'Urea', 'NPK', 'NPK Formula', 'Organik',
            'Urea', 'NPK', 'NPK Formula', 'Organik',
            'Urea', 'NPK', 'NPK Formula', 'Organik',
            'Urea', 'NPK', 'NPK Formula', 'Organik',
            'Urea', 'NPK', 'NPK Formula', 'Organik',
            'Urea', 'NPK', 'NPK Formula', 'Organik',
            'Urea', 'NPK', 'NPK Formula', 'Organik',
            'Urea', 'NPK', 'NPK Formula', 'Organik',

            // Lanjutkan untuk bulan lainnya
        ];

        // Styling Header Sub (Jenis Pupuk)
        worksheet.getRow(3).font = { bold: true, size: 12 };
        worksheet.getRow(3).alignment = { horizontal: 'center' };
        worksheet.getRow(3).eachCell((cell) => {
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFE6E6E6' } // Warna abu-abu lebih muda
            };
        });


        // 🔥 **Tambahkan Baris Sum Total**
        let totalRow = worksheet.addRow([
            'TOTAL', '', '', '', '',
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

            // January
            data.reduce((sum, r) => sum + parseFloat(r.jan_kartan_urea || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.jan_kartan_npk || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.jan_kartan_npk_formula || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.jan_kartan_organik || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.jan_ipubers_urea || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.jan_ipubers_npk || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.jan_ipubers_npk_formula || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.jan_ipubers_organik || 0), 0).toLocaleString(),

            // February
            data.reduce((sum, r) => sum + parseFloat(r.feb_kartan_urea || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.feb_kartan_npk || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.feb_kartan_npk_formula || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.feb_kartan_organik || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.feb_ipubers_urea || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.feb_ipubers_npk || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.feb_ipubers_npk_formula || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.feb_ipubers_organik || 0), 0).toLocaleString(),

            // March
            data.reduce((sum, r) => sum + parseFloat(r.mar_kartan_urea || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.mar_kartan_npk || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.mar_kartan_npk_formula || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.mar_kartan_organik || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.mar_ipubers_urea || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.mar_ipubers_npk || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.mar_ipubers_npk_formula || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.mar_ipubers_organik || 0), 0).toLocaleString(),

            // April
            data.reduce((sum, r) => sum + parseFloat(r.apr_kartan_urea || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.apr_kartan_npk || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.apr_kartan_npk_formula || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.apr_kartan_organik || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.apr_ipubers_urea || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.apr_ipubers_npk || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.apr_ipubers_npk_formula || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.apr_ipubers_organik || 0), 0).toLocaleString(),

            // May
            data.reduce((sum, r) => sum + parseFloat(r.mei_kartan_urea || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.mei_kartan_npk || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.mei_kartan_npk_formula || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.mei_kartan_organik || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.mei_ipubers_urea || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.mei_ipubers_npk || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.mei_ipubers_npk_formula || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.mei_ipubers_organik || 0), 0).toLocaleString(),

            // June
            data.reduce((sum, r) => sum + parseFloat(r.jun_kartan_urea || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.jun_kartan_npk || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.jun_kartan_npk_formula || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.jun_kartan_organik || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.jun_ipubers_urea || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.jun_ipubers_npk || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.jun_ipubers_npk_formula || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.jun_ipubers_organik || 0), 0).toLocaleString(),

            // July
            data.reduce((sum, r) => sum + parseFloat(r.jul_kartan_urea || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.jul_kartan_npk || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.jul_kartan_npk_formula || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.jul_kartan_organik || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.jul_ipubers_urea || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.jul_ipubers_npk || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.jul_ipubers_npk_formula || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.jul_ipubers_organik || 0), 0).toLocaleString(),

            // August
            data.reduce((sum, r) => sum + parseFloat(r.agu_kartan_urea || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.agu_kartan_npk || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.agu_kartan_npk_formula || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.agu_kartan_organik || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.agu_ipubers_urea || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.agu_ipubers_npk || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.agu_ipubers_npk_formula || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.agu_ipubers_organik || 0), 0).toLocaleString(),

            // September
            data.reduce((sum, r) => sum + parseFloat(r.sep_kartan_urea || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.sep_kartan_npk || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.sep_kartan_npk_formula || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.sep_kartan_organik || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.sep_ipubers_urea || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.sep_ipubers_npk || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.sep_ipubers_npk_formula || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.sep_ipubers_organik || 0), 0).toLocaleString(),

            // October
            data.reduce((sum, r) => sum + parseFloat(r.okt_kartan_urea || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.okt_kartan_npk || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.okt_kartan_npk_formula || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.okt_kartan_organik || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.okt_ipubers_urea || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.okt_ipubers_npk || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.okt_ipubers_npk_formula || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.okt_ipubers_organik || 0), 0).toLocaleString(),

            // November
            data.reduce((sum, r) => sum + parseFloat(r.nov_kartan_urea || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.nov_kartan_npk || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.nov_kartan_npk_formula || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.nov_kartan_organik || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.nov_ipubers_urea || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.nov_ipubers_npk || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.nov_ipubers_npk_formula || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.nov_ipubers_organik || 0), 0).toLocaleString(),

            // December
            data.reduce((sum, r) => sum + parseFloat(r.des_kartan_urea || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.des_kartan_npk || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.des_kartan_npk_formula || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.des_kartan_organik || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.des_ipubers_urea || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.des_ipubers_npk || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.des_ipubers_npk_formula || 0), 0).toLocaleString(),
            data.reduce((sum, r) => sum + parseFloat(r.des_ipubers_organik || 0), 0).toLocaleString(),


        ]);


        totalRow.font = { bold: true };
        totalRow.alignment = { horizontal: 'center' };
        totalRow.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFC0C0C0' } // Warna abu-abu untuk baris total
        };

        // Format Baris Total
        totalRow.font = { bold: true };
        totalRow.alignment = { horizontal: 'center' };
        totalRow.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFC0C0C0' } // Warna abu-abu untuk baris total
        };

        // Tambahkan border ke setiap sel di baris total
        totalRow.eachCell((cell) => {
            cell.border = borderStyle;
        });

        // Isi data
        data.forEach((row) => {
            worksheet.addRow([
                row.kabupaten, row.kecamatan, row.nik, row.nama_petani, row.kode_kios,
                row.urea, row.npk, row.npk_formula, row.organik, // Alokasi
                row.sisa_urea, row.sisa_npk, row.sisa_npk_formula, row.sisa_organik, // Sisa
                row.tebus_urea, row.tebus_npk, row.tebus_npk_formula, row.tebus_organik, // Tebusan
                row.jan_kartan_urea, row.jan_kartan_npk, row.jan_kartan_npk_formula, row.jan_kartan_organik, // Januari Kartan
                row.jan_ipubers_urea, row.jan_ipubers_npk, row.jan_ipubers_npk_formula, row.jan_ipubers_organik, // Januari Ipubers

                row.feb_kartan_urea, row.feb_kartan_npk, row.feb_kartan_npk_formula, row.feb_kartan_organik, // Februari Kartan
                row.feb_ipubers_urea, row.feb_ipubers_npk, row.feb_ipubers_npk_formula, row.feb_ipubers_organik, // Februari Ipubers

                row.mar_kartan_urea, row.mar_kartan_npk, row.mar_kartan_npk_formula, row.mar_kartan_organik, // Maret Kartan
                row.mar_ipubers_urea, row.mar_ipubers_npk, row.mar_ipubers_npk_formula, row.mar_ipubers_organik, // Maret Ipubers

                row.apr_kartan_urea, row.apr_kartan_npk, row.apr_kartan_npk_formula, row.apr_kartan_organik, // April Kartan
                row.apr_ipubers_urea, row.apr_ipubers_npk, row.apr_ipubers_npk_formula, row.apr_ipubers_organik, // April Ipubers

                row.mei_kartan_urea, row.mei_kartan_npk, row.mei_kartan_npk_formula, row.mei_kartan_organik, // Mei Kartan
                row.mei_ipubers_urea, row.mei_ipubers_npk, row.mei_ipubers_npk_formula, row.mei_ipubers_organik, // Mei Ipubers

                row.jun_kartan_urea, row.jun_kartan_npk, row.jun_kartan_npk_formula, row.jun_kartan_organik, // Juni Kartan
                row.jun_ipubers_urea, row.jun_ipubers_npk, row.jun_ipubers_npk_formula, row.jun_ipubers_organik, // Juni Ipubers

                row.jul_kartan_urea, row.jul_kartan_npk, row.jul_kartan_npk_formula, row.jul_kartan_organik, // Juli Kartan
                row.jul_ipubers_urea, row.jul_ipubers_npk, row.jul_ipubers_npk_formula, row.jul_ipubers_organik, // Juli Ipubers

                row.agu_kartan_urea, row.agu_kartan_npk, row.agu_kartan_npk_formula, row.agu_kartan_organik, // Agustus Kartan
                row.agu_ipubers_urea, row.agu_ipubers_npk, row.agu_ipubers_npk_formula, row.agu_ipubers_organik, // Agustus Ipubers

                row.sep_kartan_urea, row.sep_kartan_npk, row.sep_kartan_npk_formula, row.sep_kartan_organik, // September Kartan
                row.sep_ipubers_urea, row.sep_ipubers_npk, row.sep_ipubers_npk_formula, row.sep_ipubers_organik, // September Ipubers

                row.okt_kartan_urea, row.okt_kartan_npk, row.okt_kartan_npk_formula, row.okt_kartan_organik, // Oktober Kartan
                row.okt_ipubers_urea, row.okt_ipubers_npk, row.okt_ipubers_npk_formula, row.okt_ipubers_organik, // Oktober Ipubers

                row.nov_kartan_urea, row.nov_kartan_npk, row.nov_kartan_npk_formula, row.nov_kartan_organik, // November Kartan
                row.nov_ipubers_urea, row.nov_ipubers_npk, row.nov_ipubers_npk_formula, row.nov_ipubers_organik, // November Ipubers

                row.des_kartan_urea, row.des_kartan_npk, row.des_kartan_npk_formula, row.des_kartan_organik, // Desember Kartan
                row.des_ipubers_urea, row.des_ipubers_npk, row.des_ipubers_npk_formula, row.des_ipubers_organik, // Desember Ipubers

            ]);
        });

        worksheet.eachRow((row) => {
            row.eachCell((cell) => {
                cell.border = borderStyle;
            });
        });

        // Simpan file
        const filePath = path.join(baseDir, fileName);
        await workbook.xlsx.writeFile(filePath);
        console.log(`File ${fileName} berhasil digenerate`);

    } catch (error) {
        console.error(`Error generate report for ${kabupaten || 'ALL'} ${tahun || ''}:`, error);
    }
}

module.exports = { generateAllReports };