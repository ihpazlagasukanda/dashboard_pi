const db = require("../config/db");
const ExcelJS = require('exceljs');
const { Worker } = require('worker_threads');
const path = require('path');
const fs = require('fs');
const { generateReport } = require('../services/reportGenerator');

exports.downloadRekapPetani = async (req, res) => {
    try {
        const { tahun } = req.query;

        if (!tahun) {
            return res.status(400).json({ message: "Parameter tahun diperlukan" });
        }

        // List of all kabupaten we want to include
        const targetKabupaten = [
            "BOYOLALI", "KLATEN", "SUKOHARJO", "KARANGANYAR", "WONOGIRI", "SRAGEN", 
            "KOTA SURAKARTA", "SLEMAN", "BANTUL", "KULON PROGO", "GUNUNG KIDUL", 
            "KOTA YOGYAKARTA"
        ];

        const params = [tahun, ...targetKabupaten];

        // Query alokasi
        const alokasiQuery = `
            SELECT 
                kabupaten,
                kecamatan,
                CONCAT(kabupaten, ' - ', kecamatan) AS wilayah,
                COUNT(DISTINCT nik) AS alokasi
            FROM erdkk
            WHERE tahun = ? 
            AND kabupaten IN (${targetKabupaten.map(() => "?").join(",")})
            GROUP BY kabupaten, kecamatan
            ORDER BY kabupaten, kecamatan;
        `;

        // Query realisasi
        const realisasiQuery = `
            SELECT 
                kabupaten,
                kecamatan,
                CONCAT(kabupaten, ' - ', kecamatan) AS wilayah,
                COUNT(DISTINCT nik) AS realisasi
            FROM verval_summary
            WHERE tahun = ? AND kabupaten IN (${targetKabupaten.map(() => "?").join(",")})
            GROUP BY kabupaten, kecamatan
            ORDER BY kabupaten, kecamatan
        `;

        // Eksekusi query paralel
        const [alokasiData, realisasiData] = await Promise.all([
            db.query(alokasiQuery, params),
            db.query(realisasiQuery, params)
        ]);

        const alokasi = alokasiData[0] || [];
        const realisasi = realisasiData[0] || [];

        // Gabungkan data
        const mapData = {};

        alokasi.forEach(item => {
            mapData[item.wilayah] = {
                kabupaten: item.kabupaten,
                kecamatan: item.kecamatan,
                wilayah: item.wilayah,
                alokasi: item.alokasi || 0,
                realisasi: 0,
                sisa: item.alokasi || 0
            };
        });

        realisasi.forEach(item => {
            if (!mapData[item.wilayah]) {
                mapData[item.wilayah] = {
                    kabupaten: item.kabupaten,
                    kecamatan: item.kecamatan,
                    wilayah: item.wilayah,
                    alokasi: 0,
                    realisasi: item.realisasi || 0,
                    sisa: 0 - (item.realisasi || 0)
                };
            } else {
                mapData[item.wilayah].realisasi = item.realisasi || 0;

                // Calculate sisa (alokasi - realisasi)
                mapData[item.wilayah].sisa = mapData[item.wilayah].alokasi - (item.realisasi || 0);
            }
        });

        // Hitung total
        const totals = {
            alokasi: alokasi.reduce((sum, item) => sum + (item.alokasi || 0), 0),
            realisasi: realisasi.reduce((sum, item) => sum + (item.realisasi || 0), 0),
            sisa: alokasi.reduce((sum, item) => sum + (item.alokasi || 0), 0) - realisasi.reduce((sum, item) => sum + (item.realisasi || 0), 0),
        };

        // Convert mapData to array and sort by kabupaten then kecamatan
        const finalData = Object.values(mapData).sort((a, b) => {
            if (a.kabupaten < b.kabupaten) return -1;
            if (a.kabupaten > b.kabupaten) return 1;
            if (a.kecamatan < b.kecamatan) return -1;
            if (a.kecamatan > b.kecamatan) return 1;
            return 0;
        });

        // Membuat workbook Excel
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Rekap Petani All Kecamatan');

        // Header worksheet
        worksheet.mergeCells('A1:N1');
        worksheet.getCell('A1').value = `REKAPITULASI PETANI PUPUK PER KECAMATAN TAHUN ${tahun}`;
        worksheet.getCell('A1').font = { bold: true, size: 14 };
        worksheet.getCell('A1').alignment = { horizontal: 'center' };

        worksheet.mergeCells('A2:N2');
        worksheet.getCell('A2').value = `KABUPATEN: BOYOLALI, KLATEN, SUKOHARJO, KARANGANYAR, WONOGIRI, SRAGEN, KOTA SURAKARTA, SLEMAN, BANTUL, KULON PROGO, GUNUNG KIDUL, KOTA YOGYAKARTA`;
        worksheet.getCell('A2').font = { bold: true };
        worksheet.getCell('A2').alignment = { horizontal: 'center' };

        // Header tabel utama
        // worksheet.mergeCells('A3:A4');
        worksheet.getCell('A3').value = 'NO';
        worksheet.getCell('A3').alignment = { vertical: 'middle', horizontal: 'center' };

        // worksheet.mergeCells('B3:B4');
        worksheet.getCell('B3').value = 'KABUPATEN';
        worksheet.getCell('B3').alignment = { vertical: 'middle', horizontal: 'center' };

        // worksheet.mergeCells('C3:C4');
        worksheet.getCell('C3').value = 'KECAMATAN';
        worksheet.getCell('C3').alignment = { vertical: 'middle', horizontal: 'center' };

        // Header Alokasi
        // worksheet.mergeCells('D3:G3');
        worksheet.getCell('D3').value = 'TOTAL PETANI';
        worksheet.getCell('D3').alignment = { horizontal: 'center' };

        // Header Realisasi
        // worksheet.mergeCells('H3:K3');
        worksheet.getCell('E3').value = 'TOTAL PETANI TEBUS';
        worksheet.getCell('E3').alignment = { horizontal: 'center' };

        // Header Sisa Tebus
        // worksheet.mergeCells('L3:O3');
        worksheet.getCell('F3').value = 'TOTAL SISA TEBUS';
        worksheet.getCell('F3').alignment = { horizontal: 'center' };

        // Sub-header kolom
        // worksheet.getCell('D4').value = 'UREA';
        // worksheet.getCell('E4').value = 'NPK';
        // worksheet.getCell('F4').value = 'NPK FORMULA';
        // worksheet.getCell('G4').value = 'ORGANIK';
        // worksheet.getCell('H4').value = 'UREA';
        // worksheet.getCell('I4').value = 'NPK';
        // worksheet.getCell('J4').value = 'NPK FORMULA';
        // worksheet.getCell('K4').value = 'ORGANIK';
        // worksheet.getCell('L4').value = 'UREA';
        // worksheet.getCell('M4').value = 'NPK';
        // worksheet.getCell('N4').value = 'NPK FORMULA';
        // worksheet.getCell('O4').value = 'ORGANIK';

        // Styling header
        ['A3', 'B3', 'C3', 'D3', 'E3', 'F3'].forEach(cell => {
            worksheet.getCell(cell).font = { bold: true };
            worksheet.getCell(cell).fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFD3D3D3' }
            };
            worksheet.getCell(cell).border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
        });

        // Styling sub-header
        // ['D4', 'E4', 'F4', 'G4', 'H4', 'I4', 'J4', 'K4', 'L4', 'M4', 'N4', 'O4'].forEach(cell => {
        //     worksheet.getCell(cell).font = { bold: true };
        //     worksheet.getCell(cell).alignment = { horizontal: 'center' };
        //     worksheet.getCell(cell).fill = {
        //         type: 'pattern',
        //         pattern: 'solid',
        //         fgColor: { argb: 'FFE8E8E8' }
        //     };
        //     worksheet.getCell(cell).border = {
        //         top: { style: 'thin' },
        //         left: { style: 'thin' },
        //         bottom: { style: 'thin' },
        //         right: { style: 'thin' }
        //     };
        // });

        // Mengisi data
        finalData.forEach((item, index) => {
            const row = worksheet.addRow([
                index + 1,
                item.kabupaten,
                item.kecamatan,
                item.alokasi,
                item.realisasi,
                item.sisa
            ]);

            // Format angka
            ['D', 'E', 'F'].forEach(col => {
                const cell = row.getCell(col);
                cell.numFmt = '#,##0';
            });
        });

        // Menambahkan baris total
        const totalRow = worksheet.addRow([
            '',
            'TOTAL',
            '',
            totals.alokasi,
            totals.realisasi,
            totals.sisa
        ]);

        // Styling baris total
        totalRow.font = { bold: true };
        totalRow.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFC6E0B4' } // Warna hijau muda
        };

        // Format angka di baris total
        ['D', 'E', 'F'].forEach(col => {
            const cell = totalRow.getCell(col);
            cell.numFmt = '#,##0';
        });

        // Set lebar kolom
        worksheet.columns = [
            { key: 'no', width: 5 },
            { key: 'kabupaten', width: 20 },
            { key: 'kecamatan', width: 25 },
            { key: 'alokasi', width: 10 },
            { key: 'realisasi', width: 10 },
            { key: 'sisa', width: 10 }
        ];

        // Set response headers
        res.setHeader(
            'Content-Type',
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        );
        res.setHeader(
            'Content-Disposition',
            `attachment; filename="rekap_petani_kecamatan_${tahun}.xlsx"`
        );

        // Kirim workbook sebagai response
        await workbook.xlsx.write(res);
        res.end();

    } catch (error) {
        console.error(error);
        return res.status(500).json({ message: "Internal server error", error: error.message });
    }
};
