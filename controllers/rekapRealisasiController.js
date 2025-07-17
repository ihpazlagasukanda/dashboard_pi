const db = require("../config/db");
const ExcelJS = require('exceljs');
const { Worker } = require('worker_threads');
const path = require('path');
const fs = require('fs');
const { generateReport } = require('../services/reportGenerator');

exports.summaryPupuk = async (req, res) => {
    try {
        const { provinsi, kabupaten, tahun, sumber_alokasi = "erdkk" } = req.query;

        if (!tahun) {
            return res.json({
                level: !provinsi ? "provinsi" : !kabupaten ? "kabupaten" : "kecamatan",
                data: [],
                totals: {
                    urea_realisasi: 0,
                    npk_realisasi: 0,
                    npk_formula_realisasi: 0,
                    organik_realisasi: 0,
                    urea_alokasi: 0,
                    npk_alokasi: 0,
                    npk_formula_alokasi: 0,
                    organik_alokasi: 0
                }
            });
        }

        const kabupatenDIY = ["SLEMAN", "BANTUL", "GUNUNG KIDUL", "KULON PROGO", "KOTA YOGYAKARTA"];

        let selectField = "kabupaten";
        let groupField = "kabupaten";
        let whereClause = "WHERE tahun = ?";
        const params = [tahun];

        // Provinsi level
        if (!provinsi) {
            selectField = `
                CASE 
                    WHEN kabupaten IN (${kabupatenDIY.map(() => "?").join(",")}) THEN 'DI Yogyakarta'
                    ELSE 'Jawa Tengah'
                END
            `;
            groupField = "wilayah";
            params.unshift(...kabupatenDIY);
        }
        // Kabupaten level
        else if (provinsi && !kabupaten) {
            if (provinsi === "DI Yogyakarta") {
                whereClause += ` AND kabupaten IN (${kabupatenDIY.map(() => "?").join(",")})`;
                params.push(...kabupatenDIY);
            } else {
                whereClause += ` AND kabupaten NOT IN (${kabupatenDIY.map(() => "?").join(",")})`;
                params.push(...kabupatenDIY);
            }
        }
        // Kecamatan level
        else if (kabupaten) {
            whereClause += ` AND kabupaten = ?`;
            params.push(kabupaten);
            selectField = "kecamatan";
            groupField = "kecamatan";
        }

        let alokasiQuery = "";
        // Query alokasi
        if (sumber_alokasi === "sk_bupati") {
            alokasiQuery = `
        SELECT 
            ${selectField} AS wilayah,
            SUM(CASE WHEN produk = 'urea' THEN alokasi ELSE 0 END) AS alokasi_urea,
            SUM(CASE WHEN produk = 'npk' THEN alokasi ELSE 0 END) AS alokasi_npk,
            SUM(CASE WHEN produk = 'npk_formula' THEN alokasi ELSE 0 END) AS alokasi_npk_formula,
            SUM(CASE WHEN produk = 'organik' THEN alokasi ELSE 0 END) AS alokasi_organik
        FROM sk_bupati
        ${whereClause}
        GROUP BY ${groupField}
    `;
        } else {
            alokasiQuery = `
        SELECT 
            ${selectField} AS wilayah,
            SUM(urea) AS alokasi_urea,
            SUM(npk) AS alokasi_npk,
            SUM(npk_formula) AS alokasi_npk_formula,
            SUM(organik) AS alokasi_organik
        FROM erdkk
        ${whereClause}
        GROUP BY ${groupField}
    `;
        }

        // Query realisasi
        const realisasiQuery = `
            SELECT 
                ${selectField.replace(/e\./g, "")} AS wilayah,
                SUM(tebus_urea) AS realisasi_urea,
                SUM(tebus_npk) AS realisasi_npk,
                SUM(tebus_npk_formula) AS realisasi_npk_formula,
                SUM(tebus_organik) AS realisasi_organik
            FROM verval_summary
            ${whereClause.replace(/e\./g, "")}
            GROUP BY ${groupField}
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
                wilayah: item.wilayah,
                alokasi_urea: item.alokasi_urea || 0,
                alokasi_npk: item.alokasi_npk || 0,
                alokasi_npk_formula: item.alokasi_npk_formula || 0,
                alokasi_organik: item.alokasi_organik || 0,
                realisasi_urea: 0,
                realisasi_npk: 0,
                realisasi_npk_formula: 0,
                realisasi_organik: 0
            };
        });

        realisasi.forEach(item => {
            if (!mapData[item.wilayah]) {
                mapData[item.wilayah] = {
                    wilayah: item.wilayah,
                    alokasi_urea: 0,
                    alokasi_npk: 0,
                    alokasi_npk_formula: 0,
                    alokasi_organik: 0,
                    realisasi_urea: item.realisasi_urea || 0,
                    realisasi_npk: item.realisasi_npk || 0,
                    realisasi_npk_formula: item.realisasi_npk_formula || 0,
                    realisasi_organik: item.realisasi_organik || 0
                };
            } else {
                mapData[item.wilayah].realisasi_urea = item.realisasi_urea || 0;
                mapData[item.wilayah].realisasi_npk = item.realisasi_npk || 0;
                mapData[item.wilayah].realisasi_npk_formula = item.realisasi_npk_formula || 0;
                mapData[item.wilayah].realisasi_organik = item.realisasi_organik || 0;
            }
        });

        // Hitung total
        const totals = {
            urea_realisasi: 0,
            npk_realisasi: 0,
            npk_formula_realisasi: 0,
            organik_realisasi: 0,
            urea_alokasi: 0,
            npk_alokasi: 0,
            npk_formula_alokasi: 0,
            organik_alokasi: 0
        };

        const finalData = Object.values(mapData).map(item => {
            totals.urea_realisasi += item.realisasi_urea;
            totals.npk_realisasi += item.realisasi_npk;
            totals.npk_formula_realisasi += item.realisasi_npk_formula;
            totals.organik_realisasi += item.realisasi_organik;
            totals.urea_alokasi += item.alokasi_urea;
            totals.npk_alokasi += item.alokasi_npk;
            totals.npk_formula_alokasi += item.alokasi_npk_formula;
            totals.organik_alokasi += item.alokasi_organik;

            return item;
        });

        return res.json({
            level: !provinsi ? "provinsi" : !kabupaten ? "kabupaten" : "kecamatan",
            data: finalData,
            totals
        });

    } catch (error) {
        console.error(error);
        return res.status(500).json({ message: "Internal server error", error: error.message });
    }
};
// Download sum.ejs
exports.downloadSummaryPupuk = async (req, res) => {
    try {
        const { provinsi, kabupaten, tahun, sumber_alokasi = "erdkk" } = req.query;

        if (!tahun) {
            return res.status(400).json({ message: "Parameter tahun diperlukan" });
        }

        const kabupatenDIY = ["SLEMAN", "BANTUL", "GUNUNG KIDUL", "KULON PROGO", "KOTA YOGYAKARTA"];

        let selectField = "kabupaten";
        let groupField = "kabupaten";
        let whereClause = "WHERE tahun = ?";
        const params = [tahun];

        // Provinsi level
        if (!provinsi) {
            selectField = `
                CASE 
                    WHEN kabupaten IN (${kabupatenDIY.map(() => "?").join(",")}) THEN 'DI Yogyakarta'
                    ELSE 'Jawa Tengah'
                END
            `;
            groupField = "wilayah";
            params.unshift(...kabupatenDIY);
        }
        // Kabupaten level
        else if (provinsi && !kabupaten) {
            if (provinsi === "DI Yogyakarta") {
                whereClause += ` AND kabupaten IN (${kabupatenDIY.map(() => "?").join(",")})`;
                params.push(...kabupatenDIY);
            } else {
                whereClause += ` AND kabupaten NOT IN (${kabupatenDIY.map(() => "?").join(",")})`;
                params.push(...kabupatenDIY);
            }
        }
        // Kecamatan level
        else if (kabupaten) {
            whereClause += ` AND kabupaten = ?`;
            params.push(kabupaten);
            selectField = "kecamatan";
            groupField = "kecamatan";
        }

        let alokasiQuery = "";
        // Query alokasi
        if (sumber_alokasi === "sk_bupati") {
            alokasiQuery = `
        SELECT 
            ${selectField} AS wilayah,
            SUM(CASE WHEN produk = 'urea' THEN alokasi ELSE 0 END) AS alokasi_urea,
            SUM(CASE WHEN produk = 'npk' THEN alokasi ELSE 0 END) AS alokasi_npk,
            SUM(CASE WHEN produk = 'npk_formula' THEN alokasi ELSE 0 END) AS alokasi_npk_formula,
            SUM(CASE WHEN produk = 'organik' THEN alokasi ELSE 0 END) AS alokasi_organik
        FROM sk_bupati
        ${whereClause}
        GROUP BY ${groupField}
    `;
        } else {
            alokasiQuery = `
        SELECT 
            ${selectField} AS wilayah,
            SUM(urea) AS alokasi_urea,
            SUM(npk) AS alokasi_npk,
            SUM(npk_formula) AS alokasi_npk_formula,
            SUM(organik) AS alokasi_organik
        FROM erdkk
        ${whereClause}
        GROUP BY ${groupField}
    `;
        }

        // Query realisasi
        const realisasiQuery = `
            SELECT 
                ${selectField.replace(/e\./g, "")} AS wilayah,
                SUM(tebus_urea) AS realisasi_urea,
                SUM(tebus_npk) AS realisasi_npk,
                SUM(tebus_npk_formula) AS realisasi_npk_formula,
                SUM(tebus_organik) AS realisasi_organik
            FROM verval_summary
            ${whereClause.replace(/e\./g, "")}
            GROUP BY ${groupField}
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
                wilayah: item.wilayah,
                alokasi_urea: item.alokasi_urea ? Math.floor(item.alokasi_urea) : 0,
                alokasi_npk: item.alokasi_npk ? Math.floor(item.alokasi_npk) : 0,
                alokasi_npk_formula: item.alokasi_npk_formula ? Math.floor(item.alokasi_npk_formula) : 0,
                alokasi_organik: item.alokasi_organik ? Math.floor(item.alokasi_organik) : 0,
                realisasi_urea: 0,
                realisasi_npk: 0,
                realisasi_npk_formula: 0,
                realisasi_organik: 0
            };
        });

        realisasi.forEach(item => {
            if (!mapData[item.wilayah]) {
                mapData[item.wilayah] = {
                    wilayah: item.wilayah,
                    alokasi_urea: 0,
                    alokasi_npk: 0,
                    alokasi_npk_formula: 0,
                    alokasi_organik: 0,
                    realisasi_urea: item.realisasi_urea ? Math.floor(item.realisasi_urea) : 0,
                    realisasi_npk: item.realisasi_npk ? Math.floor(item.realisasi_npk) : 0,
                    realisasi_npk_formula: item.realisasi_npk_formula ? Math.floor(item.realisasi_npk_formula) : 0,
                    realisasi_organik: item.realisasi_organik ? Math.floor(item.realisasi_organik) : 0
                };
            } else {
                mapData[item.wilayah].realisasi_urea = item.realisasi_urea ? Math.floor(item.realisasi_urea) : 0;
                mapData[item.wilayah].realisasi_npk = item.realisasi_npk ? Math.floor(item.realisasi_npk) : 0;
                mapData[item.wilayah].realisasi_npk_formula = item.realisasi_npk_formula ? Math.floor(item.realisasi_npk_formula) : 0;
                mapData[item.wilayah].realisasi_organik = item.realisasi_organik ? Math.floor(item.realisasi_organik) : 0;
            }
        });

        // Hitung total
        const totals = {
            urea_alokasi: alokasi.reduce((sum, item) => sum + (item.alokasi_urea || 0), 0),
            npk_alokasi: alokasi.reduce((sum, item) => sum + (item.alokasi_npk || 0), 0),
            npk_formula_alokasi: alokasi.reduce((sum, item) => sum + (item.alokasi_npk_formula || 0), 0),
            organik_alokasi: alokasi.reduce((sum, item) => sum + (item.alokasi_organik || 0), 0),
            urea_realisasi: realisasi.reduce((sum, item) => sum + (item.realisasi_urea || 0), 0),
            npk_realisasi: realisasi.reduce((sum, item) => sum + (item.realisasi_npk || 0), 0),
            npk_formula_realisasi: realisasi.reduce((sum, item) => sum + (item.realisasi_npk_formula || 0), 0),
            organik_realisasi: realisasi.reduce((sum, item) => sum + (item.realisasi_organik || 0), 0)
        };

        const finalData = Object.values(mapData).map(item => {
            return {
                wilayah: item.wilayah,
                alokasi_urea: item.alokasi_urea || 0,
                alokasi_npk: item.alokasi_npk || 0,
                alokasi_npk_formula: item.alokasi_npk_formula || 0,
                alokasi_organik: item.alokasi_organik || 0,
                realisasi_urea: item.realisasi_urea || 0,
                realisasi_npk: item.realisasi_npk || 0,
                realisasi_npk_formula: item.realisasi_npk_formula || 0,
                realisasi_organik: item.realisasi_organik || 0
            };
        });

        // Membuat workbook Excel
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Summary Pupuk');

        // Menentukan level data
        const level = !provinsi ? "Provinsi" : !kabupaten ? "Kabupaten" : "Kecamatan";

        // Header worksheet
        worksheet.mergeCells('A1:I1');
        worksheet.getCell('A1').value = `REKAPITULASI ALOKASI DAN REALISASI PUPUK TAHUN ${tahun}`;
        worksheet.getCell('A1').font = { bold: true, size: 14 };
        worksheet.getCell('A1').alignment = { horizontal: 'center' };

        worksheet.mergeCells('A2:I2');
        worksheet.getCell('A2').value = `LEVEL: ${level}`;
        worksheet.getCell('A2').font = { bold: true };
        worksheet.getCell('A2').alignment = { horizontal: 'center' };

        // Header tabel utama
        worksheet.mergeCells('A3:A4');
        worksheet.getCell('A3').value = 'WILAYAH';
        worksheet.getCell('A3').alignment = { vertical: 'middle', horizontal: 'center' };

        // Header Alokasi
        worksheet.mergeCells('B3:E3');
        worksheet.getCell('B3').value = 'TOTAL ALOKASI';
        worksheet.getCell('B3').alignment = { horizontal: 'center' };

        // Header Realisasi
        worksheet.mergeCells('F3:I3');
        worksheet.getCell('F3').value = 'TOTAL REALISASI';
        worksheet.getCell('F3').alignment = { horizontal: 'center' };

        // Sub-header kolom
        worksheet.getCell('B4').value = 'UREA';
        worksheet.getCell('C4').value = 'NPK';
        worksheet.getCell('D4').value = 'NPK FORMULA';
        worksheet.getCell('E4').value = 'ORGANIK';
        worksheet.getCell('F4').value = 'UREA';
        worksheet.getCell('G4').value = 'NPK';
        worksheet.getCell('H4').value = 'NPK FORMULA';
        worksheet.getCell('I4').value = 'ORGANIK';

        // Styling header
        ['A3', 'B3', 'F3'].forEach(cell => {
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
        ['B4', 'C4', 'D4', 'E4', 'F4', 'G4', 'H4', 'I4'].forEach(cell => {
            worksheet.getCell(cell).font = { bold: true };
            worksheet.getCell(cell).alignment = { horizontal: 'center' };
            worksheet.getCell(cell).fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFE8E8E8' }
            };
            worksheet.getCell(cell).border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
        });

        // Mengisi data
        finalData.forEach((item, index) => {
            const row = worksheet.addRow([
                item.wilayah,
                item.alokasi_urea,
                item.alokasi_npk,
                item.alokasi_npk_formula,
                item.alokasi_organik,
                item.realisasi_urea,
                item.realisasi_npk,
                item.realisasi_npk_formula,
                item.realisasi_organik
            ]);

            // Format angka
            ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I'].forEach(col => {
                const cell = row.getCell(col);
                cell.numFmt = '#,##0';
            });
        });

        // Menambahkan baris total
        const totalRow = worksheet.addRow([
            'TOTAL',
            { formula: `SUM(B5:B${worksheet.rowCount})`, result: totals.urea_alokasi },
            { formula: `SUM(C5:C${worksheet.rowCount})`, result: totals.npk_alokasi },
            { formula: `SUM(D5:D${worksheet.rowCount})`, result: totals.npk_formula_alokasi },
            { formula: `SUM(E5:E${worksheet.rowCount})`, result: totals.organik_alokasi },
            { formula: `SUM(F5:F${worksheet.rowCount})`, result: totals.urea_realisasi },
            { formula: `SUM(G5:G${worksheet.rowCount})`, result: totals.npk_realisasi },
            { formula: `SUM(H5:H${worksheet.rowCount})`, result: totals.npk_formula_realisasi },
            { formula: `SUM(I5:I${worksheet.rowCount})`, result: totals.organik_realisasi }
        ]);

        // Styling baris total
        totalRow.font = { bold: true };
        totalRow.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFC6E0B4' } // Warna hijau muda
        };

        // Format angka di baris total
        ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I'].forEach(col => {
            const cell = totalRow.getCell(col);
            cell.numFmt = '#,##0.00'; // Format dengan 2 desimal
            cell.value = {  // Pastikan nilai yang ditampilkan sesuai
                formula: cell.formula,
                result: cell.result
            };
        });

        // Set lebar kolom
        worksheet.columns = [
            { key: 'wilayah', width: 30 }, // Kolom Wilayah
            { key: 'alokasi_urea', width: 15 },
            { key: 'alokasi_npk', width: 15 },
            { key: 'alokasi_npk_formula', width: 15 },
            { key: 'alokasi_organik', width: 15 },
            { key: 'realisasi_urea', width: 15 },
            { key: 'realisasi_npk', width: 15 },
            { key: 'realisasi_npk_formula', width: 15 },
            { key: 'realisasi_organik', width: 15 }
        ];

        // Set response headers
        res.setHeader(
            'Content-Type',
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        );
        res.setHeader(
            'Content-Disposition',
            `attachment; filename="rekap_pupuk_${level.toLowerCase()}_${tahun}.xlsx"`
        );

        // Kirim workbook sebagai response
        await workbook.xlsx.write(res);
        res.end();

    } catch (error) {
        console.error(error);
        return res.status(500).json({ message: "Internal server error", error: error.message });
    }
};
// End download sum.ejs