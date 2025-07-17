const db = require("../config/db");
const ExcelJS = require('exceljs');
const { Worker } = require('worker_threads');
const path = require('path');
const fs = require('fs');
const { generateReport } = require('../services/reportGenerator');

exports.rekapPetani = async (req, res) => {
    try {
        const { provinsi, kabupaten, tahun } = req.query;

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
                    organik_alokasi: 0,
                    urea_sisa: 0,
                    npk_sisa: 0,
                    npk_formula_sisa: 0,
                    organik_sisa: 0
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

        // Query alokasi
        const alokasiQuery = `
            SELECT 
                ${selectField} AS wilayah,
                COUNT(DISTINCT CASE WHEN urea > 0 THEN nik END) AS alokasi_urea,
                COUNT(DISTINCT CASE WHEN npk > 0 THEN nik END) AS alokasi_npk,
                COUNT(DISTINCT CASE WHEN npk_formula > 0 THEN nik END) AS alokasi_npk_formula, 
                COUNT(DISTINCT CASE WHEN organik > 0 THEN nik END) AS alokasi_organik
            FROM erdkk
            ${whereClause}
            GROUP BY ${groupField}
        `;

        // Query realisasi
        const realisasiQuery = `
            SELECT 
                ${selectField.replace(/e\./g, "")} AS wilayah,
                COUNT(DISTINCT CASE WHEN tebus_urea > 0 THEN nik END) AS realisasi_urea,
                COUNT(DISTINCT CASE WHEN tebus_npk > 0 THEN nik END) AS realisasi_npk,
                COUNT(DISTINCT CASE WHEN tebus_npk_formula > 0 THEN nik END) AS realisasi_npk_formula, 
                COUNT(DISTINCT CASE WHEN tebus_organik > 0 THEN nik END) AS realisasi_organik
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
                realisasi_organik: 0,
                sisa_urea: item.alokasi_urea || 0, // Initialize sisa with alokasi value
                sisa_npk: item.alokasi_npk || 0,
                sisa_npk_formula: item.alokasi_npk_formula || 0,
                sisa_organik: item.alokasi_organik || 0
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
                    realisasi_organik: item.realisasi_organik || 0,
                    sisa_urea: 0 - (item.realisasi_urea || 0), // If no alokasi, sisa is negative realisasi
                    sisa_npk: 0 - (item.realisasi_npk || 0),
                    sisa_npk_formula: 0 - (item.realisasi_npk_formula || 0),
                    sisa_organik: 0 - (item.realisasi_organik || 0)
                };
            } else {
                mapData[item.wilayah].realisasi_urea = item.realisasi_urea || 0;
                mapData[item.wilayah].realisasi_npk = item.realisasi_npk || 0;
                mapData[item.wilayah].realisasi_npk_formula = item.realisasi_npk_formula || 0;
                mapData[item.wilayah].realisasi_organik = item.realisasi_organik || 0;

                // Calculate sisa (alokasi - realisasi)
                mapData[item.wilayah].sisa_urea = mapData[item.wilayah].alokasi_urea - (item.realisasi_urea || 0);
                mapData[item.wilayah].sisa_npk = mapData[item.wilayah].alokasi_npk - (item.realisasi_npk || 0);
                mapData[item.wilayah].sisa_npk_formula = mapData[item.wilayah].alokasi_npk_formula - (item.realisasi_npk_formula || 0);
                mapData[item.wilayah].sisa_organik = mapData[item.wilayah].alokasi_organik - (item.realisasi_organik || 0);
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
            organik_alokasi: 0,
            urea_sisa: 0,
            npk_sisa: 0,
            npk_formula_sisa: 0,
            organik_sisa: 0
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
            totals.urea_sisa += item.sisa_urea;
            totals.npk_sisa += item.sisa_npk;
            totals.npk_formula_sisa += item.sisa_npk_formula;
            totals.organik_sisa += item.sisa_organik;

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

exports.downloadRekapPetani = async (req, res) => {
    try {
        const { provinsi, kabupaten, tahun } = req.query;

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

        // Query alokasi
        const alokasiQuery = `
            SELECT 
                ${selectField} AS wilayah,
                COUNT(DISTINCT CASE WHEN urea > 0 THEN nik END) AS alokasi_urea,
                COUNT(DISTINCT CASE WHEN npk > 0 THEN nik END) AS alokasi_npk,
                COUNT(DISTINCT CASE WHEN npk_formula > 0 THEN nik END) AS alokasi_npk_formula, 
                COUNT(DISTINCT CASE WHEN organik > 0 THEN nik END) AS alokasi_organik
            FROM erdkk
            ${whereClause}
            GROUP BY ${groupField}
        `;

        // Query realisasi
        const realisasiQuery = `
            SELECT 
                ${selectField.replace(/e\./g, "")} AS wilayah,
                COUNT(DISTINCT CASE WHEN tebus_urea > 0 THEN nik END) AS realisasi_urea,
                COUNT(DISTINCT CASE WHEN tebus_npk > 0 THEN nik END) AS realisasi_npk,
                COUNT(DISTINCT CASE WHEN tebus_npk_formula > 0 THEN nik END) AS realisasi_npk_formula, 
                COUNT(DISTINCT CASE WHEN tebus_organik > 0 THEN nik END) AS realisasi_organik
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
                realisasi_organik: 0,
                sisa_urea: item.alokasi_urea ? Math.floor(item.alokasi_urea) : 0, // Initialize sisa with alokasi value
                sisa_npk: item.alokasi_npk ? Math.floor(item.alokasi_npk) : 0,
                sisa_npk_formula: item.alokasi_npk_formula ? Math.floor(item.alokasi_npk_formula) : 0,
                sisa_organik: item.alokasi_organik ? Math.floor(item.alokasi_organik) : 0
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
                    realisasi_organik: item.realisasi_organik ? Math.floor(item.realisasi_organik) : 0,
                    sisa_urea: 0 - (item.realisasi_urea ? Math.floor(item.realisasi_urea) : 0), // If no alokasi, sisa is negative realisasi
                    sisa_npk: 0 - (item.realisasi_npk ? Math.floor(item.realisasi_npk) : 0),
                    sisa_npk_formula: 0 - (item.realisasi_npk_formula ? Math.floor(item.realisasi_npk_formula) : 0),
                    sisa_organik: 0 - (item.realisasi_organik ? Math.floor(item.realisasi_organik) : 0)
                };
            } else {
                const realUrea = item.realisasi_urea ? Math.floor(item.realisasi_urea) : 0;
                const realNpk = item.realisasi_npk ? Math.floor(item.realisasi_npk) : 0;
                const realNpkFormula = item.realisasi_npk_formula ? Math.floor(item.realisasi_npk_formula) : 0;
                const realOrganik = item.realisasi_organik ? Math.floor(item.realisasi_organik) : 0;

                mapData[item.wilayah].realisasi_urea = realUrea;
                mapData[item.wilayah].realisasi_npk = realNpk;
                mapData[item.wilayah].realisasi_npk_formula = realNpkFormula;
                mapData[item.wilayah].realisasi_organik = realOrganik;

                // Calculate sisa (alokasi - realisasi)
                mapData[item.wilayah].sisa_urea = mapData[item.wilayah].alokasi_urea - realUrea;
                mapData[item.wilayah].sisa_npk = mapData[item.wilayah].alokasi_npk - realNpk;
                mapData[item.wilayah].sisa_npk_formula = mapData[item.wilayah].alokasi_npk_formula - realNpkFormula;
                mapData[item.wilayah].sisa_organik = mapData[item.wilayah].alokasi_organik - realOrganik;
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
            organik_realisasi: realisasi.reduce((sum, item) => sum + (item.realisasi_organik || 0), 0),
            urea_sisa: alokasi.reduce((sum, item) => sum + (item.alokasi_urea || 0), 0) -
                realisasi.reduce((sum, item) => sum + (item.realisasi_urea || 0), 0),
            npk_sisa: alokasi.reduce((sum, item) => sum + (item.alokasi_npk || 0), 0) -
                realisasi.reduce((sum, item) => sum + (item.realisasi_npk || 0), 0),
            npk_formula_sisa: alokasi.reduce((sum, item) => sum + (item.alokasi_npk_formula || 0), 0) -
                realisasi.reduce((sum, item) => sum + (item.realisasi_npk_formula || 0), 0),
            organik_sisa: alokasi.reduce((sum, item) => sum + (item.alokasi_organik || 0), 0) -
                realisasi.reduce((sum, item) => sum + (item.realisasi_organik || 0), 0)
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
                realisasi_organik: item.realisasi_organik || 0,
                sisa_urea: item.sisa_urea || 0,
                sisa_npk: item.sisa_npk || 0,
                sisa_npk_formula: item.sisa_npk_formula || 0,
                sisa_organik: item.sisa_organik || 0
            };
        });

        // Membuat workbook Excel
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Summary Pupuk');

        // Menentukan level data
        const level = !provinsi ? "Provinsi" : !kabupaten ? "Kabupaten" : "Kecamatan";

        // Header worksheet
        worksheet.mergeCells('A1:M1');
        worksheet.getCell('A1').value = `REKAPITULASI ALOKASI DAN REALISASI PUPUK TAHUN ${tahun}`;
        worksheet.getCell('A1').font = { bold: true, size: 14 };
        worksheet.getCell('A1').alignment = { horizontal: 'center' };

        worksheet.mergeCells('A2:M2');
        worksheet.getCell('A2').value = `LEVEL: ${level}`;
        worksheet.getCell('A2').font = { bold: true };
        worksheet.getCell('A2').alignment = { horizontal: 'center' };

        // Header tabel utama
        worksheet.mergeCells('A3:A4');
        worksheet.getCell('A3').value = 'WILAYAH';
        worksheet.getCell('A3').alignment = { vertical: 'middle', horizontal: 'center' };

        // Header Alokasi
        worksheet.mergeCells('B3:E3');
        worksheet.getCell('B3').value = 'TOTAL PETANI';
        worksheet.getCell('B3').alignment = { horizontal: 'center' };

        // Header Realisasi
        worksheet.mergeCells('F3:I3');
        worksheet.getCell('F3').value = 'TOTAL PETANI TEBUS';
        worksheet.getCell('F3').alignment = { horizontal: 'center' };

        // Header Sisa Tebus
        worksheet.mergeCells('J3:M3');
        worksheet.getCell('J3').value = 'TOTAL SISA TEBUS';
        worksheet.getCell('J3').alignment = { horizontal: 'center' };

        // Sub-header kolom
        worksheet.getCell('B4').value = 'UREA';
        worksheet.getCell('C4').value = 'NPK';
        worksheet.getCell('D4').value = 'NPK FORMULA';
        worksheet.getCell('E4').value = 'ORGANIK';
        worksheet.getCell('F4').value = 'UREA';
        worksheet.getCell('G4').value = 'NPK';
        worksheet.getCell('H4').value = 'NPK FORMULA';
        worksheet.getCell('I4').value = 'ORGANIK';
        worksheet.getCell('J4').value = 'UREA';
        worksheet.getCell('K4').value = 'NPK';
        worksheet.getCell('L4').value = 'NPK FORMULA';
        worksheet.getCell('M4').value = 'ORGANIK';

        // Styling header
        ['A3', 'B3', 'F3', 'J3'].forEach(cell => {
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
        ['B4', 'C4', 'D4', 'E4', 'F4', 'G4', 'H4', 'I4', 'J4', 'K4', 'L4', 'M4'].forEach(cell => {
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
                item.realisasi_organik,
                item.sisa_urea,
                item.sisa_npk,
                item.sisa_npk_formula,
                item.sisa_organik
            ]);

            // Format angka
            ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'].forEach(col => {
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
            { formula: `SUM(I5:I${worksheet.rowCount})`, result: totals.organik_realisasi },
            { formula: `SUM(J5:J${worksheet.rowCount})`, result: totals.urea_sisa },
            { formula: `SUM(K5:K${worksheet.rowCount})`, result: totals.npk_sisa },
            { formula: `SUM(L5:L${worksheet.rowCount})`, result: totals.npk_formula_sisa },
            { formula: `SUM(M5:M${worksheet.rowCount})`, result: totals.organik_sisa }
        ]);

        // Styling baris total
        totalRow.font = { bold: true };
        totalRow.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFC6E0B4' } // Warna hijau muda
        };

        // Format angka di baris total
        ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'].forEach(col => {
            const cell = totalRow.getCell(col);
            cell.numFmt = '#,##0';
            cell.value = {
                formula: cell.formula,
                result: cell.result
            };
        });

        // Set lebar kolom
        worksheet.columns = [
            { key: 'wilayah', width: 30 },
            { key: 'alokasi_urea', width: 12 },
            { key: 'alokasi_npk', width: 12 },
            { key: 'alokasi_npk_formula', width: 12 },
            { key: 'alokasi_organik', width: 12 },
            { key: 'realisasi_urea', width: 12 },
            { key: 'realisasi_npk', width: 12 },
            { key: 'realisasi_npk_formula', width: 12 },
            { key: 'realisasi_organik', width: 12 },
            { key: 'sisa_urea', width: 12 },
            { key: 'sisa_npk', width: 12 },
            { key: 'sisa_npk_formula', width: 12 },
            { key: 'sisa_organik', width: 12 }
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
