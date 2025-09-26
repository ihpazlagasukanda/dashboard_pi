const db = require("../config/db");
const ExcelJS = require('exceljs');
const { Worker } = require('worker_threads');
const path = require('path');
const fs = require('fs');
const { generateReport } = require('../services/reportGenerator');

exports.rekapPetaniAll = async (req, res) => {
    try {
        const { provinsi, kabupaten, tahun } = req.query;

        if (!tahun) {
            return res.json({
                level: !provinsi ? "provinsi" : !kabupaten ? "kabupaten" : "kecamatan",
                data: [],
                totals: {
                    alokasi: 0,
                    realisasi: 0,
                    belum_tebus: 0
                }
            });
        }

        const kabupatenDIY = ["SLEMAN", "BANTUL", "GUNUNG KIDUL", "KULON PROGO", "KOTA YOGYAKARTA"];

        let selectField = "e.kabupaten AS wilayah";
        let groupField = "e.kabupaten";
        let whereClause = "WHERE e.tahun = ?";
        let params = [tahun];

        // Provinsi level
        if (!provinsi) {
            selectField = `
                CASE 
                    WHEN e.kabupaten IN (${kabupatenDIY.map(() => "?").join(",")}) THEN 'DI Yogyakarta'
                    ELSE 'Jawa Tengah'
                END AS wilayah
            `;
            groupField = "wilayah";
            // Parameter untuk CASE statement ditambahkan di depan
            params = [...kabupatenDIY, ...params];
        }
        // Kabupaten level
        else if (provinsi && !kabupaten) {
            if (provinsi === "DI Yogyakarta") {
                whereClause += ` AND e.kabupaten IN (${kabupatenDIY.map(() => "?").join(",")})`;
                params = [...params, ...kabupatenDIY];
            } else {
                whereClause += ` AND e.kabupaten NOT IN (${kabupatenDIY.map(() => "?").join(",")})`;
                params = [...params, ...kabupatenDIY];
            }
        }
        // Kecamatan level
        else if (kabupaten) {
            whereClause += ` AND e.kabupaten = ?`;
            params.push(kabupaten);
            selectField = "e.kecamatan AS wilayah";
            groupField = "e.kecamatan";
        }

        console.log("DEBUG - Query Parameters:", params);
        console.log("DEBUG - Where Clause:", whereClause);
        console.log("DEBUG - Select Field:", selectField);

        // Query alokasi - TOTAL PETANI di erdkk
        const alokasiQuery = `
            SELECT 
                ${selectField},
                COUNT(DISTINCT e.nik) AS total_alokasi
            FROM erdkk e
            ${whereClause}
            GROUP BY ${groupField}
        `;

        // Query realisasi - TOTAL PETANI yang sudah tebus
        const realisasiQuery = `
            SELECT 
                ${selectField},
                COUNT(DISTINCT e.nik) AS total_realisasi
            FROM erdkk e
            INNER JOIN verval_summary v ON e.nik = v.nik AND e.tahun = v.tahun
            ${whereClause}
            GROUP BY ${groupField}
        `;

        // Query belum tebus - PETANI yang ada di erdkk tapi TIDAK ada di verval_summary
        const belumTebusQuery = `
            SELECT 
                ${selectField},
                COUNT(DISTINCT e.nik) AS total_belum_tebus
            FROM erdkk e
            LEFT JOIN verval_summary v ON e.nik = v.nik AND e.tahun = v.tahun
            ${whereClause} AND v.nik IS NULL
            GROUP BY ${groupField}
        `;

        console.log("DEBUG - Alokasi Query:", alokasiQuery);
        console.log("DEBUG - Realisasi Query:", realisasiQuery);

        // Eksekusi query paralel
        const [alokasiData, realisasiData, belumTebusData] = await Promise.all([
            db.query(alokasiQuery, params),
            db.query(realisasiQuery, params),
            db.query(belumTebusQuery, params)
        ]);

        console.log("DEBUG - Alokasi Result:", alokasiData[0]);
        console.log("DEBUG - Realisasi Result:", realisasiData[0]);
        console.log("DEBUG - Belum Tebus Result:", belumTebusData[0]);

        const alokasi = alokasiData[0] || [];
        const realisasi = realisasiData[0] || [];
        const belumTebus = belumTebusData[0] || [];

        // Gabungkan data
        const mapData = {};

        // Inisialisasi semua wilayah dari data alokasi
        alokasi.forEach(item => {
            mapData[item.wilayah] = {
                wilayah: item.wilayah,
                alokasi: parseInt(item.total_alokasi) || 0,
                realisasi: 0,
                belum_tebus: 0
            };
        });

        // Tambahkan data realisasi
        realisasi.forEach(item => {
            if (mapData[item.wilayah]) {
                mapData[item.wilayah].realisasi = parseInt(item.total_realisasi) || 0;
            } else {
                mapData[item.wilayah] = {
                    wilayah: item.wilayah,
                    alokasi: 0,
                    realisasi: parseInt(item.total_realisasi) || 0,
                    belum_tebus: 0
                };
            }
        });

        // Tambahkan data belum tebus
        belumTebus.forEach(item => {
            if (mapData[item.wilayah]) {
                mapData[item.wilayah].belum_tebus = parseInt(item.total_belum_tebus) || 0;
            } else {
                mapData[item.wilayah] = {
                    wilayah: item.wilayah,
                    alokasi: 0,
                    realisasi: 0,
                    belum_tebus: parseInt(item.total_belum_tebus) || 0
                };
            }
        });

        // Hitung total
        const totals = {
            alokasi: 0,
            realisasi: 0,
            belum_tebus: 0
        };

        const finalData = Object.values(mapData).map(item => {
            totals.alokasi += item.alokasi;
            totals.realisasi += item.realisasi;
            totals.belum_tebus += item.belum_tebus;

            return item;
        });

        return res.json({
            level: !provinsi ? "provinsi" : !kabupaten ? "kabupaten" : "kecamatan",
            data: finalData,
            totals
        });

    } catch (error) {
        console.error("ERROR DETAIL:", error);
        return res.status(500).json({ message: "Internal server error", error: error.message });
    }
};
exports.downloadRekapPetaniAll = async (req, res) => {
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

        // Query alokasi - TOTAL PETANI di erdkk (per NIK, tanpa peduli pupuk)
        const alokasiQuery = `
            SELECT 
                ${selectField} AS wilayah,
                COUNT(DISTINCT nik) AS total_alokasi
            FROM erdkk
            ${whereClause}
            GROUP BY ${groupField}
        `;

        // Query realisasi - TOTAL PETANI yang sudah tebus (per NIK)
        const realisasiQuery = `
            SELECT 
                ${selectField} AS wilayah,
                COUNT(DISTINCT e.nik) AS total_realisasi
            FROM erdkk e
            INNER JOIN verval_summary v ON e.nik = v.nik AND e.tahun = v.tahun
            ${whereClause}
            GROUP BY ${groupField}
        `;

        // Query belum tebus - PETANI yang ada di erdkk tapi TIDAK ada di verval_summary
        const belumTebusQuery = `
            SELECT 
                ${selectField} AS wilayah,
                COUNT(DISTINCT e.nik) AS total_belum_tebus
            FROM erdkk e
            LEFT JOIN verval_summary v ON e.nik = v.nik AND e.tahun = v.tahun
            ${whereClause} AND v.nik IS NULL
            GROUP BY ${groupField}
        `;

        // Eksekusi query paralel
        const [alokasiData, realisasiData, belumTebusData] = await Promise.all([
            db.query(alokasiQuery, params),
            db.query(realisasiQuery, params),
            db.query(belumTebusQuery, params)
        ]);

        const alokasi = alokasiData[0] || [];
        const realisasi = realisasiData[0] || [];
        const belumTebus = belumTebusData[0] || [];

        // Gabungkan data
        const mapData = {};

        // Inisialisasi semua wilayah dari data alokasi
        alokasi.forEach(item => {
            mapData[item.wilayah] = {
                wilayah: item.wilayah,
                alokasi: item.total_alokasi || 0,
                realisasi: 0,
                belum_tebus: 0
            };
        });

        // Tambahkan data realisasi
        realisasi.forEach(item => {
            if (mapData[item.wilayah]) {
                mapData[item.wilayah].realisasi = item.total_realisasi || 0;
            } else {
                mapData[item.wilayah] = {
                    wilayah: item.wilayah,
                    alokasi: 0,
                    realisasi: item.total_realisasi || 0,
                    belum_tebus: 0
                };
            }
        });

        // Tambahkan data belum tebus
        belumTebus.forEach(item => {
            if (mapData[item.wilayah]) {
                mapData[item.wilayah].belum_tebus = item.total_belum_tebus || 0;
            } else {
                mapData[item.wilayah] = {
                    wilayah: item.wilayah,
                    alokasi: 0,
                    realisasi: 0,
                    belum_tebus: item.total_belum_tebus || 0
                };
            }
        });

        // Pastikan semua wilayah memiliki data yang lengkap
        Object.keys(mapData).forEach(wilayah => {
            const data = mapData[wilayah];
            // Jika belum_tebus belum diisi, hitung dari alokasi - realisasi
            if (data.belum_tebus === 0 && data.alokasi > 0) {
                data.belum_tebus = data.alokasi - data.realisasi;
            }
        });

        // Hitung total
        const totals = {
            alokasi: alokasi.reduce((sum, item) => sum + (item.total_alokasi || 0), 0),
            realisasi: realisasi.reduce((sum, item) => sum + (item.total_realisasi || 0), 0),
            belum_tebus: belumTebus.reduce((sum, item) => sum + (item.total_belum_tebus || 0), 0)
        };

        // Jika total belum tebus tidak match, hitung ulang
        if (totals.belum_tebus !== totals.alokasi - totals.realisasi) {
            totals.belum_tebus = totals.alokasi - totals.realisasi;
        }

        const finalData = Object.values(mapData).map(item => {
            return {
                wilayah: item.wilayah,
                alokasi: item.alokasi || 0,
                realisasi: item.realisasi || 0,
                belum_tebus: item.belum_tebus || 0
            };
        });

        // Membuat workbook Excel
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Rekap Petani');

        // Menentukan level data
        const level = !provinsi ? "Provinsi" : !kabupaten ? "Kabupaten" : "Kecamatan";

        // Header worksheet
        worksheet.mergeCells('A1:D1');
        worksheet.getCell('A1').value = `REKAPITULASI PETANI PUPUK BERSUBSIDI TAHUN ${tahun}`;
        worksheet.getCell('A1').font = { bold: true, size: 14 };
        worksheet.getCell('A1').alignment = { horizontal: 'center' };

        worksheet.mergeCells('A2:D2');
        worksheet.getCell('A2').value = `LEVEL: ${level}`;
        worksheet.getCell('A2').font = { bold: true };
        worksheet.getCell('A2').alignment = { horizontal: 'center' };

        // Header tabel
        worksheet.getCell('A3').value = 'WILAYAH';
        worksheet.getCell('B3').value = 'TOTAL PETANI (ALOKASI)';
        worksheet.getCell('C3').value = 'SUDAH TEBUS';
        worksheet.getCell('D3').value = 'BELUM TEBUS';

        // Styling header
        ['A3', 'B3', 'C3', 'D3'].forEach(cell => {
            worksheet.getCell(cell).font = { bold: true };
            worksheet.getCell(cell).fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFD3D3D3' }
            };
            worksheet.getCell(cell).alignment = { horizontal: 'center' };
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
                item.alokasi,
                item.realisasi,
                item.belum_tebus
            ]);

            // Format angka
            ['B', 'C', 'D'].forEach(col => {
                const cell = row.getCell(col);
                cell.numFmt = '#,##0';
                cell.alignment = { horizontal: 'right' };
            });

            // Alternating row colors
            if (index % 2 === 0) {
                row.eachCell((cell) => {
                    cell.fill = {
                        type: 'pattern',
                        pattern: 'solid',
                        fgColor: { argb: 'FFF0F0F0' }
                    };
                });
            }

            // Border untuk setiap sel
            row.eachCell((cell) => {
                cell.border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' }
                };
            });
        });

        // Menambahkan baris total
        const totalRow = worksheet.addRow([
            'TOTAL',
            { formula: `SUM(B4:B${worksheet.rowCount})`, result: totals.alokasi },
            { formula: `SUM(C4:C${worksheet.rowCount})`, result: totals.realisasi },
            { formula: `SUM(D4:D${worksheet.rowCount})`, result: totals.belum_tebus }
        ]);

        // Styling baris total
        totalRow.font = { bold: true };
        totalRow.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFC6E0B4' } // Warna hijau muda
        };

        // Format angka di baris total
        ['B', 'C', 'D'].forEach(col => {
            const cell = totalRow.getCell(col);
            cell.numFmt = '#,##0';
            cell.alignment = { horizontal: 'right' };
            cell.value = {
                formula: cell.formula,
                result: cell.result
            };
        });

        // Border untuk baris total
        totalRow.eachCell((cell) => {
            cell.border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
        });

        // Set lebar kolom
        worksheet.columns = [
            { key: 'wilayah', width: 30 },
            { key: 'alokasi', width: 20 },
            { key: 'realisasi', width: 15 },
            { key: 'belum_tebus', width: 15 }
        ];

        // Tambahkan catatan
        worksheet.mergeCells(`A${worksheet.rowCount + 2}:D${worksheet.rowCount + 2}`);
        worksheet.getCell(`A${worksheet.rowCount}`).value = 'Keterangan:';
        worksheet.getCell(`A${worksheet.rowCount}`).font = { bold: true };

        worksheet.mergeCells(`A${worksheet.rowCount + 1}:D${worksheet.rowCount + 1}`);
        worksheet.getCell(`A${worksheet.rowCount}`).value = '1. Total Petani (Alokasi): Jumlah petani yang mendapatkan alokasi pupuk bersubsidi';
        
        worksheet.mergeCells(`A${worksheet.rowCount + 1}:D${worksheet.rowCount + 1}`);
        worksheet.getCell(`A${worksheet.rowCount}`).value = '2. Sudah Tebus: Petani yang sudah melakukan penebusan pupuk';
        
        worksheet.mergeCells(`A${worksheet.rowCount + 1}:D${worksheet.rowCount + 1}`);
        worksheet.getCell(`A${worksheet.rowCount}`).value = '3. Belum Tebus: Petani yang mendapatkan alokasi tetapi belum menebus pupuk';

        // Set response headers
        res.setHeader(
            'Content-Type',
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        );
        res.setHeader(
            'Content-Disposition',
            `attachment; filename="rekap_petani_${level.toLowerCase()}_${tahun}.xlsx"`
        );

        // Kirim workbook sebagai response
        await workbook.xlsx.write(res);
        res.end();

    } catch (error) {
        console.error(error);
        return res.status(500).json({ message: "Internal server error", error: error.message });
    }
};