const db = require("../config/db");
const ExcelJS = require('exceljs');
const { Worker } = require('worker_threads');
const path = require('path');
const fs = require('fs');
const { generateReport } = require('../services/reportGenerator');

exports.getF5 = async (req, res) => {
    try {
        const { start, length, draw, produk, tahun, provinsi, kabupaten } = req.query;

        // Query utama untuk mengambil data
        let query = `
        SELECT
            kode_provinsi,
            provinsi,
            kode_kabupaten,
            kabupaten, 
            kode_distributor, 
            distributor,   
            produk,

            SUM(CASE WHEN bulan = 1 THEN stok_awal ELSE 0 END) AS jan_awal,
            SUM(CASE WHEN bulan = 1 THEN penebusan ELSE 0 END) AS jan_tebus,
            SUM(CASE WHEN bulan = 1 THEN penyaluran ELSE 0 END) AS jan_salur,
            SUM(CASE WHEN bulan = 1 THEN stok_akhir ELSE 0 END) AS jan_akhir,

            SUM(CASE WHEN bulan = 2 THEN stok_awal ELSE 0 END) AS feb_awal,
            SUM(CASE WHEN bulan = 2 THEN penebusan ELSE 0 END) AS feb_tebus,
            SUM(CASE WHEN bulan = 2 THEN penyaluran ELSE 0 END) AS feb_salur,
            SUM(CASE WHEN bulan = 2 THEN stok_akhir ELSE 0 END) AS feb_akhir,

            SUM(CASE WHEN bulan = 3 THEN stok_awal ELSE 0 END) AS mar_awal,
            SUM(CASE WHEN bulan = 3 THEN penebusan ELSE 0 END) AS mar_tebus,
            SUM(CASE WHEN bulan = 3 THEN penyaluran ELSE 0 END) AS mar_salur,
            SUM(CASE WHEN bulan = 3 THEN stok_akhir ELSE 0 END) AS mar_akhir,

            SUM(CASE WHEN bulan = 4 THEN stok_awal ELSE 0 END) AS apr_awal,
            SUM(CASE WHEN bulan = 4 THEN penebusan ELSE 0 END) AS apr_tebus,
            SUM(CASE WHEN bulan = 4 THEN penyaluran ELSE 0 END) AS apr_salur,
            SUM(CASE WHEN bulan = 4 THEN stok_akhir ELSE 0 END) AS apr_akhir,

            SUM(CASE WHEN bulan = 5 THEN stok_awal ELSE 0 END) AS mei_awal,
            SUM(CASE WHEN bulan = 5 THEN penebusan ELSE 0 END) AS mei_tebus,
            SUM(CASE WHEN bulan = 5 THEN penyaluran ELSE 0 END) AS mei_salur,
            SUM(CASE WHEN bulan = 5 THEN stok_akhir ELSE 0 END) AS mei_akhir,

            SUM(CASE WHEN bulan = 6 THEN stok_awal ELSE 0 END) AS jun_awal,
            SUM(CASE WHEN bulan = 6 THEN penebusan ELSE 0 END) AS jun_tebus,
            SUM(CASE WHEN bulan = 6 THEN penyaluran ELSE 0 END) AS jun_salur,
            SUM(CASE WHEN bulan = 6 THEN stok_akhir ELSE 0 END) AS jun_akhir,

            SUM(CASE WHEN bulan = 7 THEN stok_awal ELSE 0 END) AS jul_awal,
            SUM(CASE WHEN bulan = 7 THEN penebusan ELSE 0 END) AS jul_tebus,
            SUM(CASE WHEN bulan = 7 THEN penyaluran ELSE 0 END) AS jul_salur,
            SUM(CASE WHEN bulan = 7 THEN stok_akhir ELSE 0 END) AS jul_akhir,

            SUM(CASE WHEN bulan = 8 THEN stok_awal ELSE 0 END) AS agu_awal,
            SUM(CASE WHEN bulan = 8 THEN penebusan ELSE 0 END) AS agu_tebus,
            SUM(CASE WHEN bulan = 8 THEN penyaluran ELSE 0 END) AS agu_salur,
            SUM(CASE WHEN bulan = 8 THEN stok_akhir ELSE 0 END) AS agu_akhir,

            SUM(CASE WHEN bulan = 9 THEN stok_awal ELSE 0 END) AS sep_awal,
            SUM(CASE WHEN bulan = 9 THEN penebusan ELSE 0 END) AS sep_tebus,
            SUM(CASE WHEN bulan = 9 THEN penyaluran ELSE 0 END) AS sep_salur,
            SUM(CASE WHEN bulan = 9 THEN stok_akhir ELSE 0 END) AS sep_akhir,

            SUM(CASE WHEN bulan = 10 THEN stok_awal ELSE 0 END) AS okt_awal,
            SUM(CASE WHEN bulan = 10 THEN penebusan ELSE 0 END) AS okt_tebus,
            SUM(CASE WHEN bulan = 10 THEN penyaluran ELSE 0 END) AS okt_salur,
            SUM(CASE WHEN bulan = 10 THEN stok_akhir ELSE 0 END) AS okt_akhir,

            SUM(CASE WHEN bulan = 11 THEN stok_awal ELSE 0 END) AS nov_awal,
            SUM(CASE WHEN bulan = 11 THEN penebusan ELSE 0 END) AS nov_tebus,
            SUM(CASE WHEN bulan = 11 THEN penyaluran ELSE 0 END) AS nov_salur,
            SUM(CASE WHEN bulan = 11 THEN stok_akhir ELSE 0 END) AS nov_akhir,

            SUM(CASE WHEN bulan = 12 THEN stok_awal ELSE 0 END) AS des_awal,
            SUM(CASE WHEN bulan = 12 THEN penebusan ELSE 0 END) AS des_tebus,
            SUM(CASE WHEN bulan = 12 THEN penyaluran ELSE 0 END) AS des_salur,
            SUM(CASE WHEN bulan = 12 THEN stok_akhir ELSE 0 END) AS des_akhir
            FROM f5
        WHERE 1=1
        `;

        let countQuery = `
            SELECT COUNT(*) AS total
            FROM wcm
            WHERE 1=1
        `;

        let params = [];
        let countParams = [];

        // Tambahkan filter Kabupaten
        if (produk) {
            query += " AND produk = ?";
            countQuery += " AND produk = ?";
            params.push(produk);
            countParams.push(produk);
        }

        // Tambahkan filter Tahun
        if (tahun) {
            query += " AND tahun = ?";
            countQuery += " AND tahun = ?";
            params.push(tahun);
            countParams.push(tahun);
        }

        if (provinsi) {
            query += " AND provinsi = ?";
            countQuery += " AND provinsi = ?";
            params.push(provinsi);
            countParams.push(provinsi);
        }

        if (kabupaten) {
            query += " AND kabupaten = ?";
            countQuery += " AND kabupaten = ?";
            params.push(kabupaten);
            countParams.push(kabupaten);
        }

        query += " GROUP BY kode_provinsi, provinsi, kode_kabupaten, kabupaten, kode_distributor, distributor, produk";

        // Eksekusi count query untuk mendapatkan total data
        const [totalResult] = await db.query(countQuery, countParams);
        const total = totalResult[0].total;

        // Pagination
        const startNum = parseInt(start) || 0;
        const lengthNum = parseInt(length) || 10;
        query += " LIMIT ?, ?";
        params.push(startNum, lengthNum);

        // Eksekusi query utama
        const [data] = await db.query(query, params);

        // Kirim respons
        res.json({
            draw: draw ? parseInt(draw) : 1,
            recordsTotal: total,
            recordsFiltered: total,
            data: data
        });
    } catch (error) {
        console.error("Error fetching data:", error);
        res.status(500).json({ error: "Terjadi kesalahan dalam mengambil data" });
    }
};

exports.downloadF5 = async (req, res) => {
    try {
        const { produk, tahun, provinsi, kabupaten } = req.query;

        let query = `
            SELECT 
                produk,
                kode_provinsi,
                provinsi,
                kode_kabupaten,
                kabupaten, 
                kode_distributor, 
                distributor,
                ${[...Array(12).keys()].map(bulan => `
                    SUM(CASE WHEN bulan = ${bulan + 1} THEN stok_awal ELSE 0 END) AS stok_awal_${bulan + 1},
                    SUM(CASE WHEN bulan = ${bulan + 1} THEN penebusan ELSE 0 END) AS penebusan_${bulan + 1},
                    SUM(CASE WHEN bulan = ${bulan + 1} THEN penyaluran ELSE 0 END) AS penyaluran_${bulan + 1},
                    SUM(CASE WHEN bulan = ${bulan + 1} THEN stok_akhir ELSE 0 END) AS stok_akhir_${bulan + 1}
                `).join(",")}
            FROM f5
            WHERE 1=1`;

        let params = [];

        if (produk) {
            query += " AND produk = ?";
            params.push(produk);
        }

        if (tahun) {
            query += " AND tahun = ?";
            params.push(tahun);
        }

        if (provinsi) {
            query += " AND provinsi = ?";
            params.push(provinsi);
        }

        if (kabupaten) {
            query += " AND kabupaten = ?";
            params.push(kabupaten);
        }

        query += ` GROUP BY produk, kode_provinsi, provinsi, kode_kabupaten, kabupaten, kode_distributor, distributor
                   ORDER BY provinsi, kabupaten, kode_distributor`;

        const [data] = await db.query(query, params);

        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet("WCM");

        // Define border style here so it's accessible to helper function
        const borderStyle = {
            top: { style: "thin" },
            left: { style: "thin" },
            bottom: { style: "thin" },
            right: { style: "thin" },
        };

        const bulanList = [
            "Januari", "Februari", "Maret", "April", "Mei", "Juni",
            "Juli", "Agustus", "September", "Oktober", "November", "Desember"
        ];

        // Header Utama
        worksheet.mergeCells("A1:A2"); // Produk
        worksheet.mergeCells("B1:B2"); // Kode Provinsi
        worksheet.mergeCells("C1:C2"); // Provinsi
        worksheet.mergeCells("D1:D2"); // Kode Kabupaten
        worksheet.mergeCells("E1:E2"); // Kabupaten
        worksheet.mergeCells("F1:F2"); // Kode Distributor
        worksheet.mergeCells("G1:G2"); // Distributor
        worksheet.getCell("A1").value = "PRODUK";
        worksheet.getCell("B1").value = "KODE PROVINSI";
        worksheet.getCell("C1").value = "PROVINSI";
        worksheet.getCell("D1").value = "KODE KABUPATEN";
        worksheet.getCell("E1").value = "KABUPATEN";
        worksheet.getCell("F1").value = "KODE DISTRIBUTOR";
        worksheet.getCell("G1").value = "DISTRIBUTOR";

        // Header Bulan
        let colStart = 8; // Mulai dari kolom H
        bulanList.forEach((bulan, index) => {
            let colEnd = colStart + 3;
            worksheet.mergeCells(1, colStart, 1, colEnd);
            worksheet.getCell(1, colStart).value = bulan;
            worksheet.getCell(1, colStart).alignment = { horizontal: "center", vertical: "middle" };
            worksheet.getCell(1, colStart).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFDD4B39" } };

            // Subheader Bulan
            worksheet.getCell(2, colStart).value = "STOK AWAL";
            worksheet.getCell(2, colStart + 1).value = "TEBUS";
            worksheet.getCell(2, colStart + 2).value = "SALUR";
            worksheet.getCell(2, colStart + 3).value = "STOK AKHIR";

            colStart += 4;
        });

        // Format Header
        [1, 2].forEach(rowNum => {
            worksheet.getRow(rowNum).eachCell((cell) => {
                cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFDD4B39" } };
                cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
                cell.alignment = { horizontal: "center", vertical: "middle" };
                cell.border = borderStyle;
            });
        });

        // Atur lebar kolom
        worksheet.columns = [
            { key: "produk", width: 20 },
            { key: "kode_provinsi", width: 15 },
            { key: "provinsi", width: 20 },
            { key: "kode_kabupaten", width: 15 },
            { key: "kabupaten", width: 20 },
            { key: "kode_distributor", width: 15 },
            { key: "distributor", width: 25 },
            ...Array.from({ length: 12 * 4 }, () => ({ width: 12 }))
        ];

        // Kelompokkan data berdasarkan kabupaten
        const groupedByKabupaten = {};
        let currentKabupaten = null;
        let kabupatenStartRow = 3; // Mulai dari row 3 (setelah header)

        data.forEach((row, index) => {
            if (row.kabupaten !== currentKabupaten) {
                if (currentKabupaten !== null) {
                    // Add total for previous kabupaten
                    addKabupatenTotal(worksheet, groupedByKabupaten[currentKabupaten], kabupatenStartRow, index + 2, borderStyle);
                }
                currentKabupaten = row.kabupaten;
                kabupatenStartRow = index + 3; // +3 karena header 2 row dan data mulai row 3
                groupedByKabupaten[currentKabupaten] = [];
            }
            groupedByKabupaten[currentKabupaten].push(row);

            // Add data row
            let rowData = [
                row.produk,
                row.kode_provinsi,
                row.provinsi,
                row.kode_kabupaten,
                row.kabupaten,
                row.kode_distributor,
                row.distributor
            ];

            for (let i = 1; i <= 12; i++) {
                rowData.push(
                    parseFloat((row[`stok_awal_${i}`] || 0).toFixed(3)),
                    parseFloat((row[`penebusan_${i}`] || 0).toFixed(3)),
                    parseFloat((row[`penyaluran_${i}`] || 0).toFixed(3)),
                    parseFloat((row[`stok_akhir_${i}`] || 0).toFixed(3))
                );
            }

            let addedRow = worksheet.addRow(rowData);
            addedRow.eachCell((cell) => {
                cell.border = borderStyle;
                cell.alignment = { horizontal: "center", vertical: "middle" };
            });
        });

        // Add total for last kabupaten
        if (currentKabupaten !== null) {
            addKabupatenTotal(worksheet, groupedByKabupaten[currentKabupaten], kabupatenStartRow, data.length + 2, borderStyle);
        }

        // Kirim file Excel
        res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        res.setHeader("Content-Disposition", "attachment; filename=wcm.xlsx");
        await workbook.xlsx.write(res);
        res.end();
    } catch (error) {
        console.error(error);
        res.status(500).json({ error: "Terjadi kesalahan dalam mengunduh data" });
    }
};
// Helper function to add kabupaten total
function addKabupatenTotal(worksheet, kabupatenData, startRow, endRow, borderStyle) {
    const totalRow = ["TOTAL", "", "", "", kabupatenData[0].kabupaten, "", ""];
    const monthlyTotals = Array(12).fill().map(() => ({
        stok_awal: 0,
        penebusan: 0,
        penyaluran: 0,
        stok_akhir: 0
    }));

    // Calculate totals
    kabupatenData.forEach(row => {
        for (let i = 1; i <= 12; i++) {
            monthlyTotals[i - 1].stok_awal += row[`stok_awal_${i}`] || 0;
            monthlyTotals[i - 1].penebusan += row[`penebusan_${i}`] || 0;
            monthlyTotals[i - 1].penyaluran += row[`penyaluran_${i}`] || 0;
            monthlyTotals[i - 1].stok_akhir += row[`stok_akhir_${i}`] || 0;
        }
    });

    // Add monthly totals to row
    monthlyTotals.forEach(month => {
        totalRow.push(
            parseFloat(month.stok_awal.toFixed(3)),
            parseFloat(month.penebusan.toFixed(3)),
            parseFloat(month.penyaluran.toFixed(3)),
            parseFloat(month.stok_akhir.toFixed(3))
        );
    });

    // Add total row
    const totalRowAdded = worksheet.addRow(totalRow);
    totalRowAdded.eachCell((cell) => {
        cell.border = borderStyle;
        cell.font = { bold: true };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD3D3D3" } };
        cell.alignment = { horizontal: "center", vertical: "middle" };
    });

    // Add empty row after total
    worksheet.addRow([]);
}