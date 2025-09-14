const db = require("../config/db");
const ExcelJS = require('exceljs');
const { Worker } = require('worker_threads');
const path = require('path');
const fs = require('fs');
const { generateReport } = require('../services/reportGenerator');
const client = require('../redisClient');


exports.getDataWithPagination = async (req, res) => {
    try {
        const page = parseInt(req.query.page) || 1; // Default halaman = 1
        const limitParam = req.query.limit; // Bisa angka atau 'all'
        let query;
        let limit = 100; // Default 50 data per halaman
        let offset = 0; // Default offset

        if (limitParam === "all") {
            // Jika limit=all, ambil semua data tanpa LIMIT dan OFFSET
            query = `SELECT * FROM verval`;
        } else {
            limit = parseInt(limitParam) || 50; // Jika bukan "all", konversi limit ke angka
            offset = (page - 1) * limit;
            query = `SELECT * FROM verval LIMIT ${limit} OFFSET ${offset}`;
        }

        // Debug query di terminal
        console.log("Query:", query);

        // Eksekusi query
        const [rows] = await db.execute(query);

        // Hitung total data
        const [[{ total }]] = await db.execute(`SELECT COUNT(*) AS total FROM verval`);

        res.json({
            data: rows,
            total,
            totalPages: limitParam === "all" ? 1 : Math.ceil(total / limit),
            currentPage: limitParam === "all" ? 1 : page,
        });
    } catch (error) {
        console.error("Database Error:", error);
        res.status(500).json({ message: "Error mengambil data", error });
    }
};

exports.getAllData = async (req, res) => {
    try {
        const page = parseInt(req.query.page) || 1;
        const limit = parseInt(req.query.limit) || 1000; // Batas jumlah data per request
        const offset = (page - 1) * limit;

        const query = `SELECT id, kabupaten, kecamatan, kode_kios, nik, nama_petani, metode_penebusan, 
                      DATE_FORMAT(tanggal_tebus, '%Y-%m') AS tanggal_tebus, urea, npk, sp36, za, 
                      npk_formula, organik, organik_cair, kakao, status 
                      FROM verval 
                      LIMIT ${limit} OFFSET ${offset}`;

        console.time("Query Execution Time");
        const [rows] = await db.execute(query);
        console.timeEnd("Query Execution Time");

        // Hitung total data
        const [[{ total }]] = await db.execute(`SELECT COUNT(*) AS total FROM verval`);

        res.json({
            data: rows,
            total,
            totalPages: Math.ceil(total / limit),
            currentPage: page,
        });

    } catch (error) {
        console.error("Database Error:", error);
        res.status(500).json({ message: "Error mengambil data", error });
    }
};



// Endpoint ini khusus untuk dashboard (SUM per kabupaten per bulan)
exports.getSummaryData = async (req, res) => {
    try {
        const { tahun, provinsi, kabupaten, kecamatan } = req.query;

        let query = `
            SELECT 
                kabupaten,
                kecamatan,
                DATE_FORMAT(tanggal_tebus, '%b') AS bulan,
                SUM(urea + npk + sp36 + za + npk_formula + organik + organik_cair + kakao) AS total
            FROM verval
            WHERE 1 = 1
        `;

        const params = [];

        // Filter Tahun
        if (tahun) {
            query += " AND YEAR(tanggal_tebus) = ?";
            params.push(tahun);
        }

        // Filter Kabupaten
        if (kabupaten) {
            query += " AND kabupaten = ?";
            params.push(kabupaten);
        }

        // Filter Kecamatan
        if (kecamatan) {
            query += " AND kecamatan = ?";
            params.push(kecamatan);
        }

        query += `
            GROUP BY kabupaten, kecamatan, bulan
            ORDER BY FIELD(bulan, 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec')
        `;

        console.time("Query Execution Time");
        const [rows] = await db.execute(query, params);
        console.timeEnd("Query Execution Time");

        // Mapping Provinsi DIY atau Jawa Tengah
        const diyKabupaten = ["Kota Yogyakarta", "Bantul", "Sleman", "Gunung Kidul", "Kulon Progo"];

        const filteredRows = rows.filter(row => {
            const isDIY = diyKabupaten.some(diy => row.kabupaten.includes(diy));
            if (provinsi === 'DIY') return isDIY;
            if (provinsi === 'JATENG') return !isDIY;
            return true; // Jika tidak ada filter provinsi, tampilkan semua
        });

        console.log("Data summary yang dikirim ke frontend:", filteredRows);

        res.json({ data: filteredRows });
    } catch (error) {
        console.error("Database Error:", error);
        res.status(500).json({ message: "Error mengambil data", error });
    }
};


exports.getErdkk = async (req, res) => {
    try {
        const { start, length, draw, kabupaten, kecamatan, tahun } = req.query;

        let query = `
                SELECT tahun, kabupaten, kecamatan, desa, kode_kios, nama_kios, nik, nama_petani, 
                SUM(urea) AS urea, 
                SUM(npk) AS npk, 
                SUM(npk_formula) AS npk_formula, 
                SUM(organik) AS organik
            FROM erdkk WHERE 1=1
        `;

        let countQuery = "SELECT COUNT(DISTINCT nik) AS total FROM erdkk WHERE 1=1";
        let params = [];
        let countParams = [];

        // Filter Kabupaten
        if (kabupaten) {
            query += " AND kabupaten = ?";
            countQuery += " AND kabupaten = ?";
            params.push(kabupaten);
            countParams.push(kabupaten);
        }

        // Filter Kecamatan
        if (kecamatan) {
            query += " AND kecamatan = ?";
            countQuery += " AND kecamatan = ?";
            params.push(kecamatan);
            countParams.push(kecamatan);
        }

        if (tahun) {
            query += " AND tahun = ?";
            countQuery += " AND tahun = ?";
            params.push(tahun);
            countParams.push(tahun);
        }

        // Grouping by kode_kios & bulan
        query += `GROUP BY tahun, kabupaten, kecamatan, desa, kode_kios, nama_kios, nik, nama_petani 
        ORDER BY tahun DESC, kabupaten, kecamatan, desa`;

        // Pagination
        let startNum = parseInt(start) || 0;
        let lengthNum = parseInt(length) || 10;
        query += " LIMIT ?, ?";
        params.push(startNum, lengthNum);

        // Eksekusi Query
        const [data] = await db.query(query, params);
        const [[{ total: recordsTotal }]] = await db.query("SELECT COUNT(DISTINCT nik) AS total FROM erdkk");
        const [[{ total: recordsFiltered }]] = await db.query(countQuery, countParams);

        res.json({
            draw: draw ? parseInt(draw) : 1,
            recordsTotal: recordsTotal,
            recordsFiltered: recordsFiltered,
            data: data
        });
    } catch (error) {
        console.error(error);
        res.status(500).json({ error: 'Terjadi kesalahan dalam mengambil data' });
    }
};


// untuk dashboard 1

exports.downloadErdkk = async (req, res) => {
    try {
        const { kabupaten, kecamatan, tahun } = req.query;

        let query = `
            SELECT tahun, kabupaten, kecamatan, desa, kode_kios, nama_kios, nik, nama_petani, 
                SUM(urea) AS urea, 
                SUM(npk) AS npk, 
                SUM(npk_formula) AS npk_formula, 
                SUM(organik) AS organik
            FROM erdkk WHERE 1=1
        `;

        let params = [];

        // Filter Kabupaten
        if (kabupaten) {
            query += " AND kabupaten = ?";
            params.push(kabupaten);
        }

        // Filter Kecamatan
        if (kecamatan) {
            query += " AND kecamatan = ?";
            params.push(kecamatan);
        }

        // Filter Tahun
        if (tahun) {
            query += " AND tahun = ?";
            params.push(tahun);
        }

        // Grouping Data
        query += ` GROUP BY tahun, kabupaten, kecamatan, desa, kode_kios, nama_kios, nik, nama_petani
                    ORDER BY tahun DESC, kabupaten, kecamatan, desa`;

        // Eksekusi Query
        const [data] = await db.query(query, params);

        // Membuat workbook baru
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Data ERDKK');

        // Menambahkan Header
        worksheet.columns = [
            { header: 'Tahun', key: 'tahun', width: 10 },
            { header: 'Kabupaten', key: 'kabupaten', width: 20 },
            { header: 'Kecamatan', key: 'kecamatan', width: 20 },
            { header: 'Desa', key: 'desa', width: 20 },
            { header: 'Kode Kios', key: 'kode_kios', width: 15 },
            { header: 'Nama Kios', key: 'nama_kios', width: 20 },
            { header: 'NIK', key: 'nik', width: 20 },
            { header: 'Nama Petani', key: 'nama_petani', width: 25 },
            { header: 'Urea', key: 'urea', width: 10 },
            { header: 'NPK', key: 'npk', width: 10 },
            { header: 'NPK Formula', key: 'npk_formula', width: 15 },
            { header: 'Organik', key: 'organik', width: 10 }
        ];

        // Menambahkan data ke worksheet
        data.forEach(row => {
            worksheet.addRow(row);
        });

        // Mengatur format header agar lebih jelas
        worksheet.getRow(1).eachCell(cell => {
            cell.font = { bold: true };
            cell.alignment = { horizontal: 'center' };
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFFFCC00' } // Warna kuning untuk header
            };
        });

        // Mengirim file Excel ke response
        res.setHeader(
            'Content-Type',
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        );
        res.setHeader(
            'Content-Disposition',
            'attachment; filename=erdkk.xlsx'
        );

        await workbook.xlsx.write(res);
        res.end();
    } catch (error) {
        console.error(error);
        res.status(500).json({ error: 'Terjadi kesalahan dalam mengunduh file' });
    }
};

// batas
exports.getData = async (req, res) => {
    try {
        const { start, length, draw, kabupaten, kecamatan, metode_penebusan, bulan, tahun, bulan_awal, bulan_akhir } = req.query;

        let query = "SELECT *, DATE_FORMAT(tanggal_tebus, '%Y-%m') AS tanggal_tebus FROM verval WHERE 1=1";
        let countQuery = "SELECT COUNT(*) AS total FROM verval WHERE 1=1";
        let params = [];
        let countParams = [];

        if (kabupaten) {
            query += " AND kabupaten = ?";
            countQuery += " AND kabupaten = ?";
            params.push(kabupaten);
            countParams.push(kabupaten);
        }
        if (kecamatan) {
            query += " AND kecamatan = ?";
            countQuery += " AND kecamatan = ?";
            params.push(kecamatan);
            countParams.push(kecamatan);
        }
        if (metode_penebusan) {
            query += " AND metode_penebusan = ?";
            countQuery += " AND metode_penebusan = ?";
            params.push(metode_penebusan);
            countParams.push(metode_penebusan);
        }
        if (bulan_awal && bulan_akhir) {
            // Filter pakai rentang bulan (format: 'YYYY-MM')
            query += " AND DATE_FORMAT(tanggal_tebus, '%Y-%m') BETWEEN ? AND ?";
            params.push(bulan_awal, bulan_akhir);

        } else if (bulan && tahun) {
            // Filter berdasarkan bulan dan tahun tunggal
            query += " AND MONTH(tanggal_tebus) = ? AND YEAR(tanggal_tebus) = ?";
            params.push(bulan, tahun);

        } else if (bulan) {
            query += " AND MONTH(tanggal_tebus) = ?";
            params.push(bulan);

        } else if (tahun) {
            query += " AND YEAR(tanggal_tebus) = ?";
            params.push(tahun);
        }

        // Pagination (gunakan 0 jika null)
        let startNum = parseInt(start) || 0;
        let lengthNum = parseInt(length) || 10;
        query += " LIMIT ?, ?";
        params.push(startNum, lengthNum);

        // Eksekusi query
        const [data] = await db.query(query, params);
        const [[{ total: recordsTotal }]] = await db.query("SELECT COUNT(*) AS total FROM verval");
        const [[{ total: recordsFiltered }]] = await db.query(countQuery, countParams);

        res.json({
            draw: draw ? parseInt(draw) : 1,
            recordsTotal: recordsTotal,
            recordsFiltered: recordsFiltered,
            data: data
        });
    } catch (error) {
        console.error(error);
        res.status(500).json({ error: 'Terjadi kesalahan dalam mengambil data' });
    }
};


exports.getSalurKiosSum = async (req, res) => {
    try {
        const { kabupaten, kecamatan, tahun } = req.query;

        let query = `
            SELECT 
    SUM(urea) AS urea,
    SUM(npk) AS npk,
    SUM(npk_formula) AS npk_formula,
    SUM(organik) AS organik,

    -- Tebusan per bulan
    SUM(CASE WHEN MONTH(tanggal_tebus) = 1 THEN urea END) AS jan_tebus_urea,
    SUM(CASE WHEN MONTH(tanggal_tebus) = 1 THEN npk END) AS jan_tebus_npk,
    SUM(CASE WHEN MONTH(tanggal_tebus) = 1 THEN npk_formula END) AS jan_tebus_npk_formula,
    SUM(CASE WHEN MONTH(tanggal_tebus) = 1 THEN organik END) AS jan_tebus_organik,

    SUM(CASE WHEN MONTH(tanggal_tebus) = 2 THEN urea END) AS feb_tebus_urea,
    SUM(CASE WHEN MONTH(tanggal_tebus) = 2 THEN npk END) AS feb_tebus_npk,
    SUM(CASE WHEN MONTH(tanggal_tebus) = 2 THEN npk_formula END) AS feb_tebus_npk_formula,
    SUM(CASE WHEN MONTH(tanggal_tebus) = 2 THEN organik END) AS feb_tebus_organik,

    SUM(CASE WHEN MONTH(tanggal_tebus) = 3 THEN urea END) AS mar_tebus_urea,
    SUM(CASE WHEN MONTH(tanggal_tebus) = 3 THEN npk END) AS mar_tebus_npk,
    SUM(CASE WHEN MONTH(tanggal_tebus) = 3 THEN npk_formula END) AS mar_tebus_npk_formula,
    SUM(CASE WHEN MONTH(tanggal_tebus) = 3 THEN organik END) AS mar_tebus_organik,

    SUM(CASE WHEN MONTH(tanggal_tebus) = 4 THEN urea END) AS apr_tebus_urea,
    SUM(CASE WHEN MONTH(tanggal_tebus) = 4 THEN npk END) AS apr_tebus_npk,
    SUM(CASE WHEN MONTH(tanggal_tebus) = 4 THEN npk_formula END) AS apr_tebus_npk_formula,
    SUM(CASE WHEN MONTH(tanggal_tebus) = 4 THEN organik END) AS apr_tebus_organik,

    SUM(CASE WHEN MONTH(tanggal_tebus) = 5 THEN urea END) AS mei_tebus_urea,
    SUM(CASE WHEN MONTH(tanggal_tebus) = 5 THEN npk END) AS mei_tebus_npk,
    SUM(CASE WHEN MONTH(tanggal_tebus) = 5 THEN npk_formula END) AS mei_tebus_npk_formula,
    SUM(CASE WHEN MONTH(tanggal_tebus) = 5 THEN organik END) AS mei_tebus_organik,

    SUM(CASE WHEN MONTH(tanggal_tebus) = 6 THEN urea END) AS jun_tebus_urea,
    SUM(CASE WHEN MONTH(tanggal_tebus) = 6 THEN npk END) AS jun_tebus_npk,
    SUM(CASE WHEN MONTH(tanggal_tebus) = 6 THEN npk_formula END) AS jun_tebus_npk_formula,
    SUM(CASE WHEN MONTH(tanggal_tebus) = 6 THEN organik END) AS jun_tebus_organik,

    SUM(CASE WHEN MONTH(tanggal_tebus) = 7 THEN urea END) AS jul_tebus_urea,
    SUM(CASE WHEN MONTH(tanggal_tebus) = 7 THEN npk END) AS jul_tebus_npk,
    SUM(CASE WHEN MONTH(tanggal_tebus) = 7 THEN npk_formula END) AS jul_tebus_npk_formula,
    SUM(CASE WHEN MONTH(tanggal_tebus) = 7 THEN organik END) AS jul_tebus_organik,

    SUM(CASE WHEN MONTH(tanggal_tebus) = 8 THEN urea END) AS agu_tebus_urea,
    SUM(CASE WHEN MONTH(tanggal_tebus) = 8 THEN npk END) AS agu_tebus_npk,
    SUM(CASE WHEN MONTH(tanggal_tebus) = 8 THEN npk_formula END) AS agu_tebus_npk_formula,
    SUM(CASE WHEN MONTH(tanggal_tebus) = 8 THEN organik END) AS agu_tebus_organik,

    SUM(CASE WHEN MONTH(tanggal_tebus) = 9 THEN urea END) AS sep_tebus_urea,
    SUM(CASE WHEN MONTH(tanggal_tebus) = 9 THEN npk END) AS sep_tebus_npk,
    SUM(CASE WHEN MONTH(tanggal_tebus) = 9 THEN npk_formula END) AS sep_tebus_npk_formula,
    SUM(CASE WHEN MONTH(tanggal_tebus) = 9 THEN organik END) AS sep_tebus_organik,

    SUM(CASE WHEN MONTH(tanggal_tebus) = 10 THEN urea END) AS okt_tebus_urea,
    SUM(CASE WHEN MONTH(tanggal_tebus) = 10 THEN npk END) AS okt_tebus_npk,
    SUM(CASE WHEN MONTH(tanggal_tebus) = 10 THEN npk_formula END) AS okt_tebus_npk_formula,
    SUM(CASE WHEN MONTH(tanggal_tebus) = 10 THEN organik END) AS okt_tebus_organik,

    SUM(CASE WHEN MONTH(tanggal_tebus) = 11 THEN urea END) AS nov_tebus_urea,
    SUM(CASE WHEN MONTH(tanggal_tebus) = 11 THEN npk END) AS nov_tebus_npk,
    SUM(CASE WHEN MONTH(tanggal_tebus) = 11 THEN npk_formula END) AS nov_tebus_npk_formula,
    SUM(CASE WHEN MONTH(tanggal_tebus) = 11 THEN organik END) AS nov_tebus_organik,

    SUM(CASE WHEN MONTH(tanggal_tebus) = 12 THEN urea END) AS des_tebus_urea,
    SUM(CASE WHEN MONTH(tanggal_tebus) = 12 THEN npk END) AS des_tebus_npk,
    SUM(CASE WHEN MONTH(tanggal_tebus) = 12 THEN npk_formula END) AS des_tebus_npk_formula,
    SUM(CASE WHEN MONTH(tanggal_tebus) = 12 THEN organik END) AS des_tebus_organik


            FROM verval WHERE 1=1
        `;

        let params = [];
        if (kabupaten) {
            query += " AND kabupaten = ?";
            params.push(kabupaten);
        }
        if (kecamatan) {
            query += " AND kecamatan = ?";
            params.push(kecamatan);
        }
        if (tahun) {
            query += " AND YEAR(tanggal_tebus) = ?";
            params.push(tahun);
        }

        const [rows] = await db.execute(query, params);
        res.json(rows[0]); // Kirim hasil sum
    } catch (error) {
        console.error("Error fetching sum:", error);
        res.status(500).json({ error: "Gagal mengambil data sum" });
    }
};

exports.getSum = async (req, res) => {
    const { kabupaten, kecamatan, metode_penebusan, bulan, tahun, bulan_awal, bulan_akhir } = req.query;

    let query = `
        SELECT 
            SUM(urea) as urea, 
            SUM(npk) as npk, 
            SUM(sp36) as sp36, 
            SUM(za) as za,
            SUM(npk_formula) as npk_formula, 
            SUM(organik) as organik,
            SUM(organik_cair) as organik_cair, 
            SUM(kakao) as kakao
        FROM verval WHERE 1=1`;

    let params = [];

    if (kabupaten) {
        query += " AND kabupaten = ?";
        params.push(kabupaten);
    }
    if (kecamatan) {
        query += " AND kecamatan = ?";
        params.push(kecamatan);
    }
    if (metode_penebusan) {
        query += " AND metode_penebusan = ?";
        params.push(metode_penebusan);
    }
    if (bulan_awal && bulan_akhir) {
        // Filter pakai rentang bulan (format: 'YYYY-MM')
        query += " AND DATE_FORMAT(tanggal_tebus, '%Y-%m') BETWEEN ? AND ?";
        params.push(bulan_awal, bulan_akhir);

    } else if (bulan && tahun) {
        // Filter berdasarkan bulan dan tahun tunggal
        query += " AND MONTH(tanggal_tebus) = ? AND YEAR(tanggal_tebus) = ?";
        params.push(bulan, tahun);

    } else if (bulan) {
        query += " AND MONTH(tanggal_tebus) = ?";
        params.push(bulan);

    } else if (tahun) {
        query += " AND YEAR(tanggal_tebus) = ?";
        params.push(tahun);
    }

    try {
        const [[sumData]] = await db.query(query, params);
        res.json(sumData);
    } catch (error) {
        console.error("Error fetching sum data:", error);
        res.status(500).json({ error: "Internal Server Error" });
    }
};

exports.exportExcel = async (req, res) => {
    try {
        const { kabupaten, kecamatan, metode_penebusan, tahun, bulan, bulan_awal, bulan_akhir } = req.query;

        // 1. Gunakan cursor/pagination untuk data besar
        const pageSize = 50000; // Sesuaikan dengan kapasitas server
        let currentPage = 0;
        let allRows = [];

        // 2. Query dengan limit dan offset
        const fetchData = async (page) => {
            let query = `
                SELECT 
                    kabupaten, kecamatan, poktan, kode_kios, nik, nama_petani,
                    tanggal_tebus, urea, npk, sp36, za, npk_formula,
                    organik, organik_cair, kakao, status
                FROM verval WHERE 1=1
            `;

            let params = [];

            // Filter conditions
            if (kabupaten) {
                query += " AND kabupaten = ?";
                params.push(kabupaten);
            }
            if (kecamatan) {
                query += " AND kecamatan = ?";
                params.push(kecamatan);
            }
            if (metode_penebusan) {
                query += " AND metode_penebusan = ?";
                params.push(metode_penebusan);
            }
            if (tahun) {
                query += " AND YEAR(tanggal_tebus) = ?";
                params.push(tahun);
            }
            if (bulan) {
                query += " AND MONTH(tanggal_tebus) = ?";
                params.push(bulan);
            }
            if (bulan_awal && bulan_akhir) {
                query += " AND MONTH(tanggal_tebus) BETWEEN ? AND ?";
                params.push(bulan_awal, bulan_akhir);
            }

            // Pagination
            query += ` LIMIT ? OFFSET ?`;
            params.push(pageSize, page * pageSize);

            const [rows] = await db.query(query, params);
            return rows;
        };

        // 3. Set headers sebelum mulai proses
        res.setHeader('Content-Disposition', 'attachment; filename="verval_data.xlsx"');
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');

        // 4. Buat workbook dengan streaming
        const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({
            stream: res,
            useStyles: true,
            useSharedStrings: true
        });

        const worksheet = workbook.addWorksheet('Data Verval');

        // 5. Definisikan kolom
        worksheet.columns = [
            { header: "No", key: "no", width: 5 },
            { header: "Kabupaten", key: "kabupaten", width: 20 },
            { header: "Kecamatan", key: "kecamatan", width: 20 },
            { header: "Poktan", key: "poktan", width: 15 },
            { header: "Kode Kios", key: "kode_kios", width: 15 },
            { header: "NIK", key: "nik", width: 20 },
            { header: "Nama Petani", key: "nama_petani", width: 25 },
            { header: "Tanggal Tebus", key: "tanggal_tebus", width: 15, style: { numFmt: 'yyyy-mm-dd' } },
            { header: "Urea", key: "urea", width: 10, style: { numFmt: '#,##0' } },
            { header: "NPK", key: "npk", width: 10, style: { numFmt: '#,##0' } },
            { header: "SP36", key: "sp36", width: 10, style: { numFmt: '#,##0' } },
            { header: "ZA", key: "za", width: 10, style: { numFmt: '#,##0' } },
            { header: "NPK Formula", key: "npk_formula", width: 15, style: { numFmt: '#,##0' } },
            { header: "Organik", key: "organik", width: 10, style: { numFmt: '#,##0' } },
            { header: "Organik Cair", key: "organik_cair", width: 15, style: { numFmt: '#,##0' } },
            { header: "Kakao", key: "kakao", width: 10, style: { numFmt: '#,##0' } },
            { header: "Status", key: "status", width: 15 }
        ];

        // 6. Variabel untuk total
        let totalUrea = 0, totalNpk = 0, totalSp36 = 0, totalZa = 0,
            totalNpkFormula = 0, totalOrganik = 0, totalOrganikCair = 0, totalKakao = 0;
        let rowCount = 0;

        // 7. Proses data per halaman
        do {
            const rows = await fetchData(currentPage);
            if (rows.length === 0 && currentPage === 0) {
                return res.status(404).json({ error: "Data tidak ditemukan" });
            }

            for (const row of rows) {
                rowCount++;

                worksheet.addRow({
                    no: rowCount,
                    kabupaten: row.kabupaten,
                    kecamatan: row.kecamatan,
                    poktan: row.poktan,
                    kode_kios: row.kode_kios,
                    nik: row.nik,
                    nama_petani: row.nama_petani,
                    tanggal_tebus: row.tanggal_tebus,
                    urea: row.urea || 0,
                    npk: row.npk || 0,
                    sp36: row.sp36 || 0,
                    za: row.za || 0,
                    npk_formula: row.npk_formula || 0,
                    organik: row.organik || 0,
                    organik_cair: row.organik_cair || 0,
                    kakao: row.kakao || 0,
                    status: row.status
                }).commit(); // Langsung commit setiap baris

                // Hitung total
                totalUrea += row.urea || 0;
                totalNpk += row.npk || 0;
                totalSp36 += row.sp36 || 0;
                totalZa += row.za || 0;
                totalNpkFormula += row.npk_formula || 0;
                totalOrganik += row.organik || 0;
                totalOrganikCair += row.organik_cair || 0;
                totalKakao += row.kakao || 0;
            }

            currentPage++;

            // Beri jeda untuk menghindari blocking event loop
            if (rows.length > 0) {
                await new Promise(resolve => setTimeout(resolve, 100));
            }

        } while (allRows.length === pageSize); // Lanjutkan sampai data habis

        // 8. Tambahkan baris total
        const totalsRow = worksheet.addRow({
            no: "TOTAL",
            kabupaten: "",
            kecamatan: "",
            poktan: "",
            kode_kios: "",
            nik: "",
            nama_petani: "",
            tanggal_tebus: "",
            urea: totalUrea,
            npk: totalNpk,
            sp36: totalSp36,
            za: totalZa,
            npk_formula: totalNpkFormula,
            organik: totalOrganik,
            organik_cair: totalOrganikCair,
            kakao: totalKakao,
            status: ""
        });

        totalsRow.eachCell((cell, colNumber) => {
            if (colNumber > 7) {
                cell.font = { bold: true };
                cell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'FFFF00' }
                };
            }
        });
        totalsRow.commit();

        // 9. Finalisasi
        worksheet.commit();
        await workbook.commit();

    } catch (error) {
        console.error("Error exporting Excel:", error);
        if (!res.headersSent) {
            res.status(500).json({ error: "Internal Server Error" });
        }
    }
};

exports.getJumlahPetani = async (req, res) => {
    try {
        const { kabupaten, kecamatan, metode_penebusan, tanggal_tebus, bulan_awal, bulan_akhir } = req.query;

        let query = `SELECT COUNT(DISTINCT nik) AS jumlah_petani FROM verval WHERE 1=1`;
        let params = [];

        if (kabupaten) {
            query += ` AND kabupaten = ?`;
            params.push(kabupaten);
        }
        if (kecamatan) {
            query += ` AND kecamatan = ?`;
            params.push(kecamatan);
        }
        if (metode_penebusan) {
            query += ` AND metode_penebusan = ?`;
            params.push(metode_penebusan);
        }
        if (tanggal_tebus) {
            query += " AND DATE_FORMAT(tanggal_tebus, '%Y-%m') = ?";
            params.push(tanggal_tebus);
        }
        if (bulan_awal && bulan_akhir) {
            query += ` AND MONTH(tanggal_tebus) BETWEEN ? AND ?`;
            params.push(bulan_awal, bulan_akhir);
        }

        const [result] = await db.query(query, params);
        res.json({ jumlah_petani: result[0].jumlah_petani || 0 });
    } catch (error) {
        console.error("Error fetching jumlah petani:", error);
        res.status(500).json({ message: "Terjadi kesalahan server" });
    }
};


// Daftar kabupaten yang diizinkan
// const allowedKabupaten = [
//     "BOYOLALI", "SRAGEN", "KOTA SURAKARTA", "KOTA MAGELANG", 
//     "MAGELANG", "SUKOHARJO", "KLATEN", "WONOGIRI", "KARANGANYAR", 
//     "BANTUL", "KOTA YOGYAKARTA", "SLEMAN", "KULON PROGO", "GUNUNG KIDUL"
// ];

exports.alokasiVsTebusan = async (req, res) => {
    try {
        const { start, length, draw, kabupaten, tahun, bulan_awal, bulan_akhir } = req.query;

        const bulanIndex = {
            jan: 1, feb: 2, mar: 3, apr: 4, mei: 5, jun: 6,
            jul: 7, agu: 8, sep: 9, okt: 10, nov: 11, des: 12
        };
        const urutanBulan = Object.keys(bulanIndex);
        const batasAwal = bulan_awal ? parseInt(bulan_awal) : 1;
        const batasAkhir = bulan_akhir ? parseInt(bulan_akhir) : 12;

        // Query utama untuk mengambil data
        let query = `
            SELECT
                e.kabupaten, 
                e.kecamatan,
                e.nik, 
                e.nama_petani, 
                e.kode_kios,
                e.nama_kios,
                e.tahun,
                e.desa, 
                e.poktan,
                e.urea, 
                e.npk, 
                e.npk_formula, 
                e.organik,
                COALESCE(v.tebus_urea, 0) AS tebus_urea,
                COALESCE(v.tebus_npk, 0) AS tebus_npk,
                COALESCE(v.tebus_npk_formula, 0) AS tebus_npk_formula,
                COALESCE(v.tebus_organik, 0) AS tebus_organik,
                (e.urea - COALESCE(v.tebus_urea, 0)) AS sisa_urea,
                (e.npk - COALESCE(v.tebus_npk, 0)) AS sisa_npk,
                (e.npk_formula - COALESCE(v.tebus_npk_formula, 0)) AS sisa_npk_formula,
                (e.organik - COALESCE(v.tebus_organik, 0)) AS sisa_organik,

                -- Data tebusan per bulan
                ${urutanBulan.map(bulan => `
                    COALESCE(t.${bulan}_urea, 0) AS ${bulan}_urea,
                    COALESCE(t.${bulan}_npk, 0) AS ${bulan}_npk,
                    COALESCE(t.${bulan}_npk_formula, 0) AS ${bulan}_npk_formula,
                    COALESCE(t.${bulan}_organik, 0) AS ${bulan}_organik
                `).join(',')}
            FROM erdkk e
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

        let countQuery = `
            SELECT COUNT(*) AS total
            FROM erdkk e
            LEFT JOIN verval_summary v 
                ON e.nik = v.nik
                AND e.kabupaten = v.kabupaten
                AND e.tahun = v.tahun
                AND e.kecamatan = v.kecamatan
                AND e.kode_kios = v.kode_kios
            WHERE 1=1
        `;

        let params = [];
        let countParams = [];

        if (kabupaten) {
            query += " AND e.kabupaten = ?";
            countQuery += " AND e.kabupaten = ?";
            params.push(kabupaten);
            countParams.push(kabupaten);
        }

        if (tahun) {
            query += " AND e.tahun = ?";
            countQuery += " AND e.tahun = ?";
            params.push(tahun);
            countParams.push(tahun);
        }

        // Pagination
        const startNum = parseInt(start) || 0;
        const lengthNum = parseInt(length) || 10;
        query += " LIMIT ?, ?";
        params.push(startNum, lengthNum);

        // Eksekusi query
        const [data] = await db.query(query, params);
        const [[{ total: recordsFiltered }]] = await db.query(countQuery, countParams);
        const [[{ total: recordsTotal }]] = await db.query("SELECT COUNT(*) AS total FROM erdkk");

        // Hapus kolom bulan di luar rentang
        const filteredData = data.map(row => {
            const newRow = { ...row };
            urutanBulan.forEach(bulan => {
                const index = bulanIndex[bulan];
                if (index < batasAwal || index > batasAkhir) {
                    delete newRow[`${bulan}_urea`];
                    delete newRow[`${bulan}_npk`];
                    delete newRow[`${bulan}_npk_formula`];
                    delete newRow[`${bulan}_organik`];
                }
            });
            return newRow;
        });

        // Kirim respons
        res.json({
            draw: draw ? parseInt(draw) : 1,
            recordsTotal: recordsTotal,
            recordsFiltered: recordsFiltered,
            data: filteredData
        });
    } catch (error) {
        console.error("Error fetching data:", error);
        res.status(500).json({ error: "Terjadi kesalahan dalam mengambil data" });
    }
};


exports.getSalurKios = async (req, res) => {
    try {
        const { start, length, draw, kabupaten, kecamatan, tahun } = req.query;

        // Query utama untuk mengambil data
        let query = `
            SELECT 
    kabupaten, 
    kecamatan, 
    kode_kios,
    MIN(nama_kios) AS nama_kios,
    SUM(urea) AS urea,
    SUM(npk) AS npk,
    SUM(npk_formula) AS npk_formula,
    SUM(organik) AS organik,
SUM(CASE WHEN MONTH(tanggal_tebus) = 1 THEN urea ELSE 0 END) AS jan_urea,
SUM(CASE WHEN MONTH(tanggal_tebus) = 1 THEN npk ELSE 0 END) AS jan_npk,
SUM(CASE WHEN MONTH(tanggal_tebus) = 1 THEN npk_formula ELSE 0 END) AS jan_npk_formula,
SUM(CASE WHEN MONTH(tanggal_tebus) = 1 THEN organik ELSE 0 END) AS jan_organik,
SUM(CASE WHEN MONTH(tanggal_tebus) = 2 THEN urea ELSE 0 END) AS feb_urea,
SUM(CASE WHEN MONTH(tanggal_tebus) = 2 THEN npk ELSE 0 END) AS feb_npk,
SUM(CASE WHEN MONTH(tanggal_tebus) = 2 THEN npk_formula ELSE 0 END) AS feb_npk_formula,
SUM(CASE WHEN MONTH(tanggal_tebus) = 2 THEN organik ELSE 0 END) AS feb_organik,
SUM(CASE WHEN MONTH(tanggal_tebus) = 3 THEN urea ELSE 0 END) AS mar_urea,
SUM(CASE WHEN MONTH(tanggal_tebus) = 3 THEN npk ELSE 0 END) AS mar_npk,
SUM(CASE WHEN MONTH(tanggal_tebus) = 3 THEN npk_formula ELSE 0 END) AS mar_npk_formula,
SUM(CASE WHEN MONTH(tanggal_tebus) = 3 THEN organik ELSE 0 END) AS mar_organik,
SUM(CASE WHEN MONTH(tanggal_tebus) = 4 THEN urea ELSE 0 END) AS apr_urea,
SUM(CASE WHEN MONTH(tanggal_tebus) = 4 THEN npk ELSE 0 END) AS apr_npk,
SUM(CASE WHEN MONTH(tanggal_tebus) = 4 THEN npk_formula ELSE 0 END) AS apr_npk_formula,
SUM(CASE WHEN MONTH(tanggal_tebus) = 4 THEN organik ELSE 0 END) AS apr_organik,
SUM(CASE WHEN MONTH(tanggal_tebus) = 5 THEN urea ELSE 0 END) AS mei_urea,
SUM(CASE WHEN MONTH(tanggal_tebus) = 5 THEN npk ELSE 0 END) AS mei_npk,
SUM(CASE WHEN MONTH(tanggal_tebus) = 5 THEN npk_formula ELSE 0 END) AS mei_npk_formula,
SUM(CASE WHEN MONTH(tanggal_tebus) = 5 THEN organik ELSE 0 END) AS mei_organik,
SUM(CASE WHEN MONTH(tanggal_tebus) = 6 THEN urea ELSE 0 END) AS jun_urea,
SUM(CASE WHEN MONTH(tanggal_tebus) = 6 THEN npk ELSE 0 END) AS jun_npk,
SUM(CASE WHEN MONTH(tanggal_tebus) = 6 THEN npk_formula ELSE 0 END) AS jun_npk_formula,
SUM(CASE WHEN MONTH(tanggal_tebus) = 6 THEN organik ELSE 0 END) AS jun_organik,
SUM(CASE WHEN MONTH(tanggal_tebus) = 7 THEN urea ELSE 0 END) AS jul_urea,
SUM(CASE WHEN MONTH(tanggal_tebus) = 7 THEN npk ELSE 0 END) AS jul_npk,
SUM(CASE WHEN MONTH(tanggal_tebus) = 7 THEN npk_formula ELSE 0 END) AS jul_npk_formula,
SUM(CASE WHEN MONTH(tanggal_tebus) = 7 THEN organik ELSE 0 END) AS jul_organik,
SUM(CASE WHEN MONTH(tanggal_tebus) = 8 THEN urea ELSE 0 END) AS agu_urea,
SUM(CASE WHEN MONTH(tanggal_tebus) = 8 THEN npk ELSE 0 END) AS agu_npk,
SUM(CASE WHEN MONTH(tanggal_tebus) = 8 THEN npk_formula ELSE 0 END) AS agu_npk_formula,
SUM(CASE WHEN MONTH(tanggal_tebus) = 8 THEN organik ELSE 0 END) AS agu_organik,
SUM(CASE WHEN MONTH(tanggal_tebus) = 9 THEN urea ELSE 0 END) AS sep_urea,
SUM(CASE WHEN MONTH(tanggal_tebus) = 9 THEN npk ELSE 0 END) AS sep_npk,
SUM(CASE WHEN MONTH(tanggal_tebus) = 9 THEN npk_formula ELSE 0 END) AS sep_npk_formula,
SUM(CASE WHEN MONTH(tanggal_tebus) = 9 THEN organik ELSE 0 END) AS sep_organik,
SUM(CASE WHEN MONTH(tanggal_tebus) = 10 THEN urea ELSE 0 END) AS okt_urea,
SUM(CASE WHEN MONTH(tanggal_tebus) = 10 THEN npk ELSE 0 END) AS okt_npk,
SUM(CASE WHEN MONTH(tanggal_tebus) = 10 THEN npk_formula ELSE 0 END) AS okt_npk_formula,
SUM(CASE WHEN MONTH(tanggal_tebus) = 10 THEN organik ELSE 0 END) AS okt_organik,
SUM(CASE WHEN MONTH(tanggal_tebus) = 11 THEN urea ELSE 0 END) AS nov_urea,
SUM(CASE WHEN MONTH(tanggal_tebus) = 11 THEN npk ELSE 0 END) AS nov_npk,
SUM(CASE WHEN MONTH(tanggal_tebus) = 11 THEN npk_formula ELSE 0 END) AS nov_npk_formula,
SUM(CASE WHEN MONTH(tanggal_tebus) = 11 THEN organik ELSE 0 END) AS nov_organik,
SUM(CASE WHEN MONTH(tanggal_tebus) = 12 THEN urea ELSE 0 END) AS des_urea,
SUM(CASE WHEN MONTH(tanggal_tebus) = 12 THEN npk ELSE 0 END) AS des_npk,
SUM(CASE WHEN MONTH(tanggal_tebus) = 12 THEN npk_formula ELSE 0 END) AS des_npk_formula,
SUM(CASE WHEN MONTH(tanggal_tebus) = 12 THEN organik ELSE 0 END) AS des_organik
FROM verval
            WHERE 1=1
        `;

        let countQuery = `
            SELECT COUNT(*) AS total
            FROM verval
            WHERE 1=1
        `;

        let params = [];
        let countParams = [];

        // Tambahkan filter Kabupaten
        if (kabupaten) {
            query += " AND kabupaten = ?";
            countQuery += " AND kabupaten = ?";
            params.push(kabupaten);
            countParams.push(kabupaten);
        }

        // Tambahkan filter Kecamatan
        if (kecamatan) {
            query += " AND kecamatan = ?";
            countQuery += " AND kecamatan = ?";
            params.push(kecamatan);
            countParams.push(kecamatan);
        }

        // Tambahkan filter Tahun
        if (tahun) {
            query += " AND YEAR(tanggal_tebus) = ?";
            countQuery += " AND YEAR(tanggal_tebus) = ?";
            params.push(tahun);
            countParams.push(tahun);
        }

        query += " GROUP BY kabupaten, kecamatan, kode_kios";

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


exports.tebusanPerBulan = async (req, res) => {
    try {
        const { kabupaten, tahun } = req.query;

        let query = `SELECT * FROM tebusan_per_bulan WHERE 1=1`;
        let params = [];

        if (kabupaten) {
            query += " AND kabupaten = ?";
            params.push(kabupaten);
        }
        if (tahun) {
            query += " AND tahun = ?";
            params.push(tahun);
        }

        const [data] = await db.query(query, params);
        res.json({ data });
    } catch (error) {
        console.error("Error fetching tebusan per bulan:", error);
        res.status(500).json({ error: "Terjadi kesalahan dalam mengambil data" });
    }
};



exports.getVervalSummary = async (req, res) => {
    try {
        const { kabupaten, tahun, bulan_awal, bulan_akhir } = req.query;

        // Map index bulan ke nama kolom di tabel
        const bulanMap = [
            'jan', 'feb', 'mar', 'apr', 'mei', 'jun',
            'jul', 'agu', 'sep', 'okt', 'nov', 'des'
        ];

        const start = parseInt(bulan_awal || '1');
        const end = parseInt(bulan_akhir || '12');

        // Validasi range bulan
        const bulanRange = bulanMap.slice(start - 1, end);

        // Generator kolom SUM untuk setiap bulan dalam rentang
        const bulanSumFields = bulanRange.map(prefix => `
            SUM(COALESCE(t.${prefix}_urea, 0)) AS ${prefix}_tebus_urea,
            SUM(COALESCE(t.${prefix}_npk, 0)) AS ${prefix}_tebus_npk,
            SUM(COALESCE(t.${prefix}_npk_formula, 0)) AS ${prefix}_tebus_npk_formula,
            SUM(COALESCE(t.${prefix}_organik, 0)) AS ${prefix}_tebus_organik
        `).join(',\n');

        let query = `
            SELECT 
                SUM(combined.urea) AS total_urea,
                SUM(combined.npk) AS total_npk,
                SUM(combined.npk_formula) AS total_npk_formula,
                SUM(combined.organik) AS total_organik,

                SUM(combined.urea - COALESCE(v.tebus_urea, 0)) AS sisa_urea,
                SUM(combined.npk - COALESCE(v.tebus_npk, 0)) AS sisa_npk,
                SUM(combined.npk_formula - COALESCE(v.tebus_npk_formula, 0)) AS sisa_npk_formula,
                SUM(combined.organik - COALESCE(v.tebus_organik, 0)) AS sisa_organik,

                SUM(COALESCE(v.tebus_urea, 0)) AS total_tebus_urea,
                SUM(COALESCE(v.tebus_npk, 0)) AS total_tebus_npk,
                SUM(COALESCE(v.tebus_npk_formula, 0)) AS total_tebus_npk_formula,
                SUM(COALESCE(v.tebus_organik, 0)) AS total_tebus_organik,

                ${bulanSumFields}
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
                    v.tahun, '' AS desa, v.poktan,
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
        if (kabupaten) {
            query += " AND combined.kabupaten = ?";
            params.push(kabupaten);
        }

        if (tahun) {
            query += " AND combined.tahun = ?";
            params.push(tahun);
        }

        const [summary] = await db.query(query, params);

        res.json({ sum: summary[0] || {} });
    } catch (error) {
        console.error("Error fetching summary:", error);
        res.status(500).json({ error: "Terjadi kesalahan dalam mengambil summary" });
    }
};


exports.downloadPetaniSummary = async (req, res) => {
    const { kabupaten = 'ALL', tahun = 'ALL' } = req.query;
    const safeKabupaten = kabupaten.replace(/\s+/g, '_').toUpperCase();
    const safeTahun = tahun.toString();
    const fileName = `data_summary_${safeKabupaten}_${safeTahun}.xlsx`;

    const baseDir = path.join(__dirname, '../temp_exports');
    const filePath = path.join(baseDir, fileName);

    try {
        if (!fs.existsSync(baseDir)) fs.mkdirSync(baseDir, { recursive: true });

        if (!fs.existsSync(filePath)) {
            await generateReport(kabupaten, tahun, filePath);
        }

        res.download(filePath, fileName, err => {
            if (err) {
                console.error('Error saat download:', err);
                if (!res.headersSent) res.status(500).send('Gagal download file');
            }
        });
    } catch (error) {
        console.error('Error di downloadPetaniSummary:', error);
        res.status(500).send('Terjadi kesalahan saat proses export');
    }
};


exports.downloadSalurKios = async (req, res) => {
    try {
        const { kabupaten, tahun, kecamatan } = req.query;

        let query = `
            SELECT 
                kabupaten, 
                kecamatan, 
                kode_kios, 
                MONTH(tanggal_tebus) AS bulan, -- Ambil bulan dari tanggal_tebus
                SUM(urea) AS urea,
                SUM(npk) AS npk,
                SUM(npk_formula) AS npk_formula,
                SUM(organik) AS organik
            FROM verval
            WHERE 1=1`;
        let params = [];

        if (kabupaten) {
            query += " AND kabupaten = ?";
            params.push(kabupaten);
        }

        if (kecamatan) {
            query += " AND kecamatan = ?";
            params.push(kecamatan);
        }

        if (tahun) {
            query += " AND YEAR(tanggal_tebus) = ?";
            params.push(tahun);
        }

        query += " GROUP BY kabupaten, kecamatan, kode_kios, MONTH(tanggal_tebus)";

        const [data] = await db.query(query, params);

        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet("Salur Kios");

        const borderStyle = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };

        const bulanList = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];

        // Header Utama
        worksheet.mergeCells('A1:A2'); // Kabupaten
        worksheet.mergeCells('B1:B2'); // Kecamatan
        worksheet.mergeCells('C1:C2'); // Kode Kios
        worksheet.mergeCells('D1:G1'); // Penyaluran
        worksheet.getCell('A1').value = 'Kabupaten';
        worksheet.getCell('B1').value = 'Kecamatan';
        worksheet.getCell('C1').value = 'Kode Kios';
        worksheet.getCell('D1').value = 'Penyaluran';
        worksheet.getCell('D1').alignment = { horizontal: 'center', vertical: 'middle' };
        worksheet.getCell('D1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDD4B39' } };

        // Subheader Penyaluran
        worksheet.getCell('D2').value = 'Urea';
        worksheet.getCell('E2').value = 'NPK';
        worksheet.getCell('F2').value = 'NPK Formula';
        worksheet.getCell('G2').value = 'Organik';

        // Header Bulan
        let colStart = 8; // Mulai dari kolom H
        bulanList.forEach((bulan, index) => {
            let colEnd = colStart + 3; // Setiap bulan memiliki 4 kolom (urea, npk, npk_formula, organik)
            worksheet.mergeCells(1, colStart, 1, colEnd);
            worksheet.getCell(1, colStart).value = bulan;
            worksheet.getCell(1, colStart).alignment = { horizontal: 'center', vertical: 'middle' };
            worksheet.getCell(1, colStart).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDD4B39' } };

            // Subheader Bulan
            worksheet.getCell(2, colStart).value = 'Urea';
            worksheet.getCell(2, colStart + 1).value = 'NPK';
            worksheet.getCell(2, colStart + 2).value = 'NPK Formula';
            worksheet.getCell(2, colStart + 3).value = 'Organik';

            colStart += 4;
        });

        // Format Header
        worksheet.getRow(1).eachCell((cell) => {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDD4B39' } };
            cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
            cell.alignment = { horizontal: 'center', vertical: 'middle' };
            cell.border = borderStyle;
        });

        worksheet.getRow(2).eachCell((cell) => {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDD4B39' } };
            cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
            cell.alignment = { horizontal: 'center', vertical: 'middle' };
            cell.border = borderStyle;
        });

        // Format Data
        let formattedData = {};
        data.forEach(row => {
            let key = `${row.kabupaten}|${row.kecamatan}|${row.kode_kios}`;

            if (!formattedData[key]) {
                formattedData[key] = {
                    kabupaten: row.kabupaten,
                    kecamatan: row.kecamatan,
                    kode_kios: row.kode_kios,
                    penyaluran: { urea: 0, npk: 0, npk_formula: 0, organik: 0 },
                    bulan: Array(12).fill(0).map(() => ({ urea: 0, npk: 0, npk_formula: 0, organik: 0 }))
                };
            }

            let bulanIndex = row.bulan - 1; // Bulan dimulai dari 1 (Januari) hingga 12 (Desember)
            formattedData[key].bulan[bulanIndex] = {
                urea: row.urea || 0,
                npk: row.npk || 0,
                npk_formula: row.npk_formula || 0,
                organik: row.organik || 0
            };

            // Hitung total penyaluran per kios
            formattedData[key].penyaluran.urea += row.urea || 0;
            formattedData[key].penyaluran.npk += row.npk || 0;
            formattedData[key].penyaluran.npk_formula += row.npk_formula || 0;
            formattedData[key].penyaluran.organik += row.organik || 0;
        });

        // Hitung total untuk setiap kolom
        let totalPenyaluran = { urea: 0, npk: 0, npk_formula: 0, organik: 0 };
        let totalBulanan = Array(12).fill(0).map(() => ({ urea: 0, npk: 0, npk_formula: 0, organik: 0 }));

        Object.values(formattedData).forEach(row => {
            totalPenyaluran.urea += row.penyaluran.urea;
            totalPenyaluran.npk += row.penyaluran.npk;
            totalPenyaluran.npk_formula += row.penyaluran.npk_formula;
            totalPenyaluran.organik += row.penyaluran.organik;

            row.bulan.forEach((b, index) => {
                totalBulanan[index].urea += b.urea;
                totalBulanan[index].npk += b.npk;
                totalBulanan[index].npk_formula += b.npk_formula;
                totalBulanan[index].organik += b.organik;
            });
        });

        // Tambahkan baris total sebelum data
        const totalRow = worksheet.addRow([
            'TOTAL', '', '', // Kolom kosong untuk Kabupaten, Kecamatan, Kode Kios
            totalPenyaluran.urea,
            totalPenyaluran.npk,
            totalPenyaluran.npk_formula,
            totalPenyaluran.organik,
            ...totalBulanan.flatMap(b => [b.urea, b.npk, b.npk_formula, b.organik])
        ]);

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

        // Tambahkan data setelah baris total
        Object.values(formattedData).forEach(row => {
            let rowData = [
                row.kabupaten,
                row.kecamatan,
                row.kode_kios,
                row.penyaluran.urea,
                row.penyaluran.npk,
                row.penyaluran.npk_formula,
                row.penyaluran.organik
            ];

            // Tambahkan data bulanan
            row.bulan.forEach(b => {
                rowData.push(b.urea, b.npk, b.npk_formula, b.organik);
            });

            let addedRow = worksheet.addRow(rowData);
            addedRow.eachCell((cell) => {
                cell.border = borderStyle;
            });
        });

        // Kirim file Excel sebagai respons
        res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        res.setHeader("Content-Disposition", "attachment; filename=salur_kios.xlsx");
        await workbook.xlsx.write(res);
        res.end();
    } catch (error) {
        console.error(error);
        res.status(500).json({ error: 'Terjadi kesalahan dalam mengunduh data' });
    }
};

exports.downloadPetaniSum = async (req, res) => {
    try {
        const { kabupaten, tahun } = req.query;
        const reportsDir = path.join(__dirname, '../../public/reports');

        // Tentukan nama file berdasarkan parameter
        let fileName;
        if (!kabupaten && !tahun) {
            fileName = 'petani_summary_ALL.xlsx';
        } else {
            const safeKabupaten = kabupaten ? kabupaten.replace(/\s+/g, "_") : "ALL";
            const yearPart = tahun ? `_${tahun}` : '';
            fileName = `petani_summary_${safeKabupaten}${yearPart}.xlsx`;
        }

        const filePath = path.join(reportsDir, fileName);

        // Cek apakah file ada
        if (fs.existsSync(filePath)) {
            return res.download(filePath, fileName);
        } else {
            // Jika file tidak ada, beri tahu untuk coba lagi nanti
            return res.status(404).json({
                message: 'Laporan belum tersedia. Silakan coba lagi nanti atau hubungi admin.'
            });
        }

    } catch (error) {
        console.error("Error downloading report:", error);
        res.status(500).json({ error: "Gagal mendownload file laporan" });
    }
};



exports.downloadSalur = async (req, res) => {
    try {
        const { kabupaten, tahun, kecamatan } = req.query;

        let query = `
            SELECT 
                kabupaten, 
                kecamatan, 
                kode_kios, 
                metode_penebusan, 
                MONTH(tanggal_tebus) AS bulan,
                SUM(urea) AS urea,
                SUM(npk) AS npk,
                SUM(npk_formula) AS npk_formula,
                SUM(organik) AS organik
            FROM verval WHERE 1=1
            `;
        let params = [];

        if (kabupaten) {
            query += " AND kabupaten = ?";
            params.push(kabupaten);
        }

        if (kecamatan) {
            query += " AND kecamatan = ?";
            params.push(kecamatan);
        }

        if (tahun) {
            query += " AND YEAR(tanggal_tebus) = ?";
            params.push(tahun);
        }

        query += " GROUP BY kabupaten, kecamatan, kode_kios, metode_penebusan, MONTH(tanggal_tebus)";

        const [data] = await db.query(query, params);

        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet("Salur Kios");

        const borderStyle = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };

        const bulanList = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];

        // Header Utama
        worksheet.mergeCells('A1:A2'); // Kabupaten
        worksheet.mergeCells('B1:B2'); // Kecamatan
        worksheet.mergeCells('C1:C2'); // Kode Kios
        worksheet.mergeCells('D1:G1'); // Penyaluran
        worksheet.getCell('A1').value = 'Kabupaten';
        worksheet.getCell('B1').value = 'Kecamatan';
        worksheet.getCell('C1').value = 'Kode Kios';
        worksheet.getCell('D1').value = 'Penyaluran';
        worksheet.getCell('D1').alignment = { horizontal: 'center', vertical: 'middle' };
        worksheet.getCell('D1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDD4B39' } };

        // Subheader Penyaluran
        worksheet.getCell('D2').value = 'Urea';
        worksheet.getCell('E2').value = 'NPK';
        worksheet.getCell('F2').value = 'NPK Formula';
        worksheet.getCell('G2').value = 'Organik';

        // Header Bulan dengan Subheader Kartan dan Ipubers
        let colStart = 8; // Mulai dari kolom H
        bulanList.forEach((bulan, index) => {
            let colEnd = colStart + 7; // Setiap bulan memiliki 8 kolom (4 untuk Kartan dan 4 untuk Ipubers)
            worksheet.mergeCells(1, colStart, 1, colEnd);
            worksheet.getCell(1, colStart).value = bulan;
            worksheet.getCell(1, colStart).alignment = { horizontal: 'center', vertical: 'middle' };
            worksheet.getCell(1, colStart).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDD4B39' } };

            // Subheader Kartan dan Ipubers
            worksheet.mergeCells(2, colStart, 2, colStart + 3);
            worksheet.getCell(2, colStart).value = 'Kartan';
            worksheet.getCell(2, colStart).alignment = { horizontal: 'center', vertical: 'middle' };

            worksheet.mergeCells(2, colStart + 4, 2, colStart + 7);
            worksheet.getCell(2, colStart + 4).value = 'Ipubers';
            worksheet.getCell(2, colStart + 4).alignment = { horizontal: 'center', vertical: 'middle' };

            // Subheader untuk Kartan dan Ipubers
            worksheet.getCell(3, colStart).value = 'Urea';
            worksheet.getCell(3, colStart + 1).value = 'NPK';
            worksheet.getCell(3, colStart + 2).value = 'NPK Formula';
            worksheet.getCell(3, colStart + 3).value = 'Organik';

            worksheet.getCell(3, colStart + 4).value = 'Urea';
            worksheet.getCell(3, colStart + 5).value = 'NPK';
            worksheet.getCell(3, colStart + 6).value = 'NPK Formula';
            worksheet.getCell(3, colStart + 7).value = 'Organik';

            colStart += 8;
        });

        // Format Header
        worksheet.getRow(1).eachCell((cell) => {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDD4B39' } };
            cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
            cell.alignment = { horizontal: 'center', vertical: 'middle' };
            cell.border = borderStyle;
        });

        worksheet.getRow(2).eachCell((cell) => {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDD4B39' } };
            cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
            cell.alignment = { horizontal: 'center', vertical: 'middle' };
            cell.border = borderStyle;
        });

        worksheet.getRow(3).eachCell((cell) => {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDD4B39' } };
            cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
            cell.alignment = { horizontal: 'center', vertical: 'middle' };
            cell.border = borderStyle;
        });

        // Format Data
        let formattedData = {};
        data.forEach(row => {
            let key = `${row.kabupaten}|${row.kecamatan}|${row.kode_kios}`;

            if (!formattedData[key]) {
                formattedData[key] = {
                    kabupaten: row.kabupaten,
                    kecamatan: row.kecamatan,
                    kode_kios: row.kode_kios,
                    penyaluran: { urea: 0, npk: 0, npk_formula: 0, organik: 0 },
                    bulan: Array(12).fill(0).map(() => ({
                        kartan: { urea: 0, npk: 0, npk_formula: 0, organik: 0 },
                        ipubers: { urea: 0, npk: 0, npk_formula: 0, organik: 0 }
                    }))
                };
            }

            let bulanIndex = row.bulan - 1; // Bulan dimulai dari 1 (Januari) hingga 12 (Desember)
            if (row.metode_penebusan === 'kartan') {
                formattedData[key].bulan[bulanIndex].kartan = {
                    urea: row.urea || 0,
                    npk: row.npk || 0,
                    npk_formula: row.npk_formula || 0,
                    organik: row.organik || 0
                };
            } else if (row.metode_penebusan === 'ipubers') {
                formattedData[key].bulan[bulanIndex].ipubers = {
                    urea: row.urea || 0,
                    npk: row.npk || 0,
                    npk_formula: row.npk_formula || 0,
                    organik: row.organik || 0
                };
            }

            // Hitung total penyaluran per kios
            formattedData[key].penyaluran.urea += row.urea || 0;
            formattedData[key].penyaluran.npk += row.npk || 0;
            formattedData[key].penyaluran.npk_formula += row.npk_formula || 0;
            formattedData[key].penyaluran.organik += row.organik || 0;
        });

        // Hitung total untuk setiap kolom
        let totalPenyaluran = { urea: 0, npk: 0, npk_formula: 0, organik: 0 };
        let totalBulanan = Array(12).fill(0).map(() => ({
            kartan: { urea: 0, npk: 0, npk_formula: 0, organik: 0 },
            ipubers: { urea: 0, npk: 0, npk_formula: 0, organik: 0 }
        }));

        Object.values(formattedData).forEach(row => {
            totalPenyaluran.urea += row.penyaluran.urea;
            totalPenyaluran.npk += row.penyaluran.npk;
            totalPenyaluran.npk_formula += row.penyaluran.npk_formula;
            totalPenyaluran.organik += row.penyaluran.organik;

            row.bulan.forEach((b, index) => {
                totalBulanan[index].kartan.urea += b.kartan.urea;
                totalBulanan[index].kartan.npk += b.kartan.npk;
                totalBulanan[index].kartan.npk_formula += b.kartan.npk_formula;
                totalBulanan[index].kartan.organik += b.kartan.organik;

                totalBulanan[index].ipubers.urea += b.ipubers.urea;
                totalBulanan[index].ipubers.npk += b.ipubers.npk;
                totalBulanan[index].ipubers.npk_formula += b.ipubers.npk_formula;
                totalBulanan[index].ipubers.organik += b.ipubers.organik;
            });
        });

        // Tambahkan baris total sebelum data
        const totalRow = worksheet.addRow([
            'TOTAL', '', '', // Kolom kosong untuk Kabupaten, Kecamatan, Kode Kios
            totalPenyaluran.urea,
            totalPenyaluran.npk,
            totalPenyaluran.npk_formula,
            totalPenyaluran.organik,
            ...totalBulanan.flatMap(b => [
                b.kartan.urea, b.kartan.npk, b.kartan.npk_formula, b.kartan.organik,
                b.ipubers.urea, b.ipubers.npk, b.ipubers.npk_formula, b.ipubers.organik
            ])
        ]);

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

        // Tambahkan data setelah baris total
        Object.values(formattedData).forEach(row => {
            let rowData = [
                row.kabupaten,
                row.kecamatan,
                row.kode_kios,
                row.penyaluran.urea,
                row.penyaluran.npk,
                row.penyaluran.npk_formula,
                row.penyaluran.organik
            ];

            // Tambahkan data bulanan (kartan dan ipubers)
            row.bulan.forEach(b => {
                rowData.push(
                    b.kartan.urea, b.kartan.npk, b.kartan.npk_formula, b.kartan.organik,
                    b.ipubers.urea, b.ipubers.npk, b.ipubers.npk_formula, b.ipubers.organik
                );
            });

            let addedRow = worksheet.addRow(rowData);
            addedRow.eachCell((cell) => {
                cell.border = borderStyle;
            });
        });

        // Kirim file Excel sebagai respons
        res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        res.setHeader("Content-Disposition", "attachment; filename=salur_kios.xlsx");
        await workbook.xlsx.write(res);
        res.end();
    } catch (error) {
        console.error(error);
        res.status(500).json({ error: 'Terjadi kesalahan dalam mengunduh data' });
    }
};

// wcm

// Fungsi untuk generate cache key TANPA parameter draw
const generateCacheKey = (params) => {
  // Normalisasi parameter - abaikan draw karena selalu berubah
  const tahun = params.tahun || '';
  const bulan = params.bulan || 'ALL';
  const produk = params.produk || 'ALL';
  const kabupaten = params.kabupaten || 'ALL';
  const status = params.status || 'ALL';
  const start = params.start || '0';
  const length = params.length || '10';
  
  return `wcm_verval:${tahun}:${bulan}:${produk}:${kabupaten}:${status}:${start}:${length}`;
};

// Fungsi untuk mengecek info cache
const getCacheInfo = async (key) => {
  try {
    const ttl = await client.ttl(key);
    return { exists: ttl > -2, ttl };
  } catch (error) {
    return { exists: false, ttl: -2, error: error.message };
  }
};

exports.wcmVsVerval = async (req, res) => {
    const startTime = Date.now();
    let cacheStatus = 'MISS';
    let cacheKey = '';
    let queryTime = 0;
    let cacheCheckTime = 0;

    try {
        // Normalisasi parameter
        const tahun = req.query.tahun;
        const bulan = req.query.bulan || 'ALL';
        const produk = req.query.produk || 'ALL';
        const kabupaten = req.query.kabupaten || 'ALL';
        const status = req.query.status || 'ALL';
        const start = req.query.start || '0';
        const length = req.query.length || '10';
        const draw = req.query.draw || '1';

        // Validasi parameter wajib
        if (!tahun) {
            return res.status(400).json({ message: 'Parameter tahun wajib diisi' });
        }

        // Generate cache key TANPA draw
        cacheKey = generateCacheKey({ tahun, bulan, produk, kabupaten, status, start, length });
        
        console.log('══════════════════════════════════════════════════');
        console.log('🔍 REQUEST DITERIMA:', new Date().toLocaleString('id-ID'));
        console.log('📋 Parameter:', { tahun, bulan, produk, kabupaten, status, start, length, draw });
        console.log('🔑 Cache Key:', cacheKey);
        
        // Cek cache terlebih dahulu
        const cacheCheckStart = Date.now();
        try {
            const cachedData = await client.get(cacheKey);
            cacheCheckTime = Date.now() - cacheCheckStart;
            
            if (cachedData) {
                const cacheInfo = await getCacheInfo(cacheKey);
                const parsedData = JSON.parse(cachedData);
                
                cacheStatus = 'HIT';
                console.log('✅ ✅ ✅ DATA DIAMBIL DARI CACHE ✅ ✅ ✅');
                console.log('⏰ Cache TTL:', cacheInfo.ttl, 'detik tersisa');
                console.log('📊 Jumlah data:', parsedData.data.length, 'records');
                console.log('⚡ Waktu akses cache:', cacheCheckTime, 'ms');
                console.log('💾 Total waktu response:', Date.now() - startTime, 'ms');
                
                // Response dari cache dengan draw yang baru
                const response = {
                    draw: parseInt(draw),
                    recordsTotal: parsedData.recordsTotal,
                    recordsFiltered: parsedData.recordsFiltered,
                    data: parsedData.data,
                    cacheInfo: {
                        status: 'HIT',
                        key: cacheKey,
                        ttl: cacheInfo.ttl,
                        responseTime: Date.now() - startTime,
                        cacheTime: cacheCheckTime,
                        source: 'redis_cache'
                    }
                };
                
                res.set({
                    'X-Cache-Status': 'HIT',
                    'X-Cache-Key': cacheKey,
                    'X-Cache-TTL': cacheInfo.ttl,
                    'X-Response-Time': (Date.now() - startTime) + 'ms'
                });
                
                return res.json(response);
            } else {
                const cacheInfo = await getCacheInfo(cacheKey);
                console.log('❌ DATA TIDAK DITEMUKAN DI CACHE');
                console.log('ℹ️  Status cache:', cacheInfo.exists ? 'Exists' : 'Not exists');
                if (cacheInfo.exists) {
                    console.log('⏰ Cache TTL:', cacheInfo.ttl, 'detik');
                }
            }
        } catch (cacheError) {
            cacheCheckTime = Date.now() - cacheCheckStart;
            console.log('⚠️  Error akses cache:', cacheError.message);
            cacheStatus = 'ERROR';
        }

        // Jika tidak ada di cache, lanjutkan dengan query database
        console.log('🔍 Query database...');
        const queryStartTime = Date.now();
        
        const produkFilter = produk && produk !== 'ALL' ? produk.toUpperCase() : 'ALL';
        const statusFilter = status.toUpperCase();

        // Bangun filter conditions dan parameters
        const baseParams = [];
        const whereConditions = [];
        
        // Filter untuk CTE
        const cteConditions = {
            penebusan: [],
            stok_akhir_prev: [],
            verval: []
        };

        const cteParams = {
            penebusan: [],
            stok_akhir_prev: [],
            verval: []
        };

        // Tahun selalu ada
        cteConditions.penebusan.push('YEAR(pd.tanggal_penyaluran) = ?');
        cteParams.penebusan.push(tahun);
        
        cteConditions.stok_akhir_prev.push('(CASE WHEN w.bulan = 12 THEN w.tahun + 1 ELSE w.tahun END) = ?');
        cteParams.stok_akhir_prev.push(tahun);
        
        cteConditions.verval.push('vf.tahun = ?');
        cteParams.verval.push(tahun);

        // Filter bulan
        if (bulan && bulan !== 'ALL') {
            cteConditions.penebusan.push('MONTH(pd.tanggal_penyaluran) = ?');
            cteParams.penebusan.push(bulan);
            
            cteConditions.stok_akhir_prev.push('(CASE WHEN w.bulan = 12 THEN 1 ELSE w.bulan + 1 END) = ?');
            cteParams.stok_akhir_prev.push(bulan);
            
            cteConditions.verval.push('vf.bulan = ?');
            cteParams.verval.push(bulan);
            
            whereConditions.push('g.bulan = ?');
            baseParams.push(bulan);
        }

        // Filter produk
        if (produkFilter !== 'ALL') {
            cteConditions.penebusan.push('pd.produk = ?');
            cteParams.penebusan.push(produkFilter);
            
            cteConditions.stok_akhir_prev.push('w.produk = ?');
            cteParams.stok_akhir_prev.push(produkFilter);
            
            cteConditions.verval.push('vf.produk = ?');
            cteParams.verval.push(produkFilter);
            
            whereConditions.push('g.produk = ?');
            baseParams.push(produkFilter);
        }

        // Filter kabupaten
        if (kabupaten && kabupaten !== 'ALL') {
            cteConditions.penebusan.push('pd.kabupaten = ?');
            cteParams.penebusan.push(kabupaten);
            
            cteConditions.stok_akhir_prev.push('w.kabupaten = ?');
            cteParams.stok_akhir_prev.push(kabupaten);
            
            cteConditions.verval.push('vf.kabupaten = ?');
            cteParams.verval.push(kabupaten);
            
            whereConditions.push('g.kabupaten = ?');
            baseParams.push(kabupaten);
        }

        // Tahun untuk main query
        whereConditions.push('g.tahun = ?');
        baseParams.push(tahun);

        const whereSQL = whereConditions.length ? 'WHERE ' + whereConditions.join(' AND ') : '';

        // Query yang dioptimalkan dengan filter di setiap CTE
        const baseQuery = `
            WITH penebusan AS (
                SELECT 
                    pd.kode_kios,
                    pd.kecamatan,
                    pd.kabupaten,
                    pd.produk,
                    MONTH(pd.tanggal_penyaluran) AS bulan,
                    YEAR(pd.tanggal_penyaluran) AS tahun,
                    pd.kode_distributor,
                    pd.distributor,
                    pd.nama_kios,
                    pd.provinsi,
                    SUM(pd.qty) AS penebusan
                FROM penyaluran_do pd
                ${cteConditions.penebusan.length ? 'WHERE ' + cteConditions.penebusan.join(' AND ') : ''}
                GROUP BY 
                    pd.kode_kios, pd.kecamatan, pd.kabupaten, pd.produk,
                    MONTH(pd.tanggal_penyaluran), YEAR(pd.tanggal_penyaluran),
                    pd.kode_distributor, pd.distributor, pd.nama_kios, pd.provinsi
            ),
            stok_akhir_prev AS (
                SELECT 
                    w.kode_kios,
                    w.kecamatan,
                    w.kabupaten,
                    w.produk,
                    CASE WHEN w.bulan = 12 THEN 1 ELSE w.bulan + 1 END AS bulan,
                    CASE WHEN w.bulan = 12 THEN w.tahun + 1 ELSE w.tahun END AS tahun,
                    w.kode_distributor,
                    w.nama_kios,
                    w.provinsi,
                    w.distributor,
                    SUM(w.stok_akhir) AS stok_awal
                FROM wcm w
                WHERE w.stok_akhir > 0
                ${cteConditions.stok_akhir_prev.length ? 'AND ' + cteConditions.stok_akhir_prev.join(' AND ') : ''}
                GROUP BY 
                    w.kode_kios, w.kecamatan, w.kabupaten, w.produk,
                    CASE WHEN w.bulan = 12 THEN 1 ELSE w.bulan + 1 END,
                    CASE WHEN w.bulan = 12 THEN w.tahun + 1 ELSE w.tahun END,
                    w.kode_distributor, w.nama_kios, w.provinsi, w.distributor
            ),
            verval AS (
                SELECT 
                    vf.kode_kios,
                    vf.kecamatan,
                    vf.kabupaten,
                    vf.produk,
                    vf.tahun,
                    vf.bulan,
                    SUM(vf.penyaluran) AS total_penyaluran
                FROM verval_f6 vf
                ${cteConditions.verval.length ? 'WHERE ' + cteConditions.verval.join(' AND ') : ''}
                GROUP BY 
                    vf.kode_kios, vf.kecamatan, vf.kabupaten, vf.produk, vf.tahun, vf.bulan
            ),
            gabungan AS (
                SELECT 
                    kode_kios, kecamatan, kabupaten, produk, tahun, bulan,
                    kode_distributor, nama_kios, provinsi, distributor
                FROM penebusan
                UNION
                SELECT 
                    kode_kios, kecamatan, kabupaten, produk, tahun, bulan,
                    kode_distributor, nama_kios, provinsi, distributor
                FROM stok_akhir_prev
            )
            SELECT 
                g.provinsi,
                g.kabupaten,
                g.kecamatan,
                g.kode_kios,
                g.nama_kios,
                g.kode_distributor,
                g.distributor,
                g.tahun,
                g.bulan,
                g.produk,
                ROUND(COALESCE(sa.stok_awal, 0) * 1000, 0) AS stok_awal_wcm,
                ROUND(COALESCE(p.penebusan, 0) * 1000, 0) AS penebusan_wcm,
                ROUND(COALESCE(wcm.penyaluran, 0) * 1000, 0) AS penyaluran_wcm,
                ROUND(COALESCE(v.total_penyaluran, 0), 0) AS penyaluran_verval,
                ROUND(
                    (COALESCE(sa.stok_awal, 0) * 1000 + COALESCE(p.penebusan, 0) * 1000 - COALESCE(v.total_penyaluran, 0)),
                    0
                ) AS stok_akhir_wcm,
                ROUND(COALESCE(wcm.penyaluran, 0) * 1000, 0) - ROUND(COALESCE(v.total_penyaluran, 0), 0) AS status_penyaluran
            FROM gabungan g
            LEFT JOIN penebusan p
                ON g.kode_kios = p.kode_kios
                AND g.kecamatan = p.kecamatan
                AND g.kabupaten = p.kabupaten
                AND g.produk = p.produk
                AND g.tahun = p.tahun
                AND g.bulan = p.bulan
                AND g.kode_distributor = p.kode_distributor
            LEFT JOIN stok_akhir_prev sa
                ON g.kode_kios = sa.kode_kios
                AND g.kecamatan = sa.kecamatan
                AND g.kabupaten = sa.kabupaten
                AND g.produk = sa.produk
                AND g.tahun = sa.tahun
                AND g.bulan = sa.bulan
                AND g.kode_distributor = sa.kode_distributor
            LEFT JOIN wcm
                ON wcm.kode_kios = g.kode_kios
                AND wcm.kecamatan = g.kecamatan
                AND wcm.kabupaten = g.kabupaten
                AND wcm.produk = g.produk
                AND wcm.tahun = g.tahun
                AND wcm.bulan = g.bulan
                AND wcm.kode_distributor = g.kode_distributor
            LEFT JOIN verval v
                ON v.kode_kios = g.kode_kios
                AND v.kecamatan = g.kecamatan
                AND v.kabupaten = g.kabupaten
                AND v.produk = g.produk
                AND v.tahun = g.tahun
                AND v.bulan = g.bulan
            ${whereSQL}
        `;

        // Gabungkan semua parameter untuk base query
        const allBaseParams = [
            ...cteParams.penebusan,
            ...cteParams.stok_akhir_prev,
            ...cteParams.verval,
            ...baseParams
        ];

        // Query untuk mendapatkan total records
        const countQuery = `SELECT COUNT(*) AS total FROM (${baseQuery}) AS count_table`;
        const [countResult] = await db.query(countQuery, allBaseParams);
        const recordsTotal = countResult[0]?.total || 0;

        // Filter berdasarkan status jika diperlukan
        let filteredQuery = baseQuery;
        let filteredParams = [...allBaseParams];

        if (statusFilter !== 'ALL') {
            filteredQuery = `
                SELECT * FROM (${baseQuery}) AS filtered_table 
                WHERE status_penyaluran ${statusFilter === 'SESUAI' ? '= 0' : '!= 0'}
            `;
        }

        // Query untuk mendapatkan filtered count
        const filteredCountQuery = `SELECT COUNT(*) AS total FROM (${filteredQuery}) AS count_filtered_table`;
        const [filteredCountResult] = await db.query(filteredCountQuery, filteredParams);
        const recordsFiltered = filteredCountResult[0]?.total || 0;

        // Query akhir dengan pagination
        const finalQuery = `${filteredQuery} ORDER BY kode_kios LIMIT ? OFFSET ?`;
        filteredParams.push(parseInt(length), parseInt(start));

        const [data] = await db.query(finalQuery, filteredParams);
        
        queryTime = Date.now() - queryStartTime;
        console.log('⏱️  Waktu query database:', queryTime, 'ms');

        // Data untuk disimpan di cache (TANPA draw)
        const dataForCache = {
            recordsTotal,
            recordsFiltered,
            data
        };

        // Data untuk response (DENGAN draw)
        const response = {
            draw: parseInt(draw),
            recordsTotal,
            recordsFiltered,
            data,
            cacheInfo: {
                status: cacheStatus,
                key: cacheKey,
                responseTime: Date.now() - startTime,
                queryTime: queryTime,
                cacheCheckTime: cacheCheckTime,
                source: 'database'
            }
        };

        // Simpan ke cache dengan expiry 5 menit (300 detik) - TANPA draw
        try {
            await client.setEx(cacheKey, 300, JSON.stringify(dataForCache));
            console.log('💾 DISIMPAN KE CACHE (5 menit)');
            console.log('⏰ Akan expired dalam 300 detik');
        } catch (cacheError) {
            console.log('❌ Gagal menyimpan ke cache:', cacheError.message);
            response.cacheInfo.cacheSaveError = cacheError.message;
        }

        console.log('⚡ Total waktu response:', Date.now() - startTime, 'ms');
        console.log('══════════════════════════════════════════════════');
        
        // Header response
        res.set({
            'X-Cache-Status': cacheStatus,
            'X-Cache-Key': cacheKey,
            'X-Response-Time': (Date.now() - startTime) + 'ms',
            'X-Query-Time': queryTime + 'ms',
            'X-Cache-Check-Time': cacheCheckTime + 'ms'
        });

        res.json(response);

    } catch (error) {
        const errorTime = Date.now() - startTime;
        console.error('❌❌❌ ERROR dalam wcmVsVerval:', error);
        console.log('⏰ Waktu sampai error:', errorTime, 'ms');
        console.log('══════════════════════════════════════════════════');
        
        res.set({
            'X-Cache-Status': 'ERROR',
            'X-Error-Time': errorTime + 'ms'
        });
        
        res.status(500).json({ 
            message: 'Terjadi kesalahan saat memproses data.', 
            error: error.message,
            cacheInfo: {
                status: 'ERROR',
                key: cacheKey,
                responseTime: errorTime
            }
        });
    }
};

exports.exportExcelWcmVsVerval = async (req, res) => {
    try {
        const { tahun, bulan, produk, kabupaten, status = 'ALL' } = req.query;

        // Validasi parameter wajib
        if (!tahun) {
            return res.status(400).json({ message: 'Parameter tahun wajib diisi' });
        }

        const produkFilter = produk && produk !== 'ALL' ? produk.toUpperCase() : 'ALL';
        const statusFilter = status.toUpperCase();

        const params = [];
        let whereClauses = ['g.tahun = ?'];
        params.push(tahun);

        if (bulan && bulan !== 'ALL') {
            whereClauses.push('g.bulan = ?');
            params.push(bulan);
        }

        if (produk && produk !== 'ALL') {
            whereClauses.push('g.produk = ?');
            params.push(produkFilter);
        }

        if (kabupaten && kabupaten !== 'ALL') {
            whereClauses.push('g.kabupaten = ?');
            params.push(kabupaten);
        }

        const whereSQL = whereClauses.length > 0 ? ' WHERE ' + whereClauses.join(' AND ') : '';

        const baseQuery = `
    WITH
penebusan AS (
  SELECT 
    pd.kode_kios,
    pd.kecamatan,
    pd.kabupaten,
    pd.produk,
    MONTH(pd.tanggal_penyaluran) AS bulan,
    YEAR(pd.tanggal_penyaluran) AS tahun,
    pd.kode_distributor,
    pd.distributor,
    pd.nama_kios,
    pd.provinsi,
    SUM(pd.qty) AS penebusan
  FROM penyaluran_do pd
  GROUP BY 
    pd.kode_kios,
    pd.kecamatan,
    pd.kabupaten,
    pd.produk,
    MONTH(pd.tanggal_penyaluran),
    YEAR(pd.tanggal_penyaluran),
    pd.kode_distributor,
    pd.distributor,
    pd.nama_kios,
    pd.provinsi
),

stok_akhir_prev AS (
  SELECT 
    w.kode_kios,
    w.kecamatan,
    w.kabupaten,
    w.produk,
    CASE WHEN w.bulan = 12 THEN 1 ELSE w.bulan + 1 END AS bulan,
    CASE WHEN w.bulan = 12 THEN w.tahun + 1 ELSE w.tahun END AS tahun,
    w.kode_distributor,
    w.nama_kios,
    w.provinsi,
    w.distributor,
    SUM(w.stok_akhir) AS stok_awal
  FROM wcm w
  WHERE w.stok_akhir > 0
  GROUP BY 
    w.kode_kios,
    w.kecamatan,
    w.kabupaten,
    w.produk,
    CASE WHEN w.bulan = 12 THEN 1 ELSE w.bulan + 1 END,
    CASE WHEN w.bulan = 12 THEN w.tahun + 1 ELSE w.tahun END,
    w.kode_distributor,
    w.nama_kios,
    w.provinsi,
    w.distributor
),

verval AS (
  SELECT 
    vf.kode_kios,
    vf.kecamatan,
    vf.kabupaten,
    vf.produk,
    vf.tahun,
    vf.bulan,
    SUM(vf.penyaluran) AS total_penyaluran
  FROM verval_f6 vf
  GROUP BY 
    vf.kode_kios,
    vf.kecamatan,
    vf.kabupaten,
    vf.produk,
    vf.tahun,
    vf.bulan
),

gabungan AS (
  SELECT 
    p.kode_kios,
    p.kecamatan,
    p.kabupaten,
    p.produk,
    p.tahun,
    p.bulan,
    p.kode_distributor,
    p.nama_kios,
    p.provinsi,
    p.distributor
  FROM penebusan p

  UNION

  SELECT 
    s.kode_kios,
    s.kecamatan,
    s.kabupaten,
    s.produk,
    s.tahun,
    s.bulan,
    s.kode_distributor,
    s.nama_kios,
    s.provinsi,
    s.distributor
  FROM stok_akhir_prev s
)

SELECT 
  g.provinsi,
  g.kabupaten,
  g.kecamatan,
  g.kode_kios,
  g.nama_kios,
  g.kode_distributor,
  g.distributor,
  g.tahun,
  g.bulan,
  g.produk,

  ROUND(COALESCE(sa.stok_awal, 0) * 1000, 0) AS stok_awal_wcm,
  ROUND(COALESCE(p.penebusan, 0) * 1000, 0) AS penebusan_wcm,
  ROUND(COALESCE(wcm.penyaluran, 0) * 1000, 0) AS penyaluran_wcm,
  ROUND(COALESCE(v.total_penyaluran, 0), 0) AS penyaluran_verval,

  ROUND(
    (COALESCE(sa.stok_awal, 0) * 1000 + COALESCE(p.penebusan, 0) * 1000 - COALESCE(v.total_penyaluran, 0)),
    0
  ) AS stok_akhir_wcm,

  ROUND(COALESCE(wcm.penyaluran, 0) * 1000, 0) - ROUND(COALESCE(v.total_penyaluran, 0), 0) AS status_penyaluran

FROM gabungan g
LEFT JOIN penebusan p
  ON g.kode_kios = p.kode_kios
  AND g.kecamatan = p.kecamatan
  AND g.kabupaten = p.kabupaten
  AND g.produk = p.produk
  AND g.tahun = p.tahun
  AND g.bulan = p.bulan
  AND g.kode_distributor = p.kode_distributor

LEFT JOIN stok_akhir_prev sa
  ON g.kode_kios = sa.kode_kios
  AND g.kecamatan = sa.kecamatan
  AND g.kabupaten = sa.kabupaten
  AND g.produk = sa.produk
  AND g.tahun = sa.tahun
  AND g.bulan = sa.bulan
  AND g.kode_distributor = sa.kode_distributor

LEFT JOIN wcm
  ON wcm.kode_kios = g.kode_kios
  AND wcm.kecamatan = g.kecamatan
  AND wcm.kabupaten = g.kabupaten
  AND wcm.produk = g.produk
  AND wcm.tahun = g.tahun
  AND wcm.bulan = g.bulan
  AND wcm.kode_distributor = g.kode_distributor

LEFT JOIN verval v
  ON v.kode_kios = g.kode_kios
  AND v.kecamatan = g.kecamatan
  AND v.kabupaten = g.kabupaten
  AND v.produk = g.produk
  AND v.tahun = g.tahun
  AND v.bulan = g.bulan

${whereSQL}

ORDER BY g.kode_kios
`;
        // Filter status jika bukan ALL
        let finalQuery = baseQuery;
        if (statusFilter !== 'ALL') {
            const operator = statusFilter === 'SESUAI' ? '=' : '!=';
            finalQuery = `
        SELECT * FROM (${baseQuery}) AS filtered_table
        WHERE status_penyaluran ${operator} ?
    `;
            params.push(0);
        }


        const [data] = await db.query(finalQuery, params);

        // Create Excel workbook
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('WCM vs Verval');

        // Styling untuk header
        const headerStyle = {
            font: { bold: true, color: { argb: 'FFFFFFFF' } },
            alignment: { vertical: 'middle', horizontal: 'center', wrapText: true },
            border: {
                top: { style: 'thin', color: { argb: 'FF000000' } },
                left: { style: 'thin', color: { argb: 'FF000000' } },
                bottom: { style: 'thin', color: { argb: 'FF000000' } },
                right: { style: 'thin', color: { argb: 'FF000000' } }
            }
        };

        // Header row 1
        worksheet.mergeCells('A1:A2');
        worksheet.getCell('A1').value = 'Provinsi';
        worksheet.getCell('A1').style = { ...headerStyle, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF7F7F7F' } } };

        worksheet.mergeCells('B1:B2');
        worksheet.getCell('B1').value = 'Kabupaten';
        worksheet.getCell('B1').style = { ...headerStyle, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF7F7F7F' } } };

        worksheet.mergeCells('C1:C2');
        worksheet.getCell('C1').value = 'Kode Distributor';
        worksheet.getCell('C1').style = { ...headerStyle, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF7F7F7F' } } };

        worksheet.mergeCells('D1:D2');
        worksheet.getCell('D1').value = 'Distributor';
        worksheet.getCell('D1').style = { ...headerStyle, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF7F7F7F' } } };

        worksheet.mergeCells('E1:E2');
        worksheet.getCell('E1').value = 'Kecamatan';
        worksheet.getCell('E1').style = { ...headerStyle, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF7F7F7F' } } };

        worksheet.mergeCells('F1:F2');
        worksheet.getCell('F1').value = 'Kode Kios';
        worksheet.getCell('F1').style = { ...headerStyle, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF7F7F7F' } } };

        worksheet.mergeCells('G1:G2');
        worksheet.getCell('G1').value = 'Nama Kios';
        worksheet.getCell('G1').style = { ...headerStyle, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF7F7F7F' } } };

        worksheet.mergeCells('H1:H2');
        worksheet.getCell('H1').value = 'Produk';
        worksheet.getCell('H1').style = { ...headerStyle, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF7F7F7F' } } };

        worksheet.mergeCells('I1:I2');
        worksheet.getCell('I1').value = 'Bulan';
        worksheet.getCell('I1').style = { ...headerStyle, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF7F7F7F' } } };

        worksheet.mergeCells('J1:L1');
        worksheet.getCell('J1').value = 'WCM';
        worksheet.getCell('J1').style = { ...headerStyle, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' } } };

        worksheet.getCell('M1').value = 'VERVAL';
        worksheet.getCell('M1').style = { ...headerStyle, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFEB9C' } }, font: { ...headerStyle.font, color: { argb: 'FF000000' } } };
        worksheet.getCell('M2').value = 'Penyaluran';
        worksheet.getCell('M2').style = { ...headerStyle, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFEB9C' } }, font: { ...headerStyle.font, color: { argb: 'FF000000' } } };

        worksheet.getCell('N1').value = 'WCM';
        worksheet.getCell('N1').style = { ...headerStyle, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' } } };
        worksheet.getCell('N2').value = 'Stok Akhir';
        worksheet.getCell('N2').style = { ...headerStyle, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' } } };

        worksheet.mergeCells('O1:O2');
        worksheet.getCell('O1').value = 'Selisih Penyaluran';
        worksheet.getCell('O1').style = { ...headerStyle, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF7F7F7F' } } };

        // Header row 2 (hanya untuk kolom WCM)
        worksheet.getCell('J2').value = 'Stok Awal';
        worksheet.getCell('J2').style = { ...headerStyle, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' } } };

        worksheet.getCell('K2').value = 'Penebusan';
        worksheet.getCell('K2').style = { ...headerStyle, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' } } };

        worksheet.getCell('L2').value = 'Penyaluran';
        worksheet.getCell('L2').style = { ...headerStyle, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' } } };

        // Add data rows
        data.forEach((row, index) => {
            const dataRow = worksheet.addRow([
                row.provinsi,
                row.kabupaten,
                row.kode_distributor,
                row.distributor,
                row.kecamatan,
                row.kode_kios,
                row.nama_kios,
                row.produk,
                row.bulan,
                row.stok_awal_wcm,
                row.penebusan_wcm,
                row.penyaluran_wcm,
                row.penyaluran_verval,
                row.stok_akhir_wcm,
                row.status_penyaluran
            ]);

            // Format angka dengan pemisah ribuan
            [9, 10, 11, 12, 13].forEach(col => {
                dataRow.getCell(col).numFmt = '#,##0';
            });

            // Warna status
            if (row.status_penyaluran === 'Sesuai') {
                dataRow.getCell(15).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC6EFCE' } };
            } else {
                dataRow.getCell(15).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFC7CE' } };
            }
        });

        // Set column widths
        worksheet.columns = [
            { width: 15 }, // Provinsi
            { width: 15 }, // Kabupaten
            { width: 15 }, // Kode Distributor
            { width: 20 }, // Distributor
            { width: 15 }, // Kecamatan
            { width: 15 }, // Kode Kios
            { width: 25 }, // Nama Kios
            { width: 15 }, // Produk
            { width: 10 }, // Bulan
            { width: 12 }, // Stok Awal WCM
            { width: 12 }, // Penebusan WCM
            { width: 12 }, // Penyaluran WCM
            { width: 12 }, // Penyaluran Verval
            { width: 12 }, // Stok Akhir WCM
            { width: 12 }  // Status
        ];

        // Freeze header rows
        worksheet.views = [{ state: 'frozen', xSplit: 0, ySplit: 2 }];

        // Set response headers
        res.setHeader(
            'Content-Type',
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        );
        res.setHeader(
            'Content-Disposition',
            `attachment; filename=wcm_vs_verval_${tahun}${bulan ? '_' + bulan : ''}.xlsx`
        );

        // Write workbook to response
        await workbook.xlsx.write(res);
        res.end();

    } catch (error) {
        console.error('Error in exportExcelWcmVsVerval:', error);
        res.status(500).json({
            message: 'Terjadi kesalahan saat mengekspor data.',
            error: error.message
        });
    }
};

exports.getWcm = async (req, res) => {
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
    kecamatan,
    kode_kios,
    nama_kios,
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
FROM wcm
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

        query += " GROUP BY kode_provinsi, provinsi, kode_kabupaten, kabupaten, kode_distributor, distributor, kecamatan, kode_kios, nama_kios, produk";

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

exports.getWcmF5 = async (req, res) => {
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
FROM wcm
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

exports.downloadWcm = async (req, res) => {
    try {
        const { produk, tahun, provinsi, kabupaten } = req.query;

        let query = `
            SELECT 
                provinsi,
                kabupaten, 
                kode_distributor, 
                distributor, 
                kecamatan,
                kode_kios,
                nama_kios,
                produk,
                ${[...Array(12).keys()].map(bulan => `
                    SUM(CASE WHEN bulan = ${bulan + 1} THEN stok_awal ELSE 0 END) AS stok_awal_${bulan + 1},
                    SUM(CASE WHEN bulan = ${bulan + 1} THEN penebusan ELSE 0 END) AS penebusan_${bulan + 1},
                    SUM(CASE WHEN bulan = ${bulan + 1} THEN penyaluran ELSE 0 END) AS penyaluran_${bulan + 1},
                    SUM(CASE WHEN bulan = ${bulan + 1} THEN stok_akhir ELSE 0 END) AS stok_akhir_${bulan + 1}
                `).join(",")}
            FROM wcm
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

        query += ` GROUP BY provinsi, kabupaten, kode_distributor, distributor, kecamatan, kode_kios, nama_kios, produk
                   ORDER BY provinsi, kabupaten, kode_distributor`;

        const [data] = await db.query(query, params);

        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet("WCM");

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
        worksheet.mergeCells("A1:A2"); // Provinsi
        worksheet.mergeCells("B1:B2"); // Kabupaten
        worksheet.mergeCells("C1:C2"); // id distributor
        worksheet.mergeCells("D1:D2"); // distributor
        worksheet.mergeCells("E1:E2"); // kecamatan
        worksheet.mergeCells("F1:F2"); // kode kios
        worksheet.mergeCells("G1:G2"); // nama kios
        worksheet.mergeCells("H1:H2"); // produk
        worksheet.getCell("A1").value = "PROVINSI";
        worksheet.getCell("B1").value = "KABUPATEN";
        worksheet.getCell("C1").value = "ID DISTRIBUTOR";
        worksheet.getCell("D1").value = "DISTRIBUTOR";
        worksheet.getCell("E1").value = "KECAMATAN";
        worksheet.getCell("F1").value = "KODE PENGECER";
        worksheet.getCell("G1").value = "PENGECER";
        worksheet.getCell("H1").value = "PRODUK";

        // Header Bulan
        let colStart = 9; // Mulai dari kolom I (setelah kolom produk)
        bulanList.forEach((bulan, index) => {
            let colEnd = colStart + 3; // Setiap bulan memiliki 4 kolom (stok awal, tebus, salur, stok akhir)
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
            { key: "provinsi", width: 20 }, // Kolom A
            { key: "kabupaten", width: 20 }, // Kolom B
            { key: "kode_distributor", width: 15 }, // Kolom C
            { key: "distributor", width: 25 }, // Kolom D
            { key: "kecamatan", width: 20 }, // Kolom E
            { key: "kode_kios", width: 15 }, // Kolom F
            { key: "nama_kios", width: 25 }, // Kolom G
            { key: "produk", width: 20 }, // Kolom H
            ...Array.from({ length: 12 * 4 }, () => ({ width: 12 })) // Kolom I dan seterusnya
        ];

        // Kelompokkan data berdasarkan provinsi, kabupaten, dan distributor
        const groupedData = data.reduce((acc, row) => {
            if (!acc[row.provinsi]) {
                acc[row.provinsi] = {};
            }

            if (!acc[row.provinsi][row.kabupaten]) {
                acc[row.provinsi][row.kabupaten] = {};
            }

            if (!acc[row.provinsi][row.kabupaten][row.distributor]) {
                acc[row.provinsi][row.kabupaten][row.distributor] = [];
            }

            acc[row.provinsi][row.kabupaten][row.distributor].push(row);
            return acc;
        }, {});

        // Tambahkan data ke worksheet
        Object.keys(groupedData).forEach(provinsi => {
            Object.keys(groupedData[provinsi]).forEach(kabupaten => {
                const distributors = groupedData[provinsi][kabupaten];
                let kabupatenTotal = {
                    provinsi: '',
                    kabupaten: '',
                    kode_distributor: '',
                    distributor: `Total ${kabupaten}`,
                    kecamatan: '',
                    kode_kios: '',
                    nama_kios: '',
                    produk: '',
                    stok_awal: Array(12).fill(0),
                    penebusan: Array(12).fill(0),
                    penyaluran: Array(12).fill(0),
                    stok_akhir: Array(12).fill(0)
                };

                Object.keys(distributors).forEach(distributor => {
                    const rows = distributors[distributor];
                    let distributorTotal = {
                        provinsi: '',
                        kabupaten: '',
                        kode_distributor: '',
                        distributor: `Total ${distributor}`,
                        kecamatan: '',
                        kode_kios: '',
                        nama_kios: '',
                        produk: '',
                        stok_awal: Array(12).fill(0),
                        penebusan: Array(12).fill(0),
                        penyaluran: Array(12).fill(0),
                        stok_akhir: Array(12).fill(0)
                    };

                    rows.forEach(row => {
                        let rowData = [
                            row.provinsi,
                            row.kabupaten,
                            row.kode_distributor,
                            row.distributor,
                            row.kecamatan,
                            row.kode_kios,
                            row.nama_kios,
                            row.produk
                        ];

                        // Tambahkan data bulanan
                        for (let i = 1; i <= 12; i++) {
                            rowData.push(
                                parseFloat((row[`stok_awal_${i}`] || 0).toFixed(3)),
                                parseFloat((row[`penebusan_${i}`] || 0).toFixed(3)),
                                parseFloat((row[`penyaluran_${i}`] || 0).toFixed(3)),
                                parseFloat((row[`stok_akhir_${i}`] || 0).toFixed(3))
                            );

                            // Hitung total distributor
                            distributorTotal.stok_awal[i - 1] += row[`stok_awal_${i}`] || 0;
                            distributorTotal.penebusan[i - 1] += row[`penebusan_${i}`] || 0;
                            distributorTotal.penyaluran[i - 1] += row[`penyaluran_${i}`] || 0;
                            distributorTotal.stok_akhir[i - 1] += row[`stok_akhir_${i}`] || 0;

                            // Hitung total kabupaten
                            kabupatenTotal.stok_awal[i - 1] += row[`stok_awal_${i}`] || 0;
                            kabupatenTotal.penebusan[i - 1] += row[`penebusan_${i}`] || 0;
                            kabupatenTotal.penyaluran[i - 1] += row[`penyaluran_${i}`] || 0;
                            kabupatenTotal.stok_akhir[i - 1] += row[`stok_akhir_${i}`] || 0;
                        }

                        let addedRow = worksheet.addRow(rowData);
                        addedRow.eachCell((cell) => {
                            cell.border = borderStyle;
                            cell.alignment = { horizontal: "center", vertical: "middle" };
                        });
                    });

                    // Tambahkan baris total distributor
                    let distributorTotalRowData = [
                        distributorTotal.provinsi,
                        distributorTotal.kabupaten,
                        distributorTotal.kode_distributor,
                        distributorTotal.distributor,
                        distributorTotal.kecamatan,
                        distributorTotal.kode_kios,
                        distributorTotal.nama_kios,
                        distributorTotal.produk
                    ];

                    for (let i = 0; i < 12; i++) {
                        distributorTotalRowData.push(
                            parseFloat(distributorTotal.stok_awal[i].toFixed(3)),
                            parseFloat(distributorTotal.penebusan[i].toFixed(3)),
                            parseFloat(distributorTotal.penyaluran[i].toFixed(3)),
                            parseFloat(distributorTotal.stok_akhir[i].toFixed(3))
                        );
                    }

                    let addedDistributorTotalRow = worksheet.addRow(distributorTotalRowData);
                    addedDistributorTotalRow.eachCell((cell) => {
                        cell.border = borderStyle;
                        cell.font = { bold: true };
                        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD3D3D3" } };
                        cell.alignment = { horizontal: "center", vertical: "middle" };
                    });
                });

                // Tambahkan baris total kabupaten
                let kabupatenTotalRowData = [
                    kabupatenTotal.provinsi,
                    kabupatenTotal.kabupaten,
                    kabupatenTotal.kode_distributor,
                    kabupatenTotal.distributor,
                    kabupatenTotal.kecamatan,
                    kabupatenTotal.kode_kios,
                    kabupatenTotal.nama_kios,
                    kabupatenTotal.produk
                ];

                for (let i = 0; i < 12; i++) {
                    kabupatenTotalRowData.push(
                        parseFloat(kabupatenTotal.stok_awal[i].toFixed(3)),
                        parseFloat(kabupatenTotal.penebusan[i].toFixed(3)),
                        parseFloat(kabupatenTotal.penyaluran[i].toFixed(3)),
                        parseFloat(kabupatenTotal.stok_akhir[i].toFixed(3))
                    );
                }

                let addedKabupatenTotalRow = worksheet.addRow(kabupatenTotalRowData);
                addedKabupatenTotalRow.eachCell((cell) => {
                    cell.border = borderStyle;
                    cell.font = { bold: true };
                    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFA9A9A9" } }; // Warna abu-abu lebih gelap
                    cell.alignment = { horizontal: "center", vertical: "middle" };
                });
            });
        });

        // Kirim file Excel sebagai respons
        res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        res.setHeader("Content-Disposition", "attachment; filename=wcm.xlsx");
        await workbook.xlsx.write(res);
        res.end();
    } catch (error) {
        console.error(error);
        res.status(500).json({ error: "Terjadi kesalahan dalam mengunduh data" });
    }
};

exports.getPenyaluranDo = async (req, res) => {
    try {
        // Mendapatkan parameter dari query string
        const { start, length, draw, produk, tahun, bulan, provinsi, kabupaten } = req.query;

        // Menyiapkan query dasar
        let query = `
            SELECT 
                produsen,
                nomor_ff,
                kode_distributor,
                distributor,
                tipe_penyaluran,
                nama_order,
                nomor_so,
                kode_provinsi,
                provinsi,
                kode_kabupaten,
                kabupaten,
                kode_kecamatan,
                kecamatan,
                kode_kios,
                nama_kios,
                tanggal_penyaluran,
                produk,
                qty,
                status_ff,
                status
            FROM penyaluran_do
            WHERE 1=1
        `;

        // Menyiapkan parameter untuk query
        const params = [];

        // Filter berdasarkan produk jika ada
        if (produk) {
            query += ` AND produk = ?`;
            params.push(produk);
        }

        // Filter berdasarkan tahun jika ada
        if (tahun) {
            query += ` AND YEAR(tanggal_penyaluran) = ?`;
            params.push(tahun);
        }

        // Filter berdasarkan bulan jika ada
        if (bulan) {
            query += ` AND MONTH(tanggal_penyaluran) = ?`;
            params.push(bulan);
        }

        // Filter berdasarkan provinsi jika ada
        if (provinsi) {
            query += ` AND provinsi = ?`;
            params.push(provinsi);
        }

        // Filter berdasarkan kabupaten jika ada
        if (kabupaten) {
            query += ` AND kabupaten = ?`;
            params.push(kabupaten);
        }

        // Query untuk mendapatkan total data tanpa filter
        const countQuery = `SELECT COUNT(*) as total FROM penyaluran_do`;
        const [totalResult] = await db.query(countQuery);
        const totalRecords = totalResult[0].total;

        // Query untuk mendapatkan total data dengan filter
        const filteredCountQuery = query.replace(/SELECT.*FROM/, 'SELECT COUNT(*) as filtered FROM');
        const [filteredResult] = await db.query(filteredCountQuery, params);
        const filteredRecords = filteredResult[0].filtered;

        // Menambahkan sorting dan pagination
        query += ` ORDER BY tanggal_penyaluran DESC LIMIT ?, ?`;
        params.push(parseInt(start), parseInt(length));

        // Menjalankan query utama
        const [data] = await db.query(query, params);

        // Format tanggal sebelum dikirim ke client
        const formattedData = data.map(item => ({
            ...item,
            tanggal_penyaluran: item.tanggal_penyaluran ? new Date(item.tanggal_penyaluran).toISOString().split('T')[0] : null
        }));

        // Mengirim response dalam format yang dibutuhkan DataTables
        res.json({
            draw: parseInt(draw),
            recordsTotal: totalRecords,
            recordsFiltered: filteredRecords,
            data: formattedData
        });

    } catch (error) {
        console.error('Error saat mengambil data penyaluran DO:', error);
        res.status(500).json({
            success: false,
            message: 'Terjadi kesalahan saat mengambil data penyaluran DO'
        });
    }
};

exports.downloadWcmF5 = async (req, res) => {
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
            FROM wcm
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

exports.getLastUpdatedGlobal = async (req, res) => {
    try {
        const [result] = await db.query(`
        SELECT 
          MAX(tanggal_tebus) AS lastUpdated,
          YEAR(MAX(tanggal_tebus)) AS tahunUpdated,
          MONTHNAME(MAX(tanggal_tebus)) AS bulanUpdated,
          (
            SELECT GROUP_CONCAT(DISTINCT kabupaten SEPARATOR ', ')
            FROM verval
            WHERE tanggal_tebus = (
              SELECT MAX(tanggal_tebus)
              FROM verval
              WHERE tanggal_tebus IS NOT NULL
            )
          ) AS kabupatenUpdated
        FROM verval
        WHERE tanggal_tebus IS NOT NULL;
      `);

        res.json(result[0]); // Mengirim hasil dalam bentuk JSON
    } catch (err) {
        console.error('Error ambil last updated global:', err);
        res.status(500).send('Terjadi kesalahan server');
    }
};

// API untuk mendapatkan last updated per kabupaten
exports.getLastUpdatedPerKabupaten = async (req, res) => {
    try {
        const [result] = await db.query(`
        SELECT 
          kabupaten,
          MAX(tanggal_tebus) AS lastUpdated,
          YEAR(MAX(tanggal_tebus)) AS tahunUpdated,
          MONTHNAME(MAX(tanggal_tebus)) AS bulanUpdated
        FROM verval
        WHERE tanggal_tebus IS NOT NULL
        GROUP BY kabupaten
        ORDER BY kabupaten;
      `);

        res.json(result); // Mengirim hasil dalam bentuk JSON
    } catch (err) {
        console.error('Error ambil last updated per kabupaten:', err);
        res.status(500).send('Terjadi kesalahan server');
    }
};

// rekap poktan

exports.downloadPoktan = async (req, res) => {
    try {
        const { kabupaten } = req.query;

        //     let query = `
        //         SELECT DISTINCT
        //             e.kabupaten,
        //             e.kecamatan,
        //             e.desa,
        //             e.kode_kios,
        //             e.nama_kios,
        //             e.poktan, 
        //             p.nomor_hp
        //         FROM 
        //             erdkk e
        //         LEFT JOIN 
        //             poktan p ON TRIM(LOWER(e.poktan)) = TRIM(LOWER(p.poktan))
        //             AND e.kabupaten = p.kabupaten
        //             AND e.kecamatan = p.kecamatan
        //             AND e.desa = p.desa
        //             AND e.kode_kios = p.kode_kios
        //         WHERE 
        //             e.tahun = 2025
        //             AND (
        //                 p.poktan IS NULL 
        //                 OR p.nomor_hp IS NULL 
        //                 OR TRIM(p.nomor_hp) = ''
        //             ) AND e.kabupaten IN (
        //     'SLEMAN', 
        //     'KOTA YOGYAKARTA', 
        //     'BANTUL', 
        //     'KULON PROGO', 
        //     'GUNUNG KIDUL'
        // )
        //     `;


        let query = `
                SELECT DISTINCT
                    e.kabupaten,
                    e.kecamatan,
                    e.desa,
                    e.kode_kios,
                    e.nama_kios,
                    e.poktan, 
                    p.nomor_hp
                FROM 
                    erdkk e
                LEFT JOIN 
                    poktan p ON TRIM(LOWER(e.poktan)) = TRIM(LOWER(p.poktan))
                    AND e.kabupaten = p.kabupaten
                    AND e.kecamatan = p.kecamatan
                    AND e.desa = p.desa
                    AND e.kode_kios = p.kode_kios
                WHERE 
                    e.tahun = 2025
            `;

        // let query = `
        //     SELECT * FROM poktan WHERE nomor_hp IS NULL OR nomor_hp = ''
        // `;



        let params = [];

        if (kabupaten) {
            query += " AND e.kabupaten = ?";
            params.push(kabupaten);
        }

        query += " ORDER BY kabupaten, kecamatan ASC";

        const [data] = await db.query(query, params);

        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet("poktan");

        const borderStyle = {
            top: { style: "thin" },
            left: { style: "thin" },
            bottom: { style: "thin" },
            right: { style: "thin" },
        };

        // Tambahkan header
        worksheet.addRow(["KABUPATEN", "KECAMATAN", "KODE DESA", "DESA", "KODE KIOS", "NAMA KIOS", "NAMA POKTAN", "NAMA KETUA POKTAN", "NO HP POKTAN"]);

        // Atur lebar kolom
        worksheet.columns = [
            { key: "kabupaten", width: 20 },
            { key: "kecamatan", width: 20 },
            { key: "kode_desa", width: 15 },
            { key: "desa", width: 20 },
            { key: "kode_kios", width: 15 },
            { key: "nama_kios", width: 25 },
            { key: "poktan", width: 25 },
            { key: "nama_ketua", width: 25 },
            { key: "nomor_hp", width: 20 },
        ];

        data.forEach((row) => {
            const addedRow = worksheet.addRow([
                row.kabupaten,
                row.kecamatan,
                "",
                row.desa,
                row.kode_kios || "",
                row.nama_kios || "",
                row.poktan,
                "", // nama ketua poktan (kosong untuk saat ini)
                row.nomor_hp || "", // nomor HP poktan (kosong untuk saat ini)
            ]);
            addedRow.eachCell((cell) => {
                cell.border = borderStyle;
                cell.alignment = { horizontal: "center", vertical: "middle" };
            });
        });

        res.setHeader(
            "Content-Type",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        );
        res.setHeader("Content-Disposition", "attachment; filename=poktan_kosong.xlsx");
        await workbook.xlsx.write(res);
        res.end();
    } catch (error) {
        console.error(error);
        res.status(500).json({ error: "Terjadi kesalahan dalam mengunduh data" });
    }
};

