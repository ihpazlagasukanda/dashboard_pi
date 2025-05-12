const db = require("../config/db");
const ExcelJS = require('exceljs');

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
exports.getErdkkSummary = async (req, res) => {
    try {
        const { kabupaten, kecamatan, tahun, provinsi } = req.query; // Tambah provinsi dari frontend

        let query = `
            SELECT 
                SUM(urea) AS total_urea, 
                SUM(npk) AS total_npk, 
                SUM(npk_formula) AS total_npk_formula, 
                SUM(organik) AS total_organik
            FROM erdkk
            WHERE 1 = 1
        `;

        let params = [];

        // Mapping Provinsi ke Kabupaten
        if (provinsi) {
            if (provinsi === 'DIY') {
                query += ` AND kabupaten IN (?, ?, ?, ?, ?)`;
                params.push('SLEMAN', 'KOTA YOGYAKARTA', 'GUNUNG KIDUL', 'BANTUL', 'KULON PROGO');
            } else if (provinsi === 'JAWA TENGAH') {
                query += ` AND kabupaten NOT IN (?, ?, ?, ?, ?)`;
                params.push('SLEMAN', 'KOTA YOGYAKARTA', 'GUNUNG KIDUL', 'BANTUL', 'KULON PROGO');
            }
        }

        if (kabupaten) {
            query += " AND kabupaten = ?";
            params.push(kabupaten);
        }

        if (kecamatan) {
            query += " AND kecamatan = ?";
            params.push(kecamatan);
        }

        if (tahun) {
            query += " AND tahun = ?";
            params.push(tahun);
        }

        const [rows] = await db.execute(query, params);
        res.json(rows[0]); // Kirim hasil SUM ke frontend
    } catch (error) {
        console.error("Database Error:", error);
        res.status(500).json({ message: "Error menghitung SUM ERDKK", error });
    }
};

exports.getSkBupatiAlokasi = async (req, res) => {
    try {
        const { kabupaten, kecamatan, tahun, provinsi } = req.query; // Tambah provinsi dari frontend

        let query = `
            SELECT
    SUM(CASE WHEN produk = 'urea' THEN alokasi ELSE 0 END) AS total_urea,
    SUM(CASE WHEN produk = 'npk' THEN alokasi ELSE 0 END) AS total_npk,
    SUM(CASE WHEN produk = 'npk_formula' THEN alokasi ELSE 0 END) AS total_npk_formula,
    SUM(CASE WHEN produk = 'organik' THEN alokasi ELSE 0 END) AS total_organik,
    SUM(CASE WHEN produk = 'kakao' THEN alokasi ELSE 0 END) AS total_kakao
FROM sk_bupati
WHERE 1 = 1
        `;

        let params = [];

        // Mapping Provinsi ke Kabupaten
        if (provinsi) {
            if (provinsi === 'DIY') {
                query += ` AND kabupaten IN (?, ?, ?, ?, ?)`;
                params.push('SLEMAN', 'KOTA YOGYAKARTA', 'GUNUNG KIDUL', 'BANTUL', 'KULON PROGO');
            } else if (provinsi === 'JAWA TENGAH') {
                query += ` AND kabupaten NOT IN (?, ?, ?, ?, ?)`;
                params.push('SLEMAN', 'KOTA YOGYAKARTA', 'GUNUNG KIDUL', 'BANTUL', 'KULON PROGO');
            }
        }

        if (kabupaten) {
            query += " AND kabupaten = ?";
            params.push(kabupaten);
        }

        if (kecamatan) {
            query += " AND kecamatan = ?";
            params.push(kecamatan);
        }

        if (tahun) {
            query += " AND tahun = ?";
            params.push(tahun);
        }

        const [rows] = await db.execute(query, params);
        res.json(rows[0]); // Kirim hasil SUM ke frontend
    } catch (error) {
        console.error("Database Error:", error);
        res.status(500).json({ message: "Error menghitung SUM ERDKK", error });
    }
};

exports.getErdkkCount = async (req, res) => {
    try {
        const { provinsi, kabupaten, kecamatan, tahun } = req.query; // Tambah provinsi

        let query = `
            SELECT 
                COUNT(DISTINCT CASE WHEN urea > 0 THEN nik END) AS count_urea,
                COUNT(DISTINCT CASE WHEN npk > 0 THEN nik END) AS count_npk,
                COUNT(DISTINCT CASE WHEN npk_formula > 0 THEN nik END) AS count_npk_formula, 
                COUNT(DISTINCT CASE WHEN organik > 0 THEN nik END) AS count_organik
            FROM erdkk
            WHERE 1 = 1
        `;

        let params = [];

        // Filter Provinsi
        if (provinsi) {
            if (provinsi === 'DIY') {
                query += ` AND kabupaten IN (?, ?, ?, ?, ?)`;
                params.push('SLEMAN', 'KOTA YOGYAKARTA', 'GUNUNG KIDUL', 'BANTUL', 'KULON PROGO');
            } else if (provinsi === 'JAWA TENGAH') {
                query += ` AND kabupaten NOT IN (?, ?, ?, ?, ?)`;
                params.push('SLEMAN', 'KOTA YOGYAKARTA', 'GUNUNG KIDUL', 'BANTUL', 'KULON PROGO');
            }
        }

        if (kabupaten) {
            query += " AND kabupaten = ?";
            params.push(kabupaten);
        }

        if (kecamatan) {
            query += " AND kecamatan = ?";
            params.push(kecamatan);
        }

        if (tahun) {
            query += " AND tahun = ?";
            params.push(tahun);
        }

        const [rows] = await db.execute(query, params);
        res.json(rows[0] || { count_urea: 0, count_npk: 0, count_npk_formula: 0, count_organik: 0 });

    } catch (error) {
        console.error("Database Error:", error);
        res.status(500).json({ message: "Error menghitung SUM ERDKK", error });
    }
};


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



// API untuk SUM Penebusan berdasarkan filter tahun, kabupaten, dan kecamatan
exports.getSummaryPenebusan = async (req, res) => {
    try {
        const { provinsi, kabupaten, kecamatan, tahun } = req.query; // Tambah provinsi

        let query = `
            SELECT 
                SUM(urea) AS total_urea, 
                SUM(npk) AS total_npk, 
                SUM(npk_formula) AS total_npk_formula, 
                SUM(organik) AS total_organik
            FROM verval
            WHERE 1=1
        `;

        let params = [];

        // Mapping Provinsi ke Kabupaten
        if (provinsi) {
            if (provinsi === 'DIY') {
                query += ` AND kabupaten IN (?, ?, ?, ?, ?)`;
                params.push('SLEMAN', 'KOTA YOGYAKARTA', 'GUNUNG KIDUL', 'BANTUL', 'KULON PROGO');
            } else if (provinsi === 'JAWA TENGAH') {
                query += ` AND kabupaten NOT IN (?, ?, ?, ?, ?)`;
                params.push('SLEMAN', 'KOTA YOGYAKARTA', 'GUNUNG KIDUL', 'BANTUL', 'KULON PROGO');
            }
        }

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
        res.json(rows[0] || { total_urea: 0, total_npk: 0, total_npk_formula: 0, total_organik: 0 });

    } catch (error) {
        console.error("Error fetching sum:", error);
        res.status(500).json({ message: "Error calculating sum", error });
    }
};

exports.getPenebusanCount = async (req, res) => {
    try {
        const { provinsi, kabupaten, kecamatan, tahun } = req.query;

        let query = `
            SELECT 
                COUNT(DISTINCT CASE WHEN urea > 0 THEN nik END) AS count_urea,
                COUNT(DISTINCT CASE WHEN npk > 0 THEN nik END) AS count_npk,
                COUNT(DISTINCT CASE WHEN npk_formula > 0 THEN nik END) AS count_npk_formula, 
                COUNT(DISTINCT CASE WHEN organik > 0 THEN nik END) AS count_organik
            FROM verval
            WHERE 1=1
        `;

        let params = [];

        // Filter Provinsi (mapping kabupaten ke provinsi)
        if (provinsi) {
            if (provinsi === 'DIY') {
                query += ` AND kabupaten IN (?, ?, ?, ?, ?)`;
                params.push('SLEMAN', 'KOTA YOGYAKARTA', 'GUNUNG KIDUL', 'BANTUL', 'KULON PROGO');
            } else if (provinsi === 'JAWA TENGAH') {
                query += ` AND kabupaten NOT IN (?, ?, ?, ?, ?)`;
                params.push('SLEMAN', 'KOTA YOGYAKARTA', 'GUNUNG KIDUL', 'BANTUL', 'KULON PROGO');
            }
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

        // Filter Tahun
        if (tahun) {
            query += " AND YEAR(tanggal_tebus) = ?";
            params.push(tahun);
        }

        const [rows] = await db.execute(query, params);

        // Return hasil count dengan default 0 jika tidak ada data
        res.json(rows[0] || {
            count_urea: 0,
            count_npk: 0,
            count_npk_formula: 0,
            count_organik: 0
        });

    } catch (error) {
        console.error("Error fetching penebusan count:", error);
        res.status(500).json({
            message: "Error menghitung jumlah penebusan",
            error: error.message
        });
    }
};

// end dashboard 1

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
        const { kabupaten, kecamatan, metode_penebusan, tanggal_tebus, bulan_awal, bulan_akhir } = req.query;
        let query = "SELECT * FROM verval WHERE 1=1";
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
        if (tanggal_tebus) {
            query += " AND DATE_FORMAT(tanggal_tebus, '%Y-%m') = ?";
            params.push(tanggal_tebus);
        }
        if (bulan_awal && bulan_akhir) {
            query += " AND tanggal_tebus BETWEEN ? AND ?";
            params.push(bulan_awal, bulan_akhir);
        }

        // Ambil data dari database
        const [rows] = await db.query(query, params);

        if (rows.length === 0) {
            return res.status(404).json({ error: "Data tidak ditemukan" });
        }

        // Buat workbook dan worksheet Excel
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Data Verval');

        // 🔹 **Tentukan Header Kolom (Tanpa id & metode_penebusan)**
        worksheet.columns = [
            { header: "No", key: "no", width: 5 },
            { header: "Kabupaten", key: "kabupaten", width: 20 },
            { header: "Kecamatan", key: "kecamatan", width: 20 },
            { header: "Poktan", key: "poktan", width: 15 },
            { header: "Kode Kios", key: "kode_kios", width: 15 },
            { header: "NIK", key: "nik", width: 20 },
            { header: "Nama Petani", key: "nama_petani", width: 25 },
            { header: "Tanggal Tebus", key: "tanggal_tebus", width: 15 },
            { header: "Urea", key: "urea", width: 10 },
            { header: "NPK", key: "npk", width: 10 },
            { header: "SP36", key: "sp36", width: 10 },
            { header: "ZA", key: "za", width: 10 },
            { header: "NPK Formula", key: "npk_formula", width: 15 },
            { header: "Organik", key: "organik", width: 10 },
            { header: "Organik Cair", key: "organik_cair", width: 15 },
            { header: "Kakao", key: "kakao", width: 10 },
            { header: "Status", key: "status", width: 15 }
        ];

        // 🔹 **Tambahkan Data ke Excel (Tanpa id & metode_penebusan)**
        let totalUrea = 0, totalNpk = 0, totalSp36 = 0, totalZa = 0, totalNpkFormula = 0, totalOrganik = 0, totalOrganikCair = 0, totalKakao = 0;

        rows.forEach((row, index) => {
            worksheet.addRow({
                no: index + 1,
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
            });

            // 🔹 **Hitung Total untuk Baris Akhir**
            totalUrea += row.urea || 0;
            totalNpk += row.npk || 0;
            totalSp36 += row.sp36 || 0;
            totalZa += row.za || 0;
            totalNpkFormula += row.npk_formula || 0;
            totalOrganik += row.organik || 0;
            totalOrganikCair += row.organik_cair || 0;
            totalKakao += row.kakao || 0;
        });

        // 🔹 **Tambahkan Baris Total**
        worksheet.addRow({
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
        }).eachCell((cell, colNumber) => {
            if (colNumber > 6) { // Hanya styling bagian total pupuk
                cell.font = { bold: true };
                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFF00' } }; // Warna kuning
            }
        });

        // Set header sebelum mengirim file
        res.setHeader('Content-Disposition', 'attachment; filename="verval_data.xlsx"');
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');

        // Kirim file Excel ke response
        await workbook.xlsx.write(res);
        res.end();
    } catch (error) {
        console.error("Error exporting Excel:", error);
        res.status(500).json({ error: "Internal Server Error" });
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

exports.summaryPupuk = async (req, res) => {
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

        // Query alokasi
        const alokasiQuery = `
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
                SUM(urea) AS alokasi_urea,
                SUM(npk) AS alokasi_npk,
                SUM(npk_formula) AS alokasi_npk_formula,
                SUM(organik) AS alokasi_organik
            FROM erdkk
            ${whereClause}
            GROUP BY ${groupField}
        `;

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

exports.alokasiVsTebusan = async (req, res) => {
    try {
        const { start, length, draw, kabupaten, kecamatan, tahun } = req.query;

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
            FROM erdkk e
            LEFT JOIN verval_summary v 
                ON e.nik = v.nik
                AND e.kabupaten = v.kabupaten
                AND e.tahun = v.tahun
            LEFT JOIN tebusan_per_bulan t 
                ON e.nik = t.nik
                AND e.kabupaten = t.kabupaten
                AND e.tahun = t.tahun
            WHERE 1=1
        `;

        let countQuery = `
            SELECT COUNT(*) AS total
            FROM erdkk e
            LEFT JOIN verval_summary v 
                ON e.nik = v.nik
                AND e.kabupaten = v.kabupaten
                AND e.tahun = v.tahun
            WHERE 1=1
        `;

        let params = [];
        let countParams = [];

        // Tambahkan filter Kabupaten
        if (kabupaten) {
            query += " AND e.kabupaten = ?";
            countQuery += " AND e.kabupaten = ?";
            params.push(kabupaten);
            countParams.push(kabupaten);
        }

        if (kecamatan) {
            query += " AND e.kecamatan = ?";
            countQuery += "AND e.kecamatan = ?";
            params.push(kecamatan);
            countParams.push(kecamatan);
        }

        // Tambahkan filter Tahun
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

        // Kirim respons
        res.json({
            draw: draw ? parseInt(draw) : 1,
            recordsTotal: recordsTotal,
            recordsFiltered: recordsFiltered,
            data: data
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
        const { kabupaten, kecamatan, tahun } = req.query;

        let query = `
            SELECT 
            SUM(e.urea) AS total_urea,
            SUM(e.npk) AS total_npk,
            SUM(e.npk_formula) AS total_npk_formula,
            SUM(e.organik) AS total_organik,
            
            SUM(e.urea - COALESCE(v.tebus_urea, 0)) AS sisa_urea,
            SUM(e.npk - COALESCE(v.tebus_npk, 0)) AS sisa_npk,
            SUM(e.npk_formula - COALESCE(v.tebus_npk_formula, 0)) AS sisa_npk_formula,
            SUM(e.organik - COALESCE(v.tebus_organik, 0)) AS sisa_organik,

            SUM(COALESCE(v.tebus_urea, 0)) AS total_tebus_urea,
            SUM(COALESCE(v.tebus_npk, 0)) AS total_tebus_npk,
            SUM(COALESCE(v.tebus_npk_formula, 0)) AS total_tebus_npk_formula,
            SUM(COALESCE(v.tebus_organik, 0)) AS total_tebus_organik,

            -- Tebusan per bulan
            SUM(COALESCE(t.jan_urea, 0)) AS jan_tebus_urea,
            SUM(COALESCE(t.jan_npk, 0)) AS jan_tebus_npk,
            SUM(COALESCE(t.jan_npk_formula, 0)) AS jan_tebus_npk_formula,
            SUM(COALESCE(t.jan_organik, 0)) AS jan_tebus_organik,

            SUM(COALESCE(t.feb_urea, 0)) AS feb_tebus_urea,
            SUM(COALESCE(t.feb_npk, 0)) AS feb_tebus_npk,
            SUM(COALESCE(t.feb_npk_formula, 0)) AS feb_tebus_npk_formula,
            SUM(COALESCE(t.feb_organik, 0)) AS feb_tebus_organik,

            SUM(COALESCE(t.mar_urea, 0)) AS mar_tebus_urea,
            SUM(COALESCE(t.mar_npk, 0)) AS mar_tebus_npk,
            SUM(COALESCE(t.mar_npk_formula, 0)) AS mar_tebus_npk_formula,
            SUM(COALESCE(t.mar_organik, 0)) AS mar_tebus_organik,

            SUM(COALESCE(t.apr_urea, 0)) AS apr_tebus_urea,
            SUM(COALESCE(t.apr_npk, 0)) AS apr_tebus_npk,
            SUM(COALESCE(t.apr_npk_formula, 0)) AS apr_tebus_npk_formula,
            SUM(COALESCE(t.apr_organik, 0)) AS apr_tebus_organik,

            SUM(COALESCE(t.mei_urea, 0)) AS mei_tebus_urea,
            SUM(COALESCE(t.mei_npk, 0)) AS mei_tebus_npk,
            SUM(COALESCE(t.mei_npk_formula, 0)) AS mei_tebus_npk_formula,
            SUM(COALESCE(t.mei_organik, 0)) AS mei_tebus_organik,

            SUM(COALESCE(t.jun_urea, 0)) AS jun_tebus_urea,
            SUM(COALESCE(t.jun_npk, 0)) AS jun_tebus_npk,
            SUM(COALESCE(t.jun_npk_formula, 0)) AS jun_tebus_npk_formula,
            SUM(COALESCE(t.jun_organik, 0)) AS jun_tebus_organik,

            SUM(COALESCE(t.jul_urea, 0)) AS jul_tebus_urea,
            SUM(COALESCE(t.jul_npk, 0)) AS jul_tebus_npk,
            SUM(COALESCE(t.jul_npk_formula, 0)) AS jul_tebus_npk_formula,
            SUM(COALESCE(t.jul_organik, 0)) AS jul_tebus_organik,

            SUM(COALESCE(t.agu_urea, 0)) AS agu_tebus_urea,
            SUM(COALESCE(t.agu_npk, 0)) AS agu_tebus_npk,
            SUM(COALESCE(t.agu_npk_formula, 0)) AS agu_tebus_npk_formula,
            SUM(COALESCE(t.agu_organik, 0)) AS agu_tebus_organik,

            SUM(COALESCE(t.sep_urea, 0)) AS sep_tebus_urea,
            SUM(COALESCE(t.sep_npk, 0)) AS sep_tebus_npk,
            SUM(COALESCE(t.sep_npk_formula, 0)) AS sep_tebus_npk_formula,
            SUM(COALESCE(t.sep_organik, 0)) AS sep_tebus_organik,

            SUM(COALESCE(t.okt_urea, 0)) AS okt_tebus_urea,
            SUM(COALESCE(t.okt_npk, 0)) AS okt_tebus_npk,
            SUM(COALESCE(t.okt_npk_formula, 0)) AS okt_tebus_npk_formula,
            SUM(COALESCE(t.okt_organik, 0)) AS okt_tebus_organik,

            SUM(COALESCE(t.nov_urea, 0)) AS nov_tebus_urea,
            SUM(COALESCE(t.nov_npk, 0)) AS nov_tebus_npk,
            SUM(COALESCE(t.nov_npk_formula, 0)) AS nov_tebus_npk_formula,
            SUM(COALESCE(t.nov_organik, 0)) AS nov_tebus_organik,

            SUM(COALESCE(t.des_urea, 0)) AS des_tebus_urea,
            SUM(COALESCE(t.des_npk, 0)) AS des_tebus_npk,
            SUM(COALESCE(t.des_npk_formula, 0)) AS des_tebus_npk_formula,
            SUM(COALESCE(t.des_organik, 0)) AS des_tebus_organik

        FROM erdkk e
        LEFT JOIN verval_summary v 
            ON e.nik = v.nik 
            AND e.kabupaten = v.kabupaten
            AND e.tahun = v.tahun
            AND e.kecamatan = v.kecamatan
        LEFT JOIN tebusan_per_bulan t
            ON e.nik = t.nik
            AND e.kabupaten = t.kabupaten
            AND e.tahun = t.tahun
            AND e.kecamatan = t.kecamatan
            WHERE 1=1
        `;

        let params = [];
        if (kabupaten) {
            query += " AND e.kabupaten = ?";
            params.push(kabupaten);
        }

        if (kecamatan) {
            query += " AND e.kecamatan = ?";
            params.push(kecamatan);
        }

        if (tahun) {
            query += " AND e.tahun = ?";
            params.push(tahun);
        }

        const [summary] = await db.query(query, params);

        // Debugging: Cek hasil query sebelum dikirim
        console.log("Summary Data:", summary);

        res.json({ sum: summary[0] || {} });
    } catch (error) {
        console.error("Error fetching summary:", error);
        res.status(500).json({ error: "Terjadi kesalahan dalam mengambil summary" });
    }
};


exports.downloadPetaniSummary = async (req, res) => {
    try {
        const { kabupaten, tahun } = req.query;

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
            LEFT JOIN tebusan_per_bulan t 
                ON e.nik = t.nik
                AND e.kabupaten = t.kabupaten
                AND e.tahun = t.tahun
                AND e.kecamatan = t.kecamatan
            WHERE 1=1
        `;

        let params = [];
        if (kabupaten) {
            query += " AND e.kabupaten = ?";
            params.push(kabupaten);
        }

        if (tahun) {
            query += " AND e.tahun = ?";
            params.push(tahun);
        }

        const [data] = await db.query(query, params);

        // Buat workbook dan worksheet
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet("Summary");

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

        // **Buat Nama File Sesuai Kabupaten**
        const safeKabupaten = kabupaten ? kabupaten.replace(/\s+/g, "_") : "ALL";
        const fileName = `petani_summary_${safeKabupaten}.xlsx`;

        // Simpan dan kirim file
        res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        res.setHeader("Content-Disposition", `attachment; filename=${fileName}`);

        await workbook.xlsx.write(res);
        res.end();

    } catch (error) {
        console.error("Error generating Excel:", error);
        res.status(500).json({ error: "Gagal membuat file Excel" });
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
        const { kabupaten, tahun, kecamatan } = req.query;

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
                SUM(COALESCE(v.urea, 0)) AS tebus_urea,
    SUM(COALESCE(v.npk, 0)) AS tebus_npk,
    SUM(COALESCE(v.npk_formula, 0)) AS tebus_npk_formula,
    SUM(COALESCE(v.organik, 0)) AS tebus_organik,
                
                (e.urea - COALESCE(SUM(v.urea), 0)) AS sisa_urea,
    (e.npk - COALESCE(SUM(v.npk), 0)) AS sisa_npk,
    (e.npk_formula - COALESCE(SUM(v.npk_formula), 0)) AS sisa_npk_formula,
    (e.organik - COALESCE(SUM(v.organik), 0)) AS sisa_organik,
                
                COALESCE(SUM(CASE WHEN v.metode_penebusan = 'kartan' AND MONTH(v.tanggal_tebus) = 1 THEN v.urea ELSE 0 END), 0) AS jan_kartan_urea,
                COALESCE(SUM(CASE WHEN v.metode_penebusan = 'kartan' AND MONTH(v.tanggal_tebus) = 1 THEN v.npk ELSE 0 END), 0) AS jan_kartan_npk,
                COALESCE(SUM(CASE WHEN v.metode_penebusan = 'kartan' AND MONTH(v.tanggal_tebus) = 1 THEN v.npk_formula ELSE 0 END), 0) AS jan_kartan_npk_formula,
                COALESCE(SUM(CASE WHEN v.metode_penebusan = 'kartan' AND MONTH(v.tanggal_tebus) = 1 THEN v.organik ELSE 0 END), 0) AS jan_kartan_organik,
                COALESCE(SUM(CASE WHEN v.metode_penebusan = 'ipubers' AND MONTH(v.tanggal_tebus) = 1 THEN v.urea ELSE 0 END), 0) AS jan_ipubers_urea,
                COALESCE(SUM(CASE WHEN v.metode_penebusan = 'ipubers' AND MONTH(v.tanggal_tebus) = 1 THEN v.npk ELSE 0 END), 0) AS jan_ipubers_npk,
                COALESCE(SUM(CASE WHEN v.metode_penebusan = 'ipubers' AND MONTH(v.tanggal_tebus) = 1 THEN v.npk_formula ELSE 0 END), 0) AS jan_ipubers_npk_formula,
                COALESCE(SUM(CASE WHEN v.metode_penebusan = 'ipubers' AND MONTH(v.tanggal_tebus) = 1 THEN v.organik ELSE 0 END), 0) AS jan_ipubers_organik,

                COALESCE(SUM(CASE WHEN v.metode_penebusan = 'kartan' AND MONTH(v.tanggal_tebus) = 2 THEN v.urea ELSE 0 END), 0) AS feb_kartan_urea,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'kartan' AND MONTH(v.tanggal_tebus) = 2 THEN v.npk ELSE 0 END), 0) AS feb_kartan_npk,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'kartan' AND MONTH(v.tanggal_tebus) = 2 THEN v.npk_formula ELSE 0 END), 0) AS feb_kartan_npk_formula,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'kartan' AND MONTH(v.tanggal_tebus) = 2 THEN v.organik ELSE 0 END), 0) AS feb_kartan_organik,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'ipubers' AND MONTH(v.tanggal_tebus) = 2 THEN v.urea ELSE 0 END), 0) AS feb_ipubers_urea,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'ipubers' AND MONTH(v.tanggal_tebus) = 2 THEN v.npk ELSE 0 END), 0) AS feb_ipubers_npk,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'ipubers' AND MONTH(v.tanggal_tebus) = 2 THEN v.npk_formula ELSE 0 END), 0) AS feb_ipubers_npk_formula,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'ipubers' AND MONTH(v.tanggal_tebus) = 2 THEN v.organik ELSE 0 END), 0) AS feb_ipubers_organik,

COALESCE(SUM(CASE WHEN v.metode_penebusan = 'kartan' AND MONTH(v.tanggal_tebus) = 3 THEN v.urea ELSE 0 END), 0) AS mar_kartan_urea,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'kartan' AND MONTH(v.tanggal_tebus) = 3 THEN v.npk ELSE 0 END), 0) AS mar_kartan_npk,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'kartan' AND MONTH(v.tanggal_tebus) = 3 THEN v.npk_formula ELSE 0 END), 0) AS mar_kartan_npk_formula,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'kartan' AND MONTH(v.tanggal_tebus) = 3 THEN v.organik ELSE 0 END), 0) AS mar_kartan_organik,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'ipubers' AND MONTH(v.tanggal_tebus) = 3 THEN v.urea ELSE 0 END), 0) AS mar_ipubers_urea,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'ipubers' AND MONTH(v.tanggal_tebus) = 3 THEN v.npk ELSE 0 END), 0) AS mar_ipubers_npk,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'ipubers' AND MONTH(v.tanggal_tebus) = 3 THEN v.npk_formula ELSE 0 END), 0) AS mar_ipubers_npk_formula,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'ipubers' AND MONTH(v.tanggal_tebus) = 3 THEN v.organik ELSE 0 END), 0) AS mar_ipubers_organik,

-- Lanjutkan pola ini hingga Desember
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'kartan' AND MONTH(v.tanggal_tebus) = 4 THEN v.urea ELSE 0 END), 0) AS apr_kartan_urea,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'kartan' AND MONTH(v.tanggal_tebus) = 4 THEN v.npk ELSE 0 END), 0) AS apr_kartan_npk,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'kartan' AND MONTH(v.tanggal_tebus) = 4 THEN v.npk_formula ELSE 0 END), 0) AS apr_kartan_npk_formula,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'kartan' AND MONTH(v.tanggal_tebus) = 4 THEN v.organik ELSE 0 END), 0) AS apr_kartan_organik,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'ipubers' AND MONTH(v.tanggal_tebus) = 4 THEN v.urea ELSE 0 END), 0) AS apr_ipubers_urea,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'ipubers' AND MONTH(v.tanggal_tebus) = 4 THEN v.npk ELSE 0 END), 0) AS apr_ipubers_npk,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'ipubers' AND MONTH(v.tanggal_tebus) = 4 THEN v.npk_formula ELSE 0 END), 0) AS apr_ipubers_npk_formula,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'ipubers' AND MONTH(v.tanggal_tebus) = 4 THEN v.organik ELSE 0 END), 0) AS apr_ipubers_organik,

-- Lanjutkan pola ini hingga Desember
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'kartan' AND MONTH(v.tanggal_tebus) = 5 THEN v.urea ELSE 0 END), 0) AS mei_kartan_urea,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'kartan' AND MONTH(v.tanggal_tebus) = 5 THEN v.npk ELSE 0 END), 0) AS mei_kartan_npk,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'kartan' AND MONTH(v.tanggal_tebus) = 5 THEN v.npk_formula ELSE 0 END), 0) AS mei_kartan_npk_formula,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'kartan' AND MONTH(v.tanggal_tebus) = 5 THEN v.organik ELSE 0 END), 0) AS mei_kartan_organik,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'ipubers' AND MONTH(v.tanggal_tebus) = 5 THEN v.urea ELSE 0 END), 0) AS mei_ipubers_urea,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'ipubers' AND MONTH(v.tanggal_tebus) = 5 THEN v.npk ELSE 0 END), 0) AS mei_ipubers_npk,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'ipubers' AND MONTH(v.tanggal_tebus) = 5 THEN v.npk_formula ELSE 0 END), 0) AS mei_ipubers_npk_formula,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'ipubers' AND MONTH(v.tanggal_tebus) = 5 THEN v.organik ELSE 0 END), 0) AS mei_ipubers_organik,

-- Lanjutkan pola ini hingga Desember
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'kartan' AND MONTH(v.tanggal_tebus) = 6 THEN v.urea ELSE 0 END), 0) AS jun_kartan_urea,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'kartan' AND MONTH(v.tanggal_tebus) = 6 THEN v.npk ELSE 0 END), 0) AS jun_kartan_npk,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'kartan' AND MONTH(v.tanggal_tebus) = 6 THEN v.npk_formula ELSE 0 END), 0) AS jun_kartan_npk_formula,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'kartan' AND MONTH(v.tanggal_tebus) = 6 THEN v.organik ELSE 0 END), 0) AS jun_kartan_organik,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'ipubers' AND MONTH(v.tanggal_tebus) = 6 THEN v.urea ELSE 0 END), 0) AS jun_ipubers_urea,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'ipubers' AND MONTH(v.tanggal_tebus) = 6 THEN v.npk ELSE 0 END), 0) AS jun_ipubers_npk,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'ipubers' AND MONTH(v.tanggal_tebus) = 6 THEN v.npk_formula ELSE 0 END), 0) AS jun_ipubers_npk_formula,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'ipubers' AND MONTH(v.tanggal_tebus) = 6 THEN v.organik ELSE 0 END), 0) AS jun_ipubers_organik,

-- Lanjutkan pola ini hingga Desember
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'kartan' AND MONTH(v.tanggal_tebus) = 7 THEN v.urea ELSE 0 END), 0) AS jul_kartan_urea,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'kartan' AND MONTH(v.tanggal_tebus) = 7 THEN v.npk ELSE 0 END), 0) AS jul_kartan_npk,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'kartan' AND MONTH(v.tanggal_tebus) = 7 THEN v.npk_formula ELSE 0 END), 0) AS jul_kartan_npk_formula,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'kartan' AND MONTH(v.tanggal_tebus) = 7 THEN v.organik ELSE 0 END), 0) AS jul_kartan_organik,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'ipubers' AND MONTH(v.tanggal_tebus) = 7 THEN v.urea ELSE 0 END), 0) AS jul_ipubers_urea,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'ipubers' AND MONTH(v.tanggal_tebus) = 7 THEN v.npk ELSE 0 END), 0) AS jul_ipubers_npk,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'ipubers' AND MONTH(v.tanggal_tebus) = 7 THEN v.npk_formula ELSE 0 END), 0) AS jul_ipubers_npk_formula,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'ipubers' AND MONTH(v.tanggal_tebus) = 7 THEN v.organik ELSE 0 END), 0) AS jul_ipubers_organik,

-- Lanjutkan pola ini hingga Desember
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'kartan' AND MONTH(v.tanggal_tebus) = 8 THEN v.urea ELSE 0 END), 0) AS agu_kartan_urea,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'kartan' AND MONTH(v.tanggal_tebus) = 8 THEN v.npk ELSE 0 END), 0) AS agu_kartan_npk,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'kartan' AND MONTH(v.tanggal_tebus) = 8 THEN v.npk_formula ELSE 0 END), 0) AS agu_kartan_npk_formula,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'kartan' AND MONTH(v.tanggal_tebus) = 8 THEN v.organik ELSE 0 END), 0) AS agu_kartan_organik,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'ipubers' AND MONTH(v.tanggal_tebus) = 8 THEN v.urea ELSE 0 END), 0) AS agu_ipubers_urea,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'ipubers' AND MONTH(v.tanggal_tebus) = 8 THEN v.npk ELSE 0 END), 0) AS agu_ipubers_npk,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'ipubers' AND MONTH(v.tanggal_tebus) = 8 THEN v.npk_formula ELSE 0 END), 0) AS agu_ipubers_npk_formula,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'ipubers' AND MONTH(v.tanggal_tebus) = 8 THEN v.organik ELSE 0 END), 0) AS agu_ipubers_organik,

-- Lanjutkan pola ini hingga Desember
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'kartan' AND MONTH(v.tanggal_tebus) = 9 THEN v.urea ELSE 0 END), 0) AS sep_kartan_urea,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'kartan' AND MONTH(v.tanggal_tebus) = 9 THEN v.npk ELSE 0 END), 0) AS sep_kartan_npk,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'kartan' AND MONTH(v.tanggal_tebus) = 9 THEN v.npk_formula ELSE 0 END), 0) AS sep_kartan_npk_formula,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'kartan' AND MONTH(v.tanggal_tebus) = 9 THEN v.organik ELSE 0 END), 0) AS sep_kartan_organik,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'ipubers' AND MONTH(v.tanggal_tebus) = 9 THEN v.urea ELSE 0 END), 0) AS sep_ipubers_urea,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'ipubers' AND MONTH(v.tanggal_tebus) = 9 THEN v.npk ELSE 0 END), 0) AS sep_ipubers_npk,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'ipubers' AND MONTH(v.tanggal_tebus) = 9 THEN v.npk_formula ELSE 0 END), 0) AS sep_ipubers_npk_formula,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'ipubers' AND MONTH(v.tanggal_tebus) = 9 THEN v.organik ELSE 0 END), 0) AS sep_ipubers_organik,

-- Lanjutkan pola ini hingga Desember
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'kartan' AND MONTH(v.tanggal_tebus) = 10 THEN v.urea ELSE 0 END), 0) AS okt_kartan_urea,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'kartan' AND MONTH(v.tanggal_tebus) = 10 THEN v.npk ELSE 0 END), 0) AS okt_kartan_npk,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'kartan' AND MONTH(v.tanggal_tebus) = 10 THEN v.npk_formula ELSE 0 END), 0) AS okt_kartan_npk_formula,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'kartan' AND MONTH(v.tanggal_tebus) = 10 THEN v.organik ELSE 0 END), 0) AS okt_kartan_organik,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'ipubers' AND MONTH(v.tanggal_tebus) = 10 THEN v.urea ELSE 0 END), 0) AS okt_ipubers_urea,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'ipubers' AND MONTH(v.tanggal_tebus) = 10 THEN v.npk ELSE 0 END), 0) AS okt_ipubers_npk,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'ipubers' AND MONTH(v.tanggal_tebus) = 10 THEN v.npk_formula ELSE 0 END), 0) AS okt_ipubers_npk_formula,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'ipubers' AND MONTH(v.tanggal_tebus) = 10 THEN v.organik ELSE 0 END), 0) AS okt_ipubers_organik,

-- Lanjutkan pola ini hingga Desember
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'kartan' AND MONTH(v.tanggal_tebus) = 11 THEN v.urea ELSE 0 END), 0) AS nov_kartan_urea,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'kartan' AND MONTH(v.tanggal_tebus) = 11 THEN v.npk ELSE 0 END), 0) AS nov_kartan_npk,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'kartan' AND MONTH(v.tanggal_tebus) = 11 THEN v.npk_formula ELSE 0 END), 0) AS nov_kartan_npk_formula,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'kartan' AND MONTH(v.tanggal_tebus) = 11 THEN v.organik ELSE 0 END), 0) AS nov_kartan_organik,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'ipubers' AND MONTH(v.tanggal_tebus) = 11 THEN v.urea ELSE 0 END), 0) AS nov_ipubers_urea,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'ipubers' AND MONTH(v.tanggal_tebus) = 11 THEN v.npk ELSE 0 END), 0) AS nov_ipubers_npk,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'ipubers' AND MONTH(v.tanggal_tebus) = 11 THEN v.npk_formula ELSE 0 END), 0) AS nov_ipubers_npk_formula,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'ipubers' AND MONTH(v.tanggal_tebus) = 11 THEN v.organik ELSE 0 END), 0) AS nov_ipubers_organik,

-- Lanjutkan pola ini hingga Desember
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'kartan' AND MONTH(v.tanggal_tebus) = 12 THEN v.urea ELSE 0 END), 0) AS des_kartan_urea,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'kartan' AND MONTH(v.tanggal_tebus) = 12 THEN v.npk ELSE 0 END), 0) AS des_kartan_npk,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'kartan' AND MONTH(v.tanggal_tebus) = 12 THEN v.npk_formula ELSE 0 END), 0) AS des_kartan_npk_formula,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'kartan' AND MONTH(v.tanggal_tebus) = 12 THEN v.organik ELSE 0 END), 0) AS des_kartan_organik,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'ipubers' AND MONTH(v.tanggal_tebus) = 12 THEN v.urea ELSE 0 END), 0) AS des_ipubers_urea,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'ipubers' AND MONTH(v.tanggal_tebus) = 12 THEN v.npk ELSE 0 END), 0) AS des_ipubers_npk,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'ipubers' AND MONTH(v.tanggal_tebus) = 12 THEN v.npk_formula ELSE 0 END), 0) AS des_ipubers_npk_formula,
COALESCE(SUM(CASE WHEN v.metode_penebusan = 'ipubers' AND MONTH(v.tanggal_tebus) = 12 THEN v.organik ELSE 0 END), 0) AS des_ipubers_organik

            FROM erdkk e FORCE INDEX (idx_erdkk_nik_kabupaten_tahun)
            LEFT JOIN verval v 
                ON e.nik = v.nik
                AND e.kabupaten = v.kabupaten
                AND e.tahun = YEAR(v.tanggal_tebus)
                AND e.kecamatan = v.kecamatan
            WHERE 1=1
        `;

        let params = [];
        if (kabupaten) {
            query += " AND e.kabupaten = ?";
            params.push(kabupaten);
        }

        if (kecamatan) {
            query += " AND e.kecamatan = ?";
            params.push(kecamatan);
        }

        if (tahun) {
            query += " AND e.tahun = ?";
            params.push(tahun);
        }

        query += ` GROUP BY e.kabupaten, 
                e.kecamatan,
                e.nik, 
                e.nama_petani, 
                e.kode_kios,
                e.tahun, 
                e.urea, 
                e.npk, 
                e.npk_formula, 
                e.organik` ;

        const [data] = await db.query(query, params);

        // Buat workbook dan worksheet
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet("Summary");

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
        // **Buat Nama File Sesuai Kabupaten**
        const safeKabupaten = kabupaten ? kabupaten.replace(/\s+/g, "_") : "ALL";
        const fileName = `petani_summary_${safeKabupaten}.xlsx`;

        // Simpan dan kirim file
        res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        res.setHeader("Content-Disposition", `attachment; filename=${fileName}`);

        await workbook.xlsx.write(res);
        res.end();

    } catch (error) {
        console.error("Error generating Excel:", error);
        res.status(500).json({ error: "Gagal membuat file Excel" });
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

exports.wcmVsVerval = async (req, res) => {
    try {
        const { tahun, bulan, produk, kabupaten, status = 'ALL', start = 0, length = 10, draw = 1 } = req.query;

        const produkFilter = produk && produk !== 'ALL' ? produk.toUpperCase() : 'ALL';
        const statusFilter = status.toUpperCase();

        const params = [];
        let whereClauses = ['wcm.tahun = ?'];
        params.push(tahun);

        if (bulan && bulan !== 'ALL') {
            whereClauses.push('wcm.bulan = ?');
            params.push(bulan);
        }

        if (produk && produk !== 'ALL') {
            whereClauses.push('wcm.produk = ?');
            params.push(produkFilter);
        }

        if (kabupaten && kabupaten !== 'ALL') {
            whereClauses.push('wcm.kabupaten = ?');
            params.push(kabupaten);
        }

        const whereSQL = whereClauses.length > 0 ? ' WHERE ' + whereClauses.join(' AND ') : '';

        const baseQuery = `
            SELECT 
    wcm.provinsi,
    wcm.kabupaten, 
    wcm.kecamatan,
    wcm.kode_kios,
    wcm.nama_kios, 
    wcm.kode_distributor, 
    wcm.distributor,  
    wcm.bulan,
    wcm.produk,
    ROUND(SUM(wcm.stok_awal) * 1000, 0) AS stok_awal_wcm,
    ROUND(SUM(wcm.penebusan) * 1000, 0) AS penebusan_wcm,
    ROUND(SUM(wcm.penyaluran) * 1000, 0) AS penyaluran_wcm,
    MAX(IFNULL(verval.total_penyaluran, 0)) AS penyaluran_verval,
    ROUND(SUM(wcm.stok_akhir) * 1000, 0) AS stok_akhir_wcm,
    CASE 
        WHEN ROUND(SUM(wcm.penyaluran) * 1000, 0) = MAX(IFNULL(verval.total_penyaluran, 0)) THEN 'Sesuai'
        ELSE 'Tidak Sesuai'
    END AS status_penyaluran
FROM wcm
LEFT JOIN (
    SELECT 
        kode_kios,
        kecamatan,
        kabupaten,
        bulan,
        tahun,
        produk,
        SUM(penyaluran) AS total_penyaluran
    FROM verval_f6
    GROUP BY kode_kios, kecamatan, kabupaten, bulan, tahun, produk
) AS verval
ON wcm.kode_kios = verval.kode_kios 
AND wcm.kecamatan = verval.kecamatan
AND wcm.kabupaten = verval.kabupaten
AND wcm.bulan = verval.bulan
AND wcm.tahun = verval.tahun
AND wcm.produk = verval.produk
${whereSQL}
GROUP BY 
    wcm.provinsi,
    wcm.kabupaten, 
    wcm.kecamatan,
    wcm.kode_kios,
    wcm.nama_kios,
    wcm.kode_distributor, 
    wcm.distributor,   
    wcm.bulan,
    wcm.produk
    ORDER BY wcm.kode_kios
        `;

        // Count total records tanpa filter status
        const countQuery = `SELECT COUNT(*) AS total FROM (${baseQuery}) AS count_table`;
        const [countResult] = await db.query(countQuery, params);
        const recordsTotal = countResult[0].total;

        // Filter status jika bukan ALL
        let filteredQuery = baseQuery;
        const filteredParams = [...params];
        if (statusFilter !== 'ALL') {
            filteredQuery = `
                SELECT * FROM (${baseQuery}) AS filtered_table
                WHERE status_penyaluran = ?
            `;
            filteredParams.push(statusFilter === 'SESUAI' ? 'Sesuai' : 'Tidak Sesuai');
        }

        // Count records after filter
        const filteredCountQuery = `SELECT COUNT(*) AS total FROM (${filteredQuery}) AS count_filtered_table`;
        const [filteredCountResult] = await db.query(filteredCountQuery, filteredParams);
        const recordsFiltered = filteredCountResult[0].total;

        // Apply pagination
        const finalQuery = `
            ${filteredQuery}
            LIMIT ? OFFSET ?
        `;
        filteredParams.push(parseInt(length), parseInt(start));

        const [data] = await db.query(finalQuery, filteredParams);

        res.json({
            draw: parseInt(draw),
            recordsTotal,
            recordsFiltered,
            data
        });
    } catch (error) {
        console.error('Error in wcmVsVerval:', error);
        res.status(500).json({ message: 'Terjadi kesalahan saat memproses data.', error: error.message });
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
        let whereClauses = ['wcm.tahun = ?'];
        params.push(tahun);

        if (bulan && bulan !== 'ALL') {
            whereClauses.push('wcm.bulan = ?');
            params.push(bulan);
        }

        if (produk && produk !== 'ALL') {
            whereClauses.push('wcm.produk = ?');
            params.push(produkFilter);
        }

        if (kabupaten && kabupaten !== 'ALL') {
            whereClauses.push('wcm.kabupaten = ?');
            params.push(kabupaten);
        }

        const whereSQL = whereClauses.length > 0 ? ' WHERE ' + whereClauses.join(' AND ') : '';

        const baseQuery = `
    SELECT 
        wcm.provinsi,
        wcm.kabupaten, 
        wcm.kode_distributor, 
        wcm.distributor,   
        wcm.kecamatan,
        wcm.kode_kios,
        wcm.nama_kios,
        wcm.produk,
        wcm.bulan,
        ROUND(SUM(wcm.stok_awal) * 1000, 0) AS stok_awal_wcm,
        ROUND(SUM(wcm.penebusan) * 1000, 0) AS penebusan_wcm,
        ROUND(SUM(wcm.penyaluran) * 1000, 0) AS penyaluran_wcm,
        MAX(IFNULL(verval.total_penyaluran, 0)) AS penyaluran_verval,
        ROUND(SUM(wcm.stok_akhir) * 1000, 0) AS stok_akhir_wcm,
        CASE 
            WHEN ROUND(SUM(wcm.penyaluran) * 1000, 0) = MAX(IFNULL(verval.total_penyaluran, 0)) THEN 'Sesuai'
            ELSE 'Tidak Sesuai'
        END AS status_penyaluran
    FROM wcm
    LEFT JOIN (
        SELECT 
            kode_kios,
            kecamatan,
            kabupaten,
            bulan,
            tahun,
            produk,
            SUM(penyaluran) AS total_penyaluran
        FROM verval_f6
        GROUP BY kode_kios, kecamatan, kabupaten, bulan, tahun, produk
    ) AS verval
    ON wcm.kode_kios = verval.kode_kios 
    AND wcm.kecamatan = verval.kecamatan
    AND wcm.kabupaten = verval.kabupaten
    AND wcm.bulan = verval.bulan
    AND wcm.tahun = verval.tahun
    AND wcm.produk = verval.produk
    ${whereSQL}
    GROUP BY 
        wcm.provinsi,
        wcm.kabupaten, 
        wcm.kode_distributor, 
        wcm.distributor,   
        wcm.kecamatan,
        wcm.kode_kios,
        wcm.nama_kios,
        wcm.produk,
        wcm.bulan
    ORDER BY wcm.kabupaten, wcm.kecamatan, wcm.kode_kios
`;
        // Filter status jika bukan ALL
        let finalQuery = baseQuery;
        if (statusFilter !== 'ALL') {
            finalQuery = `
                SELECT * FROM (${baseQuery}) AS filtered_table
                WHERE status_penyaluran = ?
            `;
            params.push(statusFilter === 'SESUAI' ? 'Sesuai' : 'Tidak Sesuai');
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
        worksheet.getCell('O1').value = 'Status';
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
