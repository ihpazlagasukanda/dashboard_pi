
const db = require("../config/db");
const ExcelJS = require('exceljs');
const { Worker } = require('worker_threads');
const path = require('path');
const fs = require('fs');
const { generateReport } = require('../services/reportGenerator');

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
    SUM(CASE WHEN produk = 'kakao' THEN alokasi ELSE 0 END) AS total_npk_formula,
    SUM(CASE WHEN produk = 'organik' THEN alokasi ELSE 0 END) AS total_organik
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