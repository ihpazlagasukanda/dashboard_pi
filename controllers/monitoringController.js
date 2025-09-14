const db = require("../config/db");
const CacheUtils = require('../cacheUtils');

// Helper function untuk mapping provinsi - FIXED
const getKabupatenByProvinsi = (provinsi) => {
    if (provinsi === 'DIY') {
        return {
            type: 'IN',
            values: ['SLEMAN', 'KOTA YOGYAKARTA', 'GUNUNG KIDUL', 'BANTUL', 'KULON PROGO']
        };
    } else if (provinsi === 'JAWA TENGAH') {
        return {
            type: 'NOT IN',
            values: ['SLEMAN', 'KOTA YOGYAKARTA', 'GUNUNG KIDUL', 'BANTUL', 'KULON PROGO']
        };
    }
    return null;
};

// Generic function untuk handle query dengan caching - FIXED
const executeCachedQuery = async (req, res, queryConfig) => {
    const startTime = Date.now();
    const { prefix, queryBuilder, defaultData } = queryConfig;
    const params = { ...req.query };

    try {
        // Generate cache key
        const cacheKey = CacheUtils.generateKey(prefix, params);
        
        // Cek cache
        const cachedData = await CacheUtils.get(cacheKey);
        if (cachedData) {
            res.set({
                'X-Cache-Status': 'HIT',
                'X-Cache-Key': cacheKey,
                'X-Response-Time': `${Date.now() - startTime}ms`
            });
            return res.json(cachedData);
        }

        // Build query
        const { query, queryParams } = queryBuilder(params);
        
        // Debug log untuk melihat query dan parameters
        console.log('Executing query:', query);
        console.log('With parameters:', queryParams);
        
        // Execute query
        const [rows] = await db.execute(query, queryParams);
        const result = rows[0] || defaultData;

        // Simpan ke cache (5 menit)
        await CacheUtils.set(cacheKey, result, 300);

        res.set({
            'X-Cache-Status': 'MISS',
            'X-Cache-Key': cacheKey,
            'X-Response-Time': `${Date.now() - startTime}ms`
        });
        
        res.json(result);

    } catch (error) {
        console.error("Database Error:", error);
        res.status(500).json({ message: "Error executing query", error });
    }
};

// Query builders untuk setiap endpoint - FIXED
const queryBuilders = {
    // untuk dashboard 1 - ERDKK Summary
    erdkkSummary: (params) => {
        let query = `
            SELECT 
                SUM(urea) AS total_urea, 
                SUM(npk) AS total_npk, 
                SUM(npk_formula) AS total_npk_formula, 
                SUM(organik) AS total_organik
            FROM erdkk
            WHERE 1 = 1
        `;
        
        let queryParams = [];
        const kabupatenFilter = getKabupatenByProvinsi(params.provinsi);

        if (kabupatenFilter && kabupatenFilter.values) {
            const placeholders = kabupatenFilter.values.map(() => '?').join(',');
            query += ` AND kabupaten ${kabupatenFilter.type} (${placeholders})`;
            queryParams.push(...kabupatenFilter.values);
        }

        if (params.kabupaten) {
            query += " AND kabupaten = ?";
            queryParams.push(params.kabupaten);
        }

        if (params.kecamatan) {
            query += " AND kecamatan = ?";
            queryParams.push(params.kecamatan);
        }

        if (params.tahun) {
            query += " AND tahun = ?";
            queryParams.push(params.tahun);
        }

        return { query, queryParams };
    },

    // SK Bupati Alokasi - FIXED
    skBupatiAlokasi: (params) => {
        let query = `
            SELECT
                SUM(CASE WHEN produk = 'urea' THEN alokasi ELSE 0 END) AS total_urea,
                SUM(CASE WHEN produk = 'npk' THEN alokasi ELSE 0 END) AS total_npk,
                SUM(CASE WHEN produk = 'kakao' THEN alokasi ELSE 0 END) AS total_npk_formula,
                SUM(CASE WHEN produk = 'organik' THEN alokasi ELSE 0 END) AS total_organik
            FROM sk_bupati
            WHERE 1 = 1
        `;
        
        let queryParams = [];
        const kabupatenFilter = getKabupatenByProvinsi(params.provinsi);

        if (kabupatenFilter && kabupatenFilter.values) {
            const placeholders = kabupatenFilter.values.map(() => '?').join(',');
            query += ` AND kabupaten ${kabupatenFilter.type} (${placeholders})`;
            queryParams.push(...kabupatenFilter.values);
        }

        if (params.kabupaten) {
            query += " AND kabupaten = ?";
            queryParams.push(params.kabupaten);
        }

        if (params.kecamatan) {
            query += " AND kecamatan = ?";
            queryParams.push(params.kecamatan);
        }

        if (params.tahun) {
            query += " AND tahun = ?";
            queryParams.push(params.tahun);
        }

        return { query, queryParams };
    },

    // ERDKK Count - FIXED
    erdkkCount: (params) => {
    let query = `
        -- Hitung dari ERDKK
        SELECT 
            COUNT(DISTINCT CASE WHEN urea > 0 THEN nik END) AS count_urea,
            COUNT(DISTINCT CASE WHEN npk > 0 THEN nik END) AS count_npk,
            COUNT(DISTINCT CASE WHEN npk_formula > 0 THEN nik END) AS count_npk_formula, 
            COUNT(DISTINCT CASE WHEN organik > 0 THEN nik END) AS count_organik
        FROM erdkk
        WHERE 1 = 1
    `;
    
    let queryParams = [];
    const kabupatenFilter = getKabupatenByProvinsi(params.provinsi);

    // Filter untuk query pertama (erdkk)
    if (kabupatenFilter && kabupatenFilter.values) {
        const placeholders = kabupatenFilter.values.map(() => '?').join(',');
        query += ` AND kabupaten ${kabupatenFilter.type} (${placeholders})`;
        queryParams.push(...kabupatenFilter.values);
    }

    if (params.kabupaten) {
        query += " AND kabupaten = ?";
        queryParams.push(params.kabupaten);
    }

    if (params.kecamatan) {
        query += " AND kecamatan = ?";
        queryParams.push(params.kecamatan);
    }

    if (params.tahun) {
        query += " AND tahun = ?";
        queryParams.push(params.tahun);
    }

    query += `
        UNION ALL
        
        -- Hitung dari Verval yang TIDAK ada di ERDKK
        SELECT 
            COUNT(DISTINCT CASE WHEN v.urea > 0 THEN v.nik END) AS count_urea,
            COUNT(DISTINCT CASE WHEN v.npk > 0 THEN v.nik END) AS count_npk,
            COUNT(DISTINCT CASE WHEN v.npk_formula > 0 THEN v.nik END) AS count_npk_formula, 
            COUNT(DISTINCT CASE WHEN v.organik > 0 THEN v.nik END) AS count_organik
        FROM verval v
        LEFT JOIN erdkk e ON v.nik = e.nik 
            AND v.kabupaten = e.kabupaten 
            AND v.kecamatan = e.kecamatan
            AND YEAR(v.tanggal_tebus) = e.tahun
        WHERE e.nik IS NULL
    `;

    // Filter untuk query kedua (verval)
    if (kabupatenFilter && kabupatenFilter.values) {
        const placeholders = kabupatenFilter.values.map(() => '?').join(',');
        query += ` AND v.kabupaten ${kabupatenFilter.type} (${placeholders})`;
        queryParams.push(...kabupatenFilter.values); // Parameter sama karena filter sama
    }

    if (params.kabupaten) {
        query += " AND v.kabupaten = ?";
        queryParams.push(params.kabupaten);
    }

    if (params.kecamatan) {
        query += " AND v.kecamatan = ?";
        queryParams.push(params.kecamatan);
    }

    if (params.tahun) {
        query += " AND YEAR(v.tanggal_tebus) = ?";
        queryParams.push(params.tahun);
    }

    // Final query untuk menjumlahkan kedua hasil
    const finalQuery = `
        SELECT 
            SUM(count_urea) AS count_urea,
            SUM(count_npk) AS count_npk,
            SUM(count_npk_formula) AS count_npk_formula,
            SUM(count_organik) AS count_organik
        FROM (${query}) AS combined_results
    `;

    return { query: finalQuery, queryParams };
},

    // Summary Penebusan - FIXED
    summaryPenebusan: (params) => {
        let query = `
            SELECT 
                SUM(urea) AS total_urea, 
                SUM(npk) AS total_npk, 
                SUM(npk_formula) AS total_npk_formula, 
                SUM(organik) AS total_organik
            FROM verval
            WHERE 1=1
        `;
        
        let queryParams = [];
        const kabupatenFilter = getKabupatenByProvinsi(params.provinsi);

        if (kabupatenFilter && kabupatenFilter.values) {
            const placeholders = kabupatenFilter.values.map(() => '?').join(',');
            query += ` AND kabupaten ${kabupatenFilter.type} (${placeholders})`;
            queryParams.push(...kabupatenFilter.values);
        }

        if (params.kabupaten) {
            query += " AND kabupaten = ?";
            queryParams.push(params.kabupaten);
        }

        if (params.kecamatan) {
            query += " AND kecamatan = ?";
            queryParams.push(params.kecamatan);
        }

        if (params.tahun) {
            query += " AND YEAR(tanggal_tebus) = ?";
            queryParams.push(params.tahun);
        }

        return { query, queryParams };
    },

    // Penebusan Count - FIXED
    penebusanCount: (params) => {
        let query = `
            SELECT 
                COUNT(DISTINCT CASE WHEN urea > 0 THEN nik END) AS count_urea,
                COUNT(DISTINCT CASE WHEN npk > 0 THEN nik END) AS count_npk,
                COUNT(DISTINCT CASE WHEN npk_formula > 0 THEN nik END) AS count_npk_formula, 
                COUNT(DISTINCT CASE WHEN organik > 0 THEN nik END) AS count_organik
            FROM verval
            WHERE 1=1
        `;
        
        let queryParams = [];
        const kabupatenFilter = getKabupatenByProvinsi(params.provinsi);

        if (kabupatenFilter && kabupatenFilter.values) {
            const placeholders = kabupatenFilter.values.map(() => '?').join(',');
            query += ` AND kabupaten ${kabupatenFilter.type} (${placeholders})`;
            queryParams.push(...kabupatenFilter.values);
        }

        if (params.kabupaten) {
            query += " AND kabupaten = ?";
            queryParams.push(params.kabupaten);
        }

        if (params.kecamatan) {
            query += " AND kecamatan = ?";
            queryParams.push(params.kecamatan);
        }

        if (params.tahun) {
            query += " AND YEAR(tanggal_tebus) = ?";
            queryParams.push(params.tahun);
        }

        return { query, queryParams };
    }
};

// Default data untuk setiap endpoint
const defaultData = {
    erdkkSummary: { total_urea: 0, total_npk: 0, total_npk_formula: 0, total_organik: 0 },
    skBupatiAlokasi: { total_urea: 0, total_npk: 0, total_npk_formula: 0, total_organik: 0 },
    erdkkCount: { count_urea: 0, count_npk: 0, count_npk_formula: 0, count_organik: 0 },
    summaryPenebusan: { total_urea: 0, total_npk: 0, total_npk_formula: 0, total_organik: 0 },
    penebusanCount: { count_urea: 0, count_npk: 0, count_npk_formula: 0, count_organik: 0 }
};

// Export endpoints dengan caching
exports.getErdkkSummary = (req, res) => executeCachedQuery(req, res, {
    prefix: 'erdkk_summary',
    queryBuilder: queryBuilders.erdkkSummary,
    defaultData: defaultData.erdkkSummary
});

exports.getSkBupatiAlokasi = (req, res) => executeCachedQuery(req, res, {
    prefix: 'sk_bupati_alokasi',
    queryBuilder: queryBuilders.skBupatiAlokasi,
    defaultData: defaultData.skBupatiAlokasi
});

exports.getErdkkCount = (req, res) => executeCachedQuery(req, res, {
    prefix: 'erdkk_count',
    queryBuilder: queryBuilders.erdkkCount,
    defaultData: defaultData.erdkkCount
});

exports.getSummaryPenebusan = (req, res) => executeCachedQuery(req, res, {
    prefix: 'summary_penebusan',
    queryBuilder: queryBuilders.summaryPenebusan,
    defaultData: defaultData.summaryPenebusan
});

exports.getPenebusanCount = (req, res) => executeCachedQuery(req, res, {
    prefix: 'penebusan_count',
    queryBuilder: queryBuilders.penebusanCount,
    defaultData: defaultData.penebusanCount
});

// API untuk invalidate cache
exports.invalidateCache = async (req, res) => {
    try {
        const { pattern } = req.query;
        let clearedCount = 0;

        if (pattern) {
            clearedCount = await CacheUtils.clearByPattern(pattern);
        } else {
            // Clear semua cache dashboard
            const patterns = [
                'erdkk_summary:*',
                'sk_bupati_alokasi:*',
                'erdkk_count:*',
                'summary_penebusan:*',
                'penebusan_count:*'
            ];
            
            for (const pattern of patterns) {
                clearedCount += await CacheUtils.clearByPattern(pattern);
            }
        }

        res.json({
            message: `Cache cleared successfully`,
            clearedCount: clearedCount
        });
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
};