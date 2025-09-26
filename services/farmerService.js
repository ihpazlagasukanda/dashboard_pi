const client = require('../redisClient');
const pool = require('../config/db');

class FarmerService {
  constructor() {
    this.CACHE_PREFIX = 'farmers:cache';
    this.CACHE_TTL = 3600; // 1 jam
    this.CURRENT_YEAR = 2025; // Tahun default
  }

  // Set tahun yang aktif
  setActiveYear(year) {
    this.activeYear = parseInt(year) || this.CURRENT_YEAR;
    return this.activeYear;
  }

  // Get tahun aktif
  getActiveYear() {
    return this.activeYear || this.CURRENT_YEAR;
  }

  // Refresh data summary untuk semua tahun
  async refreshSummary() {
    try {
      console.log('Memulai refresh data farmer summary untuk semua tahun...');
      
      // Panggil stored procedure untuk semua tahun
      await pool.query('CALL RefreshFarmerSummary()');
      
      // Hapus cache lama untuk semua tahun
      const cacheKeys = await client.keys(`${this.CACHE_PREFIX}:*`);
      if (cacheKeys.length > 0) {
        await client.del(cacheKeys);
      }
      
      console.log('Refresh data selesai untuk semua tahun');
      return { success: true, message: 'Data berhasil direfresh untuk semua tahun' };
    } catch (error) {
      console.error('Error refreshing summary:', error);
      throw error;
    }
  }

  // Get statistics untuk tahun tertentu
  async getStats(year = null) {
    const activeYear = year || this.getActiveYear();
    const cacheKey = `${this.CACHE_PREFIX}:stats:${activeYear}`;
    
    // Cek cache
    const cachedData = await client.get(cacheKey);
    if (cachedData) {
      return JSON.parse(cachedData);
    }
    
    // Tentukan tabel berdasarkan tahun
    const tableName = this.getTableName(activeYear);
    
    // Query untuk statistik
    const [totalResult] = await pool.query(`SELECT COUNT(*) as total FROM ${tableName}`);
    const [kabupatenResult] = await pool.query(`
      SELECT kabupaten, COUNT(*) as count 
      FROM ${tableName}
      GROUP BY kabupaten 
      ORDER BY count DESC
    `);
    const [kiosResult] = await pool.query(`
      SELECT kode_kios, COUNT(*) as count 
      FROM ${tableName}
      GROUP BY kode_kios 
      ORDER BY count DESC 
      LIMIT 10
    `);
    
    const stats = {
      year: activeYear,
      total: totalResult[0].total,
      byKabupaten: kabupatenResult,
      topKios: kiosResult,
      lastUpdated: new Date().toISOString()
    };
    
    // Simpan ke cache
    await client.setEx(cacheKey, this.CACHE_TTL, JSON.stringify(stats));
    
    return stats;
  }

  // Get farmers dengan sorting untuk tahun tertentu
  async getFarmers(filters = {}, page = 1, limit = 20, sort = { field: 'nik', order: 'asc' }, year = null) {
    try {
      const activeYear = year || this.getActiveYear();
      
      // Generate cache key berdasarkan filter, sorting, dan tahun
      const cacheKey = this.generateCacheKey(filters, page, limit, sort, activeYear);
      
      // Cek cache dulu
      const cachedData = await client.get(cacheKey);
      if (cachedData) {
        console.log('Cache HIT untuk:', cacheKey);
        return JSON.parse(cachedData);
      }
      
      console.log('Cache MISS untuk:', cacheKey);
      
      // Tentukan tabel berdasarkan tahun
      const tableName = this.getTableName(activeYear);
      
      // Jika tidak ada di cache, query database
      let query = `SELECT * FROM ${tableName}`;
      let countQuery = `SELECT COUNT(*) as total FROM ${tableName}`;
      const params = [];
      const countParams = [];
      
      // Tambahkan filter
      if (filters.kabupaten) {
        query += ` WHERE kabupaten = ?`;
        countQuery += ` WHERE kabupaten = ?`;
        params.push(filters.kabupaten);
        countParams.push(filters.kabupaten);
      } else {
        query += ` WHERE 1=1`;
        countQuery += ` WHERE 1=1`;
      }
      
      if (filters.nik) {
        query += ` AND nik LIKE ?`;
        countQuery += ` AND nik LIKE ?`;
        params.push(`%${filters.nik}%`);
        countParams.push(`%${filters.nik}%`);
      }
      
      if (filters.nama) {
        query += ` AND nama_petani LIKE ?`;
        countQuery += ` AND nama_petani LIKE ?`;
        params.push(`%${filters.nama}%`);
        countParams.push(`%${filters.nama}%`);
      }
      
      if (filters.kode_kios) {
        query += ` AND kode_kios = ?`;
        countQuery += ` AND kode_kios = ?`;
        params.push(filters.kode_kios);
        countParams.push(filters.kode_kios);
      }
      
      // Tambahkan sorting
      const validSortFields = ['nik', 'nama_petani', 'kabupaten', 'kecamatan', 'kode_kios', 'total_sisa'];
      const sortField = validSortFields.includes(sort.field) ? sort.field : 'nik';
      const sortOrder = sort.order.toLowerCase() === 'desc' ? 'DESC' : 'ASC';
      
      query += ` ORDER BY ${sortField} ${sortOrder}`;
      
      // Tambahkan pagination
      query += ` LIMIT ? OFFSET ?`;
      params.push(parseInt(limit), (page - 1) * limit);
      
      // Eksekusi query
      const [rows] = await pool.query(query, params);
      const [countResult] = await pool.query(countQuery, countParams);
      const total = countResult[0].total;
      
      const result = {
        year: activeYear,
        total,
        page: parseInt(page),
        limit: parseInt(limit),
        totalPages: Math.ceil(total / limit),
        data: rows
      };
      
      // Simpan ke cache
      await client.setEx(cacheKey, this.CACHE_TTL, JSON.stringify(result));
      
      return result;
    } catch (error) {
      console.error('Error getting farmers:', error);
      throw error;
    }
  }

  // Generate cache key dengan sorting dan tahun
  generateCacheKey(filters, page, limit, sort, year) {
    let key = `${this.CACHE_PREFIX}:year:${year}:page:${page}:limit:${limit}:sort:${sort.field}:${sort.order}`;
    
    if (filters.kabupaten) key += `:kabupaten:${filters.kabupaten}`;
    if (filters.nik) key += `:nik:${filters.nik}`;
    if (filters.nama) key += `:nama:${filters.nama}`;
    if (filters.kode_kios) key += `:kios:${filters.kode_kios}`;
    
    return key;
  }

  // Dapatkan daftar kabupaten untuk tahun tertentu
  async getKabupatenList(year = null) {
    const activeYear = year || this.getActiveYear();
    const cacheKey = `${this.CACHE_PREFIX}:kabupaten-list:${activeYear}`;
    
    // Cek cache
    const cachedData = await client.get(cacheKey);
    if (cachedData) {
      return JSON.parse(cachedData);
    }
    
    // Tentukan tabel berdasarkan tahun
    const tableName = this.getTableName(activeYear);
    
    // Query database
    const [rows] = await pool.query(`
      SELECT DISTINCT kabupaten 
      FROM ${tableName}
      ORDER BY kabupaten
    `);
    
    const kabupatenList = rows.map(row => row.kabupaten);
    
    // Simpan ke cache
    await client.setEx(cacheKey, this.CACHE_TTL * 24, JSON.stringify(kabupatenList));
    
    return kabupatenList;
  }

  // Dapatkan semua data by kabupaten untuk tahun tertentu
  async getFarmersByKabupaten(kabupaten, year = null) {
    const activeYear = year || this.getActiveYear();
    const cacheKey = `${this.CACHE_PREFIX}:kabupaten:${kabupaten}:year:${activeYear}`;
    
    // Cek cache
    const cachedData = await client.get(cacheKey);
    if (cachedData) {
      return JSON.parse(cachedData);
    }
    
    // Tentukan tabel berdasarkan tahun
    const tableName = this.getTableName(activeYear);
    
    // Query database
    const [rows] = await pool.query(`
      SELECT * FROM ${tableName}
      WHERE kabupaten = ?
      ORDER BY nik, kode_kios
    `, [kabupaten]);
    
    // Simpan ke cache
    await client.setEx(cacheKey, this.CACHE_TTL, JSON.stringify(rows));
    
    return rows;
  }

  // Helper untuk mendapatkan nama tabel berdasarkan tahun
  getTableName(year) {
    if (year === 2023) return 'farmer_summary_2023';
    if (year === 2024) return 'farmer_summary_2024';
    if (year === 2025) return 'farmer_summary';
    return 'farmer_summary'; // default
  }

  // Dapatkan daftar tahun yang tersedia
  async getAvailableYears() {
    const cacheKey = `${this.CACHE_PREFIX}:available-years`;
    
    // Cek cache
    const cachedData = await client.get(cacheKey);
    if (cachedData) {
      return JSON.parse(cachedData);
    }
    
    // Cek tabel yang ada di database
    const years = [2025]; // Default
    
    try {
      // Cek apakah tabel 2023 dan 2024 ada
      const [tables] = await pool.query(`
        SELECT TABLE_NAME 
        FROM INFORMATION_SCHEMA.TABLES 
        WHERE TABLE_SCHEMA = DATABASE() 
        AND TABLE_NAME IN ('farmer_summary_2023', 'farmer_summary_2024')
      `);
      
      tables.forEach(table => {
        if (table.TABLE_NAME === 'farmer_summary_2023') years.push(2023);
        if (table.TABLE_NAME === 'farmer_summary_2024') years.push(2024);
      });
      
      // Urutkan tahun dari yang terbaru
      years.sort((a, b) => b - a);
      
      // Simpan ke cache
      await client.setEx(cacheKey, this.CACHE_TTL * 24, JSON.stringify(years));
      
      return years;
    } catch (error) {
      console.error('Error checking available years:', error);
      return years; // Return default jika error
    }
  }
}

module.exports = new FarmerService();