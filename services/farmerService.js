const client = require('../redisClient');
const pool = require('../config/db');

class FarmerService {
  constructor() {
    this.CACHE_PREFIX = 'farmers:cache';
    this.CACHE_TTL = 3600; // 1 jam
  }

  // Refresh data summary
  async refreshSummary() {
    try {
      console.log('Memulai refresh data farmer summary...');
      
      // Panggil stored procedure
      await pool.query('CALL RefreshFarmerSummary()');
      
      // Hapus cache lama
      const cacheKeys = await client.keys(`${this.CACHE_PREFIX}:*`);
      if (cacheKeys.length > 0) {
        await client.del(cacheKeys);
      }
      
      console.log('Refresh data selesai');
      return { success: true, message: 'Data berhasil direfresh' };
    } catch (error) {
      console.error('Error refreshing summary:', error);
      throw error;
    }
  }

  // Get statistics
  async getStats() {
    const cacheKey = `${this.CACHE_PREFIX}:stats`;
    
    // Cek cache
    const cachedData = await client.get(cacheKey);
    if (cachedData) {
      return JSON.parse(cachedData);
    }
    
    // Query untuk statistik
    const [totalResult] = await pool.query('SELECT COUNT(*) as total FROM farmer_summary WHERE tahun = 2025');
    const [kabupatenResult] = await pool.query(`
      SELECT kabupaten, COUNT(*) as count 
      FROM farmer_summary 
      WHERE tahun = 2025 
      GROUP BY kabupaten 
      ORDER BY count DESC
    `);
    const [kiosResult] = await pool.query(`
      SELECT kode_kios, COUNT(*) as count 
      FROM farmer_summary 
      WHERE tahun = 2025 
      GROUP BY kode_kios 
      ORDER BY count DESC 
      LIMIT 10
    `);
    
    const stats = {
      total: totalResult[0].total,
      byKabupaten: kabupatenResult,
      topKios: kiosResult,
      lastUpdated: new Date().toISOString()
    };
    
    // Simpan ke cache
    await client.setEx(cacheKey, this.CACHE_TTL, JSON.stringify(stats));
    
    return stats;
  }

  // Get farmers dengan sorting
  async getFarmers(filters = {}, page = 1, limit = 20, sort = { field: 'nik', order: 'asc' }) {
    try {
      // Generate cache key berdasarkan filter dan sorting
      const cacheKey = this.generateCacheKey(filters, page, limit, sort);
      
      // Cek cache dulu
      const cachedData = await client.get(cacheKey);
      if (cachedData) {
        console.log('Cache HIT untuk:', cacheKey);
        return JSON.parse(cachedData);
      }
      
      console.log('Cache MISS untuk:', cacheKey);
      
      // Jika tidak ada di cache, query database
      let query = `SELECT * FROM farmer_summary WHERE tahun = 2025`;
      let countQuery = `SELECT COUNT(*) as total FROM farmer_summary WHERE tahun = 2025`;
      const params = [];
      const countParams = [];
      
      // Tambahkan filter
      if (filters.kabupaten) {
        query += ` AND kabupaten = ?`;
        countQuery += ` AND kabupaten = ?`;
        params.push(filters.kabupaten);
        countParams.push(filters.kabupaten);
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

  // Generate cache key dengan sorting
  generateCacheKey(filters, page, limit, sort) {
    let key = `${this.CACHE_PREFIX}:page:${page}:limit:${limit}:sort:${sort.field}:${sort.order}`;
    
    if (filters.kabupaten) key += `:kabupaten:${filters.kabupaten}`;
    if (filters.nik) key += `:nik:${filters.nik}`;
    if (filters.nama) key += `:nama:${filters.nama}`;
    if (filters.kode_kios) key += `:kios:${filters.kode_kios}`;
    
    return key;
  }

  // Dapatkan daftar kabupaten
  async getKabupatenList() {
    const cacheKey = `${this.CACHE_PREFIX}:kabupaten-list`;
    
    // Cek cache
    const cachedData = await client.get(cacheKey);
    if (cachedData) {
      return JSON.parse(cachedData);
    }
    
    // Query database
    const [rows] = await pool.query(`
      SELECT DISTINCT kabupaten 
      FROM farmer_summary 
      WHERE tahun = 2025 
      ORDER BY kabupaten
    `);
    
    const kabupatenList = rows.map(row => row.kabupaten);
    
    // Simpan ke cache
    await client.setEx(cacheKey, this.CACHE_TTL * 24, JSON.stringify(kabupatenList));
    
    return kabupatenList;
  }

  // Dapatkan semua data by kabupaten
  async getFarmersByKabupaten(kabupaten) {
    const cacheKey = `${this.CACHE_PREFIX}:kabupaten:${kabupaten}`;
    
    // Cek cache
    const cachedData = await client.get(cacheKey);
    if (cachedData) {
      return JSON.parse(cachedData);
    }
    
    // Query database
    const [rows] = await pool.query(`
      SELECT * FROM farmer_summary 
      WHERE tahun = 2025 AND kabupaten = ?
      ORDER BY nik, kode_kios
    `, [kabupaten]);
    
    // Simpan ke cache
    await client.setEx(cacheKey, this.CACHE_TTL, JSON.stringify(rows));
    
    return rows;
  }
}

module.exports = new FarmerService();