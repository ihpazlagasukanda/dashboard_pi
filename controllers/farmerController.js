const farmerService = require('../services/farmerService');
const csv = require('csv-stringify');
const fs = require('fs');
const path = require('path');

exports.refreshData = async (req, res) => {
  try {
    console.log('Memulai proses refresh data...');
    const result = await farmerService.refreshSummary();
    
    res.json({
      success: true,
      message: 'Data berhasil direfresh untuk semua tahun',
      timestamp: new Date().toISOString(),
      data: result
    });
  } catch (error) {
    console.error('Error refreshing data:', error);
    res.status(500).json({ 
      success: false,
      error: 'Gagal merefresh data',
      details: error.message 
    });
  }
};

exports.getFarmers = async (req, res) => {
  try {
    const { page = 1, limit = 20, kabupaten, nik, nama, kode_kios, sortBy = 'nik', sortOrder = 'asc', year } = req.query;
    
    // Set tahun aktif jika disediakan
    if (year) {
      farmerService.setActiveYear(year);
    }
    
    const filters = {};
    if (kabupaten) filters.kabupaten = kabupaten;
    if (nik) filters.nik = nik;
    if (nama) filters.nama = nama;
    if (kode_kios) filters.kode_kios = kode_kios;
    
    const sort = { field: sortBy, order: sortOrder };
    
    const result = await farmerService.getFarmers(filters, page, limit, sort, year);
    
    res.json({
      success: true,
      ...result,
      filters: filters,
      sort: sort
    });
  } catch (error) {
    console.error('Error in getFarmers:', error);
    res.status(500).json({ 
      success: false,
      error: 'Gagal mengambil data petani',
      details: error.message 
    });
  }
};

exports.getAllFarmersByKabupaten = async (req, res) => {
  try {
    const { kabupaten, year } = req.query;
    
    if (!kabupaten) {
      return res.status(400).json({ 
        success: false,
        error: 'Parameter kabupaten diperlukan' 
      });
    }
    
    // Set tahun aktif jika disediakan
    if (year) {
      farmerService.setActiveYear(year);
    }
    
    const data = await farmerService.getFarmersByKabupaten(kabupaten, year);
    
    res.json({
      success: true,
      year: farmerService.getActiveYear(),
      kabupaten: kabupaten,
      count: data.length,
      data: data
    });
  } catch (error) {
    console.error('Error in getAllFarmersByKabupaten:', error);
    res.status(500).json({ 
      success: false,
      error: 'Gagal mengambil data petani by kabupaten',
      details: error.message 
    });
  }
};

exports.getKabupatenList = async (req, res) => {
  try {
    const { year } = req.query;
    
    // Set tahun aktif jika disediakan
    if (year) {
      farmerService.setActiveYear(year);
    }
    
    const kabupatenList = await farmerService.getKabupatenList(year);
    
    res.json({
      success: true,
      year: farmerService.getActiveYear(),
      count: kabupatenList.length,
      data: kabupatenList
    });
  } catch (error) {
    console.error('Error in getKabupatenList:', error);
    res.status(500).json({ 
      success: false,
      error: 'Gagal mengambil daftar kabupaten',
      details: error.message 
    });
  }
};

exports.getStats = async (req, res) => {
  try {
    const { year } = req.query;
    
    // Set tahun aktif jika disediakan
    if (year) {
      farmerService.setActiveYear(year);
    }
    
    const stats = await farmerService.getStats(year);
    
    res.json({
      success: true,
      data: stats
    });
  } catch (error) {
    console.error('Error in getStats:', error);
    res.status(500).json({ 
      success: false,
      error: 'Gagal mengambil statistik',
      details: error.message 
    });
  }
};

exports.exportData = async (req, res) => {
  try {
    const { kabupaten, format = 'json', year } = req.query;
    
    // Set tahun aktif jika disediakan
    if (year) {
      farmerService.setActiveYear(year);
    }
    
    let data;
    if (kabupaten) {
      data = await farmerService.getFarmersByKabupaten(kabupaten, year);
    } else {
      // Untuk export semua data, kita ambil tanpa pagination
      const result = await farmerService.getFarmers({}, 1, 1000000, { field: 'nik', order: 'asc' }, year);
      data = result.data;
    }
    
    if (format === 'csv') {
      // Set header untuk download CSV
      res.setHeader('Content-Type', 'text/csv');
      res.setHeader('Content-Disposition', `attachment; filename=petani-${kabupaten || 'all'}-${farmerService.getActiveYear()}-${new Date().toISOString().split('T')[0]}.csv`);
      
      // Convert data to CSV
      const csvData = data.map(item => ({
        tahun: item.tahun || farmerService.getActiveYear(),
        kabupaten: item.kabupaten,
        kecamatan: item.kecamatan,
        nik: item.nik,
        nama_petani: item.nama_petani,
        kode_kios: item.kode_kios,
        desa: item.desa,
        poktan: item.poktan,
        sisa_urea: item.sisa_urea,
        sisa_npk: item.sisa_npk,
        sisa_npk_formula: item.sisa_npk_formula,
        sisa_organik: item.sisa_organik,
        total_sisa: item.total_sisa
      }));
      
      csv.stringify(csvData, { header: true }, (err, output) => {
        if (err) {
          throw err;
        }
        res.send(output);
      });
    } else {
      // Default JSON export
      res.json({
        success: true,
        year: farmerService.getActiveYear(),
        kabupaten: kabupaten || 'all',
        count: data.length,
        data: data
      });
    }
  } catch (error) {
    console.error('Error in exportData:', error);
    res.status(500).json({ 
      success: false,
      error: 'Gagal export data',
      details: error.message 
    });
  }
};

// Endpoint baru untuk mendapatkan daftar tahun yang tersedia
exports.getAvailableYears = async (req, res) => {
  try {
    const years = await farmerService.getAvailableYears();
    
    res.json({
      success: true,
      data: years,
      currentYear: farmerService.getActiveYear()
    });
  } catch (error) {
    console.error('Error in getAvailableYears:', error);
    res.status(500).json({ 
      success: false,
      error: 'Gagal mengambil daftar tahun',
      details: error.message 
    });
  }
};

// Endpoint untuk mengubah tahun aktif
exports.setActiveYear = async (req, res) => {
  try {
    const { year } = req.body;
    
    if (!year) {
      return res.status(400).json({ 
        success: false,
        error: 'Parameter year diperlukan' 
      });
    }
    
    const activeYear = farmerService.setActiveYear(year);
    
    res.json({
      success: true,
      message: `Tahun aktif diubah menjadi ${activeYear}`,
      activeYear: activeYear
    });
  } catch (error) {
    console.error('Error in setActiveYear:', error);
    res.status(500).json({ 
      success: false,
      error: 'Gagal mengubah tahun aktif',
      details: error.message 
    });
  }
};