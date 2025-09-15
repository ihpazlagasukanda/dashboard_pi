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
      message: 'Data berhasil direfresh',
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
    const { page = 1, limit = 20, kabupaten, nik, nama, kode_kios, sortBy = 'nik', sortOrder = 'asc' } = req.query;
    
    const filters = {};
    if (kabupaten) filters.kabupaten = kabupaten;
    if (nik) filters.nik = nik;
    if (nama) filters.nama = nama;
    if (kode_kios) filters.kode_kios = kode_kios;
    
    const sort = { field: sortBy, order: sortOrder };
    
    const result = await farmerService.getFarmers(filters, page, limit, sort);
    
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
    const { kabupaten } = req.query;
    
    if (!kabupaten) {
      return res.status(400).json({ 
        success: false,
        error: 'Parameter kabupaten diperlukan' 
      });
    }
    
    const data = await farmerService.getFarmersByKabupaten(kabupaten);
    
    res.json({
      success: true,
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
    const kabupatenList = await farmerService.getKabupatenList();
    
    res.json({
      success: true,
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
    const stats = await farmerService.getStats();
    
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
    const { kabupaten, format = 'json' } = req.query;
    
    let data;
    if (kabupaten) {
      data = await farmerService.getFarmersByKabupaten(kabupaten);
    } else {
      // Untuk export semua data, kita ambil tanpa pagination
      const result = await farmerService.getFarmers({}, 1, 1000000);
      data = result.data;
    }
    
    if (format === 'csv') {
      // Set header untuk download CSV
      res.setHeader('Content-Type', 'text/csv');
      res.setHeader('Content-Disposition', `attachment; filename=petani-${kabupaten || 'all'}-${new Date().toISOString().split('T')[0]}.csv`);
      
      // Convert data to CSV
      const csvData = data.map(item => ({
        kabupaten: item.kabupaten,
        kecamatan: item.kecamatan,
        nik: item.nik,
        nama_petani: item.nama_petani,
        kode_kios: item.kode_kios,
        desa: item.desa,
        poktan: item.poktan,
        tahun: item.tahun,
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