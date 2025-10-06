const express = require('express');
const router = express.Router();
const multer = require('multer');
const ExcelJS = require('exceljs');
const db = require('../config/db');
const fs = require('fs');
const path = require('path');

// Konfigurasi multer untuk upload file
const storage = multer.memoryStorage();
const upload = multer({ 
  storage: storage,
  limits: {
  fileSize: 20 * 1024 * 1024 // 20MB limit
},
  fileFilter: (req, file, cb) => {
    if (file.mimetype === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' || 
        file.mimetype === 'application/vnd.ms-excel' ||
        file.mimetype === 'application/octet-stream') {
      cb(null, true);
    } else {
      cb(new Error('Hanya file Excel (.xlsx) yang diizinkan'), false);
    }
  }
});

// Middleware error handling untuk multer
const handleUploadError = (err, req, res, next) => {
  if (err instanceof multer.MulterError) {
    if (err.code === 'LIMIT_FILE_SIZE') {
      return res.status(400).json({ 
        success: false,
        error: 'Ukuran file terlalu besar. Maksimal 20MB.' 
      });
    }
  } else if (err) {
    return res.status(400).json({ 
      success: false,
      error: err.message 
    });
  }
  next();
};

const normalizeNIK = (nik) => {
  if (!nik) return '';
  let nikString = nik.toString().trim();
  nikString = nikString.replace(/['"`]/g, '');
  nikString = nikString.replace(/\D/g, '');
  if (nikString) {
    nikString = "'" + nikString;
  }
  return nikString;
};

const extractNIKFromCell = (cellValue) => {
  if (!cellValue) return '';
  if (typeof cellValue === 'string') {
    return normalizeNIK(cellValue);
  } else if (typeof cellValue === 'number') {
    return cellValue.toString().replace(/\D/g, '');
  } else if (cellValue.toString) {
    return normalizeNIK(cellValue.toString());
  }
  return '';
};

// Template validation
const validateTemplate = async (worksheet) => {
  try {
    const headerRow = worksheet.getRow(1);
    const headers = [];
    
    headerRow.eachCell((cell, colNumber) => {
      if (cell.value) {
        headers.push(cell.value.toString().trim());
      }
    });

    const requiredHeaders = ['KTP', 'Nama Petani'];
    const hasRequiredHeaders = requiredHeaders.every(header => 
      headers.some(h => h.toLowerCase().includes(header.toLowerCase()))
    );

    return hasRequiredHeaders;
  } catch (error) {
    console.error('Error validasi template:', error);
    return false;
  }
};

// Ambil data kondisi dari database
const getDatabaseConditions = async (nikList) => {
  try {
    if (nikList.length === 0) return new Map();
    
    const batchSize = 1000;
    const conditionMap = new Map();
    
    for (let i = 0; i < nikList.length; i += batchSize) {
      const batch = nikList.slice(i, i + batchSize);
      const placeholders = batch.map(() => '?').join(',');
      
      const query = `
        SELECT nik, tidak_tebus_2023, tidak_tebus_2024 
        FROM petani_belum_tebus_sumatra
        WHERE nik IN (${placeholders})
      `;
      
      const [rows] = await db.execute(query, batch.map(nik => normalizeNIK(nik)));
      
      rows.forEach(row => {
        conditionMap.set(normalizeNIK(row.nik), {
          tidak_tebus_2023: row.tidak_tebus_2023,
          tidak_tebus_2024: row.tidak_tebus_2024
        });
      });
    }
    
    return conditionMap;
  } catch (error) {
    console.error('Error getting database conditions:', error);
    return new Map();
  }
};

const allowedConditions = ['2023', '2024'];

// **ENDPOINT UTAMA YANG DIPERBAIKI: POST /api/sumatra**
router.post('/', upload.single('file'), handleUploadError, async (req, res) => {
  console.log('✅ Received POST request to /api/sumatra');
  
  let tempFilePath = null;
  
  try {
    const { conditions } = req.body;
    const file = req.file;

    if (!file) {
      return res.status(400).json({ 
        success: false,
        error: 'File harus diupload' 
      });
    }

    if (!conditions) {
      return res.status(400).json({ 
        success: false,
        error: 'Kondisi harus diisi' 
      });
    }

    const conditionList = Array.isArray(conditions) ? conditions : 
                         conditions.split(',').map(c => c.trim()).filter(c => c && c !== '');

    if (conditionList.length === 0) {
      return res.status(400).json({ 
        success: false,
        error: 'Minimal satu kondisi harus dipilih' 
      });
    }

    // Validasi kondisi menggunakan variabel yang sudah didefinisikan
    const invalidConditions = conditionList.filter(cond => !allowedConditions.includes(cond));
    
    if (invalidConditions.length > 0) {
      return res.status(400).json({ 
        success: false,
        error: `Kondisi tidak valid: ${invalidConditions.join(', ')}. Hanya boleh: ${allowedConditions.join(', ')}` 
      });
    }

    const workbook = new ExcelJS.Workbook();
    
    try {
      await workbook.xlsx.load(file.buffer);
    } catch (excelError) {
      console.error('Error loading Excel file:', excelError);
      return res.status(400).json({ 
        success: false,
        error: 'File Excel tidak valid atau corrupt' 
      });
    }
    
    const worksheet = workbook.worksheets[0];
    
    if (!worksheet) {
      return res.status(400).json({ 
        success: false,
        error: 'File Excel tidak memiliki worksheet' 
      });
    }

    const isValidTemplate = await validateTemplate(worksheet);
    if (!isValidTemplate) {
      return res.status(400).json({ 
        success: false,
        error: 'Format file tidak sesuai template. Pastikan kolom KTP dan Nama Petani ada.' 
      });
    }

    // Baca header dan cari kolom KTP
    const headerRow = worksheet.getRow(1);
    const headers = [];
    headerRow.eachCell((cell, colNumber) => {
      headers[colNumber] = cell.value ? cell.value.toString().trim() : `Column${colNumber}`;
    });

    const ktpColumnIndex = headers.findIndex(header => 
      header && header.toLowerCase().includes('ktp')
    );

    if (ktpColumnIndex === -1) {
      return res.status(400).json({ 
        success: false,
        error: 'Kolom KTP tidak ditemukan dalam file' 
      });
    }

    // **PERBAIKAN 1: Simpan semua data termasuk duplicate dalam file**
    const allRowsData = [];
    const nikSetFromDatabase = new Set(); // Untuk NIK yang ada di database dan memenuhi kondisi
    
    // Kumpulkan semua data dari file Excel (TANPA menghapus duplicate dalam file)
    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return; // Skip header
      
      const rowData = {};
      let nik = '';
      
      row.eachCell((cell, colNumber) => {
        const headerName = headers[colNumber];
        rowData[headerName] = cell.value;
        
        // Extract NIK dari kolom KTP
        if (headerName && headerName.toLowerCase().includes('ktp')) {
          nik = extractNIKFromCell(cell.value);
        }
      });
      
      // Simpan data row lengkap beserta NIK dan rowNumber
      allRowsData.push({
        rowNumber,
        data: rowData,
        nik: nik
      });
    });

    if (allRowsData.length === 0) {
      return res.status(400).json({ 
        success: false,
        error: 'Tidak ada data yang valid dalam file' 
      });
    }

    console.log(`📊 Memproses ${allRowsData.length} baris data dari file`);

    // **PERBAIKAN 2: Ambil semua NIK unik dari file untuk pengecekan database**
    const uniqueNiksFromFile = [...new Set(allRowsData.map(row => row.nik).filter(nik => nik))];
    console.log(`📋 Mengecek ${uniqueNiksFromFile.length} NIK unik di database`);

    // Ambil data kondisi dari database
    const dbConditions = await getDatabaseConditions(uniqueNiksFromFile);
    console.log(`📊 Mendapatkan ${dbConditions.size} data kondisi dari database`);

    // **PERBAIKAN 3: Identifikasi NIK yang harus dihapus (berdasarkan kondisi database)**
    const niksToRemove = new Set();
    
    dbConditions.forEach((condition, nik) => {
      const meetsCondition = conditionList.some(cond => {
        if (cond === '2023') return condition.tidak_tebus_2023 === 1;
        if (cond === '2024') return condition.tidak_tebus_2024 === 1;
        return false;
      });
      
      if (meetsCondition) {
        niksToRemove.add(nik);
      }
    });

    console.log(`🗑️  Akan menghapus ${niksToRemove.size} NIK yang memenuhi kondisi`);

    // **PERBAIKAN 4: Filter data - hapus hanya jika NIK ada di database dan memenuhi kondisi**
    const filteredRows = allRowsData.filter(row => {
      if (!row.nik) return true; // Pertahankan baris tanpa NIK
      
      const dbRecord = dbConditions.get(row.nik);
      if (!dbRecord) return true; // Pertahankan jika tidak ada di database
      
      // Hapus hanya jika memenuhi kondisi
      const shouldRemove = conditionList.some(cond => {
        if (cond === '2023') return dbRecord.tidak_tebus_2023 === 1;
        if (cond === '2024') return dbRecord.tidak_tebus_2024 === 1;
        return false;
      });
      
      return !shouldRemove; // false = hapus, true = pertahankan
    });

    // **PERBAIKAN 5: Buat workbook hasil dengan struktur persis seperti file asli**
    const resultWorkbook = new ExcelJS.Workbook();
    const resultWorksheet = resultWorkbook.addWorksheet('Hasil Filter');

    // **PERBAIKAN 6: Copy header persis seperti file asli (tanpa tambahan kolom)**
    const originalHeaders = headers.filter(header => header !== undefined);
    resultWorksheet.addRow(originalHeaders);

    // **PERBAIKAN 7: Copy data yang difilter dengan struktur persis sama**
    filteredRows.forEach(row => {
      const newRow = originalHeaders.map(header => {
        // Handle khusus untuk kolom KTP - pertahankan format asli
        if (header.toLowerCase().includes('ktp') && row.data[header]) {
          const originalValue = row.data[header];
          // Jika original value adalah string yang diapit tanda petik, pertahankan
          if (typeof originalValue === 'string' && originalValue.startsWith("'")) {
            return originalValue;
          }
          // Jika angka, tambahkan tanda petik untuk menjaga format
          return `${extractNIKFromCell(originalValue)}`;
        }
        return row.data[header] || '';
      });
      resultWorksheet.addRow(newRow);
    });

    // **PERBAIKAN 8: Pertahankan lebar kolom seperti aslinya**
    worksheet.columns.forEach((col, index) => {
      if (col.width) {
        resultWorksheet.getColumn(index + 1).width = col.width;
      }
    });

    // Styling header untuk hasil
    if (resultWorksheet.getRow(1)) {
      resultWorksheet.getRow(1).font = { bold: true };
      resultWorksheet.getRow(1).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFE6E6FA' }
      };
    }

    // Statistik proses
    const stats = {
      total_data_file: allRowsData.length,
      data_dipertahankan: filteredRows.length,
      data_dihapus: allRowsData.length - filteredRows.length,
      unique_nik_dihapus: niksToRemove.size,
      kondisi_diterapkan: conditionList
    };

    console.log('✅ Proses filter selesai dengan statistik:', stats);

    // Simpan file hasil
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    const filename = `hasil_filter_${timestamp}.xlsx`;
    const tempDir = path.join(__dirname, '../temp-sumatra');
    
    if (!fs.existsSync(tempDir)) {
      fs.mkdirSync(tempDir, { recursive: true });
    }
    
    tempFilePath = path.join(tempDir, filename);
    await resultWorkbook.xlsx.writeFile(tempFilePath);

    res.json({
      success: true,
      message: `Proses filter selesai. ${stats.data_dipertahankan} data dipertahankan, ${stats.data_dihapus} data dihapus.`,
      stats: stats,
      download_url: `/api/sumatra/download/${filename}`
    });

  } catch (error) {
    console.error('❌ Error dalam proses filter:', error);
    
    if (tempFilePath && fs.existsSync(tempFilePath)) {
      try {
        fs.unlinkSync(tempFilePath);
      } catch (cleanupError) {
        console.error('Error cleanup:', cleanupError);
      }
    }
    
    res.status(500).json({ 
      success: false,
      error: 'Terjadi kesalahan server: ' + error.message 
    });
  }
});

// **ENDPOINT DOWNLOAD: GET /api/sumatra/download/:filename**
router.get('/download/:filename', async (req, res) => {
  try {
    const filename = req.params.filename;
    
    if (!filename.match(/^[a-zA-Z0-9_.-]+$/)) {
      return res.status(400).json({
        success: false,
        error: 'Nama file tidak valid'
      });
    }

    const filepath = path.join(__dirname, '../temp-sumatra', filename);
    
    if (!fs.existsSync(filepath)) {
      return res.status(404).json({
        success: false,
        error: 'File tidak ditemukan'
      });
    }

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="hasil_filter_petani.xlsx"`);
    
    const fileStream = fs.createReadStream(filepath);
    fileStream.pipe(res);
    
    fileStream.on('end', () => {
      setTimeout(() => {
        try {
          if (fs.existsSync(filepath)) {
            fs.unlinkSync(filepath);
          }
        } catch (cleanupError) {
          console.error('Cleanup error:', cleanupError);
        }
      }, 5000);
    });
    
    fileStream.on('error', (error) => {
      console.error('File stream error:', error);
      res.status(500).json({
        success: false,
        error: 'Error streaming file'
      });
    });

  } catch (error) {
    console.error('Download error:', error);
    res.status(500).json({
      success: false,
      error: 'Error downloading file: ' + error.message
    });
  }
});

// **ENDPOINT TEMPLATE** (diperbaiki agar kolom pertama "Nama Penyuluh")
router.get('/template', async (req, res) => {
  try {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Template');

    // **PERBAIKAN: Langsung mulai dengan "Nama Penyuluh" sebagai kolom pertama**
    const headers = [
      'Nama Penyuluh', 'Kode Desa', 'Kode Kios Pengecer', 'Nama Kios Pengecer', 
      'Gapoktan', 'Nama Poktan', 'Nama Petani', 'KTP', 'Tempat Lahir', 
      'Tanggal Lahir', 'Nama Ibu Kandung', 'Alamat', 'Subsektor',
      'Komoditas MT1', 'Luas Lahan (Ha) MT1', 'Pupuk Urea (Kg) MT1', 
      'Pupuk NPK (Kg) MT1', 'Pupuk NPK Formula (Kg) MT1', 'Pupuk Organik (Kg) MT1',
      'Komoditas MT2', 'Luas Lahan (Ha) MT2', 'Pupuk Urea (Kg) MT2', 
      'Pupuk NPK (Kg) MT2', 'Pupuk NPK Formula (Kg) MT2', 'Pupuk Organik (Kg) MT2',
      'Komoditas MT3', 'Luas Lahan (Ha) MT3', 'Pupuk Urea (Kg) MT3', 
      'Pupuk NPK (Kg) MT3', 'Pupuk NPK Formula (Kg) MT3', 'Pupuk Organik (Kg) MT3'
    ];

    worksheet.addRow(headers);

    const exampleRow = [
      'Contoh Penyuluh', 'DESA001', 'KIOS001', 'Contoh Kios',
      'Gapoktan Contoh', 'Poktan Contoh', 'Contoh Petani', 
      "'1234567890123456", // Format KTP dengan tanda petik
      'Jakarta', '1990-01-01', 'Ibu Contoh', 'Alamat Contoh', 'Padi',
      'Padi', '1.0', '100', '50', '25', '10',
      'Jagung', '0.5', '50', '25', '15', '5',
      'Kedelai', '0.3', '30', '15', '10', '3'
    ];
    
    worksheet.addRow(exampleRow);

    // Styling
    const headerRow = worksheet.getRow(1);
    headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    headerRow.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF4F81BD' }
    };

    const exampleRowStyle = worksheet.getRow(2);
    exampleRowStyle.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFF0F0F0' }
    };

    // Catatan
    worksheet.addRow([]);
    worksheet.addRow(['Catatan:']);
    worksheet.addRow(['- Kolom KTP harus diisi dengan format: \'1234567890123456 (dengan tanda petik tunggal di awal)']);
    worksheet.addRow(['- Duplicate NIK dalam file akan dipertahankan selama tidak memenuhi kondisi filter']);
    worksheet.addRow(['- Data hanya dihapus jika NIK sama dengan yang di database DAN memenuhi kondisi yang dipilih']);

    // Set column widths
    worksheet.columns = headers.map(() => ({ width: 20 }));

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=template_filter_petani.xlsx');

    await workbook.xlsx.write(res);
    res.end();

  } catch (error) {
    console.error('Error generating template:', error);
    res.status(500).json({
      success: false,
      error: 'Error generating template: ' + error.message
    });
  }
});

// **ENDPOINT HEALTH CHECK**
router.get('/health', (req, res) => {
  res.json({ 
    success: true, 
    message: 'Filter petani API is running',
    timestamp: new Date().toISOString()
  });
});

module.exports = router;