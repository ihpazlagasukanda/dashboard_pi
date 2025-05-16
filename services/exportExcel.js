const ExcelJS = require('exceljs');
const { Worker } = require('worker_threads');
const path = require('path');
const db = require('../config/db');
const fs = require('fs');

exports.downloadPetaniSummary = async (req, res) => {
    const { kabupaten, tahun } = req.query;

    try {
        // 1. Buat file sementara
        const exportDir = path.join(__dirname, '../temp_exports');
        if (!fs.existsSync(exportDir)) fs.mkdirSync(exportDir);

        const fileName = `petani_${kabupaten || 'all'}_${Date.now()}.xlsx`;
        const filePath = path.join(exportDir, fileName);

        // 2. Jalankan di worker thread untuk menghindari blocking
        const worker = new Worker(path.join(__dirname, './excelWorker.js'), {
            workerData: {
                kabupaten,
                tahun,
                filePath
            }
        });

        // 3. Kirim respon langsung ke client
        res.json({
            status: 'processing',
            message: 'Export sedang diproses',
            downloadUrl: `/download-export?file=${fileName}`
        });

        // 4. Handle hasil dari worker
        worker.on('message', (message) => {
            if (message.error) {
                console.error('Worker error:', message.error);
                // Catat error ke log
                fs.appendFileSync('export_errors.log', `${new Date().toISOString()} - ${message.error}\n`);
            }
        });

    } catch (error) {
        console.error('Export error:', error);
        res.status(500).json({
            error: 'Gagal memulai proses export',
            details: error.message
        });
    }
};