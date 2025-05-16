const cron = require('node-cron');
const path = require('path');
const { generateAllReports } = require('../services/reportGenerator');

const baseDir = path.join(__dirname, '../temp_exports');

console.log('[INIT] Cron job siap...');

// CRON: Jalankan setiap hari jam 2 pagi
cron.schedule('0 2 * * *', async () => {
    console.log('[CRON] Menjalankan generateAllReports jam 2 pagi...');
    try {
        await generateAllReports();
        console.log('[CRON] Semua laporan berhasil dibuat.');
    } catch (error) {
        console.error('[CRON] Gagal generate laporan:', error);
    }
});
