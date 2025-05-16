// file: jobs/reportCron.js
const cron = require('node-cron');
const path = require('path');
const { generateAllReports } = require('../services/reportGenerator');

const baseDir = path.join(__dirname, '../temp_exports');

console.log('[INIT] Memulai script cron report...');

// langsung jalankan manual saat start (optional)
(async () => {
    try {
        console.log('[MANUAL] Generate all reports manual...');
        await generateAllReports();
        console.log('[MANUAL] Semua laporan selesai dibuat.');
    } catch (err) {
        console.error('[MANUAL] Gagal generate laporan:', err);
    }
})();

// jadwalkan otomatis tiap jam 2 pagi
cron.schedule('0 2 * * *', async () => {
    console.log('[CRON] Cronjob mulai generate all reports jam 2 pagi...');
    try {
        await generateAllReports();
        console.log('[CRON] Semua laporan berhasil dibuat.');
    } catch (err) {
        console.error('[CRON] Gagal generate laporan:', err);
    }
});
