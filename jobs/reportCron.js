const cron = require('node-cron');
const path = require('path');
const { generateAllReports, generateReport } = require('../services/reportGenerator');

const baseDir = path.join(__dirname, '../temp_exports');

// ========== MANUAL RUN ==========
// (async () => {
//     const kabupaten = 'BOYOLALI';
//     const tahun = 2025;

//     try {
//         console.log(`[MANUAL] Generate report untuk ${kabupaten} - ${tahun}`);
//         await generateReport(kabupaten, tahun, baseDir);
//         console.log('[MANUAL] Report selesai dibuat.');
//     } catch (error) {
//         console.error('[MANUAL ERROR] Gagal generate report:', error);
//     }
// })();

// ========== CRON JOB ==========
console.log('[INIT] Cron job siap...');

cron.schedule('0 2 * * *', async () => {
    console.log('[CRON] Menjalankan generateAllReports jam 2 pagi...');
    try {
        await generateAllReports();
        console.log('[CRON] Semua laporan berhasil dibuat.');
    } catch (error) {
        console.error('[CRON ERROR] Gagal generate laporan:', error);
    }
});
