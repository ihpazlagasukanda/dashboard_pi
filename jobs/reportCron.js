const cron = require('node-cron');
const path = require('path');
const { generateAllReports, generateReport } = require('../services/reportGenerator');

const baseDir = path.join(__dirname, '../temp_exports');

// (async () => {
//     const tahun = 2025;
//     const kabupatenList = [
//         'SLEMAN'
//     ];

//     for (const kabupaten of kabupatenList) {
//         try {
//             console.log(`[MANUAL] Mulai generate report untuk ${kabupaten} - ${tahun}`);
//             await generateReport(kabupaten, tahun, baseDir);
//             console.log(`[MANUAL] Selesai generate report untuk ${kabupaten}`);
//         } catch (error) {
//             console.error(`[MANUAL ERROR] Gagal generate report untuk ${kabupaten}:`, error);
//         }

//         // Opsional: Delay kecil untuk bantu pelepasan memori antar proses
//         await new Promise(resolve => setTimeout(resolve, 1000)); // 1 detik
//     }

//     console.log('[MANUAL] Semua laporan selesai diproses.');
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
