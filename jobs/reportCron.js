const cron = require('node-cron');
const path = require('path');
const { generateAllReports, generateReport } = require('../services/reportGenerator');

console.log('[INIT] Cron job siap...');
const baseDir = path.join(__dirname, '../temp_exports');

(async () => {
    const tahun = 2025;
    const kabupatenList = [
        'KOTA YOGYAKARTA',
        'SLEMAN',
        'BANTUL',
        'GUNUNG KIDUL',
        'KULON PROGO',
        'SRAGEN',
        // 'BOYOLALI',
        'KLATEN',
        'WONOGIRI',
        'KARANGANYAR',
        'KOTA SURAKARTA',
        'SUKOHARJO'
    ];

    for (const kabupaten of kabupatenList) {
        try {
            console.log(`[MANUAL] Mulai generate report untuk ${kabupaten} - ${tahun}`);
            await generateReport(kabupaten, tahun, baseDir);
            console.log(`[MANUAL] Selesai generate report untuk ${kabupaten}`);
        } catch (error) {
            console.error(`[MANUAL ERROR] Gagal generate report untuk ${kabupaten}:`, error);
        }


        await new Promise(resolve => setTimeout(resolve, 1000));
    }

    console.log('[MANUAL] Semua laporan selesai diproses.');
})();

// (async () => {
//     const tahun = 2025;
//     const kabupatenList = [
//         'BOYOLALI'
//     ];

//     for (const kabupaten of kabupatenList) {
//         try {
//             console.log(`[MANUAL] Mulai generate report untuk ${kabupaten} - ${tahun}`);
//             await generateReport(kabupaten, tahun, baseDir);
//             console.log(`[MANUAL] Selesai generate report untuk ${kabupaten}`);
//         } catch (error) {
//             console.error(`[MANUAL ERROR] Gagal generate report untuk ${kabupaten}:`, error);
//         }


//         await new Promise(resolve => setTimeout(resolve, 1000));
//     }

//     console.log('[MANUAL] Semua laporan selesai diproses.');
// })();


// ========== CRON JOB ==========

// Cron: jam 02:00 → Grup 0
cron.schedule('0 1 * * *', async () => {
  console.log('[CRON] Menjalankan Grup 0 - jam 02:00');
  await runGroup(0);
});

// Cron: jam 02:15 → Grup 1
cron.schedule('30 1 * * *', async () => {
  console.log('[CRON] Menjalankan Grup 1 - jam 02:15');
  await runGroup(1);
});

// Cron: jam 02:30 → Grup 2
cron.schedule('0 2 * * *', async () => {
  console.log('[CRON] Menjalankan Grup 2 - jam 02:30');
  await runGroup(2);
});

// Cron: jam 02:45 → Grup 3
cron.schedule('30 2 * * *', async () => {
  console.log('[CRON] Menjalankan Grup 3 - jam 02:45');
  await runGroup(3);
});

// Cron: jam 03:15 → ALL
// cron.schedule('0 3 * * *', async () => {
//   console.log('[CRON] Menjalankan Report ALL - jam 03:15');
//   await runGroup('all');
// });

async function runGroup(group) {
  try {
    await generateAllReports(group, baseDir);
    console.log(`[CRON] Laporan Grup ${group} selesai`);
  } catch (error) {
    console.error(`[CRON ERROR] Grup ${group}:`, error);
  }
}

// console.log('[INIT] Cron job siap...');

// cron.schedule('0 2 * * *', async () => {
//     console.log('[CRON] Menjalankan generateAllReports jam 2 pagi...');
//     try {
//         await generateAllReports();
//         console.log('[CRON] Semua laporan berhasil dibuat.');
//     } catch (error) {
//         console.error('[CRON ERROR] Gagal generate laporan:', error);
//     }
// });
