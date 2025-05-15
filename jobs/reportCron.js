// file: jobs/reportCron.js
const cron = require('node-cron');
const { generateAllReports } = require('../services/reportGenerator');

// Jalankan setiap hari jam 2 pagi
cron.schedule('0 2 * * *', () => {
    console.log('Memulai generate laporan harian...');
    generateAllReports();
});

console.log('Cron job untuk generate laporan telah dijadwalkan');