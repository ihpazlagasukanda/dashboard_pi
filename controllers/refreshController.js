// File: refreshCache.js
const db = require('./config/db');
const { exec } = require('child_process');

async function refreshCache() {
    try {
        console.log('Memulai refresh cache petani_summary_cache...');
        const startTime = new Date();

        // Jalankan stored procedure
        await db.query('CALL refresh_petani_summary_cache()');

        const endTime = new Date();
        const duration = (endTime - startTime) / 1000;

        console.log(`Refresh cache selesai dalam ${duration} detik`);
        process.exit(0);
    } catch (error) {
        console.error('Gagal refresh cache:', error);
        process.exit(1);
    }
}

refreshCache();