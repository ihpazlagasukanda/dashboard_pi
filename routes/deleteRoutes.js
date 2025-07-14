const express = require('express');
const router = express.Router();
const db = require('../config/db'); // Sesuaikan path ke file koneksi database kamu

router.post('/delete-verval', async (req, res) => {
    const { manajer, tahun, bulan, kabupaten } = req.body;

    // Validasi wajib
    if (!tahun || !bulan || !manajer) {
        return res.status(400).json({ message: 'Tahun, bulan, dan manajer wajib diisi.' });
    }

    try {
        // Ambil semua kabupaten milik manajer
        const [rows] = await db.execute(
            'SELECT kabupaten FROM mapping_manajer WHERE manajer = ?',
            [manajer]
        );

        const kabupatenList = rows.map(row => row.kabupaten);

        if (kabupatenList.length === 0) {
            return res.status(404).json({ message: `Tidak ditemukan kabupaten untuk manajer: ${manajer}` });
        }

        // Jika user menyertakan 1 kabupaten spesifik, pastikan kabupaten itu memang bagian dari manajer
        let filteredKabupaten = kabupatenList;
        if (kabupaten && kabupaten !== '') {
            if (!kabupatenList.includes(kabupaten)) {
                return res.status(400).json({ message: `Kabupaten ${kabupaten} tidak termasuk wilayah manajer ${manajer}` });
            }
            filteredKabupaten = [kabupaten];
        }

        // Buat query DELETE
        let query = `DELETE FROM verval WHERE YEAR(tanggal_tebus) = ? AND MONTH(tanggal_tebus) = ?`;
        const params = [tahun, bulan];

        const placeholders = filteredKabupaten.map(() => '?').join(', ');
        query += ` AND kabupaten IN (${placeholders})`;
        params.push(...filteredKabupaten);

        const [result] = await db.execute(query, params);

        res.json({
            message: `Berhasil menghapus ${result.affectedRows} data verval bulan ${bulan} tahun ${tahun} untuk manajer ${manajer} di kabupaten: ${filteredKabupaten.join(', ')}.`
        });
    } catch (error) {
        console.error('Delete error:', error);
        res.status(500).json({ message: 'Gagal menghapus data. Silakan coba lagi.' });
    }
});


router.post('/delete-penyalurando', async (req, res) => {
    const {  manajer, tahun, bulan, kabupaten } = req.body;

    if (!tahun || !bulan || !manajer) {
        return res.status(400).json({ message: 'Tahun, bulan, dan manajer wajib diisi.' });
    }

    try {
        // Ambil semua kabupaten milik manajer
        const [rows] = await db.execute(
            'SELECT kabupaten FROM mapping_manajer WHERE manajer = ?',
            [manajer]
        );

        const kabupatenList = rows.map(row => row.kabupaten);

        if (kabupatenList.length === 0) {
            return res.status(404).json({ message: `Tidak ditemukan kabupaten untuk manajer: ${manajer}` });
        }

        // Jika user menyertakan 1 kabupaten spesifik, validasi apakah termasuk wilayah manajer
        let filteredKabupaten = kabupatenList;
        if (kabupaten && kabupaten !== '') {
            if (!kabupatenList.includes(kabupaten)) {
                return res.status(400).json({ message: `Kabupaten ${kabupaten} tidak termasuk wilayah manajer ${manajer}` });
            }
            filteredKabupaten = [kabupaten];
        }

        // Buat query DELETE
        let query = `DELETE FROM penyaluran_do WHERE YEAR(tanggal_penyaluran) = ? AND MONTH(tanggal_penyaluran) = ?`;
        const params = [tahun, bulan];

        const placeholders = filteredKabupaten.map(() => '?').join(', ');
        query += ` AND kabupaten IN (${placeholders})`;
        params.push(...filteredKabupaten);

        const [result] = await db.execute(query, params);

        res.json({
            message: `Berhasil menghapus ${result.affectedRows} data penyaluran DO bulan ${bulan} tahun ${tahun} untuk manajer ${manajer} di kabupaten: ${filteredKabupaten.join(', ')}.`
        });
    } catch (error) {
        console.error('Delete error:', error);
        res.status(500).json({ message: 'Gagal menghapus data. Silakan coba lagi.' });
    }
});


module.exports = router;
