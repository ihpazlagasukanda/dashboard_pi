const express = require('express');
const router = express.Router();
const db = require('../config/db'); // Sesuaikan path ke file koneksi database kamu

router.post('/delete-verval', async (req, res) => {
    const { tahun, bulan, kabupaten } = req.body;

    if (!tahun || !bulan) {
        return res.status(400).json({ message: 'Tahun dan bulan wajib diisi.' });
    }

    try {
        let query = 'DELETE FROM verval WHERE YEAR(tanggal_tebus) = ? AND MONTH(tanggal_tebus) = ?';
        const params = [tahun, bulan];

        if (kabupaten && kabupaten !== '') {
            query += ' AND kabupaten = ?';
            params.push(kabupaten);
        }

        const [result] = await db.execute(query, params);

        res.json({
            message: `Berhasil menghapus ${result.affectedRows} data verval untuk bulan ${bulan} tahun ${tahun}${kabupaten ? ' di ' + kabupaten : ''}.`
        });
    } catch (error) {
        console.error('Delete error:', error);
        res.status(500).json({ message: 'Gagal menghapus data. Silakan coba lagi.' });
    }
});

router.post('/delete-penyalurando', async (req, res) => {
    const { tahun, bulan, kabupaten } = req.body;

    if (!tahun || !bulan) {
        return res.status(400).json({ message: 'Tahun dan bulan wajib diisi.' });
    }

    try {
        let query = 'DELETE FROM penyaluran_do WHERE YEAR(tanggal_penyaluran) = ? AND MONTH(tanggal_penyaluran) = ?';
        const params = [tahun, bulan];

        if (kabupaten && kabupaten !== '') {
            query += ' AND kabupaten = ?';
            params.push(kabupaten);
        }

        const [result] = await db.execute(query, params);

        res.json({
            message: `Berhasil menghapus ${result.affectedRows} data penyaluran do untuk bulan ${bulan} tahun ${tahun}${kabupaten ? ' di ' + kabupaten : ''}.`
        });
    } catch (error) {
        console.error('Delete error:', error);
        res.status(500).json({ message: 'Gagal menghapus data. Silakan coba lagi.' });
    }
});

module.exports = router;
