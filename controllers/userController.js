const db = require("../config/db");
const ExcelJS = require('exceljs');
const { Worker } = require('worker_threads');
const bcrypt = require('bcrypt'); // Pastikan sudah di-import di atas
const path = require('path');
const fs = require('fs');
// GET /admin/manage-users

exports.getAllUser = async (req, res) => {
    try {
        const { start, length, draw, search = '', level = '', kabupaten = '', status = '' } = req.query;
        let query = `SELECT * FROM admin WHERE 1=1`;
        let countQuery = "SELECT COUNT(*) FROM admin WHERE 1=1";
        let params = [];
        let countParams = [];

        if (search) {
          query += ` AND username LIKE ?`;
          params.push(`%${search}%`);
          countQuery += ` AND username LIKE ?`;
          countParams.push(`%${search}%`);
        }

        if (level) {
          query += ` AND level = ?`;
          params.push(level);
          countQuery += ` AND level = ?`;
          countParams.push(level);
        }

        if (kabupaten) {
            query += " AND kabupaten = ?";
            countQuery += " AND kabupaten = ?";
            params.push(kabupaten);
            countParams.push(kabupaten);
        }

        // Pagination (gunakan 0 jika null)
        let startNum = parseInt(start) || 0;
        let lengthNum = parseInt(length) || 10;
        query += " LIMIT ?, ?";
        params.push(startNum, lengthNum);

        // Eksekusi query
        const [data] = await db.query(query, params);
        const [[{ total: recordsTotal }]] = await db.query("SELECT COUNT(*) AS total FROM admin");
        const [[{ total: recordsFiltered }]] = await db.query(countQuery, countParams);

        res.json({
            draw: draw ? parseInt(draw) : 1,
            recordsTotal: recordsTotal,
            recordsFiltered: recordsFiltered,
            data: data
        });
    } catch (error) {
        console.error(error);
        res.status(500).json({ error: 'Terjadi kesalahan dalam mengambil data' });
    }
};

// POST /admin/manage-users

exports.createUser = async (req, res) => {
  try {
    const { username, password, level, kabupaten } = req.body;

    if (!username || !password) {
      return res.status(400).send('Username dan password wajib diisi');
    }

    // 🔐 Hash password sebelum disimpan
    const hashedPassword = await bcrypt.hash(password, 10); // 10 adalah saltRounds

    await db.execute(
      `INSERT INTO admin (username, password, level, kabupaten) VALUES (?, ?, ?, ?)`,
      [username, hashedPassword, level || 1, kabupaten || null]
    );

    res.redirect('/admin/manage-users');
  } catch (err) {
    console.error('❌ Error createAdmin:', err);
    res.status(500).send('Gagal tambah admin');
  }
};


// PUT /admin/manage-users/:id
exports.updateUser = async (req, res) => {
  try {
    const { id } = req.params;
    const { username, level, status, kabupaten } = req.body;
    let akses = req.body.akses;

    // 🔒 Kunci: Blokir update jika ID adalah 1 (super admin)
    if (id === '1' || id === 1) {
      return res.status(403).send('❌ Super Admin tidak boleh diubah.');
    }

    // Pastikan akses selalu dalam bentuk array
    if (!akses) {
      akses = [];
    } else if (!Array.isArray(akses)) {
      akses = [akses]; // jika hanya satu akses dipilih
    }

    const aksesString = JSON.stringify(akses);

    await db.execute(
      `UPDATE admin SET username = ?, level = ?, status = ?, kabupaten = ?, akses = ? WHERE id = ?`,
      [username, level, status, kabupaten, aksesString, id]
    );

    res.sendStatus(200);
  } catch (err) {
    console.error('❌ Error updateUser:', err);
    res.status(500).send('Gagal update user');
  }
};



exports.updateAksesUser = async (req, res) => {
  try {
    const { id } = req.params;
    let { akses } = req.body;

    // 🔒 Kunci: Blokir update jika ID adalah 1 (super admin)
    if (id === '1' || id === 1) {
      return res.status(403).send('❌ Super Admin tidak boleh diubah.');
    }

    // Normalisasi akses
    if (!akses) {
      akses = [];
    } else if (!Array.isArray(akses)) {
      akses = [akses];
    }

    const aksesString = JSON.stringify(akses);

    await db.execute(`UPDATE admin SET akses = ? WHERE id = ?`, [aksesString, id]);

    res.sendStatus(200);
  } catch (err) {
    console.error('❌ Error updateAksesUser:', err);
    res.status(500).send('Gagal memperbarui akses');
  }
};



// GET /admin/manage-users/:id
exports.getUserById = async (req, res) => {
  try {
    const { id } = req.params;
    const [rows] = await db.execute('SELECT * FROM admin WHERE id = ?', [id]);


    if (rows.length === 0) {
      return res.status(404).send('User tidak ditemukan');
    }

    res.json(rows[0]);
  } catch (err) {
    console.error('❌ Error getUserById:', err);
    res.status(500).send('Gagal mengambil data user');
  }
};



// DELETE /admin/manage-users/:id
exports.deleteUser = async (req, res) => {
  try {
    const { id } = req.params;

    if (id === '1' || id === 1) {
      return res.status(403).send('❌ Super Admin tidak bisa dihapus.');
    }

    await db.execute(`DELETE FROM admin WHERE id = ?`, [id]);

    res.sendStatus(200);
  } catch (err) {
    console.error('❌ Error deleteAdmin:', err);
    res.status(500).send('Gagal hapus admin');
  }
};
