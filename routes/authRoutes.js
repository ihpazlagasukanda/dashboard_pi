const express = require('express');
const bcrypt = require('bcrypt');
const jwt = require('jsonwebtoken');
const User = require('../models/userModel');
const db = require('../config/db');
require('dotenv').config();


const router = express.Router();

// Middleware untuk mengecek login
function authMiddleware(req, res, next) {
    const token = req.cookies.token;
    if (!token) {
        return res.redirect('/login');
    }

    try {
        const decoded = jwt.verify(token, process.env.JWT_SECRET);
        req.user = decoded; // Simpan user di req
        next();
    } catch (error) {
        res.clearCookie('token');
        return res.redirect('/login');
    }
}

// Halaman login
router.get('/login', (req, res) => {
    res.render('login');
});

// Proses login
router.post('/login', async (req, res) => {
    try {
        const { username, password } = req.body;
        console.log("🔍 Login request:", username);

        const userAgent = req.headers['user-agent'];
        const forwarded = req.headers['x-forwarded-for'];
         const ipAddress = (req.headers['x-forwarded-for'] || req.socket.remoteAddress || '').split(',')[0].trim();

        const user = await User.findByUsername(username);
        if (!user) {
            return res.status(401).send("<script>alert('User tidak ditemukan'); window.location='/login';</script>");
        }

        const isMatch = await bcrypt.compare(password, user.password);
        if (!isMatch) {
            return res.status(401).send("<script>alert('Password salah'); window.location='/login';</script>");
        }

        // ✅ Tambahkan parsing akses (dari DB bentuknya string JSON)
        let aksesArray = [];
        try {
            aksesArray = JSON.parse(user.akses || '[]');
        } catch (e) {
            console.warn('❗ Gagal parse akses:', user.akses);
        }

        
        await db.execute(
        `INSERT INTO log_login (user_id, username, ip_address, user_agent) VALUES (?, ?, ?, ?)`,
        [user.id, user.username, ipAddress, userAgent]
        );

        const tokenPayload = {
            id: user.id,
            username: user.username,
            level: user.level,
            kabupaten: user.kabupaten,
            akses: aksesArray, // ✅ ini penting
        };

        const token = jwt.sign(tokenPayload, process.env.JWT_SECRET, { expiresIn: '1h' });

        console.log("🔑 Token payload:", tokenPayload); // Debug log

        res.cookie('token', token, { httpOnly: true, secure: process.env.NODE_ENV === 'production' });
        console.log("✅ Login berhasil, redirect ke dashboard.");
        res.redirect('/');
    } catch (error) {
        console.error("❌ Kesalahan login:", error);
        res.status(500).send("Terjadi kesalahan server");
    }
});


// Logout
router.get('/logout', (req, res) => {
    res.clearCookie('token');
    res.redirect('/login');
});

module.exports = router;
