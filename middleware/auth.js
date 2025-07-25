// auth.js
const jwt = require('jsonwebtoken');
require('dotenv').config();

module.exports = (req, res, next) => {
    const token = req.cookies.token;
    if (!token) return res.redirect('/login');

    jwt.verify(token, process.env.JWT_SECRET, (err, decoded) => {
    if (err) return res.redirect('/login');

    // Pastikan akses adalah array, walau yang tersimpan adalah string
    if (typeof decoded.akses === 'string') {
        try {
            decoded.akses = JSON.parse(decoded.akses);
        } catch (e) {
            decoded.akses = []; // fallback kalau gagal parse
        }
    }

    req.user = decoded;
    res.locals.username = decoded.username;
    res.locals.userLevel = decoded.level;
    res.locals.kabupaten = decoded.kabupaten;
    res.locals.akses = decoded.akses;
    next();
});

};
