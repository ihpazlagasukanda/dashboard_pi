// auth.js
const jwt = require('jsonwebtoken');  // Hanya dideklarasikan sekali
require('dotenv').config();

module.exports = (req, res, next) => {
    const token = req.cookies.token;
    if (!token) return res.redirect('/login');

    jwt.verify(token, process.env.JWT_SECRET, (err, decoded) => {
        if (err) return res.redirect('/login');

        req.user = decoded;
        res.locals.username = decoded.username;  // Simpan username ke res.locals
        res.locals.userLevel = decoded.level;    // Simpan level ke res.locals
        next();
    });
};
