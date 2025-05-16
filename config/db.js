const mysql = require("mysql2/promise");
require("dotenv").config();

const db = mysql.createPool({
  host: process.env.DB_HOST,
  user: process.env.DB_USER,
  password: process.env.DB_PASSWORD,
  database: process.env.DB_NAME,

  // Optimasi untuk query besar
  waitForConnections: true,
  connectionLimit: 20,
  queueLimit: 0,
  decimalNumbers: true,
  namedPlaceholders: true,
  timezone: '+07:00', // WIB

  // Timeout lebih panjang
  connectTimeout: 30000
});

// Handle error koneksi
db.on('connection', (conn) => {
  conn.on('error', (err) => {
    console.error('Database error:', err);
  });
});

module.exports = db;