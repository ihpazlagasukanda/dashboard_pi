const db = require('../config/db');
const bcrypt = require('bcrypt');

const User = {
    findByUsername: async (username) => {
        try {
            console.log("🔍 Menjalankan query pencarian user:", username);
            const [results] = await db.execute("SELECT * FROM admin WHERE username = ?", [username]);

            console.log("✅ Hasil query:", results);

            if (results.length === 0) {
                console.log("⚠️ User tidak ditemukan.");
                return null;
            }

            return results[0]; // Ambil user pertama
        } catch (err) {
            console.error("❌ Kesalahan query database:", err);
            throw err;
        }
    }
};

module.exports = User;
