const winston = require('winston');

// Konfigurasi logger
const logger = winston.createLogger({
    level: process.env.NODE_ENV === 'production' ? 'warn' : 'debug', // Production hanya log warn/error
    format: winston.format.combine(
        winston.format.timestamp({ format: 'YYYY-MM-DD HH:mm:ss' }),
        winston.format.printf(({ timestamp, level, message }) => {
            return `[${timestamp}] ${level.toUpperCase()}: ${message}`;
        })
    ),
    transports: [
        new winston.transports.Console(), // Tampil di terminal
        new winston.transports.File({ filename: 'app.log' }) // Simpan ke file app.log
    ],
});

// Override bawaan console
console.log = (...args) => logger.info(args.join(' '));
console.info = (...args) => logger.info(args.join(' '));
console.warn = (...args) => logger.warn(args.join(' '));
console.error = (...args) => logger.error(args.join(' '));

module.exports = logger;
