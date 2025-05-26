const express = require('express');
const multer = require('multer');
const path = require('path');
const cookieParser = require('cookie-parser');
const fileRoutes = require('./routes/fileRoutes');
const dataRoutes = require('./routes/dataRoutes');
const authRoutes = require('./routes/authRoutes');
const authMiddleware = require('./middleware/auth');
const erdkkRoutes = require("./routes/erdkkRoutes");
const wcmRoutes = require("./routes/wcmRoutes");
const outRouter = require("./routes/outRoutes");
const skRoutes = require("./routes/skRoutes");
const triggerCron = require("./routes/triggerCron");
const uploadPenyaluranDoRoutes = require("./routes/penyaluranDoRoutes");
require('dotenv').config();

const app = express();

// Konfigurasi multer
const storage = multer.memoryStorage(); // Penyimpanan di RAM
const upload = multer({ storage: storage });

// Middleware
app.use(express.urlencoded({ extended: true }));
app.use(express.json());
app.use(cookieParser());

// Set view engine ke EJS
app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));

// Akses file statis
app.use(express.static('public'));

// Middleware untuk pengecekan level admin
function requireLevel(level) {
    return function (req, res, next) {
        if (!req.user || req.user.level < level) {
            return res.status(403).render('access-denied'); // Menampilkan halaman access-denied
        }
        next();
    };
}

require('./jobs/reportCron');

// Routes untuk halaman utama dan login
app.get('/', authMiddleware, (req, res) => {
    res.render('index', { user: req.user });
});

app.get('/login', (req, res) => {
    res.render('login');
});

// Halaman lain dengan proteksi autentikasi
app.get('/dataverval', authMiddleware, (req, res) => {
    res.render('dataTable', { user: req.user });
});

app.get('/salurkios', authMiddleware, (req, res) => {
    res.render('salurkios', { user: req.user });
});

// app.get('/dashboard2', authMiddleware, (req, res) => {
//     res.render('chart', { user: req.user });
// });

// Halaman Upload - Hanya level 2
app.get('/upload', authMiddleware, requireLevel(2), (req, res) => {
    res.render('upload', { user: req.user });
});

app.get('/upload-erdkk', authMiddleware, requireLevel(2), (req, res) => {
    res.render('upload-erdkk', { user: req.user });
});

app.get('/upload-wcm', authMiddleware, requireLevel(2), (req, res) => {
    res.render('upload-wcm', { user: req.user });
});

app.get('/upload-penyalurando', authMiddleware, requireLevel(2), (req, res) => {
    res.render('upload-penyaluranDo', { user: req.user });
});

app.get('/upload-skbupati', authMiddleware, requireLevel(2), (req, res) => {
    res.render('upload-skBupati', { user: req.user });
});

// Routes untuk halaman lain
app.get('/erdkk', authMiddleware, (req, res) => {
    res.render('erdkk', { user: req.user });
});

app.get('/esummary', authMiddleware, (req, res) => {
    res.render('alokasivstebusan', { user: req.user });
});

app.get('/summary', authMiddleware, (req, res) => {
    res.render('esummary', { user: req.user });
});

app.get('/wcm', authMiddleware, (req, res) => {
    res.render('wcm', { user: req.user });
});

app.get('/wcmf5', authMiddleware, (req, res) => {
    res.render('f5', { user: req.user });
});

app.get('/sum', authMiddleware, (req, res) => {
    res.render('sum', { user: req.user });
});

app.get('/wcmvsverval', authMiddleware, (req, res) => {
    res.render('wcmvsverval', { user: req.user });
});

app.get('/penyalurando', authMiddleware, (req, res) => {
    res.render('penyalurando', { user: req.user });
});

// app.get('/visualisasi', authMiddleware, (req, res) => {
//     res.render('visualisasi', { user: req.user });
// });

app.get('/visualisasi', authMiddleware, requireLevel(2), (req, res) => {
    res.render('visulisasi', { user: req.user });
});


// API Routes
app.use('/api/files', authMiddleware, fileRoutes);
app.use('/api', authMiddleware, dataRoutes);
app.use('/api/data', authMiddleware, erdkkRoutes);
app.use('/api/data', authMiddleware, wcmRoutes);
app.use('/api/data', authMiddleware, skRoutes);
app.use('/trigger', authMiddleware, triggerCron);
app.use('/api/data', authMiddleware, uploadPenyaluranDoRoutes);
app.use(authRoutes);  // Pastikan authRoutes ada
app.use('/upload', outRouter);  // Route untuk upload

app.use(express.json({ limit: '50mb' })); // Sesuaikan ukuran jika perlu
app.use(express.urlencoded({ extended: true, limit: '50mb' }));

// Jalankan server
const PORT = process.env.PORT || 3000;
const server = app.listen(PORT, () => {
    console.log(`Server running on port ${PORT}`);
});

// Set timeout server menjadi 10 menit (600.000 ms)
server.setTimeout(1800000); // 600.000 ms = 10 menit
