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
require('dotenv').config();

const app = express();

// Konfigurasi storage untuk multer
const storage = multer.memoryStorage(); // Simpan file di memori (RAM)
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

// Halaman login
app.get('/login', (req, res) => {
    res.render('login');
});

// Proteksi halaman hanya untuk admin yang login
app.get('/', authMiddleware, (req, res) => {
    res.render('index');
});


app.get('/upload', authMiddleware, (req, res) => {
    res.render('upload');
});

app.get('/dataverval', authMiddleware, (req, res) => {
    res.render('dataTable');
});

app.get('/salurkios', authMiddleware, (req, res) => {
    res.render('salurkios');
});

app.get('/dashboard2', authMiddleware, (req, res) => {
    res.render('chart');
});

app.get('/upload-erdkk', authMiddleware, (req, res) => {
    res.render('upload-erdkk');
});

app.get('/upload-wcm', authMiddleware, (req, res) => {
    res.render('upload-wcm');
});

app.get('/erdkk', authMiddleware, (req, res) => {
    res.render('erdkk');
});

app.get('/esummary', authMiddleware, (req, res) => {
    res.render('alokasivstebusan')
});

app.get('/summary', authMiddleware, (req, res) => {
    res.render('esummary')
});

app.get('/wcm', authMiddleware, (req, res) => {
    res.render('wcm')
});

// Menangani file upload route
app.use('/api/files', authMiddleware, fileRoutes);
app.use('/api', authMiddleware, dataRoutes);
app.use('/api/data', authMiddleware, erdkkRoutes);
app.use('/api/data', authMiddleware, wcmRoutes);
app.use('/upload', outRouter);

app.use(express.json({ limit: '50mb' })); // Sesuaikan ukuran jika perlu
app.use(express.urlencoded({ extended: true, limit: '50mb' }));
// Gunakan route untuk login/logout
app.use(authRoutes);

// 🚀 Run server
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Server berjalan port:${PORT}`);
});
