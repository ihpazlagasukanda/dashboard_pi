const express = require("express");
const router = express.Router();
const db = require("../config/db"); // Sesuaikan dengan koneksi MySQL kamu
const dataController = require("../controllers/dataController");
const { getAllData } = require("../controllers/dataController");
const { getSummaryData } = require("../controllers/dataController");
const path = require('path');
const fs = require('fs');

const { datacatalog } = require("googleapis/build/src/apis/datacatalog");

// const { generatePetaniReport } = require('../services/excelService');
// const { processLargeExport } = require('../services/bigDataService');
// const { generateExportPath } = require('../services/cacheService');

// Endpoint untuk request export
// router.post('/export', async (req, res) => {
//     try {
//         const exportPath = generateExportPath(req.body.kabupaten);

//         // Untuk data kecil (<50k rows)
//         if (req.body.smallExport) {
//             await generatePetaniReport(req.body, exportPath);
//             return res.json({
//                 downloadUrl: `/download?file=${path.basename(exportPath)}`
//             });
//         }

//         // Untuk data besar (background process)
//         processLargeExport(req.body, (progress) => {
//             req.io.emit('export-progress', { progress });
//         }).then(() => {
//             req.io.emit('export-ready', {
//                 downloadUrl: `/download?file=${path.basename(exportPath)}`
//             });
//         });

//         res.json({ message: 'Export sedang diproses...' });
//     } catch (error) {
//         res.status(500).json({ error: error.message });
//     }
// });

router.get('/download-export', (req, res) => {
    const fileName = req.query.file;

    if (!fileName || typeof fileName !== 'string') {
        return res.status(400).json({ error: 'Parameter file wajib diisi' });
    }

    const filePath = path.join(__dirname, '../temp_exports', fileName);

    if (!fs.existsSync(filePath)) {
        return res.status(404).json({ error: 'File tidak ditemukan' });
    }

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename=${fileName}`);

    const fileStream = fs.createReadStream(filePath);
    fileStream.pipe(res);

    fileStream.on('end', () => {
        try {
            fs.unlinkSync(filePath);
        } catch (err) {
            console.error('Gagal menghapus file:', err);
        }
    });
});

router.get("/data/all", getAllData);

router.get("/data/summary", getSummaryData);

router.get("/erdkk", dataController.getErdkk);

// untuk dashboard 1
//pilihan jenis alokasi
router.get("/erdkk/summary", dataController.getErdkkSummary);
router.get("/sk-bupati/alokasi", dataController.getSkBupatiAlokasi);

router.get("/verval/summary", dataController.getSummaryPenebusan);
router.get("/erdkk/count", dataController.getErdkkCount);
router.get("/verval/count", dataController.getPenebusanCount);
router.get("/download/erdkk", dataController.downloadErdkk);
// end dashboard 1

router.get('/data/jumlah-petani', dataController.getJumlahPetani);

router.get('/alokasivstebus', dataController.alokasiVsTebusan);

router.get('/summary/pupuk', dataController.summaryPupuk);
router.get('/rekap/petani', dataController.rekapPetani);


router.get('/petani-summary', dataController.getVervalSummary);

router.get('/tebusperbulan', dataController.tebusanPerBulan);

router.get('/download-petanisummary', dataController.downloadPetaniSummary);
router.get('/download-petanisum', dataController.downloadPetaniSum);

router.get('/data/salurkios', dataController.getSalurKios);

router.get('/data/salurkios/sum', dataController.getSalurKiosSum);

router.get('/download/salurkios', dataController.downloadSalurKios);
router.get('/download/salur', dataController.downloadSalur);

router.get('/data', dataController.getData);
router.get('/data/sum', dataController.getSum);
router.get('/data/download', dataController.exportExcel);


router.get('/data/wcm', dataController.getWcm);
router.get('/data/wcmf5', dataController.getWcmF5);
router.get('/download/wcm', dataController.downloadWcm);
router.get('/data/wcmvsverval', dataController.wcmVsVerval);
router.get('/data/penyaluran_do', dataController.getPenyaluranDo);
router.get('/download/wcmf5', dataController.downloadWcmF5);
router.get('/download/wcmvsverval', dataController.exportExcelWcmVsVerval);
router.get('/download/summary-pupuk', dataController.downloadSummaryPupuk);
router.get('/download/rekap-petani', dataController.downloadRekapPetani);


// Endpoint untuk last updated
router.get('/lastupdated/global', dataController.getLastUpdatedGlobal);
router.get('/lastupdated/kabupaten', dataController.getLastUpdatedPerKabupaten);

// routes/filterRoutes.js
router.get("/filters", async (req, res) => {
    try {
        const [kabupaten] = await db.query("SELECT DISTINCT kabupaten FROM verval ORDER BY kabupaten");
        const [kecamatan] = await db.query("SELECT DISTINCT kecamatan FROM verval ORDER BY kecamatan");
        const [metode_penebusan] = await db.query("SELECT DISTINCT metode_penebusan FROM verval ORDER BY metode_penebusan");
        const [tanggal_tebus] = await db.query(`
            SELECT DISTINCT DATE_FORMAT(tanggal_tebus, '%Y-%m') AS tanggal_tebus
            FROM verval
            ORDER BY tanggal_tebus
        `);

        res.json({
            kabupaten: kabupaten.map(row => row.kabupaten),
            kecamatan: kecamatan.map(row => row.kecamatan),
            metode_penebusan: metode_penebusan.map(row => row.metode_penebusan),
            tanggal_tebus: tanggal_tebus.map(row => row.tanggal_tebus)
        });

    } catch (error) {
        console.error("Error fetching filters:", error);
        res.status(500).json({ error: "Internal Server Error" });
    }
});


module.exports = router;
