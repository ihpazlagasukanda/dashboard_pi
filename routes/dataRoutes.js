const express = require("express");
const router = express.Router();
const db = require("../config/db"); // Sesuaikan dengan koneksi MySQL kamu
const dataController = require("../controllers/dataController");
const monitoringController = require("../controllers/monitoringController");
const rekapRealisasiController = require("../controllers/rekapRealisasiController");
const rekapPetaniController = require("../controllers/rekapPetaniController");
const rekapPetaniTebusController = require("../controllers/rekapPetaniTebusController");
const f5Controller = require("../controllers/f5Controller");
const { getAllData } = require("../controllers/dataController");
const { getSummaryData } = require("../controllers/dataController");
const path = require('path');
const fs = require('fs');

const { datacatalog } = require("googleapis/build/src/apis/datacatalog");
const { rekapPetani } = require("../controllers/rekapPetaniController");

// Menu Monitoring Transaksi Petani 
// pilihan jenis alokasi
router.get("/erdkk/summary", monitoringController.getErdkkSummary);
router.get("/sk-bupati/alokasi", monitoringController.getSkBupatiAlokasi);
router.get("/verval/summary", monitoringController.getSummaryPenebusan);
router.get("/erdkk/count", monitoringController.getErdkkCount);
router.get("/verval/count", monitoringController.getPenebusanCount);
router.get("/download/erdkk", dataController.downloadErdkk);
// end Monitoring Transaksi Petani

// Menu Rekap Realisasi
router.get('/summary/pupuk', rekapRealisasiController.summaryPupuk);
router.get('/download/summary-pupuk', rekapRealisasiController.downloadSummaryPupuk);
// End Rekap Realisasi

// Menu Rekap Jumlah Petani
router.get('/rekap/petani', rekapPetaniController.rekapPetani);
router.get('/download/rekap-petani', rekapPetaniController.downloadRekapPetani);
// End Rekap Jumlah Petani

// Menu Rekap Jumlah Petani
router.get('/rekap/petani/all', rekapPetaniTebusController.rekapPetaniAll);
router.get('/download/rekap-petani/all', rekapPetaniTebusController.downloadRekapPetaniAll);
// End Rekap Jumlah Petani

//Menu F5
router.get('/download/f5', f5Controller.downloadF5);
router.get('/data/f5', f5Controller.getF5);
// End Menu F5

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

router.get('/data/jumlah-petani', dataController.getJumlahPetani);

router.get('/alokasivstebus', dataController.alokasiVsTebusan);

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

router.get('/download/poktanbelumterdaftar', dataController.downloadPoktan);


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
