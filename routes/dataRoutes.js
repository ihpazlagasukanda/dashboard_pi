const express = require("express");
const router = express.Router();
const db = require("../config/db"); // Sesuaikan dengan koneksi MySQL kamu
const dataController = require("../controllers/dataController");
const { getAllData } = require("../controllers/dataController");
const { getSummaryData } = require("../controllers/dataController");
const { datacatalog } = require("googleapis/build/src/apis/datacatalog");
// Route untuk mendapatkan data dengan pagination
// router.get("/data", dataController.getDataWithPagination);

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
router.get('/download/wcm', dataController.downloadWcm);
router.get('/data/wcmvsverval', dataController.wcmVsVerval);
router.get('/download/wcmf5', dataController.downloadWcmF5);
router.get('/download/wcmvsverval', dataController.exportExcelWcmVsVerval);
router.get('/download/summary-pupuk', dataController.downloadSummaryPupuk);

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
