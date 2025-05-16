const buildPetaniQuery = (filters) => {
    const { kabupaten, tahun } = filters;
    let query = `
      SELECT 
                e.kabupaten, 
                e.kecamatan,
                e.nik, 
                e.nama_petani, 
                e.kode_kios,
                e.tahun, 
                e.urea, 
                e.npk, 
                e.npk_formula, 
                e.organik,

                -- Data tebusan total
                COALESCE(v.tebus_urea, 0) AS tebus_urea,
                COALESCE(v.tebus_npk, 0) AS tebus_npk,
                COALESCE(v.tebus_npk_formula, 0) AS tebus_npk_formula,
                COALESCE(v.tebus_organik, 0) AS tebus_organik,

                -- Perhitungan sisa
                (e.urea - COALESCE(v.tebus_urea, 0)) AS sisa_urea,
                (e.npk - COALESCE(v.tebus_npk, 0)) AS sisa_npk,
                (e.npk_formula - COALESCE(v.tebus_npk_formula, 0)) AS sisa_npk_formula,
                (e.organik - COALESCE(v.tebus_organik, 0)) AS sisa_organik,

                -- Data tebusan per bulan
                COALESCE(t.jan_urea, 0) AS jan_urea,
                COALESCE(t.feb_urea, 0) AS feb_urea,
                COALESCE(t.mar_urea, 0) AS mar_urea,
                COALESCE(t.apr_urea, 0) AS apr_urea,
                COALESCE(t.mei_urea, 0) AS mei_urea,
                COALESCE(t.jun_urea, 0) AS jun_urea,
                COALESCE(t.jul_urea, 0) AS jul_urea,
                COALESCE(t.agu_urea, 0) AS agu_urea,
                COALESCE(t.sep_urea, 0) AS sep_urea,
                COALESCE(t.okt_urea, 0) AS okt_urea,
                COALESCE(t.nov_urea, 0) AS nov_urea,
                COALESCE(t.des_urea, 0) AS des_urea,

                COALESCE(t.jan_npk, 0) AS jan_npk,
                COALESCE(t.feb_npk, 0) AS feb_npk,
                COALESCE(t.mar_npk, 0) AS mar_npk,
                COALESCE(t.apr_npk, 0) AS apr_npk,
                COALESCE(t.mei_npk, 0) AS mei_npk,
                COALESCE(t.jun_npk, 0) AS jun_npk,
                COALESCE(t.jul_npk, 0) AS jul_npk,
                COALESCE(t.agu_npk, 0) AS agu_npk,
                COALESCE(t.sep_npk, 0) AS sep_npk,
                COALESCE(t.okt_npk, 0) AS okt_npk,
                COALESCE(t.nov_npk, 0) AS nov_npk,
                COALESCE(t.des_npk, 0) AS des_npk,

                COALESCE(t.jan_npk_formula, 0) AS jan_npk_formula,
                COALESCE(t.feb_npk_formula, 0) AS feb_npk_formula,
                COALESCE(t.mar_npk_formula, 0) AS mar_npk_formula,
                COALESCE(t.apr_npk_formula, 0) AS apr_npk_formula,
                COALESCE(t.mei_npk_formula, 0) AS mei_npk_formula,
                COALESCE(t.jun_npk_formula, 0) AS jun_npk_formula,
                COALESCE(t.jul_npk_formula, 0) AS jul_npk_formula,
                COALESCE(t.agu_npk_formula, 0) AS agu_npk_formula,
                COALESCE(t.sep_npk_formula, 0) AS sep_npk_formula,
                COALESCE(t.okt_npk_formula, 0) AS okt_npk_formula,
                COALESCE(t.nov_npk_formula, 0) AS nov_npk_formula,
                COALESCE(t.des_npk_formula, 0) AS des_npk_formula,

                COALESCE(t.jan_organik, 0) AS jan_organik,
                COALESCE(t.feb_organik, 0) AS feb_organik,
                COALESCE(t.mar_organik, 0) AS mar_organik,
                COALESCE(t.apr_organik, 0) AS apr_organik,
                COALESCE(t.mei_organik, 0) AS mei_organik,
                COALESCE(t.jun_organik, 0) AS jun_organik,
                COALESCE(t.jul_organik, 0) AS jul_organik,
                COALESCE(t.agu_organik, 0) AS agu_organik,
                COALESCE(t.sep_organik, 0) AS sep_organik,
                COALESCE(t.okt_organik, 0) AS okt_organik,
                COALESCE(t.nov_organik, 0) AS nov_organik,
                COALESCE(t.des_organik, 0) AS des_organik
            FROM erdkk e FORCE INDEX (idx_erdkk_nik_kabupaten_tahun)
            LEFT JOIN verval_summary v 
                ON e.nik = v.nik
                AND e.kabupaten = v.kabupaten
                AND e.tahun = v.tahun
                AND e.kecamatan = v.kecamatan
                AND e.kode_kios = v.kode_kios
            LEFT JOIN tebusan_per_bulan t 
                ON e.nik = t.nik
                AND e.kabupaten = t.kabupaten
                AND e.tahun = t.tahun
                AND e.kecamatan = t.kecamatan
                AND e.kode_kios = t.kode_kios
    WHERE 1=1
  `;

    const params = [];
    if (tahun) {
        query += " AND e.tahun = ?";
        params.push(tahun);
    }

    if (kabupaten) {
        query += " AND e.kabupaten = ?";
        params.push(kabupaten);
    }

    return { query, params };
};

module.exports = { buildPetaniQuery };