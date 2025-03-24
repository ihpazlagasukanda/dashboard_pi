const db = require("../config/db");

const insertData = async (data) => {
  const sql = `INSERT INTO merged_data 
    (Kabupaten, Kecamatan, Kode_Kios, NIK, Nama_Petani, Metode_Penebusan, Tanggal_Tebus, Urea, NPK, SP36, ZA, NPK_Formula, Organik, Organik_Cair, Kakao) 
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`;

  return db.execute(sql, data);
};

const getAllData = async () => {
  const [rows] = await db.execute("SELECT * FROM merged_data ORDER BY Tanggal_Tebus DESC");
  return rows;
};

module.exports = { insertData, getAllData };
