import React, { useEffect, useState } from "react";
import axios from "axios";
import "bootstrap/dist/css/bootstrap.min.css";

const DataTable = () => {
  const [data, setData] = useState([]);

  useEffect(() => {
    axios
      .get("http://localhost:3000/api/data") // API dari backend
      .then((response) => setData(response.data))
      .catch((error) => console.error("Error fetching data:", error));
  }, []);

  return (
    <div className="container mt-4">
      <h2 className="mb-3 text-center">📊 Data Petani</h2>
      <div className="table-responsive">
        <table className="table table-bordered table-striped">
          <thead className="table-dark text-center">
            <tr>
              <th>Kabupaten</th>
              <th>Kecamatan</th>
              <th>Kode Kios</th>
              <th>NIK</th>
              <th>Nama Petani</th>
              <th>Metode Penebusan</th>
              <th>Tanggal Tebus</th>
              <th>Urea</th>
              <th>NPK</th>
              <th>SP36</th>
              <th>ZA</th>
              <th>NPK Formula</th>
              <th>Organik</th>
              <th>Organik Cair</th>
              <th>Kakao</th>
            </tr>
          </thead>
          <tbody>
            {data.length > 0 ? (
              data.map((row, index) => (
                <tr key={index} className="text-center">
                  <td>{row.Kabupaten}</td>
                  <td>{row.Kecamatan}</td>
                  <td>{row.Kode_Kios}</td>
                  <td>{row.NIK}</td>
                  <td>{row.Nama_Petani}</td>
                  <td>{row.Metode_Penebusan}</td>
                  <td>{row.Tanggal_Tebus}</td>
                  <td>{row.Urea}</td>
                  <td>{row.NPK}</td>
                  <td>{row.SP36}</td>
                  <td>{row.ZA}</td>
                  <td>{row.NPK_Formula}</td>
                  <td>{row.Organik}</td>
                  <td>{row.Organik_Cair}</td>
                  <td>{row.Kakao}</td>
                </tr>
              ))
            ) : (
              <tr>
                <td colSpan="15" className="text-center">
                  🔄 Memuat data...
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
};

export default DataTable;
