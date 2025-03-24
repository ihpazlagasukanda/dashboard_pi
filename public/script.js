const form = document.getElementById("uploadForm");

form.addEventListener("submit", (e) => {
  e.preventDefault();

  if (!confirm("Apakah Anda yakin ingin mengupload file ini?")) {
    alert("Upload dibatalkan.");
    return;
  }

  const fileInput = document.getElementById("fileInput");
  const metodePenebusan = document.querySelector(
    'input[name="metodePenebusan"]:checked'
  );

  if (!fileInput.files[0]) {
    alert("Silakan pilih file untuk diupload.");
    return;
  }

  if (!metodePenebusan) {
    alert("Silakan pilih metode penebusan.");
    return;
  }

  const formData = new FormData();
  // Tambahkan semua file yang dipilih
  for (let i = 0; i < fileInput.files.length; i++) {
    formData.append("files", fileInput.files[i]); // Gunakan nama field "files"
  }
  formData.append("metodePenebusan", metodePenebusan.value);

  fetch("/api/files/upload", {
    method: "POST",
    body: formData,
  })
    .then((response) => response.json())
    .then((data) => {
      if (data.success) {
        alert("✅ " + data.message); // Jika berhasil
      } else {
        alert("❌ " + data.message); // Jika gagal
      }
    })
    .catch((error) => {
      console.error("Gagal mengupload file:", error);
      alert("❌ Terjadi kesalahan saat mengupload file.");
    });
});


let currentPage = 1;
const limit = 50;

function fetchData(page) {
  fetch(`/api/data?page=${page}&limit=${limit}`)
    .then(response => response.json())
    .then(data => {
      renderTable(data.data);
      renderPagination(data.totalPages, data.currentPage);
    })
    .catch(error => console.error("Error fetching data:", error));
}

function renderTable(data) {
  const tableBody = document.getElementById("dataTable");
  tableBody.innerHTML = "";

  data.forEach(item => {
    const row = `
            <tr>
                <td>${item.kabupaten}</td>
                <td>${item.kecamatan}</td>
                <td>${item.kode_kios}</td>
                <td>${item.nik}</td>
                <td>${item.nama_petani}</td>
                <td>${item.metode_penebusan}</td>
                <td>${item.tanggal_tebus}</td>
                <td>${item.urea}</td>
                <td>${item.npk}</td>
                <td>${item.sp36}</td>
                <td>${item.za}</td>
                <td>${item.npk_formula}</td>
                <td>${item.organik}</td>
                <td>${item.organik_cair}</td>
                <td>${item.kakao}</td>
            </tr>
        `;
    tableBody.innerHTML += row;
  });
}

function renderPagination(totalPages, currentPage) {
  const pagination = document.getElementById("pagination");
  pagination.innerHTML = "";

  const prevDisabled = currentPage === 1 ? "disabled" : "";
  const nextDisabled = currentPage === totalPages ? "disabled" : "";

  pagination.innerHTML += `
        <li class="page-item ${prevDisabled}">
            <a class="page-link" href="#" onclick="fetchData(${currentPage - 1})">Previous</a>
        </li>
    `;

  let startPage = Math.max(1, currentPage - 2);
  let endPage = Math.min(totalPages, currentPage + 2);

  for (let i = startPage; i <= endPage; i++) {
    const activeClass = i === currentPage ? "active" : "";
    pagination.innerHTML += `
            <li class="page-item ${activeClass}">
                <a class="page-link" href="#" onclick="fetchData(${i})">${i}</a>
            </li>
        `;
  }

  pagination.innerHTML += `
        <li class="page-item ${nextDisabled}">
            <a class="page-link" href="#" onclick="fetchData(${currentPage + 1})">Next</a>
        </li>
    `;
}

// Load data saat halaman pertama dibuka
fetchData(1);


