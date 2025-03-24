"use strict";

// Chart

document.addEventListener("DOMContentLoaded", function () {
    let rawData = [];

    var diyList = ["BANTUL", "KOTA YOGYAKARTA", "SLEMAN", "KULON PROGO", "GUNUNG KIDUL"];
    var jatengList = ["KOTA MAGELANG", "MAGELANG", "SRAGEN", "BOYOLALI", "KLATEN", "SUKOHARJO", "KARANGANYAR", "KOTA SURAKARTA", "WONOGIRI"];

    var colors = ["#f3545d", "#fdaf4b", "#902C9C", "#1d8cf8", "#e3e3e3",
                  "#f8e71c", "#7ed321", "#50e3c2", "#4a90e2", "#9013fe",
                  "#b8e986", "#ff7f50", "#20c997"];

    function processData(kabupatenList) {
        var groupedData = {};
        rawData.forEach(entry => {
            if (kabupatenList.includes(entry.kabupaten)) {
                if (!groupedData[entry.kabupaten]) {
                    groupedData[entry.kabupaten] = { label: entry.kabupaten, data: Array(12).fill(0) };
                }
                var monthIndex = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"].indexOf(entry.bulan);
                groupedData[entry.kabupaten].data[monthIndex] = entry.total;
            }
        });
        return Object.values(groupedData).map((dataset, index) => ({
            label: dataset.label,
            borderColor: colors[index % colors.length],
            pointBackgroundColor: colors[index % colors.length],
            pointRadius: 4,
            backgroundColor: colors[index % colors.length] + '40',
            fill: true,
            borderWidth: 2,
            data: dataset.data
        }));
    }

    // ✅ Inisialisasi chart lebih awal
    var ctxDIY = document.getElementById('statisticsChartDIY').getContext('2d');
    var ctxJateng = document.getElementById('statisticsChartJateng').getContext('2d');

    var statisticsChartDIY = new Chart(ctxDIY, {
        type: 'line',
        data: {
            labels: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],
            datasets: []
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            legend: { display: true }
        }
    });

    var statisticsChartJateng = new Chart(ctxJateng, {
        type: 'line',
        data: {
            labels: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],
            datasets: []
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            legend: { display: true }
        }
    });

    fetch("http://localhost:3000/api/data/summary")
        .then(response => response.json())
        .then(responseData => {
            console.log("Data dari API:", responseData);

            let data = responseData.data;
            if (!Array.isArray(data)) {
                console.error("Data bukan array!", data);
                return;
            }

            let totalsPerKabupaten = {};

            data.forEach(row => {
                let kabupaten = row.kabupaten;
                let bulan = row.bulan; // ✅ Pakai bulan dari API
                let totalPupuk = row.total; // ✅ Pakai total dari API langsung

                if (!totalsPerKabupaten[kabupaten]) {
                    totalsPerKabupaten[kabupaten] = {};
                }

                if (!totalsPerKabupaten[kabupaten][bulan]) {
                    totalsPerKabupaten[kabupaten][bulan] = 0;
                }

                totalsPerKabupaten[kabupaten][bulan] += totalPupuk;
            });

            console.log("Total pupuk per kabupaten per bulan:", totalsPerKabupaten);

            rawData = [];
            for (let kabupaten in totalsPerKabupaten) {
                for (let bulan in totalsPerKabupaten[kabupaten]) {
                    rawData.push({
                        bulan: bulan, // ✅ Gunakan bulan dari API langsung
                        kabupaten: kabupaten,
                        total: totalsPerKabupaten[kabupaten][bulan]
                    });
                }
            }

            // ✅ Update Chart hanya jika sudah ada
            if (statisticsChartDIY) {
                statisticsChartDIY.data.datasets = processData(diyList);
                statisticsChartDIY.update();
            }

            if (statisticsChartJateng) {
                statisticsChartJateng.data.datasets = processData(jatengList);
                statisticsChartJateng.update();
            }
        })
        .catch(error => console.error("Error fetching data:", error));
});
