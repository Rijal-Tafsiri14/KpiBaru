document.addEventListener("DOMContentLoaded", () => {
  const fileInput = document.getElementById("fileInput");
  const thresholdInput = document.getElementById("thresholdInput");
  const summaryTable = document.querySelector("#summaryTable tbody");
  const failedTable = document.querySelector("#failedTable tbody");
  const pieChartCanvas = document.getElementById("pieChart");
  let pieChart;
let lastData = []; // simpan data terakhir yang diupload

  // Parsing tanggal dari Excel (baik string atau serial)
  function parseExcelDate(value) {
    if (!value) return null;
    if (typeof value === "number") {
      const d = XLSX.SSF.parse_date_code(value);
      return new Date(d.y, d.m - 1, d.d, d.H, d.M, d.S);
    } else if (typeof value === "string") {
      const [datePart, timePart] = value.split(" ");
      const [m, d, y] = datePart.split("/").map(Number);
      let hh = 0, mm = 0;
      if (timePart) [hh, mm] = timePart.split(":").map(Number);
      return new Date(2000 + y, m - 1, d, hh, mm);
    }
    return null;
  }

  // Event saat file di-upload
  fileInput.addEventListener("change", (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (event) => {
  const data = new Uint8Array(event.target.result);
  const workbook = XLSX.read(data, { type: "array" });
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];
  const json = XLSX.utils.sheet_to_json(worksheet);
  lastData = json; // simpan data ke variabel global
  processData(json);
};
thresholdInput.addEventListener("input", () => {
  if (lastData.length > 0) {
    processData(lastData);
  }
});

    reader.readAsArrayBuffer(file);
  });

  function processData(rows) {
    const threshold = Number(thresholdInput.value) || 24;
    const summaryByDate = {};
    const failedOrders = [];

    rows.forEach((row) => {
      const orderDate = parseExcelDate(row["Order Date"]);
      const lastDate = parseExcelDate(row["Last Process Date"]);
      if (!orderDate || !lastDate) return;

      const diffHrs = (lastDate - orderDate) / (1000 * 60 * 60);
      const achieve = diffHrs <= threshold;
      const dateKey = orderDate.toLocaleDateString("en-US");

      if (!summaryByDate[dateKey]) {
        summaryByDate[dateKey] = { totalOrder: 0, totalItem: 0, totalQty: 0, achieve: 0, failed: 0 };
      }

      summaryByDate[dateKey].totalOrder++;
      summaryByDate[dateKey].totalItem++;
      summaryByDate[dateKey].totalQty += row["Quantity"] || 0;

      if (achieve) {
        summaryByDate[dateKey].achieve++;
      } else {
        summaryByDate[dateKey].failed++;
        failedOrders.push({
          orderId: row["Order ID"],
          orderItemId: row["Order Item ID"],
          sku: row["SKU"] || "-",
          naming: row["Naming"] || "-",
          qty: row["Quantity"],
          orderDate: orderDate.toLocaleString(),
          lastDate: lastDate.toLocaleString(),
          diff: diffHrs.toFixed(2)
        });
      }
    });

    renderSummary(summaryByDate);
    renderFailedTable(failedOrders);
  }

  function renderSummary(data) {
    summaryTable.innerHTML = "";
    let totalAll = { totalOrder: 0, totalItem: 0, totalQty: 0, achieve: 0, failed: 0 };

    Object.keys(data).forEach((date) => {
      const s = data[date];
      const percent = ((s.achieve / s.totalOrder) * 100).toFixed(2);

      const row = `
        <tr>
          <td>${date}</td>
          <td>${s.totalOrder}</td>
          <td>${s.totalItem}</td>
          <td>${s.totalQty}</td>
          <td>${s.achieve}</td>
          <td>${s.failed}</td>
          <td>${percent}%</td>
        </tr>
      `;
      summaryTable.insertAdjacentHTML("beforeend", row);

      Object.keys(totalAll).forEach((k) => totalAll[k] += s[k]);
    });

    const totalPercent = ((totalAll.achieve / totalAll.totalOrder) * 100).toFixed(2);
    const totalRow = `
      <tr class="total-row">
        <td><b>TOTAL</b></td>
        <td>${totalAll.totalOrder}</td>
        <td>${totalAll.totalItem}</td>
        <td>${totalAll.totalQty}</td>
        <td>${totalAll.achieve}</td>
        <td>${totalAll.failed}</td>
        <td><b>${totalPercent}%</b></td>
      </tr>
    `;
    summaryTable.insertAdjacentHTML("beforeend", totalRow);

    updatePieChart(totalAll.achieve, totalAll.failed);
  }

  function renderFailedTable(data) {
  failedTable.innerHTML = "";

  // Ambil data reason lama dari localStorage (kalau ada)
  const savedReasons = JSON.parse(localStorage.getItem("reasons") || "{}");

  data.forEach((item, index) => {
    const key = `${item.orderId}_${item.orderItemId}`;
    const savedReason = savedReasons[key] || "";

    const row = `
  <tr>
    <td>${item.orderId}</td>
    <td>${item.orderItemId}</td>
    <td>${item.sku}</td>
    <td>${item.naming}</td>
    <td>${item.qty}</td>
    <td>${item.orderDate}</td>
    <td>${item.lastDate}</td>
    <td>${item.diff}</td>
    <td>
      <textarea class="reason-input"
                data-key="${key}"
                placeholder="Isi reason..."
                rows="2">${savedReason}</textarea>
    </td>
  </tr>
    `;
    failedTable.insertAdjacentHTML("beforeend", row);
  });

  // Tambahkan event listener ke semua input reason
  document.querySelectorAll(".reason-input").forEach(input => {
    input.addEventListener("input", (e) => {
      const key = e.target.dataset.key;
      savedReasons[key] = e.target.value;
      localStorage.setItem("reasons", JSON.stringify(savedReasons));
    });
  });
}


  function updatePieChart(achieve, failed) {
    const ctx = pieChartCanvas.getContext("2d");
    if (pieChart) pieChart.destroy();

    pieChart = new Chart(ctx, {
      type: "doughnut",
      data: {
        labels: ["Achieve", "Failed"],
        datasets: [{
          data: [achieve, failed],
          backgroundColor: ["#4CAF50", "#F44336"],
          borderWidth: 1
        }]
      },
      options: {
        responsive: true,
        plugins: {
          legend: { position: "bottom" },
          tooltip: { callbacks: { label: (c) => `${c.label}: ${c.formattedValue}` } }
        }
      }
    });
  }
});
// === FITUR EXPORT & IMPORT REASON ===
const downloadBtn = document.getElementById("downloadReason");
const uploadInput = document.getElementById("uploadReason");

// ✅ Download Reason ke Excel
downloadBtn.addEventListener("click", () => {
  const savedReasons = JSON.parse(localStorage.getItem("reasons") || "{}");
  const reasonArray = Object.entries(savedReasons).map(([key, reason]) => {
    const [orderId, orderItemId] = key.split("_");
    return { "Order ID": orderId, "Order Item ID": orderItemId, "Reason": reason };
  });

  if (reasonArray.length === 0) {
    alert("Belum ada reason tersimpan!");
    return;
  }

  const ws = XLSX.utils.json_to_sheet(reasonArray);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Reasons");
  XLSX.writeFile(wb, "reason_gagal.xlsx");
});

// ✅ Upload Reason dari Excel
uploadInput.addEventListener("change", (e) => {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (event) => {
    const data = new Uint8Array(event.target.result);
    const workbook = XLSX.read(data, { type: "array" });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const json = XLSX.utils.sheet_to_json(worksheet);

    const newReasons = {};
    json.forEach((r) => {
      if (r["Order ID"] && r["Order Item ID"]) {
        const key = `${r["Order ID"]}_${r["Order Item ID"]}`;
        newReasons[key] = r["Reason"] || "";
      }
    });

    localStorage.setItem("reasons", JSON.stringify(newReasons));
    alert("Reason berhasil diunggah!");
    if (lastData.length > 0) processData(lastData); // refresh tabel gagal
  };

  reader.readAsArrayBuffer(file);
});

