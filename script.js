const state = {
  rows: [],
  orderRows: [],
  deviasiRows: [],
  originalFile: null,
  deviasiLoaded: false,
  deviasiSheet: null,
  dates: [],
  activeDate: null,
  activeShift: "ALL",
  activeView: "detail",
  recapShiftVisible: false,
};

const fileInput = document.getElementById("fileInput");
const dateFilter = document.getElementById("dateFilter");
const dateInput = document.getElementById("dateInput");
const latestBtn = document.getElementById("latestBtn");
const resetBtn = document.getElementById("resetBtn");
const statusElement = document.getElementById("status");
const tableArea = document.getElementById("tableArea");
const tableInfo = document.getElementById("tableInfo");
const datePill = document.getElementById("datePill");
const cDate = document.getElementById("cDate");
const cRows = document.getElementById("cRows");
const cResult = document.getElementById("cResult");
const cShift = document.getElementById("cShift");
const shiftButtons = document.querySelectorAll(".shift-button");
const reportTabs = document.querySelectorAll(".report-tab");
const recapShiftToggle = document.getElementById("recapShiftToggle");
const shiftButtonsGroup = document.querySelector(".shift-buttons");

function toNumber(value) {
  if (value === null || value === undefined || value === "") {
    return 0;
  }

  if (typeof value === "number") {
    return Number.isFinite(value) ? value : 0;
  }

  let valueString = String(value).trim();

  if (!valueString) {
    return 0;
  }

  valueString = valueString.replace(/\s/g, "");

  if (valueString.includes(".") && valueString.includes(",")) {
    valueString = valueString.replace(/\./g, "").replace(",", ".");
  } else if (valueString.includes(",")) {
    valueString = valueString.replace(",", ".");
  }

  const number = Number(valueString);

  return Number.isFinite(number) ? number : 0;
}

function formatNumber(value) {
  return new Intl.NumberFormat("id-ID", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  }).format(toNumber(value));
}

function text(value) {
  if (value === null || value === undefined) {
    return "";
  }

  return String(value).trim();
}

function escapeHtml(value) {
  return String(value).replace(/[&<>"']/g, (character) => {
    const entities = {
      "&": "&amp;",
      "<": "&lt;",
      ">": "&gt;",
      '"': "&quot;",
      "'": "&#039;",
    };

    return entities[character];
  });
}

function showLoading(message = "Memuat data...") {
  tableArea.innerHTML = `
    <div class="loading-state">
      <div class="spinner"></div>
      <h3>${message}</h3>
      <p>Mohon tunggu sebentar</p>
    </div>`;
}

function hideLoading() {
  const loading = document.querySelector(".loading-state");
  if (loading) loading.remove();
}

function isValidDataRow(row) {
  const mesin = text(row[1]);
  const ukuran = text(row[2]);
  const merk = text(row[3]);
  const jenis = text(row[4]);

  const mesinUpper = mesin.toUpperCase();

  if (mesinUpper === "TOTAL" || mesinUpper === "JUMLAH" || mesinUpper === "SUB TOTAL" || mesinUpper === "SUBTOTAL") {
    return false;
  }

  if (!mesin) {
    return false;
  }

  if (!ukuran && !merk && !jenis) {
    return false;
  }

  return true;
}

function machineSort(a, b) {
  const mesinA = text(a.MESIN).toUpperCase();
  const mesinB = text(b.MESIN).toUpperCase();

  const matchA = mesinA.match(/^(\d+)([A-Z]*)$/);
  const matchB = mesinB.match(/^(\d+)([A-Z]*)$/);

  if (!matchA || !matchB) {
    return mesinA.localeCompare(mesinB, "id", {
      numeric: true,
    });
  }

  const numberA = Number(matchA[1]);
  const numberB = Number(matchB[1]);

  if (numberA !== numberB) {
    return numberA - numberB;
  }

  const suffixA = matchA[2];
  const suffixB = matchB[2];

  if (!suffixA && suffixB) {
    return 1;
  }

  if (suffixA && !suffixB) {
    return -1;
  }

  return suffixA.localeCompare(suffixB);
}

function hasShiftA(row) {
  return text(row.A_OPERATOR) !== "" || text(row.A_BALL) !== "" || text(row.A_HASIL) !== "" || text(row.A_TARGET) !== "";
}

function hasShiftB(row) {
  return text(row.B_OPERATOR) !== "" || text(row.B_BALL) !== "" || text(row.B_HASIL) !== "" || text(row.B_TARGET) !== "";
}

fileInput.addEventListener("change", (event) => {
  const file = event.target.files[0];

  if (!file) {
    return;
  }

  readExcel(file);
});

function readExcel(file) {
  if (typeof XLSX === "undefined") {
    hideLoading();
    alert("Library pembaca Excel belum siap. Muat ulang halaman lalu coba lagi.");
    return;
  }

  state.originalFile = file;
  state.deviasiLoaded = false;
  state.deviasiSheet = null;

  showLoading("Membaca file Excel...");

  const reader = new FileReader();

  reader.onload = (event) => {
    try {
      const workbook = XLSX.read(new Uint8Array(event.target.result), {
        type: "array",
        cellDates: true,
      });

      parseDataOrder(workbook);

      if (!workbook.SheetNames.includes("HD HASIL")) {
        throw new Error('Sheet "HD HASIL" tidak ditemukan.');
      }

      const worksheet = workbook.Sheets["HD HASIL"];

      const matrix = XLSX.utils.sheet_to_json(worksheet, {
        header: 1,
        defval: null,
        raw: true,
      });

      parseHDHasil(matrix, file.name);
      hideLoading();
    } catch (error) {
      console.error(error);
      hideLoading();
      alert("Gagal membaca file Excel:\n\n" + error.message);
    }
  };

  reader.onerror = () => {
    hideLoading();
    alert("Gagal membaca file Excel. Silakan pilih file lain dan coba kembali.");
  };

  reader.readAsArrayBuffer(file);
}

function parseDataOrder(workbook) {
  state.orderRows = [];

  if (!workbook.SheetNames.includes("DATA ORDER")) {
    return;
  }

  const sheet = workbook.Sheets["DATA ORDER"];
  const matrix = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: null,
    raw: true,
  });

  if (!matrix.length) return;

  const headerIndex = matrix.findIndex((row) => row.some((cell) => text(cell).toUpperCase().includes("MERK")));

  const start = headerIndex >= 0 ? headerIndex + 1 : 1;
  const headers = headerIndex >= 0 ? matrix[headerIndex].map(text) : [];

  function findCol(names) {
    return headers.findIndex((h) => names.some((n) => h.toUpperCase().includes(n)));
  }

  const col = {
    customer: findCol(["CUSTOMER", "CUST", "PELANGGAN"]),
    merk: findCol(["MERK"]),
    jenis: findCol(["JENIS"]),
    ukuran: findCol(["UKURAN"]),
    lebar: findCol(["LEBAR"]),
    lipatan: findCol(["LIPAT"]),
    tebal: findCol(["TEBAL"]),
    panjang: findCol(["PANJANG"]),
    chRoll: findCol(["CH PROD"]),
    chPot: findCol(["CH POT"]),
    lebarPack: findCol(["LEBAR/PACK", "LBR/PAK"]),
    isiBall: findCol(["ISI/BALL", "PAK/BALL"]),
  };

  matrix.slice(start).forEach((row) => {
    const item = {
      CUSTOMER: row[col.customer],
      MERK: row[col.merk],
      JENIS: row[col.jenis],
      UKURAN: row[col.ukuran],
      LEBAR: row[col.lebar],
      LIPATAN: row[col.lipatan],
      TEBAL: row[col.tebal],
      PANJANG: row[col.panjang],
      CH_ROLL: row[col.chRoll],
      CH_POT: row[col.chPot],
      LEBAR_PACK: row[col.lebarPack],
      ISI_BALL: row[col.isiBall],
    };

    if (text(item.MERK)) state.orderRows.push(item);
  });
}

function renderSpek() {
  const recap = getSelectedResults(state.rows.filter((row) => row.date === state.activeDate));

  const keys = [...recap.values()];
  let html = `
    <table>
      <thead>
        <tr>
          <th>No</th>
          <th>Customer</th>
          <th>Merk</th>
          <th>Jenis</th>
          <th>Ukuran</th>
          <th>Lebar</th>
          <th>Lipatan</th>
          <th>Tebal</th>
          <th>Panjang</th>
          <th>CH Roll</th>
          <th>CH Pot</th>
          <th>Lebar/Pack</th>
          <th>Isi/Ball</th>
        </tr>
      </thead>
      <tbody>
  `;

  let no = 1;

  keys.forEach((item) => {
    const row = item.row;
    const matches = state.orderRows.filter(
      (order) => text(order.MERK).toUpperCase() === text(row.MERK).toUpperCase() && text(order.JENIS).toUpperCase() === text(row.JENIS).toUpperCase() && text(order.UKURAN).toUpperCase() === text(row.UKURAN).toUpperCase(),
    );

    matches.forEach((s) => {
      html += `
      <tr>
        <td>${no++}</td>
        <td>${escapeHtml(text(s.CUSTOMER) || "-")}</td>
        <td>${escapeHtml(text(s.MERK) || "-")}</td>
        <td>${escapeHtml(text(s.JENIS) || "-")}</td>
        <td>${escapeHtml(text(s.UKURAN) || "-")}</td>
        <td>${escapeHtml(text(s.LEBAR) || "-")}</td>
        <td>${escapeHtml(text(s.LIPATAN) || "-")}</td>
        <td>${escapeHtml(text(s.TEBAL) || "-")}</td>
        <td>${escapeHtml(text(s.PANJANG) || "-")}</td>
        <td>${escapeHtml(text(s.CH_ROLL) || "-")}</td>
        <td>${escapeHtml(text(s.CH_POT) || "-")}</td>
        <td>${escapeHtml(text(s.LEBAR_PACK) || "-")}</td>
        <td>${escapeHtml(text(s.ISI_BALL) || "-")}</td>
      </tr>`;
    });
  });

  html += "</tbody></table>";

  tableArea.innerHTML = html;
  tableInfo.textContent = "Spek berdasarkan Rekap Hasil → DATA ORDER";
}

function parseHDHasil(matrix, filename) {
  const rows = [];
  let currentDate = null;

  for (let rowIndex = 4; rowIndex < matrix.length; rowIndex++) {
    const row = matrix[rowIndex] || [];
    const rawDate = row[0];

    if (rawDate !== null && rawDate !== undefined && String(rawDate).trim() !== "") {
      const number = Number(rawDate);

      if (Number.isFinite(number) && number >= 1 && number <= 31) {
        currentDate = number;
      }
    }

    if (currentDate === null) {
      continue;
    }

    if (!isValidDataRow(row)) {
      continue;
    }

    const object = {
      date: currentDate,

      MESIN: row[1],
      UKURAN: row[2],
      MERK: row[3],
      JENIS: row[4],

      A_OPERATOR: row[5],
      A_BALL: row[6],
      A_HASIL: row[7],
      A_TARGET: row[8],
      A_BS: row[9],

      B_OPERATOR: row[22],
      B_BALL: row[23],
      B_HASIL: row[24],
      B_TARGET: row[25],
      B_BS: row[26],

      TOTAL_HASIL: row[39],
    };

    rows.push(object);
  }

  state.rows = rows;

  state.dates = [...new Set(rows.map((row) => row.date))].sort((a, b) => a - b);

  state.activeDate = state.dates[state.dates.length - 1] ?? null;

  state.activeShift = "ALL";

  statusElement.textContent = `${filename} berhasil dimuat • ` + `${rows.length.toLocaleString("id-ID")} baris data • ` + `Sheet: HD HASIL`;

  setupFilters();
  updateShiftButtons();

  if (state.activeDate !== null) {
    setActiveDate(state.activeDate);
  }
}

function setupFilters() {
  dateFilter.innerHTML = "";

  state.dates.forEach((date) => {
    const option = document.createElement("option");

    option.value = date;
    option.textContent = `Tanggal ${date}`;

    dateFilter.appendChild(option);
  });

  dateFilter.disabled = false;
  dateInput.disabled = false;
  latestBtn.disabled = false;
  resetBtn.disabled = false;
}

dateFilter.addEventListener("change", (event) => {
  setActiveDate(Number(event.target.value));
});

dateInput.addEventListener("change", (event) => {
  const selectedDate = event.target.value;

  if (!selectedDate) {
    return;
  }

  const day = Number(selectedDate.slice(-2));

  if (state.dates.includes(day)) {
    setActiveDate(day);
  } else {
    alert(`Tidak ada data untuk tanggal ${day}.`);
  }
});

shiftButtons.forEach((button) => {
  button.addEventListener("click", () => {
    state.activeShift = button.dataset.shift;

    updateShiftButtons();
    renderDashboard();
  });
});

reportTabs.forEach((tab) => {
  tab.addEventListener("click", () => {
    state.activeView = tab.dataset.view;
    updateReportTabs();

    if (state.activeView === "deviasi") {
      loadDeviasiSheetJS();
    } else {
      renderDashboard();
    }
  });
});

if (recapShiftToggle) {
  recapShiftToggle.addEventListener("click", () => {
    state.recapShiftVisible = !state.recapShiftVisible;
    updateRecapShiftToggle();
    renderDashboard();
  });
}

function updateShiftButtons() {
  shiftButtons.forEach((button) => {
    button.classList.toggle("active", button.dataset.shift === state.activeShift);
  });
}

function updateReportTabs() {
  reportTabs.forEach((tab) => {
    const isActive = tab.dataset.view === state.activeView;
    tab.classList.toggle("active", isActive);
    tab.setAttribute("aria-selected", String(isActive));
  });

  updateRecapShiftToggle();
  updateActionVisibility();
}

function updateRecapShiftToggle() {
  if (!recapShiftToggle) {
    return;
  }

  const isRecap = state.activeView === "recap";

  recapShiftToggle.hidden = !isRecap;
  recapShiftToggle.textContent = state.recapShiftVisible ? "Sembunyikan Rekap Shift" : "Tampilkan Rekap Shift";
}

function updateActionVisibility() {
  const isDeviasi = state.activeView === "deviasi";
  const isDetail = state.activeView === "detail";

  if (shiftButtonsGroup) {
    shiftButtonsGroup.style.display = isDetail ? "flex" : "none";
  }

  if (recapShiftToggle) {
    recapShiftToggle.hidden = isDeviasi || state.activeView !== "recap";
  }
}

latestBtn.addEventListener("click", () => {
  if (!state.dates.length) {
    return;
  }

  const latest = state.dates[state.dates.length - 1];

  setActiveDate(latest);
});

resetBtn.addEventListener("click", () => {
  if (!state.dates.length) {
    return;
  }

  state.activeShift = "ALL";

  updateShiftButtons();

  const latest = state.dates[state.dates.length - 1];

  setActiveDate(latest);
});

function setActiveDate(date) {
  if (!state.dates.includes(date)) {
    return;
  }

  state.activeDate = date;

  dateFilter.value = String(date);

  renderDashboard();
}

function getResult(row) {
  if (state.activeShift === "A") {
    return toNumber(row.A_HASIL);
  }

  if (state.activeShift === "B") {
    return toNumber(row.B_HASIL);
  }

  return toNumber(row.TOTAL_HASIL);
}

function rowBelongsToShift(row) {
  if (state.activeShift === "ALL") {
    return true;
  }

  if (state.activeShift === "A") {
    return hasShiftA(row);
  }

  if (state.activeShift === "B") {
    return hasShiftB(row);
  }

  return true;
}

function renderDashboard() {
  if (state.activeDate === null) {
    cDate.textContent = "-";
    cRows.textContent = "0";
    cResult.textContent = "0";
    cShift.textContent = "Semua";
    tableInfo.textContent = "Belum ada data";
    datePill.textContent = "-";
    renderTable([]);
    return;
  }

  let rows = state.rows.filter((row) => row.date === state.activeDate);

  rows = rows.filter((row) => rowBelongsToShift(row));

  rows.sort(machineSort);

  const totalResult = rows.reduce((total, row) => {
    return total + getResult(row);
  }, 0);

  cDate.textContent = `Tgl ${state.activeDate}`;

  cRows.textContent = rows.length.toLocaleString("id-ID");

  cResult.textContent = formatNumber(totalResult);

  let shiftLabel = "Semua";

  if (state.activeShift === "A") {
    shiftLabel = "Shift A";
  }

  if (state.activeShift === "B") {
    shiftLabel = "Shift B";
  }

  cShift.textContent = shiftLabel;

  tableInfo.textContent = `${rows.length.toLocaleString("id-ID")} baris • ` + `Tanggal ${state.activeDate} • ` + shiftLabel;

  datePill.textContent = `TANGGAL ${state.activeDate}`;

  if (state.activeView === "recap") {
    renderRecap();
  } else if (state.activeView === "deviasi") {
    loadDeviasiSheetJS();
  } else {
    renderTable(rows);
  }
}

function displayResult(value) {
  return toNumber(value) === 0 ? "-" : formatNumber(value);
}

function renderRecap() {
  const dateRows = state.rows.filter((row) => row.date === state.activeDate);
  const selectedResults = getSelectedResults(dateRows);
  const selectedRows = [...selectedResults.values()];
  const shiftAEntries = state.rows.filter((row) => row.date === state.activeDate && hasShiftA(row));
  const shiftBEntries = state.rows.filter((row) => row.date === state.activeDate && hasShiftB(row));
  const totalRows = state.recapShiftVisible ? Math.max(selectedRows.length, shiftAEntries.length, shiftBEntries.length) : selectedRows.length;
  const selectedColumnCount = 10;
  const recapColumnCount = state.recapShiftVisible ? 4 : 0;
  const totalColumnCount = selectedColumnCount + recapColumnCount;

  tableInfo.textContent = `Rekap hasil terpilih • Tanggal ${state.activeDate} • Semua shift`;

  if (!totalRows) {
    renderTable([]);
    return;
  }

  let html = `
    <table class="recap-table">
      <thead>
        <tr class="recap-meta"><th colspan="${totalColumnCount}">Tanggal: ${state.activeDate}</th></tr>
        <tr class="recap-section">
          <th colspan="${selectedColumnCount}">HASIL TERPILIH (${selectedRows.length} MERK)</th>
          ${state.recapShiftVisible ? '<th colspan="4" class="recap-divider">REKAP SHIFT</th>' : ""}
        </tr>
        <tr>
          <th>No</th><th>Mesin</th><th>Ukuran</th><th>Jenis</th><th>Merk Terpilih</th><th>Hasil Shift A</th><th>BS A</th><th>Hasil Shift B</th><th>BS B</th><th>Total Hasil</th>
          ${state.recapShiftVisible ? '<th class="recap-divider">Merk Shift A</th><th>Merk Shift B</th><th>Total A</th><th>Total B</th>' : ""}
        </tr>
      </thead>
      <tbody>
  `;

  for (let index = 0; index < totalRows; index++) {
    const selected = selectedRows[index];
    const shiftA = shiftAEntries[index];
    const shiftB = shiftBEntries[index];
    const selectedTotal = selected ? selected.resultA + selected.resultB : 0;

    html += `
      <tr>
        <td>${selected ? index + 1 : ""}</td>
        <td>${selected ? escapeHtml(text(selected.row.MESIN) || "-") : ""}</td>
        <td>${selected ? escapeHtml(text(selected.row.UKURAN) || "-") : ""}</td>
        <td>${selected ? escapeHtml(text(selected.row.JENIS) || "-") : ""}</td>
        <td>${selected ? escapeHtml(text(selected.row.MERK) || "-") : ""}</td>
        <td class="number">${selected ? displayResult(selected.resultA) : ""}</td>
        <td class="number">${selected ? formatNumber(selected.bsA) : ""}</td>
        <td class="number">${selected ? displayResult(selected.resultB) : ""}</td>
        <td class="number">${selected ? formatNumber(selected.bsB) : ""}</td>
        <td class="number">${selected ? displayResult(selectedTotal) : ""}</td>
        ${
          state.recapShiftVisible
            ? `
          <td class="recap-divider">${shiftA ? escapeHtml(text(shiftA.MERK) || "-") : ""}</td>
          <td>${shiftB ? escapeHtml(text(shiftB.MERK) || "-") : ""}</td>
          <td class="number">${shiftA ? formatNumber(shiftA.A_HASIL) : ""}</td>
          <td class="number">${shiftB ? formatNumber(shiftB.B_HASIL) : ""}</td>
        `
            : ""
        }
      </tr>
    `;
  }

  tableArea.innerHTML = `${html}</tbody></table>`;
}

function getSelectedResults(rows) {
  const machines = new Map();

  rows.forEach((row, index) => {
    const machine = text(row.MESIN).toUpperCase();
    const brand = text(row.MERK).toUpperCase();

    if (!machine || !brand) {
      return;
    }

    if (!machines.has(machine)) {
      machines.set(machine, {
        brands: new Map(),
        totalA: 0,
        totalB: 0,
        bsA: 0,
        bsB: 0,
      });
    }

    const machineData = machines.get(machine);
    const brands = machineData.brands;

    if (!brands.has(brand)) {
      brands.set(brand, {
        row,
        rowA: null,
        rowB: null,
        firstIndex: index,
        maxA: 0,
        maxB: 0,
        indexA: null,
        indexB: null,
      });
    }

    const item = brands.get(brand);
    const resultA = toNumber(row.A_HASIL);
    const resultB = toNumber(row.B_HASIL);

    machineData.bsA += toNumber(row.A_BS);
    machineData.bsB += toNumber(row.B_BS);

    if (hasShiftA(row)) {
      machineData.totalA += resultA;

      if (item.indexA === null || resultA > item.maxA) {
        item.maxA = resultA;
        item.indexA = index;
        item.rowA = row;
      }
    }

    if (hasShiftB(row)) {
      machineData.totalB += resultB;

      if (item.indexB === null || resultB > item.maxB) {
        item.maxB = resultB;
        item.indexB = index;
        item.rowB = row;
      }
    }
  });

  const selected = new Map();

  machines.forEach((machineData) => {
    const { brands, totalA, totalB, bsA, bsB } = machineData;
    let bestA = null;
    let bestB = null;

    brands.forEach((item) => {
      if (item.indexA !== null && (!bestA || item.maxA > bestA.maxA)) {
        bestA = item;
      }

      if (item.indexB !== null && (!bestB || item.maxB > bestB.maxB)) {
        bestB = item;
      }
    });

    [bestA, bestB].filter(Boolean).forEach((item) => {
      const outputIndex = item.indexA ?? item.indexB ?? item.firstIndex;
      const existing = selected.get(outputIndex);

      if (existing) {
        existing.resultA = item === bestA ? totalA : existing.resultA;
        existing.resultB = item === bestB ? totalB : existing.resultB;
        existing.bsA = item === bestA ? bsA : existing.bsA;
        existing.bsB = item === bestB ? bsB : existing.bsB;
        return;
      }

      selected.set(outputIndex, {
        row: item.rowA ?? item.rowB ?? item.row,
        resultA: item === bestA ? totalA : 0,
        resultB: item === bestB ? totalB : 0,
        bsA: item === bestA ? bsA : 0,
        bsB: item === bestB ? bsB : 0,
      });
    });
  });

  return selected;
}

function renderTable(rows) {
  if (!rows.length) {
    tableArea.innerHTML = `
      <div class="empty-state">
        <div class="empty-icon">📭</div>
        <h3>Tidak ada data</h3>
        <p>
          Tidak ada data untuk tanggal dan shift yang dipilih.
        </p>
      </div>
    `;

    return;
  }

  let html = `
    <table>
      <thead>
        <tr>
          <th>TGL</th>
          <th>Mesin</th>
          <th>Ukuran</th>
          <th>Merk</th>
          <th>Jenis</th>
  `;

  if (state.activeShift === "A" || state.activeShift === "ALL") {
    html += `
      <th>Operator A</th>
      <th>Ball A</th>
      <th>Hasil A</th>
    `;
  }

  if (state.activeShift === "B" || state.activeShift === "ALL") {
    html += `
      <th>Operator B</th>
      <th>Ball B</th>
      <th>Hasil B</th>
    `;
  }

  html += `
          <th>Total Hasil</th>
        </tr>
      </thead>
      <tbody>
  `;

  rows.forEach((row) => {
    html += `
      <tr>
        <td>${row.date}</td>

        <td>
          ${escapeHtml(text(row.MESIN) || "-")}
        </td>

        <td>
          ${escapeHtml(text(row.UKURAN) || "-")}
        </td>

        <td>
          ${escapeHtml(text(row.MERK) || "-")}
        </td>

        <td>
          ${escapeHtml(text(row.JENIS) || "-")}
        </td>
    `;

    if (state.activeShift === "A" || state.activeShift === "ALL") {
      html += `
        <td>
          ${escapeHtml(text(row.A_OPERATOR) || "-")}
        </td>

        <td class="number">
          ${formatNumber(row.A_BALL)}
        </td>

        <td class="number">
          ${formatNumber(row.A_HASIL)}
        </td>
      `;
    }

    if (state.activeShift === "B" || state.activeShift === "ALL") {
      html += `
        <td>
          ${escapeHtml(text(row.B_OPERATOR) || "-")}
        </td>

        <td class="number">
          ${formatNumber(row.B_BALL)}
        </td>

        <td class="number">
          ${formatNumber(row.B_HASIL)}
        </td>
      `;
    }

    html += `
        <td class="number total-cell">
          ${formatNumber(getResult(row))}
        </td>

      </tr>
    `;
  });

  html += `
      </tbody>
    </table>
  `;

  tableArea.innerHTML = html;
}

function loadDeviasiSheetJS() {
  showLoading("Memuat data Deviasi...");

  try {
    if (!state.originalFile) {
      tableArea.innerHTML = `
        <div class="empty-state">
          <h3>Belum ada file</h3>
          <p>Upload file Excel terlebih dahulu untuk menampilkan Deviasi.</p>
        </div>
      `;
      return;
    }

    if (state.deviasiLoaded && state.deviasiSheet) {
      hideLoading();
      renderDeviasi();
      return;
    }

    const reader = new FileReader();

    reader.onload = (event) => {
      try {
        const workbook = XLSX.read(new Uint8Array(event.target.result), {
          type: "array",
          cellDates: true,
        });

        const sheet = workbook.Sheets["DEVIASI"];

        if (!sheet) {
          state.deviasiLoaded = true;
          state.deviasiSheet = null;
          renderDeviasi();
          return;
        }

        state.deviasiLoaded = true;
        state.deviasiSheet = sheet;
        renderDeviasi();
      } catch (error) {
        console.error(error);
        tableArea.innerHTML = `<div class="empty-state"><h3>Gagal memuat Deviasi</h3><p>${escapeHtml(error.message)}</p></div>`;
      } finally {
        hideLoading();
      }
    };

    reader.readAsArrayBuffer(state.originalFile);
  } catch (error) {
    hideLoading();
    console.error(error);
    tableArea.innerHTML = `<div class="empty-state"><h3>Gagal memuat Deviasi</h3><p>${escapeHtml(error.message)}</p></div>`;
  }
}

function buildDeviasiJumpButtons(range) {
  const sections = [
    { label: "Info Mesin", col: 0 },
    { label: "Target", col: Math.min(range.e.c, XLSX.utils.decode_col("Y")) },
    { label: "Shift A", col: Math.min(range.e.c, XLSX.utils.decode_col("AF")) },
    { label: "Downtime A", col: Math.min(range.e.c, XLSX.utils.decode_col("AK")) },
    { label: "Shift B", col: Math.min(range.e.c, XLSX.utils.decode_col("BE")) },
  ];

  const unique = [];
  const seen = new Set();
  sections.forEach((item) => {
    if (item.col >= range.s.c && item.col <= range.e.c && !seen.has(item.col)) {
      seen.add(item.col);
      unique.push(item);
    }
  });

  return unique.map((item) => `<button type="button" class="deviasi-jump-button" data-jump-col="${item.col}">${item.label}</button>`).join("");
}

function formatPercentage(value, total) {
  const totalNumber = toNumber(total);

  if (totalNumber === 0) {
    return "-";
  }

  return `${formatNumber((toNumber(value) / totalNumber) * 100)}%`;
}

function buildDeviasiRecap() {
  const selectedRows = [...getSelectedResults(state.rows.filter((row) => row.date === state.activeDate)).values()];

  if (!selectedRows.length) {
    return `
      <section class="deviasi-recap">
        <h3>Rekap Produksi untuk Deviasi</h3>
        <p>Tidak ada hasil terpilih pada tanggal aktif.</p>
      </section>
    `;
  }

  let rowsHtml = "";

  selectedRows.forEach((item, index) => {
    rowsHtml += `
      <tr>
        <td>${index + 1}</td>
        <td>${escapeHtml(text(item.row.MESIN) || "-")}</td>
        <td>${escapeHtml(text(item.row.MERK) || "-")}</td>
        <td>${escapeHtml(text(item.row.JENIS) || "-")}</td>
        <td>${escapeHtml(text(item.row.UKURAN) || "-")}</td>
        <td class="number">${displayResult(item.resultA)}</td>
        <td class="number">${formatNumber(item.bsA)}</td>
        <td class="number">${formatPercentage(item.bsA, item.resultA)}</td>
        <td class="number">${displayResult(item.resultB)}</td>
        <td class="number">${formatNumber(item.bsB)}</td>
        <td class="number">${formatPercentage(item.bsB, item.resultB)}</td>
      </tr>
    `;
  });

  return `
    <section class="deviasi-recap">
      <div class="deviasi-recap-heading">
        <div>
          <h3>Deviasi Produksi — Rekap Terpilih</h3>
          <p>Data ini diambil dari Rekap Hasil tanggal ${state.activeDate}. BS berasal dari kolom PLONG pada HD HASIL.</p>
        </div>
        <div class="deviasi-formula">BS % = BS KG ÷ Hasil KG × 100%</div>
      </div>
      <div class="deviasi-recap-wrap">
        <table class="deviasi-recap-table">
          <thead>
            <tr>
              <th rowspan="2">No</th><th rowspan="2">Mesin</th><th rowspan="2">Merk</th><th rowspan="2">Jenis</th><th rowspan="2">Ukuran</th>
              <th colspan="3" class="deviasi-shift-a">SHIFT A</th>
              <th colspan="3" class="deviasi-shift-b">SHIFT B</th>
            </tr>
            <tr>
              <th class="deviasi-shift-a">Hasil KG</th><th class="deviasi-shift-a">BS KG</th><th class="deviasi-shift-a">BS %</th>
              <th class="deviasi-shift-b">Hasil KG</th><th class="deviasi-shift-b">BS KG</th><th class="deviasi-shift-b">BS %</th>
            </tr>
          </thead>
          <tbody>${rowsHtml}</tbody>
        </table>
      </div>
    </section>
  `;
}

function setDeviasiCell(sheet, address, value) {
  const cell = XLSX.utils.encode_cell(address);
  const numeric = typeof value === "number" && Number.isFinite(value);

  sheet[cell] = numeric ? { t: "n", v: value, z: "0.00", w: formatNumber(value) } : { t: "s", v: text(value), w: text(value) };
}

function clearDeviasiCell(sheet, address) {
  const cell = XLSX.utils.encode_cell(address);
  delete sheet[cell];
}

function applyRecapToDeviasiSheet(sheet) {
  const selectedRows = [...getSelectedResults(state.rows.filter((row) => row.date === state.activeDate)).values()];
  const startRow = 5;
  const endRow = Math.max(startRow + selectedRows.length, XLSX.utils.decode_range(sheet["!ref"]).e.r + 1);

  for (let rowIndex = startRow; rowIndex < endRow; rowIndex++) {
    [0, 1, 2, 33, 38, 58, 63].forEach((column) => {
      clearDeviasiCell(sheet, { r: rowIndex, c: column });
    });
  }

  selectedRows.forEach((item, index) => {
    const rowIndex = startRow + index;
    const row = item.row;

    setDeviasiCell(sheet, { r: rowIndex, c: 0 }, row.MESIN);
    setDeviasiCell(sheet, { r: rowIndex, c: 1 }, row.MERK);
    setDeviasiCell(sheet, { r: rowIndex, c: 2 }, row.JENIS);
    setDeviasiCell(sheet, { r: rowIndex, c: 33 }, item.resultA);
    setDeviasiCell(sheet, { r: rowIndex, c: 38 }, item.bsA);
    setDeviasiCell(sheet, { r: rowIndex, c: 58 }, item.resultB);
    setDeviasiCell(sheet, { r: rowIndex, c: 63 }, item.bsB);
  });
}

function renderDeviasi() {
  if (!state.deviasiSheet) {
    tableArea.innerHTML = `<div class="empty-state"><h3>Sheet DEVIASI tidak ditemukan</h3></div>`;
    tableInfo.textContent = "Sheet DEVIASI tidak tersedia";
    return;
  }

  const sheet = state.deviasiSheet;
  applyRecapToDeviasiSheet(sheet);
  const range = XLSX.utils.decode_range(sheet["!ref"]);
  const merges = sheet["!merges"] || [];
  const cols = sheet["!cols"] || [];

  const skipMap = new Set();
  const spanMap = new Map();

  merges.forEach((merge) => {
    const key = `${merge.s.r}:${merge.s.c}`;
    spanMap.set(key, {
      rowspan: merge.e.r - merge.s.r + 1,
      colspan: merge.e.c - merge.s.c + 1,
    });

    for (let r = merge.s.r; r <= merge.e.r; r++) {
      for (let c = merge.s.c; c <= merge.e.c; c++) {
        if (r === merge.s.r && c === merge.s.c) continue;
        skipMap.add(`${r}:${c}`);
      }
    }
  });

  let colgroup = "<colgroup>";
  for (let c = range.s.c; c <= range.e.c; c++) {
    const width = cols[c]?.wpx || (cols[c]?.wch ? Math.max(64, cols[c].wch * 8) : 84);
    colgroup += `<col style="width:${width}px">`;
  }
  colgroup += "</colgroup>";

  let table = `<table class="deviasi-grid-table">${colgroup}<tbody>`;

  for (let r = range.s.r; r <= range.e.r; r++) {
    const rowType = r === 0 ? "deviasi-row-title" : r <= 4 ? "deviasi-row-header" : "deviasi-row-body";
    table += `<tr class="${rowType}">`;

    for (let c = range.s.c; c <= range.e.c; c++) {
      const posKey = `${r}:${c}`;
      if (skipMap.has(posKey)) continue;

      const cellAddress = XLSX.utils.encode_cell({ r, c });
      const cell = sheet[cellAddress];
      const value = cell ? (cell.w ?? cell.v ?? "") : "";
      const span = spanMap.get(posKey);
      const tag = r <= 4 ? "th" : "td";
      const attributes = [];
      const classes = ["deviasi-cell", `col-${c}`];

      if (span) {
        if (span.rowspan > 1) attributes.push(`rowspan="${span.rowspan}"`);
        if (span.colspan > 1) attributes.push(`colspan="${span.colspan}"`);
      }

      if (r === 1 || c === XLSX.utils.decode_col("AF")) {
        classes.push("deviasi-section-start");
      }

      if (c === XLSX.utils.decode_col("BE")) {
        classes.push("deviasi-shift-b-start");
      }

      if (c === 0) classes.push("deviasi-col-first");
      if (c === 1) classes.push("deviasi-col-merk");
      if (r === 0) classes.push("deviasi-title-cell");
      if (r > 0 && r <= 4) classes.push("deviasi-head-cell");

      table += `<${tag} class="${classes.join(" ")}" data-col="${c}" ${attributes.join(" ")}>${escapeHtml(value)}</${tag}>`;
    }

    table += "</tr>";
  }

  table += "</tbody></table>";

  const totalRows = range.e.r + 1;
  const totalCols = range.e.c + 1;
  const dataRows = Math.max(0, totalRows - 5);

  tableArea.innerHTML = `
    <section class="deviasi-view">
      <div class="deviasi-overview">
        <div>
          <h3>Sheet DEVIASI</h3>
          <p>Baris MESIN, MERK, JENIS, HASIL KG, dan BS KG diisi langsung dari Rekap Hasil. Kolom target, kapasitas, dan downtime tetap mengikuti logika sheet DEVIASI Excel. Shift A dimulai pada kolom biru, Shift B pada kolom hijau.</p>
        </div>
        <div class="deviasi-stats">
          <div class="deviasi-stat"><span>Baris Data</span><strong>${dataRows}</strong></div>
          <div class="deviasi-stat"><span>Total Kolom</span><strong>${totalCols}</strong></div>
          <div class="deviasi-stat"><span>Merge Cell</span><strong>${merges.length}</strong></div>
        </div>
      </div>

      <div class="deviasi-toolbar">
        <div class="deviasi-toolbar-label">Navigasi cepat:</div>
        <div class="deviasi-toolbar-buttons">
          ${buildDeviasiJumpButtons(range)}
        </div>
      </div>

      <div class="deviasi-grid-wrap">
        ${table}
      </div>
    </section>
  `;

  tableInfo.textContent = `Sheet DEVIASI • ${dataRows} baris data • ${totalCols} kolom`;
  bindDeviasiJumpButtons();
}

function bindDeviasiJumpButtons() {
  const scroller = tableArea.querySelector(".deviasi-grid-wrap");
  const buttons = tableArea.querySelectorAll(".deviasi-jump-button");
  if (!scroller || !buttons.length) return;

  buttons.forEach((button) => {
    button.addEventListener("click", () => {
      const targetCol = Number(button.dataset.jumpCol);
      const targetCell = scroller.querySelector(`[data-col="${targetCol}"]`);
      if (!targetCell) return;
      scroller.scrollTo({ left: Math.max(0, targetCell.offsetLeft - 24), behavior: "smooth" });
    });
  });
}

// ===============================
// EXPORT REKAP & DEVIASI
// ===============================

const exportRecapBtn = document.getElementById("exportRecapBtn");
const exportDeviasiBtn = document.getElementById("exportDeviasiBtn");

if (exportRecapBtn) {
  exportRecapBtn.addEventListener("click", exportRekap);
}

if (exportDeviasiBtn) {
  exportDeviasiBtn.addEventListener("click", exportDeviasi);
}

function exportRekap() {
  const rows = [...getSelectedResults(state.rows.filter((row) => row.date === state.activeDate)).values()];

  const data = [["No", "Mesin", "Ukuran", "Jenis", "Merk", "Hasil A", "BS A", "Hasil B", "BS B", "Total"]];

  rows.forEach((item, index) => {
    data.push([index + 1, item.row.MESIN, item.row.UKURAN, item.row.JENIS, item.row.MERK, item.resultA, item.bsA, item.resultB, item.bsB, item.resultA + item.resultB]);
  });

  const ws = XLSX.utils.aoa_to_sheet(data);
  rows.forEach((item, rowIndex) => {
    for (let column = 5; column <= 9; column++) {
      const cell = ws[XLSX.utils.encode_cell({ r: rowIndex + 1, c: column })];

      if (cell && typeof cell.v === "number") {
        cell.z = "0.00";
        cell.w = formatNumber(cell.v);
      }
    }
  });
  const wb = XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(wb, ws, "REKAP");

  XLSX.writeFile(wb, `REKAP_${state.activeDate}.xlsx`, {
    bookType: "xlsx",
    cellStyles: true,
  });
}

function exportDeviasi() {
  if (!state.originalFile) {
    alert("Upload file terlebih dahulu.");
    return;
  }

  const reader = new FileReader();

  reader.onload = (event) => {
    const workbook = XLSX.read(new Uint8Array(event.target.result), {
      type: "array",
      cellStyles: true,
    });

    const sheet = workbook.Sheets["DEVIASI"];

    if (!sheet) {
      alert("Sheet DEVIASI tidak ditemukan.");
      return;
    }

    applyRecapToDeviasiSheet(sheet);

    XLSX.writeFile(workbook, `DEVIASI_${state.activeDate}.xlsx`, {
      bookType: "xlsx",
      cellStyles: true,
    });
  };

  reader.readAsArrayBuffer(state.originalFile);
}
