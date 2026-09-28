"use strict";

const DEFAULT_ANALYSIS_DATE = "2026-02-01";
const REFERENCE_WORKBOOK = "Reference/AR-Data.xlsx";
const AGING_ORDER = ["current", "30-59 days", "60-89 days", "over 90 days"];
const AGING_LABELS = {
  current: "Current",
  "30-59 days": "30–59 days",
  "60-89 days": "60–89 days",
  "over 90 days": "Over 90 days"
};
const AGING_COLORS = ["#BC9EC0", "#D2C0E6", "#90CBF3", "#67C5E9"];
const TYPE_COLORS = ["#BC9EC0", "#8AADD9", "#67C5E9", "#A8C8B8", "#D8AA8C", "#8F86B8", "#6BA9A2"];

const appState = {
  workbookName: null,
  rawSales: [],
  rawCustomerTypes: [],
  joinedRows: [],
  processedData: [],
  analysisDate: parseInputDate(DEFAULT_ANALYSIS_DATE),
  activeTab: "one-dimensional",
  charts: {},
  errors: [],
  quality: { valid: 0, excluded: 0, reasons: {} },
  tableSearch: "",
  tableFilter: "all",
  tableSort: { key: "customer", direction: "asc" },
  source: null
};

const columnAliases = {
  customer: ["customer", "customernumber", "customerid", "cust", "custno"],
  salesPerson: ["salesperson", "salesrepresentative", "salesrep", "salespersonname"],
  balance: ["bal", "balance", "accountbalance", "amount", "outstandingbalance", "arbalance"],
  invoiceDate: ["invdate", "invoicedate", "dateofinvoice"],
  customerType: ["custtype", "customertype", "type", "customercategory"]
};

document.addEventListener("DOMContentLoaded", initializeApp);

function initializeApp() {
  bindTabs();
  bindControls();
  bindWorkflowNodes();
  bindTableControls();
  updateDateLabels();
  tryLoadReferenceWorkbook();
}

function bindTabs() {
  const tabs = [...document.querySelectorAll("[role='tab']")];
  tabs.forEach((tab, index) => {
    tab.addEventListener("click", () => switchTab(tab.dataset.tab));
    tab.addEventListener("keydown", event => {
      if (!["ArrowLeft", "ArrowRight", "Home", "End"].includes(event.key)) return;
      event.preventDefault();
      let nextIndex = index;
      if (event.key === "ArrowLeft") nextIndex = (index - 1 + tabs.length) % tabs.length;
      if (event.key === "ArrowRight") nextIndex = (index + 1) % tabs.length;
      if (event.key === "Home") nextIndex = 0;
      if (event.key === "End") nextIndex = tabs.length - 1;
      tabs[nextIndex].focus();
      switchTab(tabs[nextIndex].dataset.tab);
    });
  });
}

function switchTab(tabName) {
  appState.activeTab = tabName;
  document.querySelectorAll("[role='tab']").forEach(tab => {
    const active = tab.dataset.tab === tabName;
    tab.classList.toggle("is-active", active);
    tab.setAttribute("aria-selected", String(active));
    tab.tabIndex = active ? 0 : -1;
  });
  document.querySelectorAll("[role='tabpanel']").forEach(panel => {
    const active = panel.dataset.panel === tabName;
    panel.hidden = !active;
    panel.classList.toggle("is-active", active);
  });
  requestAnimationFrame(renderCharts);
}

function bindControls() {
  const upload = document.getElementById("workbook-upload");
  const dateInput = document.getElementById("analysis-date");

  upload.addEventListener("change", async event => {
    const file = event.target.files?.[0];
    if (!file) return;
    if (!/\.(xlsx|xls)$/i.test(file.name)) {
      showError("Choose an Excel workbook with an .xlsx or .xls extension.");
      return;
    }
    try {
      await loadWorkbook(await file.arrayBuffer(), file.name, "upload");
    } catch (error) {
      showError(error.message || "The selected workbook is not compatible with this analysis.");
      setReadyStatus();
    }
  });

  dateInput.addEventListener("change", event => {
    const nextDate = parseInputDate(event.target.value);
    if (!nextDate) {
      showError("Enter a valid analysis date.");
      event.target.value = toInputDate(appState.analysisDate);
      return;
    }
    appState.analysisDate = nextDate;
    hideError();
    updateDateLabels();
    if (appState.joinedRows.length) {
      processRecords();
      renderDashboard();
    }
  });

  document.getElementById("reset-analysis").addEventListener("click", resetAnalysis);
  document.getElementById("dismiss-error").addEventListener("click", hideError);

  const qualityButton = document.getElementById("quality-button");
  const qualityPopover = document.getElementById("quality-popover");
  qualityButton.addEventListener("click", () => {
    const opening = qualityPopover.hidden;
    qualityPopover.hidden = !opening;
    qualityButton.setAttribute("aria-expanded", String(opening));
  });
  document.getElementById("quality-close").addEventListener("click", () => {
    qualityPopover.hidden = true;
    qualityButton.setAttribute("aria-expanded", "false");
    qualityButton.focus();
  });
}

function bindWorkflowNodes() {
  document.querySelectorAll(".workflow-node").forEach(node => {
    node.addEventListener("click", () => {
      const panel = node.closest(".workflow-panel");
      panel.querySelectorAll(".workflow-node").forEach(item => item.classList.toggle("is-selected", item === node));
      const detail = panel.querySelector(".workflow-detail");
      detail.querySelector("strong").textContent = node.dataset.title;
      detail.querySelector("p").textContent = node.dataset.detail;
    });
  });
}

function bindTableControls() {
  document.getElementById("table-search").addEventListener("input", event => {
    appState.tableSearch = event.target.value.trim().toLocaleLowerCase();
    renderTransactionTable();
  });
  document.getElementById("classification-filter").addEventListener("change", event => {
    appState.tableFilter = event.target.value;
    renderTransactionTable();
  });
  document.querySelectorAll("#transaction-table [data-sort]").forEach(button => {
    button.addEventListener("click", () => {
      const key = button.dataset.sort;
      if (appState.tableSort.key === key) {
        appState.tableSort.direction = appState.tableSort.direction === "asc" ? "desc" : "asc";
      } else {
        appState.tableSort = { key, direction: "asc" };
      }
      renderTransactionTable();
    });
  });
}

async function tryLoadReferenceWorkbook() {
  if (typeof XLSX === "undefined") {
    showError("The Excel processing library could not be loaded. Check your internet connection and reload the page.");
    return;
  }
  setLoadingStatus("Loading reference workbook…");
  try {
    const response = await fetch(REFERENCE_WORKBOOK, { cache: "no-store" });
    if (!response.ok) throw new Error("Reference workbook is unavailable.");
    await loadWorkbook(await response.arrayBuffer(), "AR-Data.xlsx", "reference");
  } catch (error) {
    clearDataset();
    setEmptyStatus();
  }
}

async function loadWorkbook(arrayBuffer, workbookName, source) {
  if (typeof XLSX === "undefined") throw new Error("The Excel processing library is unavailable.");
  setLoadingStatus(`Reading ${workbookName}…`);
  hideError();

  let workbook;
  try {
    workbook = XLSX.read(arrayBuffer, { type: "array", cellDates: true, cellNF: true });
  } catch (error) {
    throw new Error("The file could not be read as an Excel workbook.");
  }

  const parsed = validateWorkbook(workbook);
  const combined = combineDatasets(parsed.salesRows, parsed.customerTypeRows, parsed.salesHasCustomerType);

  appState.workbookName = workbookName;
  appState.rawSales = parsed.salesRows;
  appState.rawCustomerTypes = parsed.customerTypeRows;
  appState.joinedRows = combined;
  appState.source = source;
  appState.tableSearch = "";
  appState.tableFilter = "all";
  document.getElementById("table-search").value = "";
  document.getElementById("classification-filter").value = "all";

  processRecords();
  if (!appState.processedData.length) {
    throw new Error("No valid AR transaction records were found. Review the Data Quality requirements and workbook fields.");
  }
  renderDashboard();
}

function validateWorkbook(workbook) {
  if (!workbook.SheetNames?.length) throw new Error("The workbook does not contain any worksheets.");

  const sheets = workbook.SheetNames.map(name => {
    const extracted = extractSheetRows(workbook.Sheets[name]);
    return { name, normalizedName: normalizeHeader(name), ...extracted };
  });

  const preferredSales = sheets.find(sheet => /breakdown.*sales|sales.*breakdown/.test(sheet.normalizedName));
  const salesSheet = preferredSales || sheets.find(sheet => hasColumns(sheet.headerMap, ["customer", "balance", "invoiceDate"]));
  if (!salesSheet) {
    throw new Error("Missing transaction worksheet or required fields: Customer, Balance, and Invoice Date.");
  }

  const requiredSales = ["customer", "salesPerson", "balance", "invoiceDate"];
  const missingSales = requiredSales.filter(field => salesSheet.headerMap[field] === undefined);
  if (missingSales.length) {
    throw new Error(`Missing required field${missingSales.length > 1 ? "s" : ""}: ${missingSales.map(displayFieldName).join(", ")}.`);
  }

  const salesHasCustomerType = salesSheet.headerMap.customerType !== undefined;
  let typeSheet = null;
  if (!salesHasCustomerType) {
    const preferredType = sheets.find(sheet => sheet !== salesSheet && /customertype|custtype/.test(sheet.normalizedName));
    typeSheet = preferredType || sheets.find(sheet => sheet !== salesSheet && hasColumns(sheet.headerMap, ["customer", "customerType"]));
    if (!typeSheet) throw new Error("Missing worksheet: Customer Type.");
    const missingType = ["customer", "customerType"].filter(field => typeSheet.headerMap[field] === undefined);
    if (missingType.length) throw new Error(`Missing required field: ${displayFieldName(missingType[0])}.`);
  }

  return {
    salesRows: mapSheetRows(salesSheet),
    customerTypeRows: typeSheet ? mapSheetRows(typeSheet) : [],
    salesHasCustomerType
  };
}

function extractSheetRows(sheet) {
  const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: null, blankrows: true });
  const headerIndex = matrix.findIndex(row => Array.isArray(row) && row.some(value => !isBlank(value)));
  if (headerIndex < 0) return { headers: [], headerMap: {}, rawRows: [] };

  const headers = matrix[headerIndex].map(value => String(value ?? "").trim());
  const headerMap = {};
  headers.forEach((header, index) => {
    const canonical = canonicalField(header);
    if (canonical && headerMap[canonical] === undefined) headerMap[canonical] = index;
  });

  const rawRows = matrix.slice(headerIndex + 1).map((values, offset) => ({
    values,
    rowNumber: headerIndex + offset + 2,
    blank: !values.some(value => !isBlank(value))
  }));
  return { headers, headerMap, rawRows };
}

function mapSheetRows(sheetInfo) {
  return sheetInfo.rawRows.map(row => {
    const mapped = { __rowNumber: row.rowNumber, __blank: row.blank };
    Object.entries(sheetInfo.headerMap).forEach(([field, index]) => { mapped[field] = row.values[index]; });
    return mapped;
  });
}

function combineDatasets(salesRows, typeRows, salesHasCustomerType) {
  if (salesHasCustomerType) return salesRows.map(row => ({ ...row }));

  const usableSales = salesRows.filter(row => !row.__blank);
  const usableTypes = typeRows.filter(row => !row.__blank);
  const positional = usableSales.length === usableTypes.length && usableSales.every((row, index) => {
    return normalizeKey(row.customer) === normalizeKey(usableTypes[index].customer);
  });

  if (positional) {
    let typeIndex = 0;
    return salesRows.map(row => {
      if (row.__blank) return { ...row };
      const typeRow = usableTypes[typeIndex++];
      return { ...row, customerType: typeRow.customerType, __joinMode: "record-position" };
    });
  }

  const typeMap = new Map();
  usableTypes.forEach(row => {
    const key = normalizeKey(row.customer);
    const value = cleanText(row.customerType);
    if (!key || !value) return;
    if (!typeMap.has(key)) typeMap.set(key, value);
  });
  return salesRows.map(row => ({
    ...row,
    customerType: row.__blank ? null : typeMap.get(normalizeKey(row.customer)) ?? null,
    __joinMode: "customer-key"
  }));
}

function processRecords() {
  const reasons = {};
  const records = [];

  appState.joinedRows.forEach(row => {
    if (row.__blank) {
      addReason(reasons, "Completely blank row");
      return;
    }

    const customer = formatCustomer(row.customer);
    const salesPerson = cleanText(row.salesPerson);
    const customerType = cleanText(row.customerType);
    const balance = parseBalance(row.balance);
    const invoiceDate = parseExcelDate(row.invoiceDate);

    if (!customer) return addReason(reasons, "Missing Customer");
    if (!salesPerson) return addReason(reasons, "Missing Sales Person");
    if (balance === null) return addReason(reasons, "Invalid Balance");
    if (!invoiceDate) return addReason(reasons, "Invalid Invoice Date");
    if (!customerType) return addReason(reasons, "Missing Customer Type");

    const aging = calculateAging(invoiceDate, appState.analysisDate);
    records.push({
      sourceRow: row.__rowNumber,
      customer,
      salesPerson,
      balance,
      invoiceDate,
      customerType,
      aging,
      classification: classifyAging(aging)
    });
  });

  appState.processedData = records;
  appState.quality = {
    valid: records.length,
    excluded: Object.values(reasons).reduce((sum, value) => sum + value, 0),
    reasons
  };
  reconcileOutputs();
}

function calculateAging(invoiceDate, analysisDate) {
  return Math.floor((utcDay(analysisDate) - utcDay(invoiceDate)) / 86400000);
}

function classifyAging(aging) {
  if (aging < 30) return "current";
  if (aging < 60) return "30-59 days";
  if (aging < 90) return "60-89 days";
  return "over 90 days";
}

function calculateKPIs() {
  const total = sum(appState.processedData.map(row => row.balance));
  const count = appState.processedData.length;
  const over90 = sum(appState.processedData.filter(row => row.classification === "over 90 days").map(row => row.balance));
  return { total, count, average: count ? total / count : 0, over90, over90Share: total ? over90 / total : 0 };
}

function buildAgingSummary() {
  const totalAR = sum(appState.processedData.map(row => row.balance));
  return AGING_ORDER.map(category => {
    const rows = appState.processedData.filter(row => row.classification === category);
    const balances = rows.map(row => row.balance);
    const categoryTotal = sum(balances);
    return {
      category,
      total: categoryTotal,
      count: rows.length,
      average: rows.length ? categoryTotal / rows.length : 0,
      maximum: rows.length ? Math.max(...balances) : 0,
      share: totalAR ? categoryTotal / totalAR : 0
    };
  });
}

function buildCustomerPivot() {
  const customers = [...new Set(appState.processedData.map(row => row.customer))].sort(naturalCompare);
  const rows = customers.map(customer => {
    const record = { customer, values: {}, total: 0 };
    AGING_ORDER.forEach(category => {
      record.values[category] = sum(appState.processedData.filter(row => row.customer === customer && row.classification === category).map(row => row.balance));
      record.total += record.values[category];
    });
    return record;
  });
  const totals = Object.fromEntries(AGING_ORDER.map(category => [category, sum(rows.map(row => row.values[category]))]));
  return { rows, totals, grandTotal: sum(rows.map(row => row.total)) };
}

function buildCustomerTypePivot() {
  const types = [...new Set(appState.processedData.map(row => row.customerType))].sort(naturalCompare);
  const rows = AGING_ORDER.map(category => {
    const record = { category, values: {}, total: 0 };
    types.forEach(type => {
      record.values[type] = sum(appState.processedData.filter(row => row.classification === category && row.customerType === type).map(row => row.balance));
      record.total += record.values[type];
    });
    return record;
  });
  const totals = Object.fromEntries(types.map(type => [type, sum(rows.map(row => row.values[type]))]));
  return { types, rows, totals, grandTotal: sum(rows.map(row => row.total)) };
}

function renderDashboard() {
  const hasData = appState.processedData.length > 0;
  document.querySelectorAll("[data-empty]").forEach(element => { element.hidden = hasData; });
  document.querySelectorAll("[data-results]").forEach(element => { element.hidden = !hasData; });
  updateDateLabels();
  renderMetadata();
  renderDataQuality();
  if (!hasData) {
    destroyAllCharts();
    return;
  }
  renderKPIs();
  renderTransactionTable();
  renderAgingSummary();
  renderCustomerPivot();
  renderCustomerTypePivot();
  renderAuditObservations();
  requestAnimationFrame(renderCharts);
}

function renderMetadata() {
  document.getElementById("header-dataset").textContent = appState.workbookName || "No workbook";
  document.getElementById("header-records").textContent = appState.workbookName ? numberFormat(appState.quality.valid) : "—";
  if (appState.workbookName) setReadyStatus();
}

function renderKPIs() {
  const kpis = calculateKPIs();
  document.getElementById("kpi-total").textContent = currency(kpis.total);
  document.getElementById("kpi-count").textContent = numberFormat(kpis.count);
  document.getElementById("kpi-average").textContent = currency(kpis.average);
  document.getElementById("kpi-over90").textContent = currency(kpis.over90);
  document.getElementById("kpi-over90-share").textContent = `${percent(kpis.over90Share)} of total AR`;
}

function renderTransactionTable() {
  const tbody = document.querySelector("#transaction-table tbody");
  if (!tbody) return;
  const { key, direction } = appState.tableSort;
  const multiplier = direction === "asc" ? 1 : -1;
  const filtered = appState.processedData.filter(row => {
    const matchesCategory = appState.tableFilter === "all" || row.classification === appState.tableFilter;
    const searchTarget = `${row.customer} ${row.salesPerson} ${row.customerType} ${row.classification}`.toLocaleLowerCase();
    return matchesCategory && (!appState.tableSearch || searchTarget.includes(appState.tableSearch));
  }).sort((a, b) => compareValues(a[key], b[key]) * multiplier);

  tbody.innerHTML = filtered.length ? filtered.map(row => `
    <tr>
      <td class="customer-cell">${escapeHtml(row.customer)}</td>
      <td>${escapeHtml(row.salesPerson)}</td>
      <td class="num">${currency(row.balance)}</td>
      <td>${formatDate(row.invoiceDate)}</td>
      <td><span class="type-chip">${escapeHtml(row.customerType)}</span></td>
      <td class="num">${numberFormat(row.aging)} days</td>
      <td><span class="classification-badge badge-${slug(row.classification)}">${AGING_LABELS[row.classification]}</span></td>
    </tr>`).join("") : `<tr><td colspan="7">No records match the current search and filter.</td></tr>`;
  document.getElementById("transaction-count").textContent = `${numberFormat(filtered.length)} of ${numberFormat(appState.processedData.length)} records`;
}

function renderAgingSummary() {
  const summary = buildAgingSummary();
  const total = sum(summary.map(row => row.total));
  const count = sum(summary.map(row => row.count));
  const maximum = appState.processedData.length ? Math.max(...appState.processedData.map(row => row.balance)) : 0;
  const tbody = document.querySelector("#aging-summary-table tbody");
  tbody.innerHTML = summary.map(row => `
    <tr>
      <th scope="row"><span class="classification-badge badge-${slug(row.category)}">${AGING_LABELS[row.category]}</span></th>
      <td class="num">${currency(row.total)}</td>
      <td class="num">${numberFormat(row.count)}</td>
      <td class="num">${currency(row.average)}</td>
      <td class="num">${currency(row.maximum)}</td>
      <td class="num">${percent(row.share)}</td>
    </tr>`).join("");
  document.querySelector("#aging-summary-table tfoot").innerHTML = `<tr><th scope="row">Total</th><td class="num">${currency(total)}</td><td class="num">${numberFormat(count)}</td><td class="num">${currency(count ? total / count : 0)}</td><td class="num">${currency(maximum)}</td><td class="num">100.0%</td></tr>`;
}

function renderCustomerPivot() {
  const pivot = buildCustomerPivot();
  const table = document.getElementById("customer-pivot-table");
  table.querySelector("thead").innerHTML = `<tr><th scope="col">Customer</th>${AGING_ORDER.map(category => `<th scope="col" class="num">${AGING_LABELS[category]}</th>`).join("")}<th scope="col" class="num">Total</th></tr>`;
  table.querySelector("tbody").innerHTML = pivot.rows.map(row => `<tr><th scope="row">${escapeHtml(row.customer)}</th>${AGING_ORDER.map(category => `<td class="num">${currency(row.values[category])}</td>`).join("")}<td class="num"><strong>${currency(row.total)}</strong></td></tr>`).join("");
  table.querySelector("tfoot").innerHTML = `<tr><th scope="row">Total</th>${AGING_ORDER.map(category => `<td class="num">${currency(pivot.totals[category])}</td>`).join("")}<td class="num">${currency(pivot.grandTotal)}</td></tr>`;
}

function renderCustomerTypePivot() {
  const pivot = buildCustomerTypePivot();
  const table = document.getElementById("customer-type-pivot-table");
  table.querySelector("thead").innerHTML = `<tr><th scope="col">Aging Classification</th>${pivot.types.map(type => `<th scope="col" class="num">${escapeHtml(type)}</th>`).join("")}<th scope="col" class="num">Total</th></tr>`;
  table.querySelector("tbody").innerHTML = pivot.rows.map(row => `<tr><th scope="row"><span class="classification-badge badge-${slug(row.category)}">${AGING_LABELS[row.category]}</span></th>${pivot.types.map(type => `<td class="num">${currency(row.values[type])}</td>`).join("")}<td class="num"><strong>${currency(row.total)}</strong></td></tr>`).join("");
  table.querySelector("tfoot").innerHTML = `<tr><th scope="row">Total</th>${pivot.types.map(type => `<td class="num">${currency(pivot.totals[type])}</td>`).join("")}<td class="num">${currency(pivot.grandTotal)}</td></tr>`;
}

function renderCharts() {
  if (!appState.processedData.length) return;
  if (typeof Chart === "undefined") {
    showChartUnavailable();
    return;
  }
  if (appState.activeTab === "one-dimensional") renderAgingChart();
  if (appState.activeTab === "two-dimensional") renderCustomerChart();
  if (appState.activeTab === "three-dimensional") renderCustomerTypeChart();
}

function renderAgingChart() {
  const summary = buildAgingSummary();
  createChart("aging", "aging-chart", {
    type: "bar",
    data: {
      labels: summary.map(row => AGING_LABELS[row.category]),
      datasets: [{ label: "AR Balance", data: summary.map(row => row.total), backgroundColor: AGING_COLORS, borderColor: AGING_COLORS, borderWidth: 1, borderRadius: 5, maxBarThickness: 58 }]
    },
    options: chartOptions({ stacked: false, legend: false })
  });
}

function renderCustomerChart() {
  const pivot = buildCustomerPivot();
  createChart("customer", "customer-chart", {
    type: "bar",
    data: {
      labels: pivot.rows.map(row => `Customer ${row.customer}`),
      datasets: AGING_ORDER.map((category, index) => ({ label: AGING_LABELS[category], data: pivot.rows.map(row => row.values[category]), backgroundColor: AGING_COLORS[index], borderColor: AGING_COLORS[index], borderWidth: 1, borderRadius: 2, maxBarThickness: 52 }))
    },
    options: chartOptions({ stacked: true, legend: true })
  });
}

function renderCustomerTypeChart() {
  const pivot = buildCustomerTypePivot();
  createChart("customerType", "customer-type-chart", {
    type: "bar",
    data: {
      labels: AGING_ORDER.map(category => AGING_LABELS[category]),
      datasets: pivot.types.map((type, index) => ({ label: type, data: pivot.rows.map(row => row.values[type]), backgroundColor: TYPE_COLORS[index % TYPE_COLORS.length], borderColor: TYPE_COLORS[index % TYPE_COLORS.length], borderWidth: 1, borderRadius: 4, maxBarThickness: 50 }))
    },
    options: chartOptions({ stacked: false, legend: true })
  });
}

function chartOptions({ stacked, legend }) {
  return {
    responsive: true,
    maintainAspectRatio: false,
    animation: { duration: 220 },
    interaction: { intersect: false, mode: "index" },
    plugins: {
      legend: { display: legend, position: "bottom", labels: { usePointStyle: true, pointStyle: "rectRounded", boxWidth: 8, boxHeight: 8, padding: 18, color: "#62666e", font: { size: 10, weight: 600 } } },
      tooltip: { backgroundColor: "#282828", padding: 11, titleFont: { size: 11 }, bodyFont: { size: 11 }, callbacks: { label: context => `${context.dataset.label}: ${currency(context.parsed.y)}` } }
    },
    scales: {
      x: { stacked, grid: { display: false }, ticks: { color: "#676b73", font: { size: 10, weight: 600 }, maxRotation: 35, autoSkip: true, maxTicksLimit: 12 } },
      y: { stacked, beginAtZero: true, grid: { color: "#eceef1" }, border: { display: false }, ticks: { color: "#777b83", font: { size: 9 }, callback: value => compactCurrency(value) }, title: { display: true, text: "AR Balance", color: "#777b83", font: { size: 9, weight: 700 } } }
    }
  };
}

function createChart(key, canvasId, config) {
  destroyChart(key);
  const canvas = document.getElementById(canvasId);
  canvas.hidden = false;
  canvas.parentElement.querySelector(".chart-unavailable")?.remove();
  appState.charts[key] = new Chart(canvas, config);
}

function destroyChart(key) {
  if (appState.charts[key]) {
    appState.charts[key].destroy();
    delete appState.charts[key];
  }
}

function destroyAllCharts() {
  Object.keys(appState.charts).forEach(destroyChart);
}

function showChartUnavailable() {
  document.querySelectorAll(".chart-wrap").forEach(wrap => {
    wrap.querySelector("canvas").hidden = true;
    if (!wrap.querySelector(".chart-unavailable")) {
      wrap.insertAdjacentHTML("beforeend", '<p class="chart-unavailable">Charts could not load. Check your internet connection and refresh.</p>');
    }
  });
}

function renderAuditObservations() {
  const kpis = calculateKPIs();
  document.getElementById("audit-one").textContent = `${currency(kpis.over90)} of outstanding receivables are over 90 days old, representing ${percent(kpis.over90Share)} of total AR.`;

  const customerPivot = buildCustomerPivot();
  const largestCustomer = [...customerPivot.rows].sort((a, b) => b.total - a.total)[0];
  const largestShare = customerPivot.grandTotal ? largestCustomer.total / customerPivot.grandTotal : 0;
  document.getElementById("audit-two").textContent = `Customer ${largestCustomer.customer} has the largest outstanding receivable balance at ${currency(largestCustomer.total)}, representing ${percent(largestShare)} of total AR.`;

  const typePivot = buildCustomerTypePivot();
  const over90 = typePivot.rows.find(row => row.category === "over 90 days");
  const rankedTypes = typePivot.types.map(type => ({ type, value: over90.values[type] })).sort((a, b) => b.value - a.value);
  const topType = rankedTypes[0];
  const over90Total = sum(rankedTypes.map(item => item.value));
  document.getElementById("audit-three").textContent = over90Total
    ? `Customer Type ${topType.type} contains the highest over-90-day balance at ${currency(topType.value)}, or ${percent(topType.value / over90Total)} of the over-90-day category.`
    : "No receivables fall into the over-90-day category for the selected analysis date.";
}

function renderDataQuality() {
  const { valid, excluded, reasons } = appState.quality;
  const qualityButton = document.getElementById("quality-button");
  qualityButton.disabled = !appState.workbookName;
  document.getElementById("quality-count").textContent = appState.workbookName ? `${numberFormat(valid)} / ${numberFormat(excluded)}` : "—";
  document.getElementById("quality-summary").textContent = appState.workbookName
    ? `${numberFormat(valid)} valid record${valid === 1 ? "" : "s"}; ${numberFormat(excluded)} excluded record${excluded === 1 ? "" : "s"}.`
    : "No workbook has been evaluated.";
  const reasonEntries = Object.entries(reasons);
  document.getElementById("quality-reasons").innerHTML = reasonEntries.length
    ? reasonEntries.map(([reason, count]) => `<li>${escapeHtml(reason)}: ${numberFormat(count)}</li>`).join("")
    : "<li>No data-quality exceptions detected.</li>";
}

function reconcileOutputs() {
  const totalAR = sum(appState.processedData.map(row => row.balance));
  const summaryTotal = sum(buildAgingSummary().map(row => row.total));
  const customerTotal = buildCustomerPivot().grandTotal;
  const typeTotal = buildCustomerTypePivot().grandTotal;
  const checks = { agingSummary: summaryTotal, customerPivot: customerTotal, customerTypePivot: typeTotal };
  Object.entries(checks).forEach(([name, value]) => {
    if (Math.abs(totalAR - value) > 0.005) console.warn(`AR reconciliation mismatch for ${name}`, { totalAR, outputTotal: value, difference: totalAR - value });
  });
}

function updateDateLabels() {
  const formatted = formatDate(appState.analysisDate);
  document.getElementById("header-date").textContent = formatted;
  document.querySelectorAll("[id^='as-of-']").forEach(element => {
    element.textContent = appState.workbookName ? `Aging calculated as of ${longDate(appState.analysisDate)}` : "Awaiting data";
  });
}

async function resetAnalysis() {
  appState.analysisDate = parseInputDate(DEFAULT_ANALYSIS_DATE);
  document.getElementById("analysis-date").value = DEFAULT_ANALYSIS_DATE;
  document.getElementById("workbook-upload").value = "";
  hideError();
  destroyAllCharts();
  clearDataset();
  renderDashboard();
  await tryLoadReferenceWorkbook();
}

function clearDataset() {
  appState.workbookName = null;
  appState.rawSales = [];
  appState.rawCustomerTypes = [];
  appState.joinedRows = [];
  appState.processedData = [];
  appState.source = null;
  appState.quality = { valid: 0, excluded: 0, reasons: {} };
  renderDashboard();
}

function setLoadingStatus(message) {
  const status = document.getElementById("dataset-status");
  status.className = "dataset-status is-loading";
  status.innerHTML = `<span class="status-mark" aria-hidden="true">···</span><span><strong>${escapeHtml(message)}</strong><small>Processing locally in your browser</small></span>`;
}

function setReadyStatus() {
  const status = document.getElementById("dataset-status");
  if (!appState.workbookName) return setEmptyStatus();
  status.className = "dataset-status is-ready";
  status.innerHTML = `<span class="status-mark" aria-hidden="true">✓</span><span><strong>${escapeHtml(appState.workbookName)}</strong><small>${numberFormat(appState.quality.valid)} valid record${appState.quality.valid === 1 ? "" : "s"}${appState.quality.excluded ? ` · ${numberFormat(appState.quality.excluded)} excluded` : ""}</small></span>`;
}

function setEmptyStatus() {
  const status = document.getElementById("dataset-status");
  status.className = "dataset-status";
  status.innerHTML = '<span class="status-mark" aria-hidden="true">○</span><span><strong>Waiting for a workbook</strong><small>Upload a compatible Excel file to begin</small></span>';
}

function showError(message) {
  const card = document.getElementById("error-card");
  document.getElementById("error-message").textContent = message;
  card.hidden = false;
}

function hideError() {
  document.getElementById("error-card").hidden = true;
}

function canonicalField(header) {
  const normalized = normalizeHeader(header);
  return Object.keys(columnAliases).find(field => columnAliases[field].includes(normalized)) || null;
}

function normalizeHeader(value) {
  return String(value ?? "").replace(/^\uFEFF/, "").trim().toLocaleLowerCase().replace(/[^a-z0-9]+/g, "");
}

function normalizeKey(value) {
  if (isBlank(value)) return "";
  if (typeof value === "number" && Number.isFinite(value)) return String(value);
  return String(value).trim().toLocaleLowerCase().replace(/\.0+$/, "");
}

function hasColumns(headerMap, fields) {
  return fields.every(field => headerMap[field] !== undefined);
}

function displayFieldName(field) {
  return ({ customer: "Customer", salesPerson: "Sales Person", balance: "Balance", invoiceDate: "Invoice Date", customerType: "Customer Type" })[field] || field;
}

function parseBalance(value) {
  if (typeof value === "number") return Number.isFinite(value) ? value : null;
  if (typeof value !== "string" || !value.trim()) return null;
  const trimmed = value.trim();
  const negative = /^\(.*\)$/.test(trimmed);
  const numeric = Number(trimmed.replace(/[,$()\s]/g, ""));
  return Number.isFinite(numeric) ? (negative ? -numeric : numeric) : null;
}

function parseExcelDate(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return new Date(Date.UTC(value.getFullYear(), value.getMonth(), value.getDate()));
  }
  if (typeof value === "number" && Number.isFinite(value)) {
    const parts = typeof XLSX !== "undefined" ? XLSX.SSF.parse_date_code(value) : null;
    return parts ? createValidatedDate(parts.y, parts.m, parts.d) : null;
  }
  if (typeof value !== "string") return null;
  const text = value.trim();
  let match = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T\s].*)?$/);
  if (match) return createValidatedDate(Number(match[1]), Number(match[2]), Number(match[3]));
  match = text.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/);
  if (match) return createValidatedDate(Number(match[3]), Number(match[1]), Number(match[2]));
  return null;
}

function createValidatedDate(year, month, day) {
  const date = new Date(Date.UTC(year, month - 1, day));
  return date.getUTCFullYear() === year && date.getUTCMonth() === month - 1 && date.getUTCDate() === day ? date : null;
}

function parseInputDate(value) {
  if (!value) return null;
  const match = String(value).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return match ? createValidatedDate(Number(match[1]), Number(match[2]), Number(match[3])) : null;
}

function formatCustomer(value) {
  if (isBlank(value)) return "";
  if (typeof value === "number" && Number.isFinite(value)) return Number.isInteger(value) ? String(value) : String(value);
  return String(value).trim().replace(/\.0+$/, "");
}

function cleanText(value) {
  return isBlank(value) ? "" : String(value).trim();
}

function isBlank(value) {
  return value === null || value === undefined || (typeof value === "string" && value.trim() === "");
}

function addReason(reasons, reason) {
  reasons[reason] = (reasons[reason] || 0) + 1;
}

function utcDay(date) {
  return Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate());
}

function toInputDate(date) {
  return `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, "0")}-${String(date.getUTCDate()).padStart(2, "0")}`;
}

function formatDate(date) {
  return new Intl.DateTimeFormat("en-US", { month: "2-digit", day: "2-digit", year: "numeric", timeZone: "UTC" }).format(date);
}

function longDate(date) {
  return new Intl.DateTimeFormat("en-US", { month: "long", day: "numeric", year: "numeric", timeZone: "UTC" }).format(date);
}

function currency(value) {
  return new Intl.NumberFormat("en-US", { style: "currency", currency: "USD", minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(value || 0);
}

function compactCurrency(value) {
  return new Intl.NumberFormat("en-US", { style: "currency", currency: "USD", notation: "compact", maximumFractionDigits: 1 }).format(value || 0);
}

function percent(value) {
  return new Intl.NumberFormat("en-US", { style: "percent", minimumFractionDigits: 1, maximumFractionDigits: 1 }).format(value || 0);
}

function numberFormat(value) {
  return new Intl.NumberFormat("en-US", { maximumFractionDigits: 0 }).format(value || 0);
}

function sum(values) {
  return values.reduce((total, value) => total + (Number(value) || 0), 0);
}

function compareValues(a, b) {
  if (a instanceof Date && b instanceof Date) return a - b;
  if (typeof a === "number" && typeof b === "number") return a - b;
  return naturalCompare(String(a ?? ""), String(b ?? ""));
}

function naturalCompare(a, b) {
  return String(a).localeCompare(String(b), undefined, { numeric: true, sensitivity: "base" });
}

function slug(value) {
  return String(value).toLocaleLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "");
}

function escapeHtml(value) {
  return String(value ?? "").replace(/[&<>'"]/g, character => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "'": "&#39;", '"': "&quot;" })[character]);
}
