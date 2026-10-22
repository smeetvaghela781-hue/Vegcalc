/**
 * VegCalc Pro — Vegetable Mandi Intake & Profit Spreadsheet Application
 * Full Excel-compatible reactivity, real-time formula calculations & local persistence
 */

// State Management
const APP_STORAGE_KEY = "vegcalc_pro_data_v2";
const APP_THEME_KEY = "vegcalc_pro_theme";

let state = {
  entries: [],
  activeCell: { rowId: null, colKey: null },
  selectedRows: new Set(),
  filterText: "",
  filterDate: "",
  currentTab: "intake",
  theme: "light"
};

// Initialize Application
document.addEventListener("DOMContentLoaded", () => {
  initTheme();
  initData();
  initQuickEntry();
  initSpreadsheet();
  initFormulaBar();
  initRibbonEvents();
  initMandiCalculator();
  initAutocomplete();
  renderApp();
});

/* =========================================================
   1. Data Initialization & Persistence
   ========================================================= */
function initData() {
  const saved = localStorage.getItem(APP_STORAGE_KEY);
  if (saved) {
    try {
      state.entries = JSON.parse(saved);
    } catch (e) {
      console.error("Failed to parse saved data, loading defaults", e);
      state.entries = JSON.parse(JSON.stringify(DEFAULT_INITIAL_ENTRIES));
    }
  } else {
    // Load default template matching user's Veg_Business_Tracker.xlsx Intake sheet
    state.entries = JSON.parse(JSON.stringify(DEFAULT_INITIAL_ENTRIES));
    saveData();
  }
}

function saveData() {
  localStorage.setItem(APP_STORAGE_KEY, JSON.stringify(state.entries));
}

function initTheme() {
  const savedTheme = localStorage.getItem(APP_THEME_KEY) || "light";
  state.theme = savedTheme;
  applyTheme();
}

function applyTheme() {
  if (state.theme === "dark") {
    document.body.classList.remove("theme-light");
    document.body.classList.add("theme-dark");
    document.getElementById("theme-icon").textContent = "☀️";
  } else {
    document.body.classList.remove("theme-dark");
    document.body.classList.add("theme-light");
    document.getElementById("theme-icon").textContent = "🌙";
  }
}

/* =========================================================
   2. Calculations & Computed Formulas
   ========================================================= */
function computeRow(entry) {
  const qty = parseFloat(entry.qty) || 0;
  const totalPrice = parseFloat(entry.totalPrice) || 0;
  const retailPrice = parseFloat(entry.retailPrice) || 0;

  // Formula [E]: Price per kg = Total Price / Quantity
  const pricePerKg = qty > 0 ? totalPrice / qty : 0;

  // Formula [G]: Profit per kg = Retail Price - Cost per kg
  const profitPerKg = retailPrice > 0 ? retailPrice - pricePerKg : 0;

  // Formula [H]: Total Expected Profit = Profit per kg * Quantity
  const totalProfit = profitPerKg * qty;

  // Formula [I]: Margin % = (Profit per kg / Retail Price) * 100
  const marginPct = retailPrice > 0 ? (profitPerKg / retailPrice) * 100 : 0;

  // Total Projected Revenue
  const totalRevenue = retailPrice * qty;

  return {
    ...entry,
    qty,
    totalPrice,
    retailPrice,
    pricePerKg,
    profitPerKg,
    totalProfit,
    marginPct,
    totalRevenue
  };
}

function getComputedEntries() {
  return state.entries.map(computeRow);
}

function getFilteredEntries() {
  const computed = getComputedEntries();
  return computed.filter(entry => {
    // Text search filter
    if (state.filterText) {
      const q = state.filterText.toLowerCase().trim();
      const matchItem = (entry.item || "").toLowerCase().includes(q);
      const matchSupplier = (entry.supplier || "").toLowerCase().includes(q);
      if (!matchItem && !matchSupplier) return false;
    }
    // Date filter
    if (state.filterDate) {
      if (entry.date !== state.filterDate) return false;
    }
    return true;
  });
}

function getTotals(entries) {
  let totalQty = 0;
  let totalCost = 0;
  let totalRevenue = 0;
  let totalProfit = 0;

  entries.forEach(e => {
    totalQty += e.qty;
    totalCost += e.totalPrice;
    totalRevenue += e.totalRevenue;
    totalProfit += e.totalProfit;
  });

  const avgCostPerKg = totalQty > 0 ? totalCost / totalQty : 0;
  const avgProfitPerKg = totalQty > 0 ? totalProfit / totalQty : 0;
  const avgMargin = totalRevenue > 0 ? (totalProfit / totalRevenue) * 100 : 0;

  return {
    count: entries.length,
    totalQty,
    totalCost,
    totalRevenue,
    totalProfit,
    avgCostPerKg,
    avgProfitPerKg,
    avgMargin
  };
}

/* =========================================================
   3. Rendering: KPIs, Grid & Totals
   ========================================================= */
function renderApp() {
  renderKPIs();
  renderSpreadsheet();
  renderFormulaBar();
  renderStatusBar();
}

function formatCurrency(val) {
  if (val === undefined || val === null || isNaN(val)) return "₹0.00";
  return "₹" + Number(val).toLocaleString("en-IN", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function formatNumber(val, decimals = 2) {
  if (val === undefined || val === null || isNaN(val)) return "0.00";
  return Number(val).toLocaleString("en-IN", { minimumFractionDigits: decimals, maximumFractionDigits: decimals });
}

function renderKPIs() {
  const computed = getFilteredEntries();
  const totals = getTotals(computed);

  document.getElementById("kpi-total-cost").textContent = formatCurrency(totals.totalCost);
  document.getElementById("kpi-total-weight").textContent = `${formatNumber(totals.totalQty, 1)} kg across ${totals.count} items`;

  document.getElementById("kpi-total-revenue").textContent = formatCurrency(totals.totalRevenue);
  document.getElementById("kpi-avg-cost-rate").textContent = `Avg Buy: ${formatCurrency(totals.avgCostPerKg)} / kg`;

  const profitEl = document.getElementById("kpi-total-profit");
  profitEl.textContent = (totals.totalProfit >= 0 ? "+" : "") + formatCurrency(totals.totalProfit);
  if (totals.totalProfit < 0) {
    profitEl.classList.add("profit-red");
  } else {
    profitEl.classList.remove("profit-red");
  }

  document.getElementById("kpi-profit-per-kg").textContent = `${totals.avgProfitPerKg >= 0 ? "+" : ""}${formatCurrency(totals.avgProfitPerKg)} avg gain / kg`;

  document.getElementById("kpi-avg-margin").textContent = `${totals.avgMargin.toFixed(1)}%`;

  // Best margin item
  if (computed.length > 0) {
    const sorted = [...computed].sort((a, b) => b.marginPct - a.marginPct);
    const best = sorted[0];
    if (best && best.marginPct > 0) {
      document.getElementById("kpi-highest-item").textContent = `Best: ${best.item} (${best.marginPct.toFixed(0)}%)`;
    } else {
      document.getElementById("kpi-highest-item").textContent = `Best item: —`;
    }
  } else {
    document.getElementById("kpi-highest-item").textContent = `Best item: —`;
  }
}

function renderSpreadsheet() {
  const tbody = document.getElementById("spreadsheet-tbody");
  const filtered = getFilteredEntries();
  const emptyState = document.getElementById("empty-state");

  if (filtered.length === 0) {
    tbody.innerHTML = "";
    emptyState.style.display = "flex";
    renderFooterTotals([]);
    return;
  }

  emptyState.style.display = "none";
  let html = "";

  filtered.forEach((entry, idx) => {
    const rowNum = idx + 5; // Starting at row 5 matching Excel sheet row 5!
    const isSelected = state.selectedRows.has(entry.id);

    html += `
      <tr class="spreadsheet-row ${isSelected ? "row-selected" : ""}" data-id="${entry.id}" data-row="${rowNum}">
        <!-- Row Index Cell -->
        <td class="row-num-cell" data-action="select-row" title="Row ${rowNum}">
          <span>${rowNum}</span>
        </td>

        <!-- Column A: Date -->
        <td class="cell-input cell-editable" data-col="date" data-id="${entry.id}" data-coord="A${rowNum}" tabindex="0">
          ${entry.date || ""}
        </td>

        <!-- Column B: Item -->
        <td class="cell-input cell-editable font-semibold" data-col="item" data-id="${entry.id}" data-coord="B${rowNum}" tabindex="0">
          ${escapeHtml(entry.item || "")}
        </td>

        <!-- Column C: Quantity (kg) -->
        <td class="cell-input cell-editable num-cell" data-col="qty" data-id="${entry.id}" data-coord="C${rowNum}" tabindex="0">
          ${entry.qty > 0 ? formatNumber(entry.qty, 1) : ""}
        </td>

        <!-- Column D: Total Price Paid (₹) -->
        <td class="cell-input cell-editable num-cell font-mono" data-col="totalPrice" data-id="${entry.id}" data-coord="D${rowNum}" tabindex="0">
          ${entry.totalPrice > 0 ? formatCurrency(entry.totalPrice) : ""}
        </td>

        <!-- Column E: Price per kg (₹) [Formula D/C] -->
        <td class="cell-formula num-cell font-mono" data-col="pricePerKg" data-id="${entry.id}" data-coord="E${rowNum}" title="Formula: =D${rowNum}/C${rowNum}">
          ${entry.pricePerKg > 0 ? formatCurrency(entry.pricePerKg) : "—"}
        </td>

        <!-- Column F: Retail Price per kg (₹) -->
        <td class="cell-input cell-editable num-cell font-mono" data-col="retailPrice" data-id="${entry.id}" data-coord="F${rowNum}" tabindex="0">
          ${entry.retailPrice > 0 ? formatCurrency(entry.retailPrice) : ""}
        </td>

        <!-- Column G: Profit per kg (₹) [Formula F-E] -->
        <td class="cell-formula num-cell font-mono ${entry.profitPerKg >= 0 ? "profit-green" : "profit-red"}" data-col="profitPerKg" data-id="${entry.id}" data-coord="G${rowNum}" title="Formula: =F${rowNum}-E${rowNum}">
          ${entry.retailPrice > 0 ? (entry.profitPerKg >= 0 ? "+" : "") + formatCurrency(entry.profitPerKg) : "—"}
        </td>

        <!-- Column H: Total Profit (₹) [Formula G*C] -->
        <td class="cell-formula num-cell font-mono ${entry.totalProfit >= 0 ? "profit-green" : "profit-red"}" data-col="totalProfit" data-id="${entry.id}" data-coord="H${rowNum}" title="Formula: =G${rowNum}*C${rowNum}">
          ${entry.retailPrice > 0 ? (entry.totalProfit >= 0 ? "+" : "") + formatCurrency(entry.totalProfit) : "—"}
        </td>

        <!-- Column I: Margin % [Formula G/F] -->
        <td class="cell-formula num-cell font-mono" data-col="marginPct" data-id="${entry.id}" data-coord="I${rowNum}" title="Formula: =G${rowNum}/F${rowNum}">
          ${entry.retailPrice > 0 ? entry.marginPct.toFixed(1) + "%" : "—"}
        </td>

        <!-- Column J: Supplier / Notes -->
        <td class="cell-input cell-editable" data-col="supplier" data-id="${entry.id}" data-coord="J${rowNum}" tabindex="0">
          ${escapeHtml(entry.supplier || "")}
        </td>

        <!-- Actions -->
        <td class="row-actions-cell col-actions">
          <button class="btn-row-action" data-action="delete-row" data-id="${entry.id}" title="Delete Row">
            🗑️
          </button>
        </td>
      </tr>
    `;
  });

  tbody.innerHTML = html;
  renderFooterTotals(filtered);
  restoreActiveCellFocus();
}

function renderFooterTotals(entries) {
  const totals = getTotals(entries);
  document.getElementById("footer-count").textContent = `${totals.count} items`;
  document.getElementById("footer-total-qty").textContent = `${formatNumber(totals.totalQty, 1)} kg`;
  document.getElementById("footer-total-cost").textContent = formatCurrency(totals.totalCost);
  document.getElementById("footer-avg-cost").textContent = totals.avgCostPerKg > 0 ? `${formatCurrency(totals.avgCostPerKg)} /kg` : "—";
  document.getElementById("footer-avg-profit").textContent = totals.avgProfitPerKg > 0 ? `+${formatCurrency(totals.avgProfitPerKg)} /kg` : "—";
  
  const footerProfit = document.getElementById("footer-total-profit");
  footerProfit.textContent = (totals.totalProfit >= 0 ? "+" : "") + formatCurrency(totals.totalProfit);
  if (totals.totalProfit < 0) {
    footerProfit.classList.add("profit-red");
  } else {
    footerProfit.classList.remove("profit-red");
  }

  document.getElementById("footer-avg-margin").textContent = `${totals.avgMargin.toFixed(1)}%`;
}

function renderStatusBar() {
  const computed = getFilteredEntries();
  const totals = getTotals(computed);

  document.getElementById("status-row-info").textContent = `Showing ${computed.length} entries of ${state.entries.length}`;
  document.getElementById("stat-sum").textContent = `SUM PROFIT: ${formatCurrency(totals.totalProfit)}`;
  document.getElementById("stat-avg").textContent = `AVG MARGIN: ${totals.avgMargin.toFixed(1)}%`;
  document.getElementById("stat-count").textContent = `ITEMS: ${totals.count}`;
}

function escapeHtml(str) {
  if (!str) return "";
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

/* =========================================================
   4. Interactive Spreadsheet & Keyboard Navigation
   ========================================================= */
const EDITABLE_COLUMNS = ["date", "item", "qty", "totalPrice", "retailPrice", "supplier"];

function initSpreadsheet() {
  const tbody = document.getElementById("spreadsheet-tbody");

  // Click on Cell
  tbody.addEventListener("click", (e) => {
    const td = e.target.closest("td");
    if (!td) return;

    if (e.target.dataset.action === "delete-row") {
      const id = e.target.dataset.id;
      deleteRow(id);
      return;
    }

    if (td.dataset.action === "select-row") {
      const tr = td.closest("tr");
      const id = tr.dataset.id;
      toggleSelectRow(id);
      return;
    }

    const rowId = td.dataset.id;
    const colKey = td.dataset.col;

    if (rowId && colKey) {
      setActiveCell(rowId, colKey, td);
    }
  });

  // Double Click on Cell -> Immediate Edit Mode
  tbody.addEventListener("dblclick", (e) => {
    const td = e.target.closest("td");
    if (!td) return;
    const rowId = td.dataset.id;
    const colKey = td.dataset.col;

    if (EDITABLE_COLUMNS.includes(colKey)) {
      startInlineEditing(td, rowId, colKey);
    }
  });

  // Keyboard Navigation on Table
  tbody.addEventListener("keydown", (e) => {
    const td = e.target.closest("td");
    if (!td || td.querySelector(".cell-inline-input")) return;

    const rowId = td.dataset.id;
    const colKey = td.dataset.col;

    if (e.key === "Enter" || e.key === "F2") {
      if (EDITABLE_COLUMNS.includes(colKey)) {
        e.preventDefault();
        startInlineEditing(td, rowId, colKey);
      }
    } else if (e.key === "Tab") {
      e.preventDefault();
      moveActiveCell(rowId, colKey, e.shiftKey ? "left" : "right");
    } else if (e.key === "ArrowDown") {
      e.preventDefault();
      moveActiveCell(rowId, colKey, "down");
    } else if (e.key === "ArrowUp") {
      e.preventDefault();
      moveActiveCell(rowId, colKey, "up");
    } else if (e.key === "ArrowLeft") {
      e.preventDefault();
      moveActiveCell(rowId, colKey, "left");
    } else if (e.key === "ArrowRight") {
      e.preventDefault();
      moveActiveCell(rowId, colKey, "right");
    } else if (e.key === "Delete" || e.key === "Backspace") {
      if (EDITABLE_COLUMNS.includes(colKey)) {
        e.preventDefault();
        updateEntryValue(rowId, colKey, "");
      }
    }
  });
}

function setActiveCell(rowId, colKey, tdElement) {
  // Clear existing active styles
  document.querySelectorAll(".cell-active").forEach(el => el.classList.remove("cell-active"));

  state.activeCell = { rowId, colKey };
  if (tdElement) {
    tdElement.classList.add("cell-active");
    tdElement.focus();
  }

  renderFormulaBar();
}

function restoreActiveCellFocus() {
  if (!state.activeCell.rowId || !state.activeCell.colKey) return;
  const td = document.querySelector(`td[data-id="${state.activeCell.rowId}"][data-col="${state.activeCell.colKey}"]`);
  if (td) {
    td.classList.add("cell-active");
  }
}

function moveActiveCell(rowId, colKey, direction) {
  const filtered = getFilteredEntries();
  const rowIndex = filtered.findIndex(e => e.id === rowId);
  if (rowIndex === -1) return;

  const colOrder = ["date", "item", "qty", "totalPrice", "pricePerKg", "retailPrice", "profitPerKg", "totalProfit", "marginPct", "supplier"];
  let colIndex = colOrder.indexOf(colKey);

  let nextRowIndex = rowIndex;
  let nextColIndex = colIndex;

  if (direction === "down") {
    nextRowIndex = Math.min(filtered.length - 1, rowIndex + 1);
  } else if (direction === "up") {
    nextRowIndex = Math.max(0, rowIndex - 1);
  } else if (direction === "right") {
    if (colIndex < colOrder.length - 1) {
      nextColIndex = colIndex + 1;
    } else if (rowIndex < filtered.length - 1) {
      nextRowIndex = rowIndex + 1;
      nextColIndex = 0;
    }
  } else if (direction === "left") {
    if (colIndex > 0) {
      nextColIndex = colIndex - 1;
    } else if (rowIndex > 0) {
      nextRowIndex = rowIndex - 1;
      nextColIndex = colOrder.length - 1;
    }
  }

  const nextEntry = filtered[nextRowIndex];
  const nextCol = colOrder[nextColIndex];
  const nextTd = document.querySelector(`td[data-id="${nextEntry.id}"][data-col="${nextCol}"]`);

  if (nextTd) {
    setActiveCell(nextEntry.id, nextCol, nextTd);
  }
}

function startInlineEditing(td, rowId, colKey) {
  const entry = state.entries.find(e => e.id === rowId);
  if (!entry) return;

  const originalVal = entry[colKey] !== undefined ? entry[colKey] : "";
  const inputType = (colKey === "qty" || colKey === "totalPrice" || colKey === "retailPrice") ? "number" : (colKey === "date" ? "date" : "text");

  td.innerHTML = `<input class="cell-inline-input" type="${inputType}" step="any" value="${escapeHtml(originalVal)}">`;
  const input = td.querySelector(".cell-inline-input");
  input.focus();
  if (inputType === "text" || inputType === "number") {
    input.select();
  }

  const commit = () => {
    const newVal = input.value;
    updateEntryValue(rowId, colKey, newVal);
  };

  input.addEventListener("blur", () => {
    commit();
  }, { once: true });

  input.addEventListener("keydown", (e) => {
    if (e.key === "Enter") {
      e.preventDefault();
      commit();
      moveActiveCell(rowId, colKey, "down");
    } else if (e.key === "Tab") {
      e.preventDefault();
      commit();
      moveActiveCell(rowId, colKey, e.shiftKey ? "left" : "right");
    } else if (e.key === "Escape") {
      e.preventDefault();
      renderSpreadsheet();
    }
  });
}

function updateEntryValue(rowId, colKey, value) {
  const entry = state.entries.find(e => e.id === rowId);
  if (!entry) return;

  if (colKey === "qty" || colKey === "totalPrice" || colKey === "retailPrice") {
    entry[colKey] = parseFloat(value) || 0;
  } else {
    entry[colKey] = value;
  }

  saveData();
  renderApp();
}

function deleteRow(id) {
  state.entries = state.entries.filter(e => e.id !== id);
  state.selectedRows.delete(id);
  saveData();
  renderApp();
  showToast("Vegetable entry removed", "info");
}

function toggleSelectRow(id) {
  if (state.selectedRows.has(id)) {
    state.selectedRows.delete(id);
  } else {
    state.selectedRows.add(id);
  }
  renderSpreadsheet();
}

/* =========================================================
   5. Excel Formula Bar Synchronization
   ========================================================= */
function initFormulaBar() {
  const input = document.getElementById("formula-bar-input");

  input.addEventListener("keydown", (e) => {
    if (e.key === "Enter") {
      e.preventDefault();
      if (!state.activeCell.rowId || !state.activeCell.colKey) return;
      if (EDITABLE_COLUMNS.includes(state.activeCell.colKey)) {
        updateEntryValue(state.activeCell.rowId, state.activeCell.colKey, input.value);
        input.blur();
      }
    }
  });
}

function renderFormulaBar() {
  const activeCoordEl = document.getElementById("formula-active-cell");
  const formulaInput = document.getElementById("formula-bar-input");

  if (!state.activeCell.rowId || !state.activeCell.colKey) {
    activeCoordEl.textContent = "A5";
    formulaInput.value = "";
    formulaInput.disabled = true;
    return;
  }

  const filtered = getFilteredEntries();
  const entryIdx = filtered.findIndex(e => e.id === state.activeCell.rowId);
  const rowNum = entryIdx !== -1 ? entryIdx + 5 : 5;
  const colLetterMap = {
    date: "A",
    item: "B",
    qty: "C",
    totalPrice: "D",
    pricePerKg: "E",
    retailPrice: "F",
    profitPerKg: "G",
    totalProfit: "H",
    marginPct: "I",
    supplier: "J"
  };

  const colLetter = colLetterMap[state.activeCell.colKey] || "A";
  activeCoordEl.textContent = `${colLetter}${rowNum}`;

  const entry = state.entries.find(e => e.id === state.activeCell.rowId);
  if (!entry) return;

  if (state.activeCell.colKey === "pricePerKg") {
    formulaInput.value = `=D${rowNum}/C${rowNum}`;
    formulaInput.disabled = true;
  } else if (state.activeCell.colKey === "profitPerKg") {
    formulaInput.value = `=F${rowNum}-E${rowNum}`;
    formulaInput.disabled = true;
  } else if (state.activeCell.colKey === "totalProfit") {
    formulaInput.value = `=G${rowNum}*C${rowNum}`;
    formulaInput.disabled = true;
  } else if (state.activeCell.colKey === "marginPct") {
    formulaInput.value = `=(G${rowNum}/F${rowNum})*100`;
    formulaInput.disabled = true;
  } else {
    formulaInput.value = entry[state.activeCell.colKey] !== undefined ? entry[state.activeCell.colKey] : "";
    formulaInput.disabled = false;
  }
}

/* =========================================================
   6. Rapid Mandi Quick Entry Bar
   ========================================================= */
function initQuickEntry() {
  const form = document.getElementById("quick-entry-form");
  const dateInput = document.getElementById("qe-date");
  const itemInput = document.getElementById("qe-item");
  const qtyInput = document.getElementById("qe-qty");
  const priceInput = document.getElementById("qe-price");
  const retailInput = document.getElementById("qe-retail");
  const notesInput = document.getElementById("qe-notes");

  // Set today's date by default (2026-10-22)
  const todayStr = new Date().toISOString().split("T")[0];
  dateInput.value = todayStr;

  // Real-time Live calculation preview inside Quick Bar
  function updateQuickLiveBadges() {
    const qty = parseFloat(qtyInput.value) || 0;
    const price = parseFloat(priceInput.value) || 0;
    const retail = parseFloat(retailInput.value) || 0;

    const costPerKgBadge = document.getElementById("qe-cost-per-kg-badge");
    const profitBadge = document.getElementById("qe-profit-badge");

    const costPerKg = qty > 0 ? price / qty : 0;
    const profitPerKg = retail > 0 ? retail - costPerKg : 0;
    const totalProfit = profitPerKg * qty;
    const margin = retail > 0 ? (profitPerKg / retail) * 100 : 0;

    costPerKgBadge.textContent = costPerKg > 0 ? `${formatCurrency(costPerKg)} /kg` : "₹0.00 /kg";

    if (retail > 0) {
      profitBadge.textContent = `${profitPerKg >= 0 ? "+" : ""}${formatCurrency(profitPerKg)} /kg (${margin.toFixed(0)}%)`;
      if (profitPerKg < 0) {
        profitBadge.classList.add("negative");
      } else {
        profitBadge.classList.remove("negative");
      }
    } else {
      profitBadge.textContent = "+₹0.00 /kg";
      profitBadge.classList.remove("negative");
    }
  }

  qtyInput.addEventListener("input", updateQuickLiveBadges);
  priceInput.addEventListener("input", updateQuickLiveBadges);
  retailInput.addEventListener("input", updateQuickLiveBadges);

  // Form Submission
  form.addEventListener("submit", (e) => {
    e.preventDefault();
    const item = itemInput.value.trim();
    const qty = parseFloat(qtyInput.value) || 0;
    const price = parseFloat(priceInput.value) || 0;
    const retail = parseFloat(retailInput.value) || 0;
    const date = dateInput.value || todayStr;
    const notes = notesInput.value.trim();

    if (!item) {
      showToast("Please enter vegetable name", "error");
      itemInput.focus();
      return;
    }
    if (qty <= 0) {
      showToast("Please enter a valid quantity", "error");
      qtyInput.focus();
      return;
    }

    const newEntry = {
      id: "v-row-" + Date.now(),
      date,
      item,
      qty,
      totalPrice: price,
      retailPrice: retail,
      supplier: notes
    };

    state.entries.push(newEntry);
    saveData();
    renderApp();

    showToast(`Added ${item} (${qty} kg) to spreadsheet`, "success");

    // Reset fields for next entry
    itemInput.value = "";
    qtyInput.value = "";
    priceInput.value = "";
    retailInput.value = "";
    notesInput.value = "";
    updateQuickLiveBadges();
    itemInput.focus();
  });

  // Populate Fast Select Chips
  const chipsContainer = document.getElementById("quick-veg-chips");
  const popularVegs = VEGETABLE_CATALOG.slice(0, 11); // Potato, Onion, Tomato, Chilli, Dudhi, Fulavar, etc.
  
  chipsContainer.innerHTML = `<span class="chips-label">Fast Select:</span>` + popularVegs.map(veg => `
    <button type="button" class="veg-chip" data-name="${veg.en}">
      <span>${veg.icon}</span>
      <span>${veg.en}</span>
    </button>
  `).join("");

  chipsContainer.addEventListener("click", (e) => {
    const chip = e.target.closest(".veg-chip");
    if (!chip) return;
    const name = chip.dataset.name;
    itemInput.value = name;
    qtyInput.focus();
    updateQuickLiveBadges();
  });

  // Toggle quick entry panel
  document.getElementById("btn-quick-entry-toggle").addEventListener("click", () => {
    const panel = document.getElementById("quick-entry-panel");
    panel.style.display = panel.style.display === "none" ? "block" : "none";
    if (panel.style.display !== "none") {
      itemInput.focus();
    }
  });

  document.getElementById("btn-close-quick-entry").addEventListener("click", () => {
    document.getElementById("quick-entry-panel").style.display = "none";
  });
}

/* =========================================================
   7. Smart Vegetable Autocomplete
   ========================================================= */
function initAutocomplete() {
  const itemInput = document.getElementById("qe-item");
  const dropdown = document.getElementById("autocomplete-list");

  itemInput.addEventListener("input", () => {
    const q = itemInput.value.toLowerCase().trim();
    if (!q) {
      dropdown.style.display = "none";
      return;
    }

    const matches = VEGETABLE_CATALOG.filter(v => {
      return v.en.toLowerCase().includes(q) ||
             v.gu.toLowerCase().includes(q) ||
             v.hi.toLowerCase().includes(q);
    }).slice(0, 7);

    if (matches.length === 0) {
      dropdown.style.display = "none";
      return;
    }

    dropdown.innerHTML = matches.map(v => `
      <div class="autocomplete-item" data-name="${v.en}">
        <div class="ac-left">
          <span class="ac-icon">${v.icon}</span>
          <div>
            <div class="ac-title">${v.en}</div>
            <div class="ac-sub">${v.gu} • ${v.hi}</div>
          </div>
        </div>
      </div>
    `).join("");

    dropdown.style.display = "block";
  });

  dropdown.addEventListener("click", (e) => {
    const item = e.target.closest(".autocomplete-item");
    if (!item) return;
    itemInput.value = item.dataset.name;
    dropdown.style.display = "none";
    document.getElementById("qe-qty").focus();
  });

  document.addEventListener("click", (e) => {
    if (!itemInput.contains(e.target) && !dropdown.contains(e.target)) {
      dropdown.style.display = "none";
    }
  });
}

/* =========================================================
   8. Ribbon Toolbar Actions: Excel Import/Export & Filters
   ========================================================= */
function initRibbonEvents() {
  // Add Blank Row
  document.getElementById("btn-add-row").addEventListener("click", () => {
    const todayStr = new Date().toISOString().split("T")[0];
    const newEntry = {
      id: "v-row-" + Date.now(),
      date: todayStr,
      item: "New Item",
      qty: 1,
      totalPrice: 0,
      retailPrice: 0,
      supplier: ""
    };
    state.entries.push(newEntry);
    saveData();
    renderApp();
    setActiveCell(newEntry.id, "item");
    showToast("Added new blank row", "info");
  });

  // Empty State Add Button
  document.getElementById("btn-empty-add").addEventListener("click", () => {
    document.getElementById("btn-add-row").click();
  });

  // Export Excel (.xlsx) Native File with Formulas!
  document.getElementById("btn-export-excel").addEventListener("click", exportToExcel);

  // Export CSV
  document.getElementById("btn-export-csv").addEventListener("click", exportToCSV);

  // Import Excel File (.xlsx / .csv)
  document.getElementById("excel-file-input").addEventListener("change", importExcelFile);

  // Print Sheet
  document.getElementById("btn-print").addEventListener("click", () => {
    window.print();
  });

  // Reset Data Template
  document.getElementById("btn-reset-data").addEventListener("click", () => {
    if (confirm("Reset to default Veg Business Tracker template data? This will restore the 11 original Mandi entries.")) {
      state.entries = JSON.parse(JSON.stringify(DEFAULT_INITIAL_ENTRIES));
      saveData();
      renderApp();
      showToast("Restored Veg Business Tracker template data", "success");
    }
  });

  // Theme Toggle
  document.getElementById("btn-theme-toggle").addEventListener("click", () => {
    state.theme = state.theme === "light" ? "dark" : "light";
    localStorage.setItem(APP_THEME_KEY, state.theme);
    applyTheme();
  });

  // Search Filter
  const searchInput = document.getElementById("filter-search");
  const clearSearchBtn = document.getElementById("btn-clear-search");

  searchInput.addEventListener("input", () => {
    state.filterText = searchInput.value;
    clearSearchBtn.style.display = state.filterText ? "block" : "none";
    renderApp();
  });

  clearSearchBtn.addEventListener("click", () => {
    searchInput.value = "";
    state.filterText = "";
    clearSearchBtn.style.display = "none";
    renderApp();
    searchInput.focus();
  });

  // Date Filter
  const dateFilterInput = document.getElementById("filter-date");
  dateFilterInput.addEventListener("change", () => {
    state.filterDate = dateFilterInput.value;
    document.getElementById("btn-date-all").classList.remove("btn-active");
    document.getElementById("btn-date-today").classList.remove("btn-active");
    renderApp();
  });

  document.getElementById("btn-date-all").addEventListener("click", () => {
    dateFilterInput.value = "";
    state.filterDate = "";
    document.getElementById("btn-date-all").classList.add("btn-active");
    document.getElementById("btn-date-today").classList.remove("btn-active");
    renderApp();
  });

  document.getElementById("btn-date-today").addEventListener("click", () => {
    const today = "2026-10-22";
    dateFilterInput.value = today;
    state.filterDate = today;
    document.getElementById("btn-date-today").classList.add("btn-active");
    document.getElementById("btn-date-all").classList.remove("btn-active");
    renderApp();
  });

  // Sheet Tabs
  document.querySelectorAll(".sheet-tab").forEach(tab => {
    tab.addEventListener("click", () => {
      document.querySelectorAll(".sheet-tab").forEach(t => t.classList.remove("active"));
      tab.classList.add("active");
      const tabName = tab.dataset.tab;
      state.currentTab = tabName;
      if (tabName === "summary" || tabName === "analytics") {
        openDailySummaryModal();
      }
    });
  });

  // Global Keyboard Shortcuts (Alt+N for Quick Entry)
  document.addEventListener("keydown", (e) => {
    if (e.altKey && e.key.toLowerCase() === "n") {
      e.preventDefault();
      document.getElementById("btn-quick-entry-toggle").click();
    }
  });
}

/* =========================================================
   9. Excel File (.xlsx) Export & Import Engine
   ========================================================= */
function exportToExcel() {
  if (typeof XLSX === "undefined") {
    showToast("Excel export library loading...", "error");
    return;
  }

  const computed = getComputedEntries();
  const rows = [];

  // Sheet Title Header
  rows.push(["Vegetable Intake (Purchases & Profit Tracker)"]);
  rows.push(["Excel Generated by VegCalc Pro", "Date: " + new Date().toISOString().split("T")[0]]);
  rows.push([]); // blank row

  // Table Column Headers (Row 4 in Excel sheet)
  rows.push([
    "Date",
    "Item",
    "Quantity (kg)",
    "Total Price Paid (₹)",
    "Price per kg (₹)",
    "Retail Price per kg (₹)",
    "Profit per kg (₹)",
    "Total Profit (₹)",
    "Margin %",
    "Supplier / Notes"
  ]);

  // Data Rows with dynamic Excel formulas
  computed.forEach((entry, idx) => {
    const r = idx + 5; // starting at row 5
    rows.push([
      entry.date || "",
      entry.item || "",
      entry.qty,
      entry.totalPrice,
      { t: "n", f: `IF(OR(C${r}="",C${r}=0),"",D${r}/C${r})` },
      entry.retailPrice,
      { t: "n", f: `IF(OR(F${r}="",E${r}=""),"",F${r}-E${r})` },
      { t: "n", f: `IF(OR(G${r}="",C${r}=""),"",G${r}*C${r})` },
      { t: "n", f: `IF(OR(F${r}="",F${r}=0),"",G${r}/F${r})` },
      entry.supplier || ""
    ]);
  });

  // Totals Row
  const lastDataRow = computed.length + 4;
  rows.push([]);
  rows.push([
    "TOTAL",
    `${computed.length} items`,
    { t: "n", f: `SUM(C5:C${lastDataRow})` },
    { t: "n", f: `SUM(D5:D${lastDataRow})` },
    "",
    "",
    "",
    { t: "n", f: `SUM(H5:H${lastDataRow})` },
    "",
    ""
  ]);

  const ws = XLSX.utils.aoa_to_sheet(rows);

  // Set Column Widths
  ws["!cols"] = [
    { wch: 14 }, // Date
    { wch: 22 }, // Item
    { wch: 15 }, // Qty
    { wch: 20 }, // Total Paid
    { wch: 18 }, // Price/kg
    { wch: 20 }, // Retail Price
    { wch: 16 }, // Profit/kg
    { wch: 18 }, // Total Profit
    { wch: 12 }, // Margin %
    { wch: 26 }  // Supplier
  ];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Intake");

  XLSX.writeFile(wb, `Veg_Business_Intake_${new Date().toISOString().split("T")[0]}.xlsx`);
  showToast("Excel spreadsheet exported successfully!", "success");
}

function exportToCSV() {
  const computed = getComputedEntries();
  const headers = ["Date", "Item", "Quantity (kg)", "Total Price Paid (INR)", "Price per kg (INR)", "Retail Price (INR)", "Profit per kg (INR)", "Total Profit (INR)", "Margin Pct", "Supplier / Notes"];
  const csvRows = [headers.join(",")];

  computed.forEach(e => {
    const row = [
      `"${e.date || ""}"`,
      `"${(e.item || "").replace(/"/g, '""')}"`,
      e.qty,
      e.totalPrice,
      e.pricePerKg.toFixed(2),
      e.retailPrice,
      e.profitPerKg.toFixed(2),
      e.totalProfit.toFixed(2),
      e.marginPct.toFixed(1) + "%",
      `"${(e.supplier || "").replace(/"/g, '""')}"`
    ];
    csvRows.push(row.join(","));
  });

  const blob = new Blob([csvRows.join("\n")], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `Veg_Business_Intake_${new Date().toISOString().split("T")[0]}.csv`;
  a.click();
  URL.revokeObjectURL(url);
  showToast("CSV file exported", "success");
}

function importExcelFile(e) {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (evt) => {
    try {
      const data = new Uint8Array(evt.target.result);
      const workbook = XLSX.read(data, { type: "array" });

      // Look for sheet named "Intake" or first sheet
      const sheetName = workbook.SheetNames.includes("Intake") ? "Intake" : workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

      // Find header row containing "Item" or "Quantity"
      let headerRowIdx = -1;
      for (let i = 0; i < Math.min(jsonData.length, 10); i++) {
        const row = jsonData[i].map(c => String(c).toLowerCase());
        if (row.some(c => c.includes("item") || c.includes("quantity") || c.includes("price"))) {
          headerRowIdx = i;
          break;
        }
      }

      if (headerRowIdx === -1) {
        showToast("Could not recognize Intake table columns in this file", "error");
        return;
      }

      const headers = jsonData[headerRowIdx].map(h => String(h).trim().toLowerCase());
      const colIdxMap = {
        date: headers.findIndex(h => h.includes("date")),
        item: headers.findIndex(h => h.includes("item")),
        qty: headers.findIndex(h => h.includes("quantity") || h.includes("qty") || h.includes("kg")),
        totalPrice: headers.findIndex(h => h.includes("total price") || h.includes("price paid") || h.includes("cost")),
        retailPrice: headers.findIndex(h => h.includes("retail") || h.includes("selling")),
        supplier: headers.findIndex(h => h.includes("supplier") || h.includes("notes"))
      };

      const newEntries = [];
      for (let i = headerRowIdx + 1; i < jsonData.length; i++) {
        const row = jsonData[i];
        if (!row || row.length === 0) continue;

        const itemName = colIdxMap.item !== -1 && row[colIdxMap.item] ? String(row[colIdxMap.item]).trim() : "";
        if (!itemName || itemName.toLowerCase() === "total" || itemName.toLowerCase().includes("example")) continue;

        let dateVal = colIdxMap.date !== -1 && row[colIdxMap.date] ? row[colIdxMap.date] : "";
        if (typeof dateVal === "number") {
          // Convert Excel serial date
          const dateObj = new Date((dateVal - (25567 + 2)) * 86400 * 1000);
          dateVal = dateObj.toISOString().split("T")[0];
        }

        const qty = colIdxMap.qty !== -1 ? parseFloat(row[colIdxMap.qty]) || 0 : 0;
        const totalPrice = colIdxMap.totalPrice !== -1 ? parseFloat(row[colIdxMap.totalPrice]) || 0 : 0;
        const retailPrice = colIdxMap.retailPrice !== -1 ? parseFloat(row[colIdxMap.retailPrice]) || 0 : 0;
        const supplier = colIdxMap.supplier !== -1 && row[colIdxMap.supplier] ? String(row[colIdxMap.supplier]).trim() : "";

        newEntries.push({
          id: "v-row-" + Date.now() + "-" + i,
          date: String(dateVal || "2026-10-22"),
          item: itemName,
          qty,
          totalPrice,
          retailPrice,
          supplier
        });
      }

      if (newEntries.length > 0) {
        state.entries = newEntries;
        saveData();
        renderApp();
        showToast(`Imported ${newEntries.length} items from ${file.name}`, "success");
      } else {
        showToast("No valid vegetable rows found in sheet", "error");
      }
    } catch (err) {
      console.error(err);
      showToast("Error reading Excel file: " + err.message, "error");
    }
  };
  reader.readAsArrayBuffer(file);
}

/* =========================================================
   10. Mandi Spot Bargain Calculator
   ========================================================= */
function initMandiCalculator() {
  const modal = document.getElementById("mandi-calc-modal");
  const btnToggle = document.getElementById("btn-calc-modal-toggle");
  const btnClose = document.getElementById("btn-close-calc-modal");
  const btnCancel = document.getElementById("btn-calc-cancel");
  const btnAdd = document.getElementById("btn-calc-add-to-sheet");

  const nameInput = document.getElementById("mc-name");
  const qtyInput = document.getElementById("mc-qty");
  const totalInput = document.getElementById("mc-total-price");
  const retailInput = document.getElementById("mc-retail");

  function updateModalCalc() {
    const qty = parseFloat(qtyInput.value) || 0;
    const total = parseFloat(totalInput.value) || 0;
    const retail = parseFloat(retailInput.value) || 0;

    const costPerKg = qty > 0 ? total / qty : 0;
    const profitPerKg = retail > 0 ? retail - costPerKg : 0;
    const totalProfit = profitPerKg * qty;
    const margin = retail > 0 ? (profitPerKg / retail) * 100 : 0;

    document.getElementById("mc-res-cost").textContent = `${formatCurrency(costPerKg)} / kg`;
    document.getElementById("mc-res-profit-kg").textContent = `${profitPerKg >= 0 ? "+" : ""}${formatCurrency(profitPerKg)} / kg`;
    document.getElementById("mc-res-total-profit").textContent = `${totalProfit >= 0 ? "+" : ""}${formatCurrency(totalProfit)}`;
    document.getElementById("mc-res-margin").textContent = `${margin.toFixed(1)}%`;
  }

  [qtyInput, totalInput, retailInput].forEach(inp => inp.addEventListener("input", updateModalCalc));

  btnToggle.addEventListener("click", () => {
    modal.style.display = "flex";
    nameInput.focus();
    updateModalCalc();
  });

  const closeModal = () => { modal.style.display = "none"; };
  btnClose.addEventListener("click", closeModal);
  btnCancel.addEventListener("click", closeModal);

  btnAdd.addEventListener("click", () => {
    const item = nameInput.value.trim() || "Vegetable";
    const qty = parseFloat(qtyInput.value) || 0;
    const total = parseFloat(totalInput.value) || 0;
    const retail = parseFloat(retailInput.value) || 0;

    const newEntry = {
      id: "v-row-" + Date.now(),
      date: new Date().toISOString().split("T")[0],
      item,
      qty,
      totalPrice: total,
      retailPrice: retail,
      supplier: "Spot Bargain"
    };

    state.entries.push(newEntry);
    saveData();
    renderApp();
    closeModal();
    showToast(`Added ${item} bargain to spreadsheet!`, "success");
  });
}

/* =========================================================
   11. Daily Summary Modal View
   ========================================================= */
function openDailySummaryModal() {
  const modal = document.getElementById("daily-summary-modal");
  const content = document.getElementById("daily-summary-content");
  const computed = getFilteredEntries();
  const totals = getTotals(computed);

  let html = `
    <div style="display: grid; grid-template-columns: repeat(4, 1fr); gap: 12px; margin-bottom: 20px;">
      <div style="background: #f8fafc; padding: 12px; border-radius: 8px; border: 1px solid #e2e8f0;">
        <div style="font-size: 0.72rem; color: #64748b; font-weight: 700;">TOTAL ITEMS</div>
        <div style="font-size: 1.3rem; font-weight: 800; font-family: var(--font-mono);">${totals.count}</div>
      </div>
      <div style="background: #f8fafc; padding: 12px; border-radius: 8px; border: 1px solid #e2e8f0;">
        <div style="font-size: 0.72rem; color: #64748b; font-weight: 700;">TOTAL WEIGHT</div>
        <div style="font-size: 1.3rem; font-weight: 800; font-family: var(--font-mono);">${formatNumber(totals.totalQty, 1)} kg</div>
      </div>
      <div style="background: #f8fafc; padding: 12px; border-radius: 8px; border: 1px solid #e2e8f0;">
        <div style="font-size: 0.72rem; color: #64748b; font-weight: 700;">TOTAL MANDI COST</div>
        <div style="font-size: 1.3rem; font-weight: 800; font-family: var(--font-mono); color: #b45309;">${formatCurrency(totals.totalCost)}</div>
      </div>
      <div style="background: #f0fdf4; padding: 12px; border-radius: 8px; border: 1px solid #bbf7d0;">
        <div style="font-size: 0.72rem; color: #166534; font-weight: 700;">ESTIMATED PROFIT</div>
        <div style="font-size: 1.3rem; font-weight: 800; font-family: var(--font-mono); color: #047857;">${formatCurrency(totals.totalProfit)}</div>
      </div>
    </div>

    <h4 style="margin-bottom: 10px; font-size: 0.95rem;">Profit & Margin Ranking by Item</h4>
    <table style="width: 100%; border-collapse: collapse; font-size: 0.82rem; border: 1px solid #e2e8f0;">
      <thead>
        <tr style="background: #f1f5f9; text-align: left;">
          <th style="padding: 8px; border: 1px solid #e2e8f0;">Item Name</th>
          <th style="padding: 8px; border: 1px solid #e2e8f0; text-align: right;">Bought (kg)</th>
          <th style="padding: 8px; border: 1px solid #e2e8f0; text-align: right;">Cost / kg</th>
          <th style="padding: 8px; border: 1px solid #e2e8f0; text-align: right;">Retail / kg</th>
          <th style="padding: 8px; border: 1px solid #e2e8f0; text-align: right;">Profit / kg</th>
          <th style="padding: 8px; border: 1px solid #e2e8f0; text-align: right;">Total Profit</th>
          <th style="padding: 8px; border: 1px solid #e2e8f0; text-align: right;">Margin %</th>
        </tr>
      </thead>
      <tbody>
  `;

  const sorted = [...computed].sort((a, b) => b.totalProfit - a.totalProfit);
  sorted.forEach(item => {
    html += `
      <tr style="border-bottom: 1px solid #e2e8f0;">
        <td style="padding: 8px; font-weight: 600;">${escapeHtml(item.item)}</td>
        <td style="padding: 8px; text-align: right; font-family: var(--font-mono);">${formatNumber(item.qty, 1)} kg</td>
        <td style="padding: 8px; text-align: right; font-family: var(--font-mono);">${formatCurrency(item.pricePerKg)}</td>
        <td style="padding: 8px; text-align: right; font-family: var(--font-mono);">${formatCurrency(item.retailPrice)}</td>
        <td style="padding: 8px; text-align: right; font-family: var(--font-mono); color: #059669; font-weight: 700;">+${formatCurrency(item.profitPerKg)}</td>
        <td style="padding: 8px; text-align: right; font-family: var(--font-mono); color: #047857; font-weight: 800;">+${formatCurrency(item.totalProfit)}</td>
        <td style="padding: 8px; text-align: right; font-family: var(--font-mono);"><span style="background: #f5f3ff; color: #7c3aed; padding: 2px 6px; border-radius: 4px; font-weight: 600;">${item.marginPct.toFixed(1)}%</span></td>
      </tr>
    `;
  });

  html += `
      </tbody>
    </table>
  `;

  content.innerHTML = html;
  modal.style.display = "flex";

  const closeModal = () => {
    modal.style.display = "none";
    document.querySelectorAll(".sheet-tab").forEach(t => t.classList.remove("active"));
    document.querySelector('.sheet-tab[data-tab="intake"]').classList.add("active");
  };

  document.getElementById("btn-close-summary-modal").onclick = closeModal;
  document.getElementById("btn-close-summary-btn").onclick = closeModal;
  document.getElementById("btn-print-summary").onclick = () => window.print();
}

/* =========================================================
   12. Notification Toast
   ========================================================= */
function showToast(message, type = "info") {
  const container = document.getElementById("toast-container");
  const toast = document.createElement("div");
  toast.className = `toast ${type === "error" ? "toast-error" : ""}`;
  toast.innerHTML = `<span>${type === "error" ? "⚠️" : "✅"}</span> <span>${message}</span>`;
  container.appendChild(toast);

  setTimeout(() => {
    toast.style.opacity = "0";
    toast.style.transform = "translateX(100%)";
    toast.style.transition = "all 0.3s ease";
    setTimeout(() => toast.remove(), 300);
  }, 2800);
}

/* =========================================================
   13. PWA Service Worker & Mobile App Handlers
   ========================================================= */
let deferredPrompt = null;

// Register Service Worker for 100% Offline Capability
if ("serviceWorker" in navigator) {
  window.addEventListener("load", () => {
    navigator.serviceWorker.register("./service-worker.js")
      .then((reg) => {
        console.log("[PWA] Service worker registered successfully:", reg.scope);
      })
      .catch((err) => {
        console.log("[PWA] Service worker registration failed:", err);
      });
  });
}

// Capture Android Chrome Install Prompt
window.addEventListener("beforeinstallprompt", (e) => {
  e.preventDefault();
  deferredPrompt = e;
  const directInstallBtn = document.getElementById("btn-pwa-direct-install");
  if (directInstallBtn) {
    directInstallBtn.style.display = "inline-flex";
  }
});

// Mobile Modal Handlers
document.addEventListener("DOMContentLoaded", () => {
  const mobileModal = document.getElementById("mobile-app-modal");
  const btnOpenMobile = document.getElementById("btn-mobile-app");
  const btnCloseModal = document.getElementById("btn-close-mobile-modal");
  const btnCloseModal2 = document.getElementById("btn-close-mobile-btn");
  const btnCopyLink = document.getElementById("btn-copy-mobile-link");
  const btnDirectInstall = document.getElementById("btn-pwa-direct-install");

  if (btnOpenMobile && mobileModal) {
    btnOpenMobile.addEventListener("click", () => {
      mobileModal.style.display = "flex";
    });
  }

  const closeMobile = () => {
    if (mobileModal) mobileModal.style.display = "none";
  };

  if (btnCloseModal) btnCloseModal.addEventListener("click", closeMobile);
  if (btnCloseModal2) btnCloseModal2.addEventListener("click", closeMobile);

  if (btnCopyLink) {
    btnCopyLink.addEventListener("click", () => {
      const link = "http://192.168.43.63:5174/";
      navigator.clipboard.writeText(link).then(() => {
        showToast("Phone link copied to clipboard!", "success");
      }).catch(() => {
        showToast("Link: " + link, "info");
      });
    });
  }

  if (btnDirectInstall) {
    btnDirectInstall.addEventListener("click", async () => {
      if (deferredPrompt) {
        deferredPrompt.prompt();
        const { outcome } = await deferredPrompt.userChoice;
        console.log(`[PWA] User response to install prompt: ${outcome}`);
        deferredPrompt = null;
        closeMobile();
      } else {
        showToast("Follow the steps below to install on your browser!", "info");
      }
    });
  }
});
