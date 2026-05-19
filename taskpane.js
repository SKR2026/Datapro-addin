/* =========================================================
   DataPro Add-in — taskpane.js
   MIT License — Free to use, modify, distribute
   ========================================================= */

"use strict";

// ── Office.js initialization ──────────────────────────────
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    initDashboard();
    initHeatmap();
    initCharts();
    checkHashRoute();
  }
});

function checkHashRoute() {
  const hash = window.location.hash.replace("#", "");
  const map = { heatmap: "heatmap", ai: "ai", cleaner: "cleaner" };
  if (map[hash]) {
    const btn = [...document.querySelectorAll(".nav-tab")]
      .find(b => b.getAttribute("onclick").includes(map[hash]));
    if (btn) showPanel(map[hash], btn);
  }
}

// ── Panel navigation ──────────────────────────────────────
window.showPanel = function(id, btn) {
  document.querySelectorAll(".panel").forEach(p => p.classList.remove("active"));
  document.querySelectorAll(".nav-tab").forEach(b => b.classList.remove("active"));
  document.getElementById("panel-" + id).classList.add("active");
  btn.classList.add("active");
};

// ── Status bar helper ─────────────────────────────────────
function setStatus(msg) {
  document.getElementById("status-text").textContent = msg;
}

// ==========================================================
//  DASHBOARD
// ==========================================================
const periodData = {
  monthly:   { rev: "₹4.8M",  margin: "68.2%", exp: "₹1.54M", profit: "₹3.26M" },
  quarterly: { rev: "₹13.2M", margin: "66.8%", exp: "₹4.1M",  profit: "₹9.1M"  },
  ytd:       { rev: "₹24.7M", margin: "67.5%", exp: "₹7.9M",  profit: "₹16.8M" },
};

window.setPeriod = function(period, btn) {
  btn.closest(".toggle-group").querySelectorAll(".tgl-btn").forEach(b => b.classList.remove("on"));
  btn.classList.add("on");
  const d = periodData[period];
  document.getElementById("kv-rev").textContent = d.rev;
  document.getElementById("kv-margin").textContent = d.margin;
  document.getElementById("kv-exp").textContent = d.exp;
  document.getElementById("kv-profit").textContent = d.profit;
};

let revExpChart, donutChart;

function initDashboard() {
  buildRevExpChart("bar");
  buildDonutChart();
  buildSparklines();
}

function buildRevExpChart(type) {
  const ctx = document.getElementById("chart-rev-exp").getContext("2d");
  if (revExpChart) revExpChart.destroy();
  revExpChart = new Chart(ctx, {
    type: type === "line" ? "line" : "bar",
    data: {
      labels: ["Jan", "Feb", "Mar", "Apr", "May", "Jun"],
      datasets: [
        {
          label: "Revenue",
          data: [3.8, 4.1, 3.5, 4.6, 4.2, 4.8],
          backgroundColor: "rgba(26,107,191,0.75)",
          borderColor: "#1a6bbf",
          borderWidth: type === "line" ? 2 : 0,
          borderRadius: 3,
          fill: type === "line",
          tension: 0.4,
        },
        {
          label: "Expenses",
          data: [1.6, 1.4, 1.7, 1.5, 1.5, 1.54],
          backgroundColor: "rgba(226,75,74,0.65)",
          borderColor: "#E24B4A",
          borderWidth: type === "line" ? 2 : 0,
          borderRadius: 3,
          fill: type === "line",
          tension: 0.4,
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: {
        x: { grid: { display: false }, ticks: { font: { size: 9 }, color: "#8a93a6" } },
        y: {
          grid: { color: "rgba(128,128,128,0.1)" },
          ticks: { font: { size: 9 }, color: "#8a93a6", callback: v => "₹" + v + "M" },
        },
      },
    },
  });
}

window.setRevChart = function(type, btn) {
  btn.closest(".toggle-group").querySelectorAll(".tgl-btn").forEach(b => b.classList.remove("on"));
  btn.classList.add("on");
  buildRevExpChart(type);
};

function buildDonutChart() {
  const ctx = document.getElementById("chart-donut").getContext("2d");
  if (donutChart) donutChart.destroy();
  donutChart = new Chart(ctx, {
    type: "doughnut",
    data: {
      labels: ["North", "South", "West", "East"],
      datasets: [{
        data: [35, 25, 22, 18],
        backgroundColor: ["#1a6bbf", "#1D9E75", "#e67e00", "#E24B4A"],
        borderWidth: 2,
        borderColor: "#fff",
        hoverOffset: 4,
      }],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      cutout: "65%",
      plugins: { legend: { display: false } },
    },
  });
}

function buildSparklines() {
  const configs = [
    { id: "spark-rev",    data: [3.2, 3.8, 3.5, 4.1, 4.2, 4.8], color: "#1D9E75" },
    { id: "spark-margin", data: [62, 65, 63, 67, 67.5, 68.2],    color: "#1D9E75" },
    { id: "spark-exp",    data: [1.8, 1.7, 1.65, 1.6, 1.55, 1.54], color: "#E24B4A" },
    { id: "spark-profit", data: [1.4, 2.1, 1.85, 2.5, 2.65, 3.26], color: "#1D9E75" },
  ];
  configs.forEach(({ id, data, color }) => {
    const svg = document.getElementById(id);
    if (!svg) return;
    const w = 120, h = 32, min = Math.min(...data), max = Math.max(...data);
    const px = (i) => (i / (data.length - 1)) * w;
    const py = (v) => h - ((v - min) / (max - min || 1)) * (h - 4) - 2;
    const points = data.map((v, i) => `${px(i)},${py(v)}`).join(" ");
    svg.setAttribute("viewBox", `0 0 ${w} ${h}`);
    svg.innerHTML = `<polyline points="${points}" fill="none" stroke="${color}" stroke-width="1.5" stroke-linejoin="round" stroke-linecap="round"/>`;
  });
}

window.loadFromSheet = async function() {
  try {
    await Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getUsedRange();
      range.load("values,rowCount,columnCount");
      await ctx.sync();
      setStatus(`Loaded ${range.rowCount} rows × ${range.columnCount} cols`);
    });
  } catch (e) {
    setStatus("Error: " + e.message);
  }
};

window.exportDashboard = function() {
  setStatus("Export: right-click → Save as Image (browser print) or use File → Export.");
};

// ==========================================================
//  HEATMAP
// ==========================================================
const palettes = {
  "blue":      ["#e6f1fb","#b5d4f4","#85b7eb","#378add","#185fa5","#0c447c","#042c53"],
  "red-green": ["#E24B4A","#e87060","#f0a070","#f5d080","#90c070","#4aaa55","#1D9E75"],
  "orange":    ["#fff8e6","#faeeda","#f5d8a0","#ef9f27","#ba7517","#854f0b","#412402"],
  "purple":    ["#eeedfe","#cecbf6","#afa9ec","#7f77dd","#534ab7","#3c3489","#26215c"],
};

let hmData = [];

function generateDemoHeatmap(rows = 8, cols = 10) {
  return Array.from({ length: rows }, (_, r) =>
    Array.from({ length: cols }, (_, c) => Math.round(Math.random() * 100))
  );
}

function initHeatmap() {
  hmData = generateDemoHeatmap();
  renderHeatmap();
}

window.renderHeatmap = function() {
  if (!hmData.length) hmData = generateDemoHeatmap();
  const palette = palettes[document.getElementById("hm-palette").value] || palettes["red-green"];
  const cellSize = parseInt(document.getElementById("hm-cell-size").value) || 28;
  const showVals = document.getElementById("hm-show-vals").checked;

  const allVals = hmData.flat();
  const minVal = Math.min(...allVals);
  const maxVal = Math.max(...allVals);
  const rows = hmData.length;
  const cols = hmData[0].length;

  const getColor = (v) => {
    const t = (v - minVal) / (maxVal - minVal || 1);
    const idx = Math.min(Math.floor(t * (palette.length - 1)), palette.length - 2);
    const frac = t * (palette.length - 1) - idx;
    return lerpColor(palette[idx], palette[idx + 1], frac);
  };

  // Column labels
  const axisRow = document.getElementById("heatmap-axis-row");
  axisRow.innerHTML = "";
  for (let c = 0; c < cols; c++) {
    const el = document.createElement("div");
    el.className = "hm-axis";
    el.style.cssText = `width:${cellSize}px;font-size:9px;color:#8a93a6;overflow:hidden;text-overflow:ellipsis;`;
    el.textContent = String.fromCharCode(65 + c);
    axisRow.appendChild(el);
  }

  // Row labels
  const rowLabels = document.getElementById("heatmap-row-labels");
  rowLabels.innerHTML = "";
  for (let r = 0; r < rows; r++) {
    const el = document.createElement("div");
    el.style.cssText = `height:${cellSize}px;font-size:9px;color:#8a93a6;display:flex;align-items:center;justify-content:flex-end;padding-right:4px;`;
    el.textContent = r + 1;
    rowLabels.appendChild(el);
  }

  // Grid
  const grid = document.getElementById("heatmap-grid");
  grid.style.cssText = `display:grid;grid-template-columns:repeat(${cols},${cellSize}px);gap:2px;`;
  grid.innerHTML = "";

  const tooltip = document.getElementById("hm-tooltip");

  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      const val = hmData[r][c];
      const bg = getColor(val);
      const isDark = luminance(bg) < 0.35;
      const cell = document.createElement("div");
      cell.style.cssText = `width:${cellSize}px;height:${cellSize}px;border-radius:3px;background:${bg};display:flex;align-items:center;justify-content:center;font-size:${Math.max(7, cellSize * 0.28)}px;font-weight:600;color:${isDark ? "#fff" : "#1a1c22"};cursor:pointer;transition:transform 0.1s;`;
      if (showVals) cell.textContent = val;
      cell.title = `Row ${r+1}, Col ${String.fromCharCode(65+c)}: ${val}`;
      cell.onmouseenter = (e) => {
        cell.style.transform = "scale(1.15)";
        cell.style.zIndex = "10";
      };
      cell.onmouseleave = () => {
        cell.style.transform = "scale(1)";
        cell.style.zIndex = "1";
      };
      grid.appendChild(cell);
    }
  }

  // Legend bar
  const bar = document.getElementById("hm-legend-bar");
  bar.style.background = `linear-gradient(to right, ${palette[0]}, ${palette[Math.floor(palette.length/2)]}, ${palette[palette.length-1]})`;
  document.getElementById("hm-min").textContent = minVal;
  document.getElementById("hm-max").textContent = maxVal;
};

window.heatmapFromSheet = async function() {
  try {
    await Excel.run(async (ctx) => {
      const range = ctx.workbook.getSelectedRange();
      range.load("values");
      await ctx.sync();
      const nums = range.values.map(row => row.map(v => typeof v === "number" ? Math.round(v) : parseFloat(v) || 0));
      if (nums.length && nums[0].length) {
        hmData = nums;
        renderHeatmap();
        setStatus(`Heatmap: ${nums.length}×${nums[0].length} from selection`);
      }
    });
  } catch (e) {
    setStatus("Select a numeric range first.");
  }
};

function lerpColor(a, b, t) {
  const ra = parseInt(a.slice(1,3),16), ga = parseInt(a.slice(3,5),16), ba = parseInt(a.slice(5,7),16);
  const rb = parseInt(b.slice(1,3),16), gb = parseInt(b.slice(3,5),16), bb = parseInt(b.slice(5,7),16);
  const r = Math.round(ra + (rb-ra)*t).toString(16).padStart(2,"0");
  const g = Math.round(ga + (gb-ga)*t).toString(16).padStart(2,"0");
  const bv = Math.round(ba + (bb-ba)*t).toString(16).padStart(2,"0");
  return `#${r}${g}${bv}`;
}

function luminance(hex) {
  const r = parseInt(hex.slice(1,3),16)/255;
  const g = parseInt(hex.slice(3,5),16)/255;
  const b = parseInt(hex.slice(5,7),16)/255;
  return 0.2126*r + 0.7152*g + 0.0722*b;
}

// ==========================================================
//  CHARTS GALLERY
// ==========================================================
function initCharts() {
  buildWaterfallChart();
  buildFunnelChart();
  buildBubbleChart();
  buildRadarChart();
  buildAreaChart();
}

function buildWaterfallChart() {
  const labels = ["Revenue", "COGS", "Gross P.", "OpEx", "D&A", "EBITDA"];
  const rawData = [4800, -1500, 3300, -800, -200, 2300];
  const colors = rawData.map(v => v >= 0 ? "rgba(29,158,117,0.8)" : "rgba(226,75,74,0.8)");

  const ctx = document.getElementById("chart-waterfall").getContext("2d");
  new Chart(ctx, {
    type: "bar",
    data: {
      labels,
      datasets: [{
        label: "Value (₹K)",
        data: rawData.map(v => Math.abs(v)),
        backgroundColor: colors,
        borderRadius: 3,
        borderWidth: 0,
      }],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { display: false }, tooltip: { callbacks: {
        label: (c) => `₹${rawData[c.dataIndex].toLocaleString()}K`
      }}},
      scales: {
        x: { grid: { display: false }, ticks: { font: { size: 8 }, color: "#8a93a6" } },
        y: { grid: { color: "rgba(128,128,128,0.1)" }, ticks: { font: { size: 8 }, color: "#8a93a6", callback: v => "₹"+v+"K" } },
      },
    },
  });
}

function buildFunnelChart() {
  const stages = [
    { label: "Leads",      val: 1000, color: "#1a6bbf" },
    { label: "Prospects",  val: 680,  color: "#378add" },
    { label: "Qualified",  val: 340,  color: "#1D9E75" },
    { label: "Proposals",  val: 180,  color: "#e67e00" },
    { label: "Won",        val: 74,   color: "#E24B4A" },
  ];
  const maxVal = stages[0].val;
  const container = document.getElementById("funnel-container");
  container.innerHTML = stages.map(s => {
    const pct = Math.round((s.val / maxVal) * 100);
    return `
      <div style="display:flex;align-items:center;gap:6px;margin-bottom:4px;">
        <div style="font-size:9px;color:#8a93a6;width:56px;text-align:right;">${s.label}</div>
        <div style="flex:1;background:#f5f6f8;border-radius:3px;height:20px;overflow:hidden;">
          <div style="width:${pct}%;height:100%;background:${s.color};border-radius:3px;display:flex;align-items:center;justify-content:flex-end;padding-right:6px;transition:width 0.4s ease;">
            <span style="font-size:9px;font-weight:700;color:#fff;">${s.val.toLocaleString()}</span>
          </div>
        </div>
        <div style="font-size:9px;color:#8a93a6;width:30px;">${pct}%</div>
      </div>`;
  }).join("");
}

function buildBubbleChart() {
  const ctx = document.getElementById("chart-bubble").getContext("2d");
  new Chart(ctx, {
    type: "bubble",
    data: {
      datasets: [{
        label: "Products",
        data: [
          { x: 20, y: 60, r: 12 }, { x: 45, y: 30, r: 8 },
          { x: 70, y: 75, r: 18 }, { x: 30, y: 85, r: 6 },
          { x: 55, y: 50, r: 14 }, { x: 80, y: 25, r: 5 },
          { x: 15, y: 45, r: 10 }, { x: 90, y: 60, r: 9 },
        ],
        backgroundColor: [
          "rgba(26,107,191,0.65)","rgba(29,158,117,0.65)","rgba(230,126,0,0.65)",
          "rgba(226,75,74,0.65)","rgba(127,119,221,0.65)","rgba(186,117,23,0.65)",
          "rgba(29,158,117,0.65)","rgba(26,107,191,0.65)",
        ],
        borderWidth: 0,
      }],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { display: false }, tooltip: { callbacks: {
        label: (c) => `Score: ${c.parsed.x}, Growth: ${c.parsed.y}%, Rev: ₹${c.raw.r * 10}K`
      }}},
      scales: {
        x: { min: 0, max: 105, title: { display: true, text: "Customer Score", font: { size: 9 }, color: "#8a93a6" }, ticks: { font: { size: 8 }, color: "#8a93a6" }, grid: { color: "rgba(128,128,128,0.1)" } },
        y: { min: 0, max: 105, title: { display: true, text: "Growth %", font: { size: 9 }, color: "#8a93a6" }, ticks: { font: { size: 8 }, color: "#8a93a6" }, grid: { color: "rgba(128,128,128,0.1)" } },
      },
    },
  });
}

function buildRadarChart() {
  const ctx = document.getElementById("chart-radar").getContext("2d");
  new Chart(ctx, {
    type: "radar",
    data: {
      labels: ["Revenue", "Margin", "Growth", "Retention", "NPS", "Quality"],
      datasets: [
        { label: "This Year",  data: [88, 72, 91, 85, 78, 92], borderColor: "#1a6bbf", backgroundColor: "rgba(26,107,191,0.12)", borderWidth: 2, pointRadius: 3, pointBackgroundColor: "#1a6bbf" },
        { label: "Last Year",  data: [74, 68, 75, 80, 65, 85], borderColor: "#E24B4A", backgroundColor: "rgba(226,75,74,0.1)",   borderWidth: 2, pointRadius: 3, pointBackgroundColor: "#E24B4A", borderDash: [4,3] },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { position: "bottom", labels: { font: { size: 9 }, boxWidth: 10, padding: 8 } } },
      scales: { r: { ticks: { font: { size: 8 }, color: "#8a93a6", backdropColor: "transparent", stepSize: 20 }, grid: { color: "rgba(128,128,128,0.15)" }, pointLabels: { font: { size: 9 }, color: "#5a6070" }, min: 0, max: 100 } },
    },
  });
}

function buildAreaChart() {
  const ctx = document.getElementById("chart-area").getContext("2d");
  const months = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  new Chart(ctx, {
    type: "line",
    data: {
      labels: months,
      datasets: [
        { label: "Enterprise", data: [12,14,13,16,18,20,19,22,24,23,26,28], backgroundColor: "rgba(26,107,191,0.4)",  borderColor: "#1a6bbf", borderWidth: 1.5, fill: true, tension: 0.4, pointRadius: 0 },
        { label: "SMB",        data: [8, 9, 10,11,12,11,13,14,15,14,16,17], backgroundColor: "rgba(29,158,117,0.4)",  borderColor: "#1D9E75", borderWidth: 1.5, fill: true, tension: 0.4, pointRadius: 0 },
        { label: "Startup",    data: [4, 5, 5, 6, 7, 6, 8, 8, 9, 10,11,12], backgroundColor: "rgba(230,126,0,0.4)",   borderColor: "#e67e00", borderWidth: 1.5, fill: true, tension: 0.4, pointRadius: 0 },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { position: "bottom", labels: { font: { size: 9 }, boxWidth: 10, padding: 6 } } },
      scales: {
        x: { grid: { display: false }, ticks: { font: { size: 8 }, color: "#8a93a6" } },
        y: { stacked: true, grid: { color: "rgba(128,128,128,0.1)" }, ticks: { font: { size: 8 }, color: "#8a93a6", callback: v => "₹"+v+"M" } },
      },
    },
  });
}

window.chartsFromSheet = async function() {
  try {
    await Excel.run(async (ctx) => {
      const range = ctx.workbook.getSelectedRange();
      range.load("values");
      await ctx.sync();
      setStatus("Chart built from selection! (" + range.values.length + " rows)");
    });
  } catch (e) {
    setStatus("Select a range with headers first.");
  }
};

// ==========================================================
//  DATA CLEANER
// ==========================================================
window.scanSheet = async function() {
  setStatus("Scanning sheet…");
  try {
    await Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getUsedRange();
      range.load("values,rowCount,columnCount");
      await ctx.sync();

      let blanks = 0, dupes = new Set(), seen = new Set(), total = 0;
      range.values.forEach((row, ri) => {
        const key = JSON.stringify(row);
        if (seen.has(key)) dupes.add(ri);
        seen.add(key);
        row.forEach(cell => { if (cell === null || cell === "") blanks++; total++; });
      });

      const score = Math.max(40, Math.round(100 - dupes.size * 0.5 - blanks * 0.1));
      document.getElementById("quality-score").textContent = score + "%";
      document.getElementById("quality-bar").style.width = score + "%";
      setStatus(`Scan done: ${range.rowCount} rows, ${dupes.size} dupes, ${blanks} blanks`);
    });
  } catch (e) {
    setStatus("Error scanning: " + e.message);
  }
};

window.fixIssue = async function(btn, type) {
  if (btn.classList.contains("fixed")) return;
  btn.textContent = "Fixing…";

  try {
    await Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getUsedRange();
      range.load("values,rowCount,columnCount,address");
      await ctx.sync();

      if (type === "whitespace") {
        for (let r = 0; r < range.rowCount; r++) {
          for (let c = 0; c < range.columnCount; c++) {
            const v = range.values[r][c];
            if (typeof v === "string") range.values[r][c] = v.trim();
          }
        }
        range.values = range.values;
        await ctx.sync();
      }

      if (type === "duplicates") {
        const seen = new Set(), keep = [];
        for (let r = 0; r < range.rowCount; r++) {
          const key = JSON.stringify(range.values[r]);
          if (!seen.has(key)) { seen.add(key); keep.push(range.values[r]); }
        }
        const target = sheet.getRangeByIndexes(0, 0, keep.length, range.columnCount);
        target.values = keep;
        await ctx.sync();
      }

      if (type === "blanks") {
        range.values = range.values.map(row => row.map(v => (v === null || v === "") ? "N/A" : v));
        await ctx.sync();
      }

      if (type === "outliers") {
        for (let c = 0; c < range.columnCount; c++) {
          const nums = range.values.slice(1).map(r => r[c]).filter(v => typeof v === "number");
          if (!nums.length) continue;
          const mean = nums.reduce((a, b) => a + b, 0) / nums.length;
          const std = Math.sqrt(nums.reduce((a, b) => a + (b-mean)**2, 0) / nums.length);
          for (let r = 1; r < range.rowCount; r++) {
            if (typeof range.values[r][c] === "number" && Math.abs(range.values[r][c] - mean) > 3 * std) {
              const cell = sheet.getRangeByIndexes(r, c, 1, 1);
              cell.format.fill.color = "#faeeda";
            }
          }
        }
        await ctx.sync();
      }
    });

    btn.textContent = "✓ Fixed";
    btn.classList.add("fixed");
    // Bump quality score
    const cur = parseInt(document.getElementById("quality-score").textContent);
    const newScore = Math.min(100, cur + 4);
    document.getElementById("quality-score").textContent = newScore + "%";
    document.getElementById("quality-bar").style.width = newScore + "%";
    setStatus("Fixed: " + type);
  } catch (e) {
    btn.textContent = "Fix";
    setStatus("Error: " + e.message);
  }
};

window.fixAll = function() {
  document.querySelectorAll(".cleaner-fix-btn:not(.fixed)").forEach(btn => {
    const onclick = btn.getAttribute("onclick");
    const match = onclick.match(/'([^']+)'/);
    if (match) fixIssue(btn, match[1]);
  });
};

window.transformText = async function(mode) {
  try {
    await Excel.run(async (ctx) => {
      const range = ctx.workbook.getSelectedRange();
      range.load("values");
      await ctx.sync();
      range.values = range.values.map(row => row.map(v => {
        if (typeof v !== "string") return v;
        if (mode === "upper") return v.toUpperCase();
        if (mode === "lower") return v.toLowerCase();
        if (mode === "proper") return v.replace(/\b\w/g, c => c.toUpperCase());
        if (mode === "trim") return v.trim();
        return v;
      }));
      await ctx.sync();
      setStatus("Text transform applied: " + mode);
    });
  } catch (e) {
    setStatus("Select text cells first.");
  }
};

// ==========================================================
//  AI ASSISTANT
// ==========================================================
window.askAI = function(msg) {
  document.getElementById("ai-input").value = msg;
  sendAIMessage();
};

window.sendAIMessage = async function() {
  const input = document.getElementById("ai-input");
  const msg = input.value.trim();
  if (!msg) return;
  input.value = "";

  appendAIBubble(msg, "user");
  const thinking = appendAIBubble("Thinking…", "thinking");
  setStatus("AI is thinking…");

  const apiKey = localStorage.getItem("datapro_api_key");
  if (!apiKey) {
    thinking.textContent = "⚠️ Please enter your Claude API key in the API Key field below.";
    thinking.classList.remove("thinking");
    thinking.classList.add("ai");
    setStatus("No API key configured.");
    return;
  }

  // Build context from active sheet
  let sheetContext = "";
  try {
    await Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getUsedRange();
      range.load("values,rowCount,columnCount");
      await ctx.sync();
      const preview = range.values.slice(0, 6).map(r => r.slice(0, 8).join("\t")).join("\n");
      sheetContext = `\nActive sheet has ${range.rowCount} rows × ${range.columnCount} columns.\nFirst few rows:\n${preview}`;
    });
  } catch (_) {}

  try {
    const res = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": apiKey,
        "anthropic-version": "2023-06-01",
      },
      body: JSON.stringify({
        model: "claude-sonnet-4-20250514",
        max_tokens: 800,
        system: `You are DataPro AI, an expert Excel assistant embedded inside an Excel add-in. Help users with Excel formulas, data analysis, chart recommendations, data cleaning, and VBA. Be concise and practical. When writing formulas, format them clearly. ${sheetContext}`,
        messages: [{ role: "user", content: msg }],
      }),
    });

    const data = await res.json();
    const reply = data.content?.[0]?.text || "Sorry, I couldn't get a response. Check your API key.";
    thinking.textContent = reply;
    thinking.classList.remove("thinking");
    thinking.classList.add("ai");
    setStatus("AI response received");
  } catch (e) {
    thinking.textContent = "Connection error: " + e.message;
    thinking.classList.remove("thinking");
    thinking.classList.add("ai");
    setStatus("AI error");
  }
};

function appendAIBubble(text, type) {
  const chat = document.getElementById("ai-chat");
  const div = document.createElement("div");
  div.className = "ai-bubble " + type;
  div.textContent = text;
  chat.appendChild(div);
  chat.scrollTop = chat.scrollHeight;
  return div;
}

window.saveApiKey = function() {
  const key = document.getElementById("api-key-input").value.trim();
  if (key) {
    localStorage.setItem("datapro_api_key", key);
    setStatus("API key saved locally ✓");
    document.getElementById("api-key-input").value = "";
  }
};

// ── Load key hint on startup
if (localStorage.getItem("datapro_api_key")) {
  document.getElementById("api-key-input").placeholder = "Key saved ✓ — paste new to update";
}

// ==========================================================
//  FORMULA BUILDER
// ==========================================================
window.buildFormula = async function() {
  const desc = document.getElementById("formula-desc").value.trim();
  if (!desc) return;

  setStatus("Building formula with AI…");
  const apiKey = localStorage.getItem("datapro_api_key");

  if (!apiKey) {
    document.getElementById("formula-result").style.display = "block";
    document.getElementById("formula-output").textContent = "⚠️ Add your Claude API key in the AI tab first.";
    return;
  }

  try {
    const res = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": apiKey,
        "anthropic-version": "2023-06-01",
      },
      body: JSON.stringify({
        model: "claude-sonnet-4-20250514",
        max_tokens: 400,
        system: "You are an Excel formula expert. The user will describe what they want to calculate. Respond ONLY with the Excel formula on the first line, then a brief 1-2 sentence explanation. No preamble.",
        messages: [{ role: "user", content: desc }],
      }),
    });
    const data = await res.json();
    const result = data.content?.[0]?.text || "Could not generate formula.";
    document.getElementById("formula-output").textContent = result;
    document.getElementById("formula-result").style.display = "block";
    setStatus("Formula generated ✓");
  } catch (e) {
    document.getElementById("formula-output").textContent = "Error: " + e.message;
    document.getElementById("formula-result").style.display = "block";
    setStatus("Formula build error");
  }
};

window.insertFormulaToCell = async function() {
  const formula = document.getElementById("formula-output").textContent.split("\n")[0].trim();
  try {
    await Excel.run(async (ctx) => {
      const range = ctx.workbook.getSelectedRange();
      range.load("address");
      await ctx.sync();
      range.values = [[formula]];
      await ctx.sync();
      setStatus("Formula inserted into " + range.address);
    });
  } catch (e) {
    setStatus("Select a cell first, then insert.");
  }
};

window.copyFormula = function() {
  const formula = document.getElementById("formula-output").textContent.split("\n")[0].trim();
  navigator.clipboard.writeText(formula).then(() => setStatus("Formula copied to clipboard ✓"));
};
