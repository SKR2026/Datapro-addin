# DataPro Add-in for Excel
### Free, Open Source · MIT License · No Restrictions

A full-featured Excel add-in with smart interactive dashboards, AI-powered tools, heatmaps, and a data cleaning wizard. Zero license fees, zero restrictions.

---

## ✨ Features

| Panel | Features |
|---|---|
| 📊 **Dashboard** | Live KPI cards, sparklines, revenue vs expenses bar/line chart, regional donut chart. Monthly / Quarterly / YTD toggle. |
| 🔥 **Heatmap** | Read from selected Excel range or demo data. 4 color palettes, adjustable cell size, hover tooltips. |
| 📈 **Charts Gallery** | Waterfall (P&L Bridge), Funnel (Sales Pipeline), Bubble (Product Matrix), Radar (KPI Scorecard), Stacked Area (Segment Trend) |
| 🧹 **Data Cleaner** | Detects and fixes: duplicate rows, empty cells, whitespace, invalid dates, type mismatches, statistical outliers. Text transform (UPPER/lower/Proper/Trim). Quality score tracker. |
| 🤖 **AI Assistant** | Claude-powered chat. Reads your active sheet for context. Formula chips for quick queries. |
| 𝑓 **Formula Builder** | Describe in plain English → AI generates the Excel formula. Insert directly into the selected cell. Quick reference for XLOOKUP, SUMIFS, FILTER, LAMBDA. |

---

## 🚀 Quick Start

### Option A — Preview (no install needed)
Open `demo.html` in any browser. Fully functional except live Excel data reads.

### Option B — Install in Excel (Development)

**Prerequisites**
- Node.js 18+
- Microsoft Excel (Windows/Mac) or Excel on the Web

**Steps**

```bash
# 1. Clone / download
git clone https://github.com/your-repo/datapro-addin
cd datapro-addin

# 2. Install dependencies
npm install

# 3. Start dev server (HTTPS on port 3000)
npm start

# 4. Sideload the add-in in Excel
```

**Sideloading on Windows**
1. Open Excel → File → Options → Trust Center → Trust Center Settings
2. Trusted Add-in Catalogs → Add `https://localhost:3000`
3. Insert → My Add-ins → Shared Folder → DataPro Add-in

**Sideloading on Mac**
```bash
cp manifest.xml ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/
```
Then restart Excel → Insert → My Add-ins → DataPro Add-in

**Sideloading on Excel Web**
1. Insert → Add-ins → Upload My Add-in
2. Upload `manifest.xml`

---

## 🤖 AI Features Setup

1. Get a free Claude API key from [console.anthropic.com](https://console.anthropic.com)
2. Open the **AI tab** in the add-in
3. Paste your key in the API Key field and click Save
4. The key is stored in `localStorage` — never sent to any server

---

## 📁 Project Structure

```
datapro-addin/
├── manifest.xml              # Office Add-in manifest (register with Excel)
├── package.json              # npm dependencies
├── webpack.config.js         # Build configuration
├── demo.html                 # Standalone browser preview (no Excel needed)
│
└── src/
    ├── taskpane/
    │   ├── taskpane.html     # Main add-in UI (6 panels)
    │   └── taskpane.js       # All logic: dashboard, heatmap, charts, cleaner, AI
    └── commands/
        ├── commands.html     # Ribbon command host page
        └── commands.js       # Ribbon button handlers
```

---

## 🔧 Build for Production

```bash
npm run build
# Output: /dist/ — deploy to any static host (GitHub Pages, Netlify, Azure Static Web Apps)
```

Update the `DefaultValue` URLs in `manifest.xml` to your production domain.

---

## 📊 Chart Types Included

- **Bar / Line** — Revenue vs Expenses with toggle
- **Donut** — Regional breakdown  
- **Waterfall** — P&L bridge (positive=green, negative=red)
- **Funnel** — Sales pipeline with stage percentages
- **Bubble** — Product matrix (X=score, Y=growth, size=revenue)
- **Radar** — KPI scorecard multi-series with year comparison
- **Stacked Area** — Segment trend over 12 months
- **Sparklines** — Inline SVG sparklines in KPI cards
- **Heatmap** — Interactive grid with 4 color palettes

---

## 🔌 Extending with APIs / Data Sources

Edit `src/taskpane/taskpane.js` — the `loadFromSheet()` function shows the pattern:

```javascript
await Excel.run(async (ctx) => {
  const sheet = ctx.workbook.worksheets.getActiveWorksheet();
  const range = sheet.getUsedRange();
  range.load("values");
  await ctx.sync();
  // range.values is a 2D array — use it to update charts
});
```

---

## 📜 License

MIT — free for personal and commercial use, no attribution required.

```
Copyright (c) 2025 DataPro Contributors

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software.
```
