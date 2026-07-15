# E-Faktur CSV Converter Implementation Plan

> **For Hermes:** Use subagent-driven-development skill to implement this plan task-by-task.

**Goal:** Membangun website statis yang memungkinkan pengguna menyalin (paste) dari Excel ke dalam tabel dengan dua kolom: (1) Nomor Faktur Pajak, (2) Bulan Faktur Pajak. Aplikasi akan memvalidasi data dan mengekspornya sebagai CSV untuk sistem E-Faktur AlatPajak.id.
**Architecture:** Single-page application murni klien (Static Site), berjalan di browser tanpa backend.
**Tech Stack:** HTML5, Geist Font + Custom CSS (Vercel-style Clean Minimal), Jspreadsheet CE, PapaParse.
**Target Deployment:** GitHub Repo → Cloudflare Pages → `converter.alatpajak.id`.

---

## Design System: Vercel Clean Minimal

- **Font:** Geist Sans (display/body) + Geist Mono (code/technical labels) via Google Fonts
- **Color Palette:** `#ffffff` (bg), `#171717` (primary text), `#4d4d4d` (secondary), `#666666` (muted), `#ebebeb` (borders/shadows), `#fafafa` (surface tint)
- **Shadow-as-border:** `0px 0px 0px 1px rgba(0,0,0,0.08)` menggantikan CSS border tradisional
- **Card Shadow Stack:** Multi-layer shadow untuk depth yang subtle
- **Letter-spacing:** Negative di heading (-0.96px di 24px), normal di body
- **Weights:** 400 (body), 500 (UI/interactive), 600 (headings)
- **Radius:** 6px (buttons), 8px (cards), 9999px (pill badges only)

---

### Task 1: Setup Struktur Proyek & HTML Dasar (Vercel Style)

**Objective:** Membuat `index.html` dengan Geist font, custom CSS Vercel-style, dan struktur UI yang bersih.

**Files:**
- Create: `index.html`
- Create: `js/app.js`
- Create: `css/style.css`

**Step 1: Buat index.html**

```html
<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>E-Faktur CSV Converter | AlatPajak.id</title>
    
    <!-- Geist Font (Vercel) -->
    <link href="https://fonts.googleapis.com/css2?family=Geist:wght@400;500;600&family=Geist+Mono:wght@400;500&display=swap" rel="stylesheet">
    
    <!-- Jspreadsheet CE CSS -->
    <link rel="stylesheet" href="https://bossanova.uk/jspreadsheet/v4/jexcel.css" type="text/css" />
    <link rel="stylesheet" href="https://jsuites.net/v4/jsuites.css" type="text/css" />
    
    <!-- Custom CSS (Vercel-style) -->
    <link rel="stylesheet" href="css/style.css">
</head>
<body>
    <!-- Main Container -->
    <main class="container">
        <!-- Header Section -->
        <header class="header">
            <h1 class="heading">E-Faktur CSV Converter</h1>
            <p class="subtext">Paste data Excel ke tabel. Validasi otomatis Nomor Faktur 17 digit & Bulan Faktur Pajak. Download CSV siap upload Coretax.</p>
        </header>

        <!-- Toolbar -->
        <div class="toolbar">
            <button id="btn-clear" class="btn btn-ghost">Bersihkan Tabel</button>
            <button id="btn-export" class="btn btn-primary">⬇️ Download CSV</button>
        </div>

        <!-- Validation Alert -->
        <div id="validation-alert" class="alert hidden" role="alert"></div>

        <!-- Spreadsheet Card -->
        <section class="card">
            <div id="spreadsheet-container">
                <div id="spreadsheet"></div>
            </div>
            <p class="hint">💡 Copy dari Excel → Paste di sini (Ctrl+V). Baris error akan berwarna merah muda.</p>
        </section>

        <!-- Footer -->
        <footer class="footer">
            <p>Dibuat untuk <a href="https://alatpajak.id" target="_blank" rel="noopener">AlatPajak.id</a> · E-Faktur Automation</p>
        </footer>
    </main>

    <!-- Scripts -->
    <script src="https://bossanova.uk/jspreadsheet/v4/jexcel.js"></script>
    <script src="https://jsuites.net/v4/jsuites.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/PapaParse/5.4.1/papaparse.min.js"></script>
    <script src="js/app.js"></script>
</body>
</html>
```

**Step 2: Buat css/style.css (Vercel Clean Minimal)**

```css
/* ===== CSS Custom Properties (Vercel Design Tokens) ===== */
:root {
    --color-bg: #ffffff;
    --color-text: #171717;
    --color-text-secondary: #4d4d4d;
    --color-text-muted: #666666;
    --color-border: rgba(0, 0, 0, 0.08);
    --color-border-light: #ebebeb;
    --color-surface-tint: #fafafa;
    --color-error-bg: #fef2f2;
    --color-error-text: #b91c1c;
    --color-error-border: #fecaca;
    --color-primary: #171717;
    --color-primary-hover: #000000;
    --color-ghost-bg: #ffffff;
    --color-ghost-hover: #f5f5f5;
    --color-focus: hsla(212, 100%, 48%, 1);
    
    --font-sans: 'Geist', system-ui, -apple-system, 'Segoe UI', Roboto, sans-serif;
    --font-mono: 'Geist Mono', ui-monospace, SFMono-Regular, Menlo, Monaco, 'Liberation Mono', 'Courier New', monospace;
    
    --shadow-border: 0px 0px 0px 1px var(--color-border);
    --shadow-card: 0px 0px 0px 1px var(--color-border), 0px 2px 2px rgba(0,0,0,0.04), 0px 8px 8px -8px rgba(0,0,0,0.04), 0px 0px 0px 1px var(--color-surface-tint);
    --shadow-card-hover: 0px 0px 0px 1px var(--color-border), 0px 4px 4px rgba(0,0,0,0.06), 0px 12px 12px -8px rgba(0,0,0,0.06), 0px 0px 0px 1px var(--color-surface-tint);
    
    --radius-sm: 6px;
    --radius-md: 8px;
    --radius-pill: 9999px;
    
    --space-1: 4px;
    --space-2: 8px;
    --space-3: 12px;
    --space-4: 16px;
    --space-5: 24px;
    --space-6: 32px;
    --space-8: 48px;
    --space-10: 64px;
}

/* ===== Reset & Base ===== */
*, *::before, *::after { box-sizing: border-box; }
html { font-size: 16px; -webkit-font-smoothing: antialiased; -moz-osx-font-smoothing: grayscale; }
body {
    margin: 0;
    font-family: var(--font-sans);
    font-size: 16px;
    font-weight: 400;
    line-height: 1.56;
    color: var(--color-text);
    background: var(--color-bg);
    min-height: 100vh;
}

/* ===== Layout ===== */
.container {
    max-width: 800px;
    margin: 0 auto;
    padding: var(--space-6) var(--space-4) var(--space-10);
}

@media (min-width: 768px) {
    .container { padding: var(--space-10) var(--space-6) var(--space-10); }
}

/* ===== Header ===== */
.header {
    text-align: center;
    margin-bottom: var(--space-8);
}

.heading {
    font-family: var(--font-sans);
    font-size: 32px;
    font-weight: 600;
    line-height: 1.25;
    letter-spacing: -1.28px;
    color: var(--color-text);
    margin: 0 0 var(--space-3);
}

.subtext {
    font-size: 18px;
    font-weight: 400;
    line-height: 1.56;
    color: var(--color-text-secondary);
    margin: 0;
    max-width: 560px;
    margin-left: auto;
    margin-right: auto;
}

/* ===== Toolbar ===== */
.toolbar {
    display: flex;
    justify-content: space-between;
    align-items: center;
    gap: var(--space-3);
    margin-bottom: var(--space-4);
    flex-wrap: wrap;
}

.btn {
    font-family: var(--font-sans);
    font-size: 14px;
    font-weight: 500;
    line-height: 1.43;
    padding: 10px 16px;
    border-radius: var(--radius-sm);
    border: none;
    cursor: pointer;
    transition: background-color 0.15s ease, box-shadow 0.15s ease;
    display: inline-flex;
    align-items: center;
    gap: 6px;
}

.btn:focus-visible {
    outline: 2px solid var(--color-focus);
    outline-offset: 2px;
}

.btn-primary {
    background: var(--color-primary);
    color: #ffffff;
}

.btn-primary:hover {
    background: var(--color-primary-hover);
}

.btn-ghost {
    background: var(--color-ghost-bg);
    color: var(--color-text);
    box-shadow: var(--shadow-border);
}

.btn-ghost:hover {
    background: var(--color-ghost-hover);
}

/* ===== Alert ===== */
.alert {
    background: var(--color-error-bg);
    color: var(--color-error-text);
    border: 1px solid var(--color-error-border);
    border-radius: var(--radius-md);
    padding: var(--space-4) var(--space-4);
    font-size: 14px;
    line-height: 1.5;
}

.alert strong { font-weight: 600; display: block; margin-bottom: var(--space-2); }
.alert ul { margin: 0; padding-left: 20px; }
.alert li { margin-bottom: 4px; }

.hidden { display: none !important; }

/* ===== Card ===== */
.card {
    background: var(--color-bg);
    border-radius: var(--radius-md);
    box-shadow: var(--shadow-card);
    padding: var(--space-5);
    transition: box-shadow 0.2s ease;
}

.card:hover { box-shadow: var(--shadow-card-hover); }

/* ===== Spreadsheet Container ===== */
#spreadsheet-container {
    min-height: 420px;
    overflow-x: auto;
    border-radius: 4px;
    /* Jspreadsheet akan render di dalam sini */
}

/* Override Jspreadsheet styles untuk konsistensi */
#spreadsheet-container .jexcel {
    font-family: var(--font-sans) !important;
    font-size: 14px !important;
}

#spreadsheet-container .jexcel thead th {
    background: var(--color-surface-tint) !important;
    color: var(--color-text) !important;
    font-weight: 500 !important;
    font-size: 13px !important;
    letter-spacing: -0.32px !important;
    border-bottom: 1px solid var(--color-border-light) !important;
    padding: 10px 12px !important;
}

#spreadsheet-container .jexcel tbody td {
    border-right: 1px solid var(--color-border-light) !important;
    border-bottom: 1px solid var(--color-border-light) !important;
    padding: 8px 12px !important;
    vertical-align: middle !important;
}

#spreadsheet-container .jexcel tbody td:last-child {
    border-right: none !important;
}

#spreadsheet-container .jexcel tbody tr:last-child td {
    border-bottom: none !important;
}

/* Highlight error cell (applied via JS) */
#spreadsheet-container .jexcel tbody td.error-highlight {
    background-color: #fef2f2 !important;
    border-color: #fecaca !important;
}

/* Focus state untuk cell editing */
#spreadsheet-container .jexcel tbody td.editor input,
#spreadsheet-container .jexcel tbody td.editor textarea {
    font-family: var(--font-sans) !important;
    font-size: 14px !important;
    border: 1px solid var(--color-focus) !important;
    box-shadow: 0 0 0 2px rgba(147, 197, 253, 0.3) !important;
    outline: none !important;
}

/* ===== Hint ===== */
.hint {
    font-size: 13px;
    font-weight: 400;
    color: var(--color-text-muted);
    margin: var(--space-4) 0 0;
    text-align: center;
}

/* ===== Footer ===== */
.footer {
    margin-top: var(--space-10);
    padding-top: var(--space-6);
    border-top: 1px solid var(--color-border-light);
    text-align: center;
}

.footer p {
    font-size: 13px;
    color: var(--color-text-muted);
    margin: 0;
}

.footer a {
    color: var(--color-text-secondary);
    text-decoration: none;
    font-weight: 500;
    transition: color 0.15s;
}

.footer a:hover { color: var(--color-text); text-decoration: underline; }
```

**Step 3: Commit**
```bash
git init
git add .
git commit -m "feat: setup Vercel-style clean minimal UI with Geist font"
```

---

### Task 2: Inisialisasi Spreadsheet Dua Kolom

**Objective:** Tabel interaktif dengan kolom "Nomor Faktur Pajak" & "Bulan Faktur Pajak".

**Files:** Modify `js/app.js`

**Step 1:**
```javascript
document.addEventListener('DOMContentLoaded', function() {
    const columns = [
        { type: 'text', title: 'Nomor Faktur Pajak', width: 280 },
        { type: 'text', title: 'Bulan Faktur Pajak', width: 220 }
    ];

    const table = jspreadsheet(document.getElementById('spreadsheet'), {
        data: [['', '']], 
        columns: columns,
        minDimensions: [2, 50],
        allowInsertRow: true,
        allowManualInsertRow: true,
        allowInsertColumn: false,
        allowDeleteRow: true,
        wordWrap: true,
    });

    document.getElementById('btn-clear').addEventListener('click', () => {
        if(confirm('Bersihkan tabel?')) table.setData([['', '']]);
    });

    window.efakturTable = table;
});
```

**Step 2: Commit**
```bash
git commit -am "feat: init spreadsheet with two columns"
```

---

### Task 3: Ekspor CSV & Validator Pintar (Vercel-style Alert)

**Objective:** Validasi 17 digit & nama bulan Indonesia. Block download jika error, highlight cell merah muda, alert box di atas tabel.

**Files:** Modify `js/app.js`

**Step 1: Tambahkan setelah inisialisasi table**
```javascript
const INDONESIAN_MONTHS = new Set([
    'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
    'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'
]);

document.getElementById('btn-export').addEventListener('click', () => {
    const table = window.efakturTable;
    let rawData = table.getData();
    let alertBox = document.getElementById('validation-alert');
    let errors = [];
    let validData = [];
    
    // Reset error highlighting di seluruh tabel
    // Jspreadsheet: loop through rows and clear style
    const totalRows = table.getRowCount();
    for (let r = 0; r < totalRows; r++) {
        table.setStyle(`A${r + 1}`, 'background-color', '');
        table.setStyle(`B${r + 1}`, 'background-color', '');
    }
    // Atau pakai resetStyle jika tersedia
    if (typeof table.resetStyle === 'function') table.resetStyle();

    rawData.forEach((row, index) => {
        let noFaktur = String(row[0] || '').trim();
        let bulan = String(row[1] || '').trim();
        let rowNum = index + 1;

        // Skip baris benar-benar kosong (kedua kolom kosong)
        if (noFaktur === '' && bulan === '') return;

        let rowHasError = false;

        // Validasi Nomor Faktur: 17 digit angka
        if (!/^\d{17}$/.test(noFaktur)) {
            errors.push(`Baris ${rowNum}: Nomor Faktur Pajak harus 17 digit angka.`);
            table.setStyle(`A${rowNum}`, 'background-color', '#fef2f2');
            rowHasError = true;
        }
        
        // Validasi Bulan: harus nama bulan Indonesia lengkap
        if (!INDONESIAN_MONTHS.has(bulan)) {
            errors.push(`Baris ${rowNum}: Bulan Faktur Pajak tidak valid (Harus ejaan lengkap, misal: Januari, Februari, dst.).`);
            table.setStyle(`B${rowNum}`, 'background-color', '#fef2f2');
            rowHasError = true;
        }

        if (!rowHasError) {
            validData.push([noFaktur, bulan]);
        }
    });

    // Jika ada error: tampilkan alert, block download
    if (errors.length > 0) {
        alertBox.innerHTML = `<strong>⚠️ Ditemukan ${errors.length} error. Perbaiki sel berwarna merah muda sebelum download:</strong><ul class="list-disc ml-5 mt-2 text-sm">${errors.slice(0, 5).map(e => `<li>${e}</li>`).join('')}${errors.length > 5 ? `<li>...dan ${errors.length - 5} lainnya</li>` : ''}</ul>`;
        alertBox.classList.remove('hidden');
        return;
    }

    // Jika tidak ada data valid
    if (validData.length === 0) {
        alertBox.innerHTML = 'Tabel masih kosong. Tempel data dari Excel terlebih dahulu.';
        alertBox.classList.remove('hidden');
        return;
    }

    // Sukses: sembunyikan alert
    alertBox.classList.add('hidden');

    // Header CSV sesuai permintaan
    const csvData = [['Nomor Faktur Pajak', 'Bulan Faktur Pajak']].concat(validData);
    const csvString = Papa.unparse(csvData, { 
        quotes: false, 
        delimiter: ",", 
        newline: "\r\n" 
    });

    // Trigger download
    const blob = new Blob([csvString], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.setAttribute("href", url);
    link.setAttribute("download", `Data_Upload_EFaktur_${new Date().getTime()}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
});
```

**Step 2: Commit**
```bash
git commit -am "feat: robust validation, error highlighting, and CSV export"
```

---

### Task 4: Deploy ke Cloudflare Pages

**Objective:** Repo GitHub + Cloudflare Pages + Custom domain `converter.alatpajak.id`.

**Langkah (Agen menjalankan via terminal):**
```bash
# 1. Buat repo GitHub & push
gh config set git_protocol https
gh repo create mashiker/alatpajak-converter --public --source="E:/kantor_archisight/alatpajak-converter" --push

# 2. Deploy Cloudflare Pages
export CLOUDFLARE_API_TOKEN=$(tr -d '\r\n ' < 'E:/.secrets/cloudflare_apitoken.md')
export CLOUDFLARE_ACCOUNT_ID=$(tr -d '\r\n ' < 'E:/.secrets/cloudflare_accountid.md')
npx wrangler pages deploy . --project-name alatpajak-converter

# 3. Setup DNS CNAME converter.alatpajak.id → alatpajak-converter.pages.dev
# (via Cloudflare API curl - lihat Panduan-Akses-API-Cloudflare.md)
```

**Step: Verifikasi**
```bash
curl -I https://converter.alatpajak.id
# Target: HTTP 200, SSL aktif, halaman terbuka
```

---