/**
 * تحليل صرف الأدوية – v6 (المرحلة المعقدة)
 * سنة كاملة · شهور متعددة · أيام متعددة · بحث أصناف
 */

const DB_NAME = 'DispensingAnalyticsDB';
const DB_VERSION = 7;
const STORE_NAME = 'dispensing';
const MONTH_NAMES = ['', 'يناير', 'فبراير', 'مارس', 'أبريل', 'مايو', 'يونيو', 'يوليو', 'أغسطس', 'سبتمبر', 'أكتوبر', 'نوفمبر', 'ديسمبر'];

let currentType = 'free';
let db = null;
let charts = [];
let gMode = 'month';
let iMode = 'month';
let allItemsList = []; // {code, name}

function formatPct(part, whole) {
  if (!whole || !part) return '0%';
  const p = (part / whole) * 100;
  if (p === 0) return '0%';
  if (p >= 10) return p.toFixed(1) + '%';
  if (p >= 1) return p.toFixed(2) + '%';
  if (p >= 0.1) return p.toFixed(2) + '%';
  if (p >= 0.01) return p.toFixed(3) + '%';
  return p.toFixed(4) + '%';
}

const PALETTE = [
  '#0d9488', '#14b8a6', '#2dd4bf', '#5eead4', '#0991b3',
  '#0284c7', '#0369a1', '#7c3aed', '#a855f7', '#c026d3',
  '#db2777', '#e11d48', '#ea580c', '#ca8a04', '#65a30d'
];

// ─── DB ──────────────────────────────────────────────────────
function openDB() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, DB_VERSION);
    req.onupgradeneeded = (e) => {
      const database = e.target.result;
      if (!database.objectStoreNames.contains(STORE_NAME)) {
        database.createObjectStore(STORE_NAME, { keyPath: 'id' });
      }
    };
    req.onsuccess = (e) => resolve(e.target.result);
    req.onerror = (e) => reject(e.target.error);
  });
}

function saveRecord(record) {
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, 'readwrite');
    tx.objectStore(STORE_NAME).put(record);
    tx.oncomplete = () => resolve();
    tx.onerror = (e) => reject(e.target.error);
  });
}

function getRecords(type, year, month, day = null) {
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, 'readonly');
    const req = tx.objectStore(STORE_NAME).getAll();
    req.onsuccess = () => {
      let results = req.result.filter(r => r.type === type && r.year === year && r.month === month);
      if (day !== null && day !== '') results = results.filter(r => r.day === Number(day));
      resolve(results);
    };
    req.onerror = (e) => reject(e.target.error);
  });
}

function getAllForYear(type, year) {
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, 'readonly');
    const req = tx.objectStore(STORE_NAME).getAll();
    req.onsuccess = () => resolve(req.result.filter(r => r.type === type && r.year === year));
    req.onerror = (e) => reject(e.target.error);
  });
}

function getAllRecordsForType(type) {
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, 'readonly');
    const req = tx.objectStore(STORE_NAME).getAll();
    req.onsuccess = () => resolve(req.result.filter(r => r.type === type));
    req.onerror = (e) => reject(e.target.error);
  });
}

function deleteRecords(type, year, month, day = null) {
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, 'readwrite');
    const store = tx.objectStore(STORE_NAME);
    const req = store.getAll();
    req.onsuccess = () => {
      const toDelete = req.result.filter(r => {
        if (r.type !== type || r.year !== year || r.month !== month) return false;
        if (day !== null && day !== '') return r.day === Number(day);
        return true;
      });
      toDelete.forEach(r => store.delete(r.id));
      tx.oncomplete = () => resolve(toDelete.length);
    };
    tx.onerror = (e) => reject(e.target.error);
  });
}

function getUploadedDays(type, year, month) {
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, 'readonly');
    const req = tx.objectStore(STORE_NAME).getAll();
    req.onsuccess = () => {
      const days = new Set();
      let hasMonthly = false;
      req.result.forEach(r => {
        if (r.type === type && r.year === year && r.month === month) {
          if (r.day === 0 || r.isMonthly) hasMonthly = true;
          else days.add(r.day);
        }
      });
      resolve({ days: Array.from(days).sort((a, b) => a - b), hasMonthly });
    };
    req.onerror = (e) => reject(e.target.error);
  });
}

// ─── Excel parse ─────────────────────────────────────────────
function parseDailySheet(workbook) {
  const sheetName = workbook.SheetNames.find(n => n.includes('منصرف'))
    || workbook.SheetNames.find(n => n !== 'Sheet3' && n !== 'TOTAL')
    || workbook.SheetNames[1];
  if (!sheetName) throw new Error('لم يتم العثور على ورقة «منصرف روشتات»');
  const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: null });
  const data = [];
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row || row.length < 3) continue;
    const code = row[0], name = row[1], qty = row[2];
    if (!code || String(name) === '#N/A' || name == null || qty == null || qty === '') continue;
    const q = Number(qty);
    if (!q || q <= 0) continue;
    data.push({ code: String(code).trim().toLowerCase(), name: String(name).trim(), qty: q });
  }
  return data;
}

function parseMonthlySheet(workbook) {
  let sheetName = workbook.SheetNames.find(n => n.trim() === 'الأصناف المنصرفة')
    || workbook.SheetNames.find(n => n.includes('الأصناف المنصرفة') && !n.includes('منقول'))
    || workbook.SheetNames.find(n => n.includes('منصرفة') && !n.includes('منقول'))
    || workbook.SheetNames.find(n => n.includes('جميع الأصناف') && !n.includes('منقول'))
    || workbook.SheetNames[0];
  const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: null });
  if (rows.length < 2) throw new Error('الورقة الشهرية فارغة');
  const header = rows[0].map(h => h != null ? String(h).trim() : '');
  const codeIdx = header.findIndex(h => h === 'كود' || h.toLowerCase() === 'code');
  const nameIdx = header.findIndex(h => h.includes('اسم'));
  const totalIdx = header.findIndex(h => h.toLowerCase() === 'total' || h === 'الإجمالي');
  const dayCols = {};
  for (let d = 1; d <= 31; d++) {
    const idx = header.findIndex(h => h === String(d));
    if (idx !== -1) dayCols[d] = idx;
  }
  const items = [];
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row || !row[codeIdx]) continue;
    const code = String(row[codeIdx]).trim().toLowerCase();
    const name = nameIdx !== -1 ? String(row[nameIdx] || '').trim() : code;
    const daily = {};
    let sum = 0;
    Object.entries(dayCols).forEach(([day, idx]) => {
      const v = Number(row[idx]) || 0;
      if (v > 0) { daily[day] = v; sum += v; }
    });
    const total = totalIdx !== -1 ? (Number(row[totalIdx]) || 0) : sum;
    const finalTotal = total > 0 ? total : sum;
    if (finalTotal <= 0) continue;
    items.push({ code, name, daily, total: finalTotal });
  }
  return items;
}


function parseCatalog(workbook) {
  // Full item list (with or without dispensing)
  let sheetName = workbook.SheetNames.find(n => n.trim() === 'Sheet3')
    || workbook.SheetNames.find(n => n.includes('جميع الأصناف') && !n.includes('منقول'))
    || workbook.SheetNames.find(n => /sheet\s*3/i.test(n))
    || workbook.SheetNames.find(n => n.toLowerCase().includes('code') || n.includes('كود'));
  // fallback: first sheet that has كود + اسم columns
  const trySheets = sheetName ? [sheetName, ...workbook.SheetNames] : workbook.SheetNames;
  const seen = new Set();
  const items = [];
  for (const sn of trySheets) {
    if (!sn || seen.has(sn)) continue;
    seen.add(sn);
    const sheet = workbook.Sheets[sn];
    if (!sheet) continue;
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });
    if (!rows || rows.length < 2) continue;
    const header = rows[0].map(h => h != null ? String(h).trim() : '');
    let codeIdx = header.findIndex(h => h === 'كود' || h.toLowerCase() === 'code' || h === 'الكود');
    let nameIdx = header.findIndex(h => h.includes('اسم'));
    // daily sheet3 often has no header - col0=code col1=name
    const hasHeader = codeIdx !== -1 || nameIdx !== -1;
    if (!hasHeader) {
      // heuristic: many rows with string in col0 and col1
      let ok = 0;
      for (let i = 0; i < Math.min(rows.length, 20); i++) {
        if (rows[i] && rows[i][0] != null && rows[i][1] != null) ok++;
      }
      if (ok < 5) continue;
      codeIdx = 0; nameIdx = 1;
      const start = (String(rows[0][0]).includes('كود') || String(rows[0][0]).toLowerCase() === 'code') ? 1 : 0;
      for (let i = start; i < rows.length; i++) {
        const row = rows[i];
        if (!row || row[0] == null) continue;
        const code = String(row[0]).trim().toLowerCase();
        if (!code || code === '#n/a' || code === 'كود') continue;
        const name = row[1] != null ? String(row[1]).trim() : code;
        if (name === '#N/A') continue;
        items.push({ code, name });
      }
      if (items.length > 10) break;
      items.length = 0;
      continue;
    }
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (!row || row[codeIdx] == null) continue;
      const code = String(row[codeIdx]).trim().toLowerCase();
      if (!code || code === '#n/a') continue;
      const name = nameIdx !== -1 && row[nameIdx] != null ? String(row[nameIdx]).trim() : code;
      if (name === '#N/A') continue;
      items.push({ code, name });
    }
    if (items.length > 5) break;
    items.length = 0;
  }
  // dedupe
  const map = new Map();
  items.forEach(it => {
    if (!map.has(it.code)) map.set(it.code, it.name);
    else if (it.name.length > map.get(it.code).length) map.set(it.code, it.name);
  });
  return Array.from(map.entries()).map(([code, name]) => ({ code, name }));
}

async function mergeCatalog(type, newItems) {
  if (!newItems || !newItems.length) return;
  const id = `catalog_${type}`;
  const existing = await new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, 'readonly');
    const req = tx.objectStore(STORE_NAME).get(id);
    req.onsuccess = () => resolve(req.result);
    req.onerror = (e) => reject(e.target.error);
  });
  const map = new Map();
  if (existing && existing.items) {
    existing.items.forEach(it => map.set(it.code, it.name));
  }
  newItems.forEach(it => {
    if (!map.has(it.code)) map.set(it.code, it.name);
    else if (it.name && it.name.length > (map.get(it.code) || '').length) map.set(it.code, it.name);
  });
  await saveRecord({
    id,
    type,
    isCatalog: true,
    items: Array.from(map.entries()).map(([code, name]) => ({ code, name, qty: 0 })),
    uploadedAt: new Date().toISOString()
  });
}

async function getCatalog(type) {
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, 'readonly');
    const req = tx.objectStore(STORE_NAME).get(`catalog_${type}`);
    req.onsuccess = () => resolve(req.result ? (req.result.items || []) : []);
    req.onerror = (e) => reject(e.target.error);
  });
}


function aggregateDaily(rows) {
  const map = new Map();
  rows.forEach(r => {
    if (!map.has(r.code)) map.set(r.code, { code: r.code, name: r.name, qty: 0 });
    const item = map.get(r.code);
    item.qty += r.qty;
    if (r.name && r.name.length > (item.name || '').length) item.name = r.name;
  });
  return Array.from(map.values()).filter(i => i.qty > 0).sort((a, b) => b.qty - a.qty);
}

// ─── Helpers ─────────────────────────────────────────────────
function showMsg(text, type = 'info') {
  const el = document.getElementById('message');
  el.className = `msg show msg-${type}`;
  el.textContent = text;
  clearTimeout(el._t);
  el._t = setTimeout(() => el.classList.remove('show'), 5000);
}

function showUploadSuccess(text) {
  const el = document.getElementById('uploadSuccess');
  el.textContent = text;
  el.classList.add('show');
  clearTimeout(el._t);
  el._t = setTimeout(() => el.classList.remove('show'), 5000);
}

function fillDays(selectId, year, month) {
  const sel = document.getElementById(selectId);
  if (!sel) return;
  const current = sel.value;
  const isUpload = selectId === 'uploadDay';
  sel.innerHTML = isUpload ? '<option value="">-- ملف شهري --</option>' : '<option value="">-- الشهر كله --</option>';
  if (month) {
    const n = new Date(year, month, 0).getDate();
    for (let d = 1; d <= n; d++) {
      const o = document.createElement('option');
      o.value = d; o.textContent = String(d);
      sel.appendChild(o);
    }
    if (current) sel.value = current;
  }
  refreshCustomSelect(selectId);
}

function customConfirm(message) {
  return new Promise((resolve) => {
    const overlay = document.createElement('div');
    overlay.className = 'modal-overlay';
    overlay.innerHTML = `<div class="modal-box"><p class="modal-msg">${message}</p>
      <div class="modal-actions">
        <button class="btn btn-ghost" data-act="cancel">إلغاء</button>
        <button class="btn btn-danger-solid" data-act="ok">تأكيد الحذف</button>
      </div></div>`;
    document.body.appendChild(overlay);
    requestAnimationFrame(() => overlay.classList.add('show'));
    overlay.addEventListener('click', (e) => {
      const act = e.target.dataset?.act;
      if (act === 'ok') { close(); resolve(true); }
      else if (act === 'cancel' || e.target === overlay) { close(); resolve(false); }
    });
    function close() { overlay.classList.remove('show'); setTimeout(() => overlay.remove(), 200); }
  });
}

function destroyCharts() {
  charts.forEach(c => c.destroy());
  charts = [];
}

function getChecked(containerId) {
  return Array.from(document.querySelectorAll(`#${containerId} input:checked`)).map(i => Number(i.value));
}

function buildCheckGrid(containerId, count, labels) {
  const el = document.getElementById(containerId);
  el.innerHTML = '';
  for (let i = 1; i <= count; i++) {
    const lab = document.createElement('label');
    lab.className = 'check-item';
    lab.innerHTML = `<input type="checkbox" value="${i}" /><span>${labels ? labels[i] : i}</span>`;
    lab.querySelector('input').addEventListener('change', () => {
      lab.classList.toggle('checked', lab.querySelector('input').checked);
    });
    el.appendChild(lab);
  }
  const actions = document.createElement('div');
  actions.className = 'check-actions';
  actions.innerHTML = `<button type="button" data-a="all">تحديد الكل</button><button type="button" data-a="none">إلغاء الكل</button>`;
  actions.addEventListener('click', (e) => {
    const a = e.target.dataset?.a;
    if (!a) return;
    el.querySelectorAll('input').forEach(inp => {
      inp.checked = a === 'all';
      inp.closest('.check-item').classList.toggle('checked', a === 'all');
    });
  });
  el.appendChild(actions);
}

// ─── Charts ──────────────────────────────────────────────────
function renderBarChart(canvasId, labels, values, opts = {}) {
  const canvas = document.getElementById(canvasId);
  if (!canvas) return;
  const horizontal = opts.horizontal !== false;
  const showAll = opts.showAllLabels !== false;
  const n = labels.length;
  const isNarrow = window.innerWidth < 480;
  const insideLabels = opts.insideLabels !== false && horizontal;

  const parent = canvas.parentElement;
  if (parent) {
    parent.style.minWidth = '';
    parent.style.width = '100%';
    parent.setAttribute('dir', 'ltr');
    if (horizontal) {
      const rowH = isNarrow ? 28 : 32;
      parent.style.height = Math.max(260, n * rowH) + 'px';
    } else {
      parent.style.height = (opts.height || 300) + 'px';
    }
  }
  canvas.setAttribute('dir', 'ltr');

  const fontSize = isNarrow ? 10 : 11;
  const maxBar = isNarrow ? 26 : 36;
  const displayLabels = labels.map(l => String(l));
  const colors = values.map((_, i) => PALETTE[i % PALETTE.length]);
  const grand = (opts.grandTotal && opts.grandTotal > 0)
    ? opts.grandTotal
    : (values.reduce((a, b) => a + b, 0) || 1);

  const insidePlugin = {
    id: 'insideBarLabels_' + canvasId,
    afterDatasetsDraw(chart) {
      if (!insideLabels) return;
      const { ctx } = chart;
      const meta = chart.getDatasetMeta(0);
      const area = chart.chartArea;
      ctx.save();
      ctx.font = `600 ${isNarrow ? 10 : 11}px "IBM Plex Sans Arabic", sans-serif`;
      ctx.textBaseline = 'middle';
      meta.data.forEach((bar, i) => {
        const text = String(labels[i] || '');
        const { x, y, base } = bar.getProps(['x', 'y', 'base'], true);
        // بعد dir=ltr: القاعدة يسار والعمود يتمدد يمين
        const start = Math.min(x, base);
        const end = Math.max(x, base);
        const barW = end - start;
        const tw = ctx.measureText(text).width;
        const pad = 8;
        const fitsInside = tw + pad * 2 <= barW - 2;
        const spaceAfter = Math.max(0, area.right - end - 4);
        if (fitsInside) {
          ctx.fillStyle = '#ffffff';
          ctx.textAlign = 'left';
          ctx.fillText(text, start + pad, y);
        } else if (spaceAfter >= 24) {
          ctx.fillStyle = '#134e4a';
          ctx.textAlign = 'left';
          let draw = text;
          while (draw.length > 4 && ctx.measureText(draw).width > spaceAfter) draw = draw.slice(0, -1);
          if (draw !== text) draw = draw.slice(0, Math.max(3, draw.length - 1)) + '…';
          ctx.fillText(draw, end + 5, y);
        } else if (barW > 20) {
          ctx.fillStyle = '#ffffff';
          ctx.textAlign = 'left';
          let draw = text;
          const maxW = Math.max(barW - pad * 2, 16);
          while (draw.length > 3 && ctx.measureText(draw).width > maxW) draw = draw.slice(0, -1);
          if (draw !== text) draw = draw.slice(0, Math.max(3, draw.length - 1)) + '…';
          ctx.fillText(draw, start + pad, y);
        }
      });
      ctx.restore();
    }
  };

  const chart = new Chart(canvas, {
    type: 'bar',
    plugins: [insidePlugin],
    data: {
      labels: insideLabels ? labels.map(() => '') : displayLabels,
      datasets: [{
        label: 'الكمية', data: values,
        backgroundColor: colors.map(c => c + 'dd'), borderColor: colors,
        borderWidth: 0, borderRadius: 7, borderSkipped: false, maxBarThickness: maxBar
      }]
    },
    options: {
      indexAxis: horizontal ? 'y' : 'x',
      responsive: true,
      maintainAspectRatio: false,
      locale: 'en-US',
      layout: { padding: { left: 2, right: insideLabels ? 8 : 4, top: 2, bottom: 2 } },
      plugins: {
        legend: { display: false },
        tooltip: {
          backgroundColor: 'rgba(15,118,110,.95)',
          titleFont: { family: 'IBM Plex Sans Arabic', size: 12 },
          bodyFont: { family: 'IBM Plex Sans Arabic', size: 12 },
          padding: 10, cornerRadius: 10,
          callbacks: {
            title: (items) => {
              const i = items[0]?.dataIndex ?? 0;
              return labels[i] || '';
            },
            label: (ctx) => {
              return ` الكمية: ${Number(ctx.raw).toLocaleString('en')}  ·  ${formatPct(ctx.raw, grand)} من الإجمالي`;
            }
          }
        }
      },
      scales: {
        x: {
          beginAtZero: true,
          ticks: { font: { family: 'IBM Plex Sans Arabic', size: fontSize }, color: '#64748b', autoSkip: !showAll, maxRotation: 0 },
          grid: { color: 'rgba(0,0,0,.05)' }
        },
        y: {
          display: !insideLabels,
          ticks: {
            font: { family: 'IBM Plex Sans Arabic', size: fontSize },
            color: '#334155',
            autoSkip: false
          },
          grid: { color: 'rgba(0,0,0,.04)' },
          afterFit: (!insideLabels && horizontal) ? function (scale) {
            const maxLen = Math.max(...displayLabels.map(l => String(l).length), 8);
            const charW = isNarrow ? 6.2 : 7.2;
            scale.width = Math.min(Math.max(maxLen * charW, 130), isNarrow ? 200 : 280);
          } : undefined
        }
      }
    }
  });
  charts.push(chart);
  return chart;
}

function renderDoughnut(canvasId, labels, values, opts = {}) {
  const canvas = document.getElementById(canvasId);
  if (!canvas) return;
  const colors = values.map((_, i) => PALETTE[i % PALETTE.length]);
  const total = values.reduce((a, b) => a + b, 0) || 1;
  const chart = new Chart(canvas, {
    type: 'doughnut',
    data: {
      labels,
      datasets: [{ data: values, backgroundColor: colors.map(c => c + 'dd'), borderColor: '#fff', borderWidth: 2, hoverOffset: 6 }]
    },
    options: {
      responsive: true, maintainAspectRatio: false, cutout: '55%',
      plugins: {
        legend: { display: false },
        tooltip: {
          backgroundColor: 'rgba(15,118,110,.95)',
          titleFont: { family: 'IBM Plex Sans Arabic', size: 12 },
          bodyFont: { family: 'IBM Plex Sans Arabic', size: 12 },
          padding: 10, cornerRadius: 10,
          callbacks: {
            label: (ctx) => {
              const name = labels[ctx.dataIndex] || '';
              return ` ${name}: ${Number(ctx.raw).toLocaleString('en')} (${formatPct(ctx.raw, total)})`;
            }
          }
        }
      }
    }
  });
  charts.push(chart);
  return chart;
}

function buildDonutLegendHtml(labels, values) {
  const total = values.reduce((a, b) => a + b, 0) || 1;
  return `<div class="donut-legend">${labels.map((l, i) => {
    const color = PALETTE[i % PALETTE.length];
    const pct = formatPct(values[i], total);
    return `<div class="donut-legend-item">
      <span class="donut-swatch" style="background:${color}"></span>
      <span class="donut-legend-text"><strong>${l}</strong>
        <span class="donut-legend-meta">${Number(values[i]).toLocaleString('en')} وحدة · ${pct}</span>
      </span>
    </div>`;
  }).join('')}</div>`;
}

function daysInMonth(year, month) {
  return new Date(year, month, 0).getDate();
}

function fillMonthDays(year, month, dailyMap) {
  const last = daysInMonth(year, month);
  const labels = [];
  const values = [];
  for (let d = 1; d <= last; d++) {
    labels.push(String(d));
    values.push(Number(dailyMap[d]) || 0);
  }
  return { labels, values, last };
}

function renderLineChart(canvasId, labels, values) {
  const canvas = document.getElementById(canvasId);
  if (!canvas) return;
  const p = canvas.parentElement;
  if (p) {
    // كل نقطة يوم لها مساحة ثابتة → سكرول أفقي عند الحاجة
    const minW = Math.max(labels.length * 36, 360);
    p.style.minWidth = minW + 'px';
    p.style.height = p.style.height || '280px';
  }
  const chart = new Chart(canvas, {
    type: 'line',
    data: {
      labels,
      datasets: [{
        label: 'الكمية', data: values,
        borderColor: '#0d9488', backgroundColor: 'rgba(13,148,136,.12)',
        borderWidth: 2.5, fill: true, tension: 0.3,
        pointRadius: labels.length > 60 ? 2 : 4,
        pointBackgroundColor: '#0f766e', pointBorderColor: '#fff', pointBorderWidth: 2
      }]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          backgroundColor: 'rgba(15,118,110,.95)',
          titleFont: { family: 'IBM Plex Sans Arabic', size: 13 },
          bodyFont: { family: 'IBM Plex Sans Arabic', size: 13 },
          padding: 12, cornerRadius: 10,
          callbacks: { label: (ctx) => ' الكمية: ' + Number(ctx.raw).toLocaleString('en') }
        }
      },
      scales: {
        x: { ticks: { font: { family: 'IBM Plex Sans Arabic', size: 10 }, color: '#64748b', autoSkip: false, maxRotation: 0 }, grid: { color: 'rgba(0,0,0,.04)' } },
        y: { beginAtZero: true, ticks: { font: { family: 'IBM Plex Sans Arabic', size: 11 }, color: '#64748b' }, grid: { color: 'rgba(0,0,0,.05)' } }
      }
    }
  });
  charts.push(chart);
}

// ─── Export ──────────────────────────────────────────────────

function tableSearchBox(id, placeholder) {
  return `<div class="table-search-wrap">
    <input type="search" id="${id}" class="table-search" placeholder="${placeholder}" autocomplete="off" />
  </div>`;
}

function bindTableSearch(inputId, tbodyId, countId, total) {
  const input = document.getElementById(inputId);
  if (!input) return;
  input.addEventListener('input', () => {
    const q = input.value.trim().toLowerCase();
    let vis = 0;
    document.querySelectorAll('#' + tbodyId + ' tr').forEach(tr => {
      if (tr.classList.contains('total-row')) {
        tr.style.display = q ? 'none' : '';
        return;
      }
      const hay = (tr.dataset.q || tr.textContent || '').toLowerCase();
      const ok = !q || hay.includes(q);
      tr.style.display = ok ? '' : 'none';
      if (ok) vis++;
    });
    const c = document.getElementById(countId);
    if (c) c.textContent = q ? `(${vis} من ${total})` : `(${total})`;
  });
}

function exportExcel(rows, sheetName, filename) {
  try {
    if (!rows || !rows.length) { showMsg('لا توجد بيانات للتصدير', 'info'); return; }
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, (sheetName || 'بيانات').slice(0, 31));
    const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = filename || 'export.xlsx';
    document.body.appendChild(a); a.click();
    setTimeout(() => { document.body.removeChild(a); URL.revokeObjectURL(url); }, 400);
    showMsg('تم تصدير Excel', 'success');
  } catch (e) { console.error(e); showMsg('فشل التصدير: ' + e.message, 'error'); }
}

function exportChartPNG(canvasId, filename) {
  const canvas = document.getElementById(canvasId);
  if (!canvas) { showMsg('لا يوجد رسم', 'info'); return; }
  try {
    const tmp = document.createElement('canvas');
    tmp.width = canvas.width; tmp.height = canvas.height;
    const ctx = tmp.getContext('2d');
    ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, tmp.width, tmp.height);
    ctx.drawImage(canvas, 0, 0);
    const a = document.createElement('a');
    a.href = tmp.toDataURL('image/png'); a.download = filename || 'chart.png';
    document.body.appendChild(a); a.click();
    setTimeout(() => document.body.removeChild(a), 300);
    showMsg('تم تصدير الرسم', 'success');
  } catch (e) { console.error(e); showMsg('فشل تصدير الرسم', 'error'); }
}

window.doExportChart = (id, file) => exportChartPNG(id, file);


// ─── Aggregate month total qty ───────────────────────────────
async function monthTotalQty(type, year, month) {
  let recs = await getRecords(type, year, month, 0);
  if (!recs.length) recs = await getRecords(type, year, month);
  let total = 0;
  const seen = new Set();
  recs.forEach(rec => {
    // prefer monthly aggregate (day=0)
    if (rec.day === 0 || rec.isMonthly) {
      (rec.items || []).forEach(it => { if (it.qty > 0) total += it.qty; });
      seen.add('m');
    }
  });
  if (!seen.has('m')) {
    // sum unique day records
    const byCode = new Map();
    recs.forEach(rec => {
      (rec.items || []).forEach(it => {
        if (!it.qty) return;
        byCode.set(it.code, (byCode.get(it.code) || 0) + it.qty);
      });
    });
    byCode.forEach(v => total += v);
  }
  return total;
}

async function getMonthItems(type, year, month) {
  let recs = await getRecords(type, year, month, 0);
  if (!recs.length) recs = await getRecords(type, year, month);
  const map = new Map();
  recs.forEach(rec => {
    (rec.items || []).forEach(it => {
      if (!it.qty || it.qty <= 0) return;
      if (!map.has(it.code)) map.set(it.code, { code: it.code, name: it.name, qty: 0 });
      const e = map.get(it.code);
      e.qty += it.qty;
      if (it.name && it.name.length > (e.name || '').length) e.name = it.name;
    });
  });
  return Array.from(map.values()).filter(i => i.qty > 0).sort((a, b) => b.qty - a.qty);
}

// ─── Upload status ───────────────────────────────────────────
async function refreshUploadStatus() {
  const year = Number(document.getElementById('uploadYear').value);
  const month = Number(document.getElementById('uploadMonth').value);
  const container = document.getElementById('uploadStatus');
  container.innerHTML = '';
  if (!month) return;
  const { days, hasMonthly } = await getUploadedDays(currentType, year, month);
  if (hasMonthly) {
    const badge = document.createElement('span');
    badge.className = 'month-badge';
    badge.innerHTML = '✓ الشهر كامل مرفوع';
    container.appendChild(badge);
    const delBtn = document.createElement('button');
    delBtn.className = 'btn btn-danger';
    delBtn.innerHTML = '🗑 حذف بيانات الشهر';
    delBtn.onclick = async () => {
      if (!(await customConfirm('حذف كل بيانات هذا الشهر؟'))) return;
      await deleteRecords(currentType, year, month);
      showMsg('تم الحذف', 'success');
      refreshUploadStatus(); populateItems();
    };
    container.appendChild(delBtn);
  } else {
    const n = new Date(year, month, 0).getDate();
    for (let d = 1; d <= n; d++) {
      const chip = document.createElement('span');
      if (days.includes(d)) {
        chip.className = 'day-chip uploaded';
        chip.innerHTML = `${d} <button class="trash-btn">🗑</button>`;
        chip.querySelector('.trash-btn').onclick = async (e) => {
          e.stopPropagation();
          if (!(await customConfirm(`حذف يوم ${d}؟`))) return;
          await deleteRecords(currentType, year, month, d);
          showMsg(`تم حذف يوم ${d}`, 'success');
          refreshUploadStatus(); populateItems();
        };
      } else {
        chip.className = 'day-chip';
        chip.textContent = d;
      }
      container.appendChild(chip);
    }
  }
}

// ─── Item search ─────────────────────────────────────────────
async function populateItems() {
  const map = new Map();
  // 1) Full catalog (includes items with zero dispensing)
  try {
    const catalog = await getCatalog(currentType);
    catalog.forEach(it => {
      if (!it || !it.code) return;
      const code = String(it.code).trim().toLowerCase();
      const name = (it.name && String(it.name).trim()) || code;
      if (name === '#N/A') return;
      map.set(code, name);
    });
  } catch (_) {}
  // 2) Merge any items from dispensing records
  const records = await getAllRecordsForType(currentType);
  records.forEach(rec => {
    if (rec.isCatalog) return;
    (rec.items || []).forEach(it => {
      if (!it || !it.code) return;
      const code = String(it.code).trim().toLowerCase();
      const name = (it.name && String(it.name).trim()) || code;
      if (name === '#N/A') return;
      if (!map.has(code)) map.set(code, name);
      else if (name.length > (map.get(code) || '').length) map.set(code, name);
    });
  });
  allItemsList = Array.from(map.entries())
    .map(([code, name]) => ({ code, name }))
    .sort((a, b) => a.name.localeCompare(b.name, 'ar'));
  renderItemDropdown(document.getElementById('itemSearch')?.value || '');
}

function renderItemDropdown(query) {
  const dd = document.getElementById('itemDropdown');
  const q = (query || '').trim().toLowerCase();
  const filtered = q
    ? allItemsList.filter(i => i.name.toLowerCase().includes(q) || i.code.includes(q))
    : allItemsList;
  if (!filtered.length) {
    dd.innerHTML = '<div class="search-empty">لا توجد نتائج</div>';
    return;
  }
  dd.innerHTML = filtered.map(i =>
    `<div class="search-option" data-code="${i.code}" data-name="${i.name}"><span class="so-name">${i.name}</span><span class="so-code">${i.code}</span></div>`
  ).join('');
}

function setupItemSearch() {
  const input = document.getElementById('itemSearch');
  const hidden = document.getElementById('itemCode');
  const dd = document.getElementById('itemDropdown');
  const clearBtn = document.getElementById('itemClear');

  function updateClear() {
    if (!clearBtn) return;
    clearBtn.classList.toggle('show', !!(input.value || hidden.value));
  }

  function clearItem() {
    input.value = '';
    hidden.value = '';
    updateClear();
    renderItemDropdown('');
    dd.classList.add('show');
    input.focus();
  }

  if (clearBtn) clearBtn.addEventListener('click', (e) => {
    e.preventDefault();
    e.stopPropagation();
    clearItem();
  });

  input.addEventListener('focus', () => {
    // if selected text is full name, show all list when focusing with empty query intent
    renderItemDropdown(hidden.value ? '' : input.value);
    dd.classList.add('show');
    updateClear();
  });
  input.addEventListener('input', () => {
    hidden.value = '';
    renderItemDropdown(input.value);
    dd.classList.add('show');
    updateClear();
  });
  dd.addEventListener('click', (e) => {
    const opt = e.target.closest('.search-option');
    if (!opt) return;
    hidden.value = opt.dataset.code;
    input.value = `${opt.dataset.name} (${opt.dataset.code})`;
    dd.classList.remove('show');
    updateClear();
  });
  document.addEventListener('click', (e) => {
    if (!e.target.closest('.search-wrap')) dd.classList.remove('show');
  });
  updateClear();
}

// ─── Mode switching ──────────────────────────────────────────
function setGMode(mode) {
  gMode = mode;
  document.querySelectorAll('#gModePills .mode-pill').forEach(p => {
    p.classList.toggle('active', p.dataset.mode === mode);
  });
  document.getElementById('gMonthWrap').classList.toggle('hidden', mode !== 'month');
  document.getElementById('gDayWrap').classList.toggle('hidden', mode !== 'month');
  document.getElementById('gMonthsCheck').classList.toggle('show', mode === 'months');
}

function setIMode(mode) {
  iMode = mode;
  document.querySelectorAll('#iModePills .mode-pill').forEach(p => {
    p.classList.toggle('active', p.dataset.mode === mode);
  });
  document.getElementById('iMonthWrap').classList.toggle('hidden', mode === 'year' || mode === 'months');
  document.getElementById('iDayWrap').classList.toggle('hidden', mode !== 'month');
  document.getElementById('iMonthsCheck').classList.toggle('show', mode === 'months');
  document.getElementById('iDaysCheck').classList.toggle('show', mode === 'days');
  if (mode === 'days') refreshIDaysCheck();
}

function refreshIDaysCheck() {
  const year = Number(document.getElementById('itemYear').value);
  const month = Number(document.getElementById('itemMonth').value) || 7;
  const n = new Date(year, month, 0).getDate();
  buildCheckGrid('iDaysCheck', n, null);
  // show month selector for days mode
  document.getElementById('iMonthWrap').classList.remove('hidden');
}

// ─── Upload ──────────────────────────────────────────────────
async function handleUpload() {
  const year = Number(document.getElementById('uploadYear').value);
  const month = Number(document.getElementById('uploadMonth').value);
  const dayVal = document.getElementById('uploadDay').value;
  const fileInput = document.getElementById('fileInput');
  if (!month) { showMsg('اختر الشهر', 'error'); return; }
  if (!fileInput.files.length) { showMsg('اختر ملف Excel', 'error'); return; }
  try {
    showMsg('جاري القراءة...', 'info');
    const buffer = await fileInput.files[0].arrayBuffer();
    const workbook = XLSX.read(buffer, { type: 'array' });
    if (!dayVal) {
      const catalogItems = parseCatalog(workbook);
      await mergeCatalog(currentType, catalogItems);
      await deleteRecords(currentType, year, month);
      const items = parseMonthlySheet(workbook);
      if (!items.length) { showMsg('لا أصناف منصرفة', 'error'); return; }
      const byDay = {};
      items.forEach(item => {
        Object.entries(item.daily).forEach(([d, qty]) => {
          if (qty > 0) {
            if (!byDay[d]) byDay[d] = [];
            byDay[d].push({ code: item.code, name: item.name, qty });
          }
        });
      });
      await saveRecord({
        id: `${currentType}_${year}_${month}_0`,
        type: currentType, year, month, day: 0, isMonthly: true,
        items: items.map(i => ({ code: i.code, name: i.name, qty: i.total, daily: i.daily })),
        uploadedAt: new Date().toISOString()
      });
      for (const [d, dayItems] of Object.entries(byDay)) {
        await saveRecord({
          id: `${currentType}_${year}_${month}_${d}`,
          type: currentType, year, month, day: Number(d), isMonthly: false,
          items: aggregateDaily(dayItems), uploadedAt: new Date().toISOString()
        });
      }
      showUploadSuccess(`✓ تم رفع الملف الشهري (${items.length} صنف)`);
    } else {
      const day = Number(dayVal);
      const catalogItems = parseCatalog(workbook);
      await mergeCatalog(currentType, catalogItems);
      const rows = parseDailySheet(workbook);
      if (!rows.length) { showMsg('لا بيانات صالحة', 'error'); return; }
      const items = aggregateDaily(rows);
      await deleteRecords(currentType, year, month, day);
      await saveRecord({
        id: `${currentType}_${year}_${month}_${day}`,
        type: currentType, year, month, day, isMonthly: false,
        items, uploadedAt: new Date().toISOString()
      });
      showUploadSuccess(`✓ تم رفع يوم ${day} (${items.length} صنف)`);
    }
    fileInput.value = '';
    const pill = document.getElementById('fileLabel');
    pill.textContent = 'لم يتم اختيار ملف';
    pill.classList.remove('has-file');
    pill.title = '';
    await refreshUploadStatus();
    await populateItems();
  } catch (err) {
    console.error(err);
    showMsg('خطأ: ' + err.message, 'error');
  }
}


function prevYearMonth(year, month) {
  if (month <= 1) return { year: year - 1, month: 12 };
  return { year, month: month - 1 };
}

async function collectMonthItems(type, year, month) {
  let recs = await getRecords(type, year, month, 0);
  if (!recs.length) recs = await getRecords(type, year, month);
  recs = recs.filter(r => !r.isCatalog);
  const map = new Map();
  recs.forEach(rec => {
    (rec.items || []).forEach(it => {
      if (!it || !it.code) return;
      const q = Number(it.qty) || 0;
      if (q <= 0) return;
      const code = String(it.code).trim().toLowerCase();
      if (!map.has(code)) map.set(code, { code, name: it.name || code, qty: 0 });
      const e = map.get(code);
      e.qty += q;
      if (it.name && String(it.name).length > (e.name || '').length) e.name = it.name;
    });
  });
  return Array.from(map.values()).sort((a, b) => b.qty - a.qty);
}

async function collectMonthDailyTotals(type, year, month) {
  const totals = {};
  let recs = await getRecords(type, year, month, 0);
  const monthly = recs.filter(r => r.isMonthly || r.day === 0);
  if (monthly.length) {
    monthly.forEach(rec => {
      (rec.items || []).forEach(it => {
        if (!it.daily) return;
        Object.entries(it.daily).forEach(([d, q]) => {
          const v = Number(q) || 0;
          if (v > 0) totals[d] = (totals[d] || 0) + v;
        });
      });
    });
  }
  if (!Object.keys(totals).length) {
    recs = await getRecords(type, year, month);
    recs.filter(r => !r.isMonthly && r.day > 0).forEach(rec => {
      const s = (rec.items || []).reduce((a, it) => a + (Number(it.qty) || 0), 0);
      if (s > 0) totals[String(rec.day)] = (totals[String(rec.day)] || 0) + s;
    });
  }
  return totals;
}

function paretoItems(items, totalQty, threshold) {
  let acc = 0;
  const out = [];
  for (const it of items) {
    acc += it.qty;
    out.push(it);
    if (acc >= totalQty * threshold) break;
  }
  return { count: out.length, qty: acc, pct: totalQty ? (acc / totalQty) * 100 : 0, list: out };
}

function findAnomalousDays(dailyTotals) {
  const entries = Object.entries(dailyTotals).map(([d, q]) => ({ day: Number(d), qty: Number(q) || 0 })).filter(x => x.qty > 0);
  if (entries.length < 3) return { avg: 0, days: [] };
  const avg = entries.reduce((s, x) => s + x.qty, 0) / entries.length;
  const days = entries.filter(x => x.qty >= avg * 2)
    .map(x => ({ ...x, avg, ratio: x.qty / avg }))
    .sort((a, b) => b.qty - a.qty);
  return { avg, days };
}

function changeLabel(curr, prev) {
  if (!prev) return { text: 'لا يوجد شهر سابق', cls: 'chg-flat' };
  const diff = curr - prev;
  const pct = prev ? (diff / prev) * 100 : 0;
  if (Math.abs(pct) < 0.5) return { text: 'ثابت تقريباً', cls: 'chg-flat' };
  if (pct > 0) return { text: `+${pct.toFixed(1)}% عن السابق`, cls: 'chg-up' };
  return { text: `${pct.toFixed(1)}% عن السابق`, cls: 'chg-down' };
}

let lastAdminReport = null;

async function buildMonthInsights(items, totalQty, year, month, dayVal) {
  const box = document.getElementById('gInsights');
  if (!box) return;
  if (dayVal) { box.innerHTML = ''; lastAdminReport = null; return; }

  const prev = prevYearMonth(year, month);
  const prevItems = await collectMonthItems(currentType, prev.year, prev.month);
  const prevMap = new Map(prevItems.map(i => [i.code, i.qty]));
  const prevTotal = prevItems.reduce((s, i) => s + i.qty, 0);

  const daily = await collectMonthDailyTotals(currentType, year, month);
  const anom = findAnomalousDays(daily);
  const p80 = paretoItems(items, totalQty, 0.8);
  const p50 = paretoItems(items, totalQty, 0.5);

  const catalog = await getCatalog(currentType);
  const moved = new Set(items.map(i => i.code));
  const idle = (catalog || []).filter(c => c && c.code && !moved.has(String(c.code).toLowerCase()))
    .map(c => ({ code: String(c.code).toLowerCase(), name: c.name || c.code }));

  const chg = changeLabel(totalQty, prevTotal);

  const compared = items.map(it => {
    const pq = prevMap.get(it.code) || 0;
    let delta = 'جديد';
    if (pq > 0) {
      const p = ((it.qty - pq) / pq) * 100;
      delta = (p >= 0 ? '+' : '') + p.toFixed(1) + '%';
    }
    return { ...it, prev: pq, delta };
  });

  const spikeDays = new Set(anom.days.map(d => d.day));
  const dailyRows = Object.entries(daily).map(([d, q]) => ({
    day: Number(d), qty: Number(q) || 0, flag: spikeDays.has(Number(d))
  })).sort((a, b) => a.day - b.day);

  lastAdminReport = {
    year, month, monthName: MONTH_NAMES[month], type: currentType,
    totalQty, itemCount: items.length, prevTotal, prevMonth: prev,
    chg, p50, p80, anom, idle, items: compared, top20: items.slice(0, 20),
    dailyRows, catalogCount: (catalog || []).length
  };

  const anomHtml = anom.days.length
    ? `<div class="table-wrap"><table>
        <thead><tr><th>اليوم</th><th>الكمية</th><th>كم ضعف المتوسط</th></tr></thead>
        <tbody>${anom.days.map(d => `<tr>
          <td>${d.day}</td><td class="qty-cell">${d.qty.toLocaleString('en')}</td>
          <td>${d.ratio.toFixed(1)}×</td></tr>`).join('')}</tbody></table></div>
       <div class="avg-box">المتوسط اليومي للأيام التي فيها حركة
         <strong>${Math.round(anom.avg).toLocaleString('en')}</strong>
       </div>`
    : `<div class="no-data" style="padding:.6rem">لا توجد أيام أعلى من ضعف المتوسط</div>`;

  const idleHtml = idle.length
    ? `<div class="table-search-wrap">${tableSearchBox('idleSearch', 'ابحث في الراكد...')}</div>
       <div class="table-wrap"><table>
        <thead><tr><th>#</th><th>الكود</th><th>اسم الصنف</th></tr></thead>
        <tbody id="idleBody">${idle.map((it, i) => `<tr data-q="${(it.name + ' ' + it.code).toLowerCase()}">
          <td>${i + 1}</td><td><code>${it.code}</code></td><td>${it.name}</td></tr>`).join('')}</tbody>
       </table></div>`
    : `<div class="no-data" style="padding:.6rem">لا توجد أصناف راكدة (أو لم يُرفع كتالوج الأصناف بعد)</div>`;

  box.innerHTML = `
    <div class="admin-bar">
      <button type="button" class="btn btn-primary" id="btnAdminReport">📄 تصدير التقرير الإداري</button>
    </div>
    <div class="insight-block">
      <div class="insight-title">مقارنة بالشهر السابق (${MONTH_NAMES[prev.month]} ${prev.year})</div>
      <div class="insight-cards">
        <div class="insight-card"><div class="k">هذا الشهر</div><div class="v">${totalQty.toLocaleString('en')}</div></div>
        <div class="insight-card"><div class="k">الشهر السابق</div><div class="v">${prevTotal ? prevTotal.toLocaleString('en') : '—'}</div></div>
        <div class="insight-card"><div class="k">التغيير</div><div class="v ${chg.cls}" style="font-size:1rem">${chg.text}</div></div>
      </div>
    </div>
    <div class="insight-block">
      <div class="insight-title">باريتو — تركيز الصرف</div>
      <div class="insight-cards">
        <button type="button" class="insight-card clickable" id="pareto50Card">
          <div class="k">50% من الوحدات المنصرفة تتمثل في</div>
          <div class="v">${p50.count}</div>
          <div class="s">صنف · اضغط للتفاصيل</div>
        </button>
        <button type="button" class="insight-card clickable" id="pareto80Card">
          <div class="k">80% من الوحدات المنصرفة تتمثل في</div>
          <div class="v">${p80.count}</div>
          <div class="s">من أصل ${items.length} · اضغط للتفاصيل</div>
        </button>
      </div>
      <div id="paretoDetail" class="pareto-detail" style="display:none"></div>
    </div>
    <div class="insight-block">
      <div class="insight-title">أيام شاذة (ضعف المتوسط أو أكثر)</div>
      ${anomHtml}
    </div>
    <div class="insight-block">
      <div class="insight-title">راكد هذا الشهر <small id="idleCount">(${idle.length})</small></div>
      ${idleHtml}
    </div>`;

  const b1 = document.getElementById('btnAdminReport');
  const b2 = document.getElementById('btnAdminReportTop');
  if (b1) b1.onclick = exportAdminReport;
  if (b2) b2.onclick = exportAdminReport;
  if (idle.length) bindTableSearch('idleSearch', 'idleBody', 'idleCount', idle.length);

  const p50c = document.getElementById('pareto50Card');
  const p80c = document.getElementById('pareto80Card');
  if (p50c) p50c.onclick = () => showParetoDetail(50, p50, totalQty);
  if (p80c) p80c.onclick = () => showParetoDetail(80, p80, totalQty);
}

function showParetoDetail(level, data, totalQty) {
  const el = document.getElementById('paretoDetail');
  if (!el) return;
  const list = data.list || [];
  if (!list.length) {
    el.style.display = 'block';
    el.innerHTML = '<div class="no-data">لا توجد بيانات</div>';
    return;
  }
  // destroy previous pareto chart if any
  charts = charts.filter(ch => {
    if (ch.canvas && ch.canvas.id === 'paretoDonut') {
      try { ch.destroy(); } catch (_) {}
      return false;
    }
    return true;
  });
  el.style.display = 'block';
  el.innerHTML = `
    <div class="pareto-detail-head">
      <strong>${level}% من الوحدات ← ${list.length} صنفاً</strong>
      <span>إجماليهم ${Math.round(data.qty).toLocaleString('en')} · ${formatPct(data.qty, totalQty)} من الشهر</span>
    </div>
    <div class="chart-box" style="margin-top:.5rem">
      <div style="height:240px"><canvas id="paretoDonut"></canvas></div>
    </div>
    <div class="table-search-wrap" style="margin-top:.5rem">${tableSearchBox('paretoSearch', 'ابحث في قائمة باريتو...')}</div>
    <div class="table-wrap"><table>
      <thead><tr><th>#</th><th>الكود</th><th>الصنف</th><th>الكمية</th><th>النسبة</th></tr></thead>
      <tbody id="paretoBody">${list.map((it, i) => `<tr data-q="${(it.name + ' ' + it.code).toLowerCase()}">
        <td>${i + 1}</td><td><code>${it.code}</code></td><td>${it.name}</td>
        <td class="qty-cell">${it.qty.toLocaleString('en')}</td>
        <td class="pct-cell">${formatPct(it.qty, totalQty)}</td>
      </tr>`).join('')}</tbody>
    </table></div>`;
  // doughnut of these items relative to their own sum (distribution among them)
  renderDoughnut('paretoDonut', list.map(i => i.name), list.map(i => i.qty));
  const leg = document.createElement('div');
  leg.innerHTML = buildDonutLegendHtml(list.map(i => i.name), list.map(i => i.qty));
  el.querySelector('.chart-box').appendChild(leg);
  bindTableSearch('paretoSearch', 'paretoBody', null, list.length);
  el.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

function downloadXlsx(wb, filename) {
  try {
    XLSX.writeFile(wb, filename);
    return true;
  } catch (e1) {
    try {
      const buf = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = filename;
      a.rel = 'noopener';
      document.body.appendChild(a);
      a.click();
      setTimeout(() => { a.remove(); URL.revokeObjectURL(url); }, 1500);
      return true;
    } catch (e2) {
      console.error(e1, e2);
      showMsg('تعذر حفظ الملف على هذا الجهاز: ' + (e2.message || e1.message), 'error');
      return false;
    }
  }
}

function exportAdminReport() {
  const r = lastAdminReport;
  if (!r) { showMsg('اعرض التحليل أولاً', 'info'); return; }
  try {
    const wb = XLSX.utils.book_new();
    const typeName = r.type === 'free' ? 'مجاني' : 'مدعم';
    const now = new Date();
    const stamp = now.toISOString().slice(0, 16).replace('T', ' ');
    const summary = [
      { البند: 'نوع الصرف', القيمة: typeName },
      { البند: 'السنة', القيمة: r.year },
      { البند: 'الشهر', القيمة: r.month + ' — ' + r.monthName },
      { البند: 'تاريخ استخراج التقرير', القيمة: stamp },
      { البند: 'إجمالي الوحدات', القيمة: r.totalQty },
      { البند: 'عدد الأصناف ذات حركة', القيمة: r.itemCount },
      { البند: 'شهر المقارنة', القيمة: r.prevMonth.month + ' / ' + r.prevMonth.year },
      { البند: 'إجمالي الشهر السابق', القيمة: r.prevTotal || 0 },
      { البند: 'التغيير عن السابق', القيمة: r.chg.text },
      { البند: 'أصناف تغطي 50% من الوحدات', القيمة: r.p50.count },
      { البند: 'كمية الـ 50%', القيمة: Math.round(r.p50.qty) },
      { البند: 'أصناف تغطي 80% من الوحدات', القيمة: r.p80.count },
      { البند: 'كمية الـ 80%', القيمة: Math.round(r.p80.qty) },
      { البند: 'متوسط الأيام ذات الحركة', القيمة: r.anom.avg ? Math.round(r.anom.avg) : 0 },
      { البند: 'عدد الأيام الشاذة', القيمة: r.anom.days.length },
      { البند: 'عدد الأصناف الراكدة', القيمة: r.idle.length },
      { البند: 'حجم الكتالوج', القيمة: r.catalogCount || r.idle.length + r.itemCount }
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(summary), 'ملخص تنفيذي');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(r.top20.map((it, i) => ({
      الترتيب: i + 1, الكود: it.code, الصنف: it.name, الكمية: it.qty,
      'النسبة من الإجمالي %': +((it.qty / r.totalQty) * 100).toFixed(4)
    }))), 'أعلى 20');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(r.items.map((it, i) => ({
      الترتيب: i + 1, الكود: it.code, الصنف: it.name,
      'كمية هذا الشهر': it.qty,
      'النسبة %': +((it.qty / r.totalQty) * 100).toFixed(4),
      'كمية الشهر السابق': it.prev,
      التغيير: it.delta
    }))), 'كل الأصناف');
    const paretoRows = [];
    let acc = 0;
    r.items.forEach((it, i) => {
      acc += it.qty;
      paretoRows.push({
        الترتيب: i + 1, الكود: it.code, الصنف: it.name, الكمية: it.qty,
        'تجميعي': acc,
        'تجميعي %': +((acc / r.totalQty) * 100).toFixed(2),
        'وصل لـ 80%': acc >= r.totalQty * 0.8 ? 'نعم' : ''
      });
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(paretoRows), 'باريتو تفصيلي');
    const dayRows = (r.dailyRows || []).map(d => ({
      اليوم: d.day, الكمية: d.qty,
      'النسبة من الشهر %': r.totalQty ? +((d.qty / r.totalQty) * 100).toFixed(3) : 0,
      شاذ: d.flag ? 'نعم' : ''
    }));
    if (dayRows.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(dayRows), 'حركة كل الأيام');
    if (r.anom.days.length) {
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(r.anom.days.map(d => ({
        اليوم: d.day, الكمية: d.qty, المتوسط: Math.round(d.avg),
        'كم ضعف المتوسط': +d.ratio.toFixed(2)
      }))), 'أيام شاذة');
    }
    if (r.idle.length) {
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(r.idle.map((it, i) => ({
        '#': i + 1, الكود: it.code, الصنف: it.name, الحالة: 'راكد — لا حركة هذا الشهر'
      }))), 'راكد');
    }
    const ok = downloadXlsx(wb, `تقرير_إداري_${typeName}_${r.year}_${r.month}.xlsx`);
    if (ok) showMsg('تم تجهيز التقرير الإداري', 'ok');
  } catch (err) {
    console.error(err);
    showMsg('خطأ في التقرير: ' + err.message, 'error');
  }
}


// ─── General: single month (existing) ────────────────────────
async function analyzeGeneralMonth(year, month, dayVal) {
  let records;
  if (dayVal) records = await getRecords(currentType, year, month, dayVal);
  else {
    records = await getRecords(currentType, year, month, 0);
    if (!records.length) records = await getRecords(currentType, year, month);
  }
  if (!records.length) {
    document.getElementById('gCharts').innerHTML = `<div class="empty-state"><div class="icon">📭</div><p>لا توجد بيانات</p></div>`;
    return;
  }
  const { days: uploadedDays, hasMonthly } = await getUploadedDays(currentType, year, month);
  let periodLabel = dayVal ? `يوم ${dayVal}` : (hasMonthly ? 'شهر' : (uploadedDays.length === 1 ? `يوم ${uploadedDays[0]}` : `${uploadedDays.length || ''} أيام`));

  const map = new Map();
  records.forEach(rec => {
    (rec.items || []).forEach(it => {
      if (!it.qty || it.qty <= 0) return;
      if (!map.has(it.code)) map.set(it.code, { code: it.code, name: it.name, qty: 0 });
      const e = map.get(it.code);
      e.qty += it.qty;
      if (it.name && it.name.length > (e.name || '').length) e.name = it.name;
    });
  });
  const items = Array.from(map.values()).filter(i => i.qty > 0).sort((a, b) => b.qty - a.qty);
  if (!items.length) {
    document.getElementById('gCharts').innerHTML = `<div class="empty-state"><div class="icon">🔍</div><p>لا أصناف</p></div>`;
    return;
  }
  renderGeneralItems(items, periodLabel, year, month, dayVal);
  try {
    await buildMonthInsights(items, items.reduce((s,i)=>s+i.qty,0), year, month, dayVal);
  } catch (e) {
    console.error(e);
    showMsg('جزء التقرير: ' + e.message, 'error');
  }
}

function renderGeneralItems(items, periodLabel, year, month, dayVal) {
  const totalQty = items.reduce((s, i) => s + i.qty, 0);
  const topN = 20;
  const top = items.slice(0, topN);
  const rest = items.slice(topN);

  document.getElementById('gSummary').style.display = 'grid';
  document.getElementById('gSummary').innerHTML = `
    <div class="summary-card"><div class="lbl">عدد الأصناف</div><div class="val">${items.length}</div></div>
    <div class="summary-card"><div class="lbl">إجمالي الوحدات</div><div class="val">${totalQty.toLocaleString('en')}</div></div>
    <div class="summary-card"><div class="lbl">الفترة</div><div class="val" style="font-size:1.15rem">${periodLabel}</div></div>`;
  const insightsHost = document.getElementById('gInsights');
  if (insightsHost && !dayVal) {
    insightsHost.innerHTML = `<div class="admin-bar"><button type="button" class="btn btn-primary" id="btnAdminReportTop">📄 تصدير التقرير الإداري</button></div>`;
  }

  const topLabels = top.map(i => i.name + '  ' + formatPct(i.qty, totalQty));
  const restLabels = rest.map(i => i.name + '  ' + formatPct(i.qty, totalQty));

  document.getElementById('gCharts').innerHTML = `
    <div class="chart-box">
      <div class="chart-head"><h3>أكثر ${Math.min(topN, top.length)} أصنافاً</h3>
        <button type="button" class="btn-mini" onclick="doExportChart('chartTop','top.png')">🖼</button></div>
      <div class="chart-scroll-y"><div class="chart-inner" style="width:100%"><canvas id="chartTop"></canvas></div></div>
    </div>
    <div class="chart-box full">
      <div class="chart-head"><h3>توزيع أعلى ${Math.min(20, top.length)} (نسب من مجموعهم)</h3>
        <button type="button" class="btn-mini" onclick="doExportChart('chartDonut','donut.png')">🖼</button></div>
      <div style="height:260px"><canvas id="chartDonut"></canvas></div>
      <div id="donutLegend"></div>
    </div>
    ${rest.length ? `<div class="chart-box full">
      <div class="chart-head"><h3>باقي الأصناف (${rest.length})</h3>
        <button type="button" class="btn-mini" onclick="doExportChart('chartRest','rest.png')">🖼</button></div>
      <div class="chart-scroll-y"><div class="chart-inner" style="width:100%"><canvas id="chartRest"></canvas></div></div>
    </div>` : ''}`;

  renderBarChart('chartTop', topLabels, top.map(i => i.qty), { grandTotal: totalQty });
  renderDoughnut('chartDonut', top.map(i => i.name), top.map(i => i.qty));
  const leg = document.getElementById('donutLegend');
  if (leg) leg.innerHTML = buildDonutLegendHtml(top.map(i => i.name), top.map(i => i.qty));
  if (rest.length) {
    renderBarChart('chartRest', restLabels, rest.map(i => i.qty), { grandTotal: totalQty });
  }

  const excelRows = items.map((it, i) => ({
    '#': i + 1, الكود: it.code, 'اسم الصنف': it.name, الكمية: it.qty,
    'النسبة %': formatPct(it.qty, totalQty).replace('%','')
  }));
  document.getElementById('gTable').innerHTML = `
    <div class="table-toolbar">
      <span class="table-title">جدول الأصناف <small id="gTableCount">(${items.length})</small></span>
      <button type="button" class="btn-mini" id="btnGExcel">📗 Excel</button>
    </div>
    <div class="table-search-wrap">
      <input type="search" id="gTableSearch" class="table-search" placeholder="ابحث بالاسم أو الكود..." autocomplete="off" />
    </div>
    <div class="table-wrap"><table>
      <thead><tr><th>#</th><th>الكود</th><th>اسم الصنف</th><th>الكمية</th><th>النسبة</th></tr></thead>
      <tbody id="gTableBody">${items.map((it, idx) => `<tr data-q="${(it.name + ' ' + it.code).toLowerCase()}">
        <td>${idx + 1}</td><td><code>${it.code}</code></td><td>${it.name}</td>
        <td class="qty-cell">${it.qty.toLocaleString('en')}</td>
        <td class="pct-cell">${formatPct(it.qty, totalQty)}</td>
      </tr>`).join('')}</tbody>
    </table></div>`;
  document.getElementById('btnGExcel').onclick = () =>
    exportExcel(excelRows, 'أصناف', `صرف_${year}_${month || 'year'}.xlsx`);
  bindTableSearch('gTableSearch', 'gTableBody', 'gTableCount', items.length);
}

// ─── General: year / multi-months ────────────────────────────
async function analyzeGeneralYearOrMonths(year, months) {
  // months: array of 1-12, or all 1-12 for full year
  const monthQtys = [];
  let yearTotal = 0;
  for (const m of months) {
    const q = await monthTotalQty(currentType, year, m);
    monthQtys.push({ month: m, name: MONTH_NAMES[m], qty: q });
    yearTotal += q;
  }
  // if full year mode, ensure all 12 shown (already are)
  const allZero = monthQtys.every(m => m.qty === 0);
  if (allZero) {
    document.getElementById('gCharts').innerHTML = `<div class="empty-state"><div class="icon">📭</div><p>لا توجد بيانات لهذه الفترة</p></div>`;
    document.getElementById('gSummary').style.display = 'none';
    return;
  }

  const monthsWithData = monthQtys.filter(m => m.qty > 0).length;
  document.getElementById('gSummary').style.display = 'grid';
  document.getElementById('gSummary').innerHTML = `
    <div class="summary-card"><div class="lbl">شهور بها بيانات</div><div class="val">${monthsWithData}</div>
      <div class="note">من أصل ${months.length} شهر معروض</div></div>
    <div class="summary-card"><div class="lbl">إجمالي الوحدات</div><div class="val">${yearTotal.toLocaleString('en')}</div></div>
    <div class="summary-card"><div class="lbl">السنة</div><div class="val">${year}</div></div>`;

  document.getElementById('gCharts').innerHTML = `
    <div class="chart-box full">
      <div class="chart-head"><h3>الكميات المنصرفة حسب الشهر</h3>
        <button type="button" class="btn-mini" onclick="doExportChart('chartMonths','months.png')">🖼</button></div>
      <div class="chart-scroll"><div style="height:320px;min-width:${Math.max(months.length * 50, 500)}px"><canvas id="chartMonths"></canvas></div></div>
    </div>`;

  renderBarChart('chartMonths',
    monthQtys.map(m => String(m.month)),
    monthQtys.map(m => m.qty),
    { horizontal: false, showAllLabels: true }
  );

  const excelRows = monthQtys.map(m => ({
    الشهر: m.month, الاسم: m.name, الكمية: m.qty,
    'النسبة %': yearTotal ? +((m.qty / yearTotal) * 100).toFixed(1) : 0
  }));
  document.getElementById('gTable').innerHTML = `
    <div class="table-toolbar"><span class="table-title">جدول الشهور <small id="gTableCount">(${monthQtys.length})</small></span>
      <button type="button" class="btn-mini" id="btnGExcel">📗 Excel</button></div>
    ${tableSearchBox('gTableSearch', 'ابحث باسم الشهر أو رقمه...')}
    <div class="table-wrap"><table>
      <thead><tr><th>الشهر</th><th>الاسم</th><th>الكمية</th><th>النسبة من الإجمالي</th></tr></thead>
      <tbody id="gTableBody">${monthQtys.map(m => `<tr class="${m.qty === 0 ? 'zero-row' : ''}" data-q="${m.month} ${m.name}">
        <td>${m.month}</td><td>${m.name}</td>
        <td class="qty-cell">${m.qty.toLocaleString('en')}</td>
        <td class="pct-cell">${yearTotal ? formatPct(m.qty, yearTotal) : '0%'}</td>
      </tr>`).join('')}</tbody>
    </table></div>`;
  document.getElementById('btnGExcel').onclick = () =>
    exportExcel(excelRows, 'شهور', `صرف_شهور_${year}.xlsx`);
  bindTableSearch('gTableSearch', 'gTableBody', 'gTableCount', monthQtys.length);
}

async function handleGeneralAnalyze() {
  const year = Number(document.getElementById('gYear').value);
  destroyCharts();
  document.getElementById('gCharts').innerHTML = '';
  document.getElementById('gTable').innerHTML = '';
    const gi = document.getElementById('gInsights'); if (gi) gi.innerHTML = '';
  document.getElementById('gSummary').style.display = 'none';

  try {
    if (gMode === 'month') {
      const month = Number(document.getElementById('gMonth').value);
      const dayVal = document.getElementById('gDay').value;
      if (!month) { showMsg('اختر الشهر', 'error'); return; }
      await analyzeGeneralMonth(year, month, dayVal);
    } else if (gMode === 'months') {
      const months = getChecked('gMonthsCheck');
      if (!months.length) { showMsg('اختر شهراً واحداً على الأقل', 'error'); return; }
      await analyzeGeneralYearOrMonths(year, months.sort((a, b) => a - b));
    } else {
      // full year 1-12
      await analyzeGeneralYearOrMonths(year, [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]);
    }
  } catch (err) {
    console.error(err);
    showMsg('خطأ: ' + err.message, 'error');
  }
}

// ─── Item analysis ───────────────────────────────────────────
async function getItemDailyForMonth(type, year, month, code) {
  let records = await getRecords(type, year, month, 0);
  if (!records.length) records = await getRecords(type, year, month);
  const dailyMap = {};
  let name = code;
  records.forEach(rec => {
    (rec.items || []).forEach(it => {
      if (it.code !== code) return;
      if (it.name) name = it.name;
      if (it.daily && typeof it.daily === 'object') {
        Object.entries(it.daily).forEach(([d, q]) => {
          dailyMap[d] = (dailyMap[d] || 0) + Number(q);
        });
      } else if (rec.day && rec.day > 0) {
        dailyMap[rec.day] = (dailyMap[rec.day] || 0) + (it.qty || 0);
      } else if (it.qty) {
        // only total, no daily – skip for daily map
      }
    });
  });
  if (!Object.keys(dailyMap).length) {
    const dayRecs = await getRecords(type, year, month);
    dayRecs.forEach(rec => {
      if (!rec.day || rec.day === 0) return;
      (rec.items || []).forEach(it => {
        if (it.code !== code) return;
        if (it.name) name = it.name;
        dailyMap[rec.day] = (dailyMap[rec.day] || 0) + (it.qty || 0);
      });
    });
  }
  return { name, dailyMap };
}

async function getDayTotalAll(type, year, month, day) {
  let total = 0;
  const dayRecs = await getRecords(type, year, month, day);
  if (dayRecs.length) {
    dayRecs.forEach(rec => (rec.items || []).forEach(it => { if (it.qty > 0) total += it.qty; }));
  } else {
    const monthRecs = await getRecords(type, year, month, 0);
    monthRecs.forEach(rec => {
      (rec.items || []).forEach(it => {
        if (it.daily && it.daily[day]) total += Number(it.daily[day]) || 0;
      });
    });
  }
  return total;
}

async function handleItemAnalyze() {
  const code = document.getElementById('itemCode').value;
  const year = Number(document.getElementById('itemYear').value);
  if (!code) { showMsg('اختر صنفاً من قائمة البحث', 'error'); return; }

  destroyCharts();
  document.getElementById('iResult').innerHTML = '';

  try {
    if (iMode === 'month') {
      const month = Number(document.getElementById('itemMonth').value);
      const dayVal = document.getElementById('itemDay').value;
      if (!month) { showMsg('اختر الشهر', 'error'); return; }
      await analyzeItemMonth(code, year, month, dayVal);
    } else if (iMode === 'months') {
      const months = getChecked('iMonthsCheck');
      if (!months.length) { showMsg('اختر شهراً واحداً على الأقل', 'error'); return; }
      await analyzeItemMonths(code, year, months.sort((a, b) => a - b));
    } else if (iMode === 'year') {
      await analyzeItemYear(code, year);
    } else if (iMode === 'days') {
      const month = Number(document.getElementById('itemMonth').value);
      const days = getChecked('iDaysCheck');
      if (!month) { showMsg('اختر الشهر', 'error'); return; }
      if (!days.length) { showMsg('اختر يوماً واحداً على الأقل', 'error'); return; }
      await analyzeItemDays(code, year, month, days.sort((a, b) => a - b));
    }
  } catch (err) {
    console.error(err);
    showMsg('خطأ: ' + err.message, 'error');
  }
}

async function analyzeItemMonth(code, year, month, dayVal) {
  const { name, dailyMap } = await getItemDailyForMonth(currentType, year, month, code);
  const entries = Object.entries(dailyMap).map(([d, q]) => [Number(d), q]).filter(([, q]) => q > 0).sort((a, b) => a[0] - b[0]);
  const total = entries.reduce((s, [, q]) => s + q, 0);

  if (dayVal) {
    const dayNum = Number(dayVal);
    const dayQty = dailyMap[dayNum] || 0;
    if (dayQty <= 0) {
      document.getElementById('iResult').innerHTML = `<div class="single-item-view">
        <div class="item-name">${name}</div><div class="item-code">${code}</div>
        <div class="no-data">لا توجد حركة صرف في يوم ${dayNum}</div></div>`;
      return;
    }
    const dayTot = await getDayTotalAll(currentType, year, month, dayNum);
    const pct = dayTot > 0 ? ((dayQty / Math.max(dayTot, dayQty)) * 100).toFixed(1) : '0';
    document.getElementById('iResult').innerHTML = `<div class="single-item-view">
      <div class="item-name">${name}</div><div class="item-code">${code}</div>
      <div class="day-first">يوم ${dayNum}</div>
      <div class="qty-label" style="margin-top:.35rem;margin-bottom:.15rem">الكمية المنصرفة</div>
      <div class="big-qty">${dayQty.toLocaleString('en')}</div>
      <div class="pct-badge">${pct}% من إجمالي صرف اليوم (${Math.max(dayTot, dayQty).toLocaleString('en')})</div>
    </div>`;
    return;
  }

  if (!entries.length) {
    document.getElementById('iResult').innerHTML = `<div class="single-item-view">
      <div class="item-name">${name}</div><div class="item-code">${code}</div>
      <div class="no-data">لا توجد حركة صرف في هذا الشهر</div></div>`;
    return;
  }

  // day totals for %
  const dayTotals = {};
  for (const [d] of entries) {
    dayTotals[d] = await getDayTotalAll(currentType, year, month, d);
  }

  const tableRows = entries.map(([d, q]) => {
    const tot = Math.max(dayTotals[d] || 0, q);
    return { day: d, qty: q, tot, pct: ((q / tot) * 100).toFixed(1) };
  });

  document.getElementById('iResult').innerHTML = `
    <div class="single-item-view">
      <div class="item-name">${name}</div><div class="item-code">${code}</div>
      <div class="big-qty">${total.toLocaleString('en')}</div>
      <div class="qty-label">إجمالي الشهر</div>
    </div>
    <div class="chart-box full" style="margin-top:1rem">
      <div class="chart-head"><h3>حركة الصرف · كل أيام الشهر</h3>
        <button type="button" class="btn-mini" onclick="doExportChart('chartItemDaily','item.png')">🖼</button></div>
      <div class="chart-scroll"><div id="wrapItemDaily" style="height:280px"><canvas id="chartItemDaily"></canvas></div></div>
    </div>
    <div class="table-toolbar" style="margin-top:1rem">
      <span class="table-title">تفاصيل الأيام <small id="iTableCount">(${tableRows.length})</small></span>
      <button type="button" class="btn-mini" id="btnIExcel">📗 Excel</button>
    </div>
    ${tableSearchBox('iTableSearch', 'ابحث برقم اليوم...')}
    <div class="table-wrap"><table>
      <thead><tr><th>اليوم</th><th>الكمية</th><th>إجمالي اليوم</th><th>النسبة</th></tr></thead>
      <tbody id="iTableBody">${tableRows.map(r => `<tr data-q="يوم ${r.day} ${r.day}">
        <td>${r.day}</td><td class="qty-cell">${r.qty.toLocaleString('en')}</td>
        <td>${r.tot.toLocaleString('en')}</td><td class="pct-cell">${r.pct}%</td>
      </tr>`).join('')}
      <tr class="total-row" style="font-weight:700;background:#f0fdfa"><td>الإجمالي</td><td class="qty-cell">${total.toLocaleString('en')}</td><td colspan="2"></td></tr>
      </tbody></table></div>`;

  const filled = fillMonthDays(year, month, dailyMap);
  renderLineChart('chartItemDaily', filled.labels, filled.values);
  document.getElementById('btnIExcel').onclick = () =>
    exportExcel(tableRows.map(r => ({ اليوم: r.day, الكمية: r.qty, 'إجمالي اليوم': r.tot, 'النسبة %': +r.pct })),
      name, `صرف_${code}_${year}_${month}.xlsx`);
  bindTableSearch('iTableSearch', 'iTableBody', 'iTableCount', tableRows.length);
}

async function analyzeItemMonths(code, year, months) {
  let name = code;
  const monthData = [];
  let grand = 0;
  for (const m of months) {
    const { name: n, dailyMap } = await getItemDailyForMonth(currentType, year, m, code);
    if (n) name = n;
    const q = Object.values(dailyMap).reduce((s, v) => s + v, 0);
    monthData.push({ month: m, name: MONTH_NAMES[m], qty: q });
    grand += q;
  }
  if (grand === 0) {
    document.getElementById('iResult').innerHTML = `<div class="single-item-view">
      <div class="item-name">${name}</div><div class="item-code">${code}</div>
      <div class="no-data">لا توجد حركة صرف في الشهور المحددة</div></div>`;
    return;
  }
  document.getElementById('iResult').innerHTML = `
    <div class="single-item-view">
      <div class="item-name">${name}</div><div class="item-code">${code}</div>
      <div class="big-qty">${grand.toLocaleString('en')}</div>
      <div class="qty-label">إجمالي الشهور المحددة</div>
    </div>
    <div class="chart-box full" style="margin-top:1rem">
      <div class="chart-head"><h3>الكمية حسب الشهر</h3>
        <button type="button" class="btn-mini" onclick="doExportChart('chartItemMonths','im.png')">🖼</button></div>
      <div class="chart-scroll"><div style="height:300px;min-width:${Math.max(months.length * 55, 400)}px"><canvas id="chartItemMonths"></canvas></div></div>
    </div>
    <div class="table-toolbar" style="margin-top:1rem">
      <span class="table-title">جدول الشهور <small id="iTableCount">(${monthData.length})</small></span>
      <button type="button" class="btn-mini" id="btnIExcel">📗 Excel</button>
    </div>
    ${tableSearchBox('iTableSearch', 'ابحث بالشهر...')}
    <div class="table-wrap"><table>
      <thead><tr><th>الشهر</th><th>الاسم</th><th>الكمية</th><th>النسبة</th></tr></thead>
      <tbody id="iTableBody">${monthData.map(m => `<tr class="${m.qty === 0 ? 'zero-row' : ''}" data-q="${m.month} ${m.name}">
        <td>${m.month}</td><td>${m.name}</td>
        <td class="qty-cell">${m.qty.toLocaleString('en')}</td>
        <td class="pct-cell">${grand ? formatPct(m.qty, grand) : '0%'}</td>
      </tr>`).join('')}</tbody></table></div>`;

  renderBarChart('chartItemMonths', monthData.map(m => String(m.month)), monthData.map(m => m.qty), { horizontal: false, showAllLabels: true });
  document.getElementById('btnIExcel').onclick = () =>
    exportExcel(monthData.map(m => ({ الشهر: m.month, الاسم: m.name, الكمية: m.qty, 'النسبة %': grand ? +((m.qty / grand) * 100).toFixed(1) : 0 })),
      name, `صرف_${code}_${year}_شهور.xlsx`);
  bindTableSearch('iTableSearch', 'iTableBody', 'iTableCount', monthData.length);
}

async function analyzeItemYear(code, year) {
  let name = code;
  const monthData = [];
  const allDaily = []; // {label, qty} for full year timeline
  let grand = 0;

  for (let m = 1; m <= 12; m++) {
    const { name: n, dailyMap } = await getItemDailyForMonth(currentType, year, m, code);
    if (n && n !== code) name = n;
    const q = Object.values(dailyMap).reduce((s, v) => s + v, 0);
    monthData.push({ month: m, name: MONTH_NAMES[m], qty: q });
    grand += q;
    Object.entries(dailyMap).forEach(([d, qty]) => {
      if (qty > 0) allDaily.push({ label: `${m}/${d}`, month: m, day: Number(d), qty });
    });
  }
  allDaily.sort((a, b) => a.month - b.month || a.day - b.day);

  if (grand === 0) {
    document.getElementById('iResult').innerHTML = `<div class="single-item-view">
      <div class="item-name">${name}</div><div class="item-code">${code}</div>
      <div class="no-data">لا توجد حركة صرف خلال السنة</div></div>`;
    return;
  }

  document.getElementById('iResult').innerHTML = `
    <div class="single-item-view">
      <div class="item-name">${name}</div><div class="item-code">${code}</div>
      <div class="big-qty">${grand.toLocaleString('en')}</div>
      <div class="qty-label">إجمالي السنة ${year}</div>
    </div>
    <div class="chart-box full" style="margin-top:1rem">
      <div class="chart-head"><h3>الكمية حسب الشهر (1–12)</h3>
        <button type="button" class="btn-mini" onclick="doExportChart('chartItemYearM','y-m.png')">🖼</button></div>
      <div class="chart-scroll"><div style="height:300px;min-width:560px"><canvas id="chartItemYearM"></canvas></div></div>
    </div>
    ${allDaily.length ? `
    <div class="chart-box full" style="margin-top:1rem">
      <div class="chart-head"><h3>خط زمني يومي لكل أيام السنة</h3>
        <button type="button" class="btn-mini" onclick="doExportChart('chartItemYearD','y-d.png')">🖼</button></div>
      <div class="chart-scroll"><div style="height:280px"><canvas id="chartItemYearD"></canvas></div></div>
    </div>` : ''}
    <div class="table-toolbar" style="margin-top:1rem">
      <span class="table-title">جدول الشهور <small id="iTableCount">(${monthData.length})</small></span>
      <button type="button" class="btn-mini" id="btnIExcel">📗 Excel</button>
    </div>
    ${tableSearchBox('iTableSearch', 'ابحث بالشهر...')}
    <div class="table-wrap"><table>
      <thead><tr><th>الشهر</th><th>الاسم</th><th>الكمية</th><th>النسبة</th></tr></thead>
      <tbody id="iTableBody">${monthData.map(m => `<tr class="${m.qty === 0 ? 'zero-row' : ''}" data-q="${m.month} ${m.name}">
        <td>${m.month}</td><td>${m.name}</td>
        <td class="qty-cell">${m.qty.toLocaleString('en')}</td>
        <td class="pct-cell">${formatPct(m.qty, grand)}</td>
      </tr>`).join('')}</tbody></table></div>`;

  renderBarChart('chartItemYearM', monthData.map(m => String(m.month)), monthData.map(m => m.qty), { horizontal: false, showAllLabels: true });
  if (allDaily.length) {
    renderLineChart('chartItemYearD', allDaily.map(d => d.label), allDaily.map(d => d.qty));
  }
  document.getElementById('btnIExcel').onclick = () =>
    exportExcel(monthData.map(m => ({ الشهر: m.month, الاسم: m.name, الكمية: m.qty, 'النسبة %': +((m.qty / grand) * 100).toFixed(1) })),
      name, `صرف_${code}_${year}.xlsx`);
  bindTableSearch('iTableSearch', 'iTableBody', 'iTableCount', monthData.length);
}

async function analyzeItemDays(code, year, month, days) {
  const { name, dailyMap } = await getItemDailyForMonth(currentType, year, month, code);
  const rows = days.map(d => {
    const q = dailyMap[d] || 0;
    return { day: d, qty: q };
  });
  const total = rows.reduce((s, r) => s + r.qty, 0);
  if (total === 0) {
    document.getElementById('iResult').innerHTML = `<div class="single-item-view">
      <div class="item-name">${name}</div><div class="item-code">${code}</div>
      <div class="no-data">لا توجد حركة صرف في الأيام المحددة</div></div>`;
    return;
  }

  const enriched = [];
  for (const r of rows) {
    const tot = r.qty > 0 ? await getDayTotalAll(currentType, year, month, r.day) : 0;
    const t = Math.max(tot, r.qty);
    enriched.push({ ...r, tot: t, pct: t ? ((r.qty / t) * 100).toFixed(1) : '0' });
  }

  document.getElementById('iResult').innerHTML = `
    <div class="single-item-view">
      <div class="item-name">${name}</div><div class="item-code">${code}</div>
      <div class="big-qty">${total.toLocaleString('en')}</div>
      <div class="qty-label">إجمالي الأيام المحددة</div>
    </div>
    <div class="chart-box full" style="margin-top:1rem">
      <div class="chart-head"><h3>الكمية في الأيام المحددة</h3>
        <button type="button" class="btn-mini" onclick="doExportChart('chartItemDays','days.png')">🖼</button></div>
      <div class="chart-scroll"><div style="height:280px;min-width:${Math.max(days.length * 40, 360)}px"><canvas id="chartItemDays"></canvas></div></div>
    </div>
    <div class="table-toolbar" style="margin-top:1rem">
      <span class="table-title">تفاصيل الأيام <small id="iTableCount">(${enriched.length})</small></span>
      <button type="button" class="btn-mini" id="btnIExcel">📗 Excel</button>
    </div>
    ${tableSearchBox('iTableSearch', 'ابحث برقم اليوم...')}
    <div class="table-wrap"><table>
      <thead><tr><th>اليوم</th><th>الكمية</th><th>إجمالي اليوم</th><th>النسبة</th></tr></thead>
      <tbody id="iTableBody">${enriched.map(r => `<tr class="${r.qty === 0 ? 'zero-row' : ''}" data-q="يوم ${r.day} ${r.day}">
        <td>${r.day}</td><td class="qty-cell">${r.qty.toLocaleString('en')}</td>
        <td>${r.tot.toLocaleString('en')}</td><td class="pct-cell">${r.pct}%</td>
      </tr>`).join('')}</tbody></table></div>`;

  renderBarChart('chartItemDays', rows.map(r => String(r.day)), rows.map(r => r.qty), { horizontal: false, showAllLabels: true });
  document.getElementById('btnIExcel').onclick = () =>
    exportExcel(enriched.map(r => ({ اليوم: r.day, الكمية: r.qty, 'إجمالي اليوم': r.tot, 'النسبة %': +r.pct })),
      name, `صرف_${code}_${year}_${month}_أيام.xlsx`);
  bindTableSearch('iTableSearch', 'iTableBody', 'iTableCount', enriched.length);
}


// ─── Custom select (unified look) ────────────────────────────
const customSelects = new Map(); // id -> { refresh }

function enhanceSelect(selectId, opts = {}) {
  const sel = document.getElementById(selectId);
  if (!sel || sel.dataset.enhanced === '1') return;
  sel.dataset.enhanced = '1';
  sel.classList.add('cselect-native');

  const wrap = document.createElement('div');
  wrap.className = 'cselect';
  wrap.dataset.for = selectId;

  const trigger = document.createElement('button');
  trigger.type = 'button';
  trigger.className = 'cselect-trigger';
  trigger.innerHTML = `<span class="cs-val"></span><span class="cs-arrow">▾</span>`;

  const menu = document.createElement('div');
  menu.className = 'cselect-menu';

  sel.parentNode.insertBefore(wrap, sel);
  wrap.appendChild(trigger);
  wrap.appendChild(menu);
  wrap.appendChild(sel);

  function syncLabel() {
    const opt = sel.options[sel.selectedIndex];
    trigger.querySelector('.cs-val').textContent = opt ? opt.textContent : (opts.placeholder || '—');
  }

  function rebuildMenu() {
    menu.innerHTML = '';
    Array.from(sel.options).forEach((opt, idx) => {
      const div = document.createElement('div');
      div.className = 'cselect-opt' + (opt.selected ? ' selected' : '');
      div.textContent = opt.textContent;
      div.dataset.value = opt.value;
      div.dataset.idx = idx;
      div.addEventListener('click', (e) => {
        e.stopPropagation();
        sel.selectedIndex = idx;
        sel.dispatchEvent(new Event('change', { bubbles: true }));
        syncLabel();
        wrap.classList.remove('open');
        rebuildMenu();
      });
      menu.appendChild(div);
    });
  }

  trigger.addEventListener('click', (e) => {
    e.preventDefault();
    e.stopPropagation();
    // close others
    document.querySelectorAll('.cselect.open').forEach(el => {
      if (el !== wrap) el.classList.remove('open');
    });
    document.getElementById('itemDropdown')?.classList.remove('show');
    rebuildMenu();
    wrap.classList.toggle('open');
  });

  sel.addEventListener('change', syncLabel);

  syncLabel();
  rebuildMenu();

  customSelects.set(selectId, {
    refresh() { rebuildMenu(); syncLabel(); }
  });
}

function refreshCustomSelect(selectId) {
  const cs = customSelects.get(selectId);
  if (cs) cs.refresh();
}

document.addEventListener('click', (e) => {
  if (!e.target.closest('.cselect')) {
    document.querySelectorAll('.cselect.open').forEach(el => el.classList.remove('open'));
  }
});


// ─── Init ────────────────────────────────────────────────────
async function init() {
  db = await openDB();
  const monthLabelsNum = MONTH_NAMES.map((n, i) => i ? `${i} – ${n}` : '');
  buildCheckGrid('gMonthsCheck', 12, monthLabelsNum);
  buildCheckGrid('iMonthsCheck', 12, monthLabelsNum);

  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      currentType = btn.dataset.type;
      refreshUploadStatus();
      populateItems();
    });
  });

  document.querySelectorAll('.section-tab').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.section-tab').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      document.querySelectorAll('.section-panel').forEach(p => p.classList.remove('active'));
      const sec = document.getElementById('sec-' + btn.dataset.section);
      if (sec) sec.classList.add('active');
    });
  });

  document.querySelectorAll('#gModePills .mode-pill').forEach(p => {
    p.addEventListener('click', () => setGMode(p.dataset.mode));
  });
  document.querySelectorAll('#iModePills .mode-pill').forEach(p => {
    p.addEventListener('click', () => setIMode(p.dataset.mode));
  });

  document.getElementById('uploadYear').addEventListener('change', () => {
    fillDays('uploadDay', document.getElementById('uploadYear').value, document.getElementById('uploadMonth').value);
    refreshUploadStatus();
  });
  document.getElementById('uploadMonth').addEventListener('change', () => {
    fillDays('uploadDay', document.getElementById('uploadYear').value, document.getElementById('uploadMonth').value);
    refreshUploadStatus();
  });
  document.getElementById('gYear').addEventListener('change', () => {
    fillDays('gDay', document.getElementById('gYear').value, document.getElementById('gMonth').value);
  });
  document.getElementById('gMonth').addEventListener('change', () => {
    fillDays('gDay', document.getElementById('gYear').value, document.getElementById('gMonth').value);
  });
  document.getElementById('itemYear').addEventListener('change', () => {
    fillDays('itemDay', document.getElementById('itemYear').value, document.getElementById('itemMonth').value);
    if (iMode === 'days') refreshIDaysCheck();
  });
  document.getElementById('itemMonth').addEventListener('change', () => {
    fillDays('itemDay', document.getElementById('itemYear').value, document.getElementById('itemMonth').value);
    if (iMode === 'days') refreshIDaysCheck();
  });

  document.getElementById('fileInput').addEventListener('change', () => {
    const pill = document.getElementById('fileLabel');
    const files = document.getElementById('fileInput').files;
    if (files.length) {
      pill.textContent = files[0].name;
      pill.classList.add('has-file');
      pill.title = files[0].name;
    } else {
      pill.textContent = 'لم يتم اختيار ملف';
      pill.classList.remove('has-file');
      pill.title = '';
    }
  });

  document.getElementById('btnUpload').addEventListener('click', handleUpload);
  document.getElementById('btnGeneral').addEventListener('click', handleGeneralAnalyze);
  document.getElementById('btnItem').addEventListener('click', handleItemAnalyze);

  setupItemSearch();
  fillDays('uploadDay', 2026, 7);
  fillDays('gDay', 2026, 7);
  fillDays('itemDay', 2026, 7);

  // Enhance all selects to unified custom UI
  ['uploadYear','uploadMonth','uploadDay','gYear','gMonth','gDay','itemYear','itemMonth','itemDay']
    .forEach(id => enhanceSelect(id));

  await refreshUploadStatus();
  await populateItems();
}

init().catch(err => {
  console.error(err);
  showMsg('فشل التهيئة', 'error');
});
