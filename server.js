'use strict';

require('dotenv').config();

const express  = require('express');
const path     = require('path');
const fs       = require('fs');
const os       = require('os');
const XLSX     = require('xlsx');
const ExcelJS  = require('exceljs');

const yadisk             = require('./yadisk');
const mailer             = require('./mailer');
const { startScheduler, archiveToday } = require('./scheduler');

const app  = express();
const PORT = process.env.PORT || 3000;

const DATA_DIR = path.join(__dirname, 'data');
fs.mkdirSync(path.join(DATA_DIR, 'archive'), { recursive: true });

app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// ─── Drivers database ─────────────────────────────────────────────

function loadDrivers() {
  const filePath = path.join(__dirname, 'drivers.xlsx');
  if (!fs.existsSync(filePath)) return [];
  const wb = XLSX.readFile(filePath);
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
  return rows.slice(1).map(r => ({
    key:   String(r[0] || '').trim(),
    fio:   String(r[1] || '').trim(),
    brand: String(r[2] || '').trim(),
    plate: String(r[3] || '').trim(),
    phone: String(r[4] || '').trim(),
  })).filter(d => d.key);
}

const DRIVERS = loadDrivers();

// ─── Helpers ──────────────────────────────────────────────────────

function todayFilename() {
  const now = new Date();
  const dd   = String(now.getDate()).padStart(2, '0');
  const mm   = String(now.getMonth() + 1).padStart(2, '0');
  const yyyy = now.getFullYear();
  return `Отчёт_${dd}.${mm}.${yyyy}.xlsx`;
}

function todayPath() {
  return path.join(DATA_DIR, 'today.xlsx');
}

function styleHeaderRow(ws) {
  const COLS = 4;
  const row = ws.getRow(1);
  for (let c = 1; c <= COLS; c++) {
    const cell = row.getCell(c);
    cell.font      = { bold: true, color: { argb: 'FFFFFFFF' }, size: 12 };
    cell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' } };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
    cell.border    = {
      top: { style: 'thin', color: { argb: 'FF2F5496' } }, bottom: { style: 'thin', color: { argb: 'FF2F5496' } },
      left: { style: 'thin', color: { argb: 'FF2F5496' } }, right: { style: 'thin', color: { argb: 'FF2F5496' } },
    };
  }
  row.height = 22;
}

function styleDataRow(row, rowNumber) {
  const isEven = rowNumber % 2 === 0;
  const COLS = 4;
  for (let c = 1; c <= COLS; c++) {
    const cell = row.getCell(c);
    cell.fill   = { type: 'pattern', pattern: 'solid', fgColor: { argb: isEven ? 'FFD9E1F2' : 'FFFFFFFF' } };
    cell.border = {
      top: { style: 'thin', color: { argb: 'FFBFBFBF' } }, bottom: { style: 'thin', color: { argb: 'FFBFBFBF' } },
      left: { style: 'thin', color: { argb: 'FFBFBFBF' } }, right: { style: 'thin', color: { argb: 'FFBFBFBF' } },
    };
    cell.alignment = { vertical: 'middle', horizontal: c === 1 ? 'left' : 'center' };
  }
  row.height = 18;
}

// Append one record row to today.xlsx (create with headers if missing)
async function appendRecord(record) {
  const filePath = todayPath();

  // Read existing data rows using XLSX (handles any xlsx format, avoids exceljs read bugs)
  let existingRows = [];
  if (fs.existsSync(filePath)) {
    const existingWb = XLSX.readFile(filePath);
    const existingWs = existingWb.Sheets[existingWb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(existingWs, { header: 1 });
    existingRows = rows.slice(1); // skip header
  }

  // Build new workbook with exceljs (for styling)
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Данные');

  // Add header row first, then set column widths
  ws.addRow(['ФИО', 'Дата', 'Время', 'Количество коробов']);
  ws.getColumn(1).width = 30;
  ws.getColumn(2).width = 14;
  ws.getColumn(3).width = 10;
  ws.getColumn(4).width = 22;
  styleHeaderRow(ws);

  // Re-add existing data rows
  existingRows.forEach(row => {
    const newRow = ws.addRow([row[0], row[1], row[2], row[3]]);
    styleDataRow(newRow, ws.rowCount);
  });

  // Add new record
  const newRow = ws.addRow([record.fio, record.date, record.time, record.box_count]);
  styleDataRow(newRow, ws.rowCount);

  const buffer = await wb.xlsx.writeBuffer();
  fs.writeFileSync(filePath, buffer);
  return buffer;
}

// ─── Helpers ──────────────────────────────────────────────────────

// If today.xlsx is missing locally (e.g. after cloud restart), restore it from YaDisk
async function ensureTodayFile() {
  const filePath = todayPath();
  if (fs.existsSync(filePath)) return;

  try {
    const buffer = await yadisk.downloadFile(todayFilename());
    if (buffer) {
      fs.writeFileSync(filePath, buffer);
      console.log('[Recovery] Файл восстановлен с Яндекс Диска');
    }
  } catch (err) {
    console.error('[Recovery] Не удалось скачать файл:', err.message);
  }
}

// ─── Routes ───────────────────────────────────────────────────────

// Driver lookup by key (surname)
app.get('/api/driver', (req, res) => {
  const q = String(req.query.q || '').trim().toLowerCase();
  if (!q) return res.json([]);
  const matches = DRIVERS.filter(d => d.key.toLowerCase().includes(q));
  res.json(matches);
});

// Driver form submission
app.post('/api/submit', async (req, res) => {
  const { driver_key, date: dateRaw, time: timeRaw, box_count } = req.body;

  if (!driver_key || !box_count) {
    return res.status(400).json({ error: 'Заполните все поля' });
  }

  const driver = DRIVERS.find(d => d.key.toLowerCase() === String(driver_key).trim().toLowerCase());
  if (!driver) {
    return res.status(400).json({ error: 'Водитель не найден в базе' });
  }

  const count = parseInt(box_count);
  if (isNaN(count) || count < 1) {
    return res.status(400).json({ error: 'Некорректное количество коробов (не менее 1)' });
  }

  // Use client date/time (arrival); fallback to server time
  let date, time;
  if (dateRaw && timeRaw) {
    const p = String(dateRaw).split('-');
    date = p.length === 3 ? `${p[2]}.${p[1]}.${p[0]}` : dateRaw;
    time = timeRaw;
  } else {
    const now = new Date();
    date = `${String(now.getDate()).padStart(2,'0')}.${String(now.getMonth()+1).padStart(2,'0')}.${now.getFullYear()}`;
    time = `${String(now.getHours()).padStart(2,'0')}:${String(now.getMinutes()).padStart(2,'0')}`;
  }

  try {
    await ensureTodayFile();
    const buffer = await appendRecord({ fio: driver.fio, date, time, box_count: count });

    // Upload to YaDisk in background (don't block the response)
    yadisk.uploadFile(buffer, todayFilename()).catch(err =>
      console.error('[YaDisk] Ошибка загрузки:', err.message)
    );

    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Manual send (simple admin endpoint)
app.post('/api/send', async (req, res) => {
  const fp = todayPath();
  if (!fs.existsSync(fp)) {
    return res.status(400).json({ error: 'Нет данных за сегодня' });
  }
  try {
    const result = await mailer.sendReport(fp);
    archiveToday();
    res.json({ success: true, message: `Отчёт отправлен (${result.count} записей, ${result.totalBoxes} коробов)` });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ─── Start ────────────────────────────────────────────────────────

yadisk.ensureFolder().catch(err => console.error('[YaDisk] init:', err.message));
startScheduler();

app.listen(PORT, '0.0.0.0', () => {
  const ips = [];
  for (const list of Object.values(os.networkInterfaces())) {
    for (const iface of list) {
      if (iface.family === 'IPv4' && !iface.internal) ips.push(iface.address);
    }
  }

  console.log('\n═══════════════════════════════════════════');
  console.log('  Сервер запущен!');
  console.log('───────────────────────────────────────────');
  console.log(`  Форма (ПК):    http://localhost:${PORT}`);
  console.log(`  Форма (телефон): http://${ips[0] || '?'}:${PORT}`);
  console.log('═══════════════════════════════════════════\n');
});
