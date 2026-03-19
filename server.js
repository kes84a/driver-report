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
const { startScheduler } = require('./scheduler');

const app  = express();
const PORT = process.env.PORT || 3000;

const DATA_DIR = path.join(__dirname, 'data');
fs.mkdirSync(DATA_DIR, { recursive: true });

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
    key:    String(r[0] || '').trim(),
    fio:    String(r[1] || '').trim(),
    brand:  String(r[2] || '').trim(),
    plate:  String(r[3] || '').trim(),
    phone:  String(r[4] || '').trim(),
    region: String(r[5] || 'МСК').trim() || 'МСК',
  })).filter(d => d.key);
}

const DRIVERS = loadDrivers();

// ─── File helpers ──────────────────────────────────────────────────
// date format: DD.MM.YYYY

function remoteFilename(date) {
  return `Отчёт_${date}.xlsx`;
}

function localPath(date) {
  return path.join(DATA_DIR, `report_${date}.xlsx`);
}

function serverDateDDMMYYYY() {
  const now = new Date();
  const dd   = String(now.getDate()).padStart(2, '0');
  const mm   = String(now.getMonth() + 1).padStart(2, '0');
  const yyyy = now.getFullYear();
  return `${dd}.${mm}.${yyyy}`;
}

// ─── Styling ───────────────────────────────────────────────────────

function styleHeaderRow(ws) {
  const COLS = 6;
  const row = ws.getRow(1);
  for (let c = 1; c <= COLS; c++) {
    const cell = row.getCell(c);
    cell.font      = { bold: true, color: { argb: 'FFFFFFFF' }, size: 12 };
    cell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' } };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
    cell.border    = {
      top:    { style: 'thin', color: { argb: 'FF2F5496' } },
      bottom: { style: 'thin', color: { argb: 'FF2F5496' } },
      left:   { style: 'thin', color: { argb: 'FF2F5496' } },
      right:  { style: 'thin', color: { argb: 'FF2F5496' } },
    };
  }
  row.height = 22;
}

function styleDataRow(row, rowNumber) {
  const isEven = rowNumber % 2 === 0;
  const COLS = 6;
  for (let c = 1; c <= COLS; c++) {
    const cell = row.getCell(c);
    cell.fill   = { type: 'pattern', pattern: 'solid', fgColor: { argb: isEven ? 'FFD9E1F2' : 'FFFFFFFF' } };
    cell.border = {
      top:    { style: 'thin', color: { argb: 'FFBFBFBF' } },
      bottom: { style: 'thin', color: { argb: 'FFBFBFBF' } },
      left:   { style: 'thin', color: { argb: 'FFBFBFBF' } },
      right:  { style: 'thin', color: { argb: 'FFBFBFBF' } },
    };
    cell.alignment = { vertical: 'middle', horizontal: c === 6 ? 'left' : 'center' };
  }
  row.height = 18;
}

// ─── Core write function ───────────────────────────────────────────

// Reads existing rows from a local xlsx, appends a new record, writes styled file.
// record: { date, time, box_count, region, plate, fio }
async function appendRecord(record) {
  const fpath = localPath(record.date);

  let existingRows = [];
  if (fs.existsSync(fpath)) {
    const existingWb = XLSX.readFile(fpath);
    const existingWs = existingWb.Sheets[existingWb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(existingWs, { header: 1 });
    existingRows = rows.slice(1); // skip header
  }

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Данные');

  ws.addRow(['Дата', 'Время', 'Количество коробов', 'Регион', 'Номер ТС', 'Водитель']);
  ws.getColumn(1).width = 14;
  ws.getColumn(2).width = 10;
  ws.getColumn(3).width = 22;
  ws.getColumn(4).width = 10;
  ws.getColumn(5).width = 16;
  ws.getColumn(6).width = 32;
  styleHeaderRow(ws);

  existingRows.forEach(row => {
    const newRow = ws.addRow([row[0], row[1], row[2], row[3], row[4], row[5]]);
    styleDataRow(newRow, ws.rowCount);
  });

  const newRow = ws.addRow([record.date, record.time, record.box_count, record.region, record.plate, record.fio]);
  styleDataRow(newRow, ws.rowCount);

  const buffer = await wb.xlsx.writeBuffer();
  fs.writeFileSync(fpath, buffer);
  return buffer;
}

// ─── Ensure file exists locally (download from YaDisk if needed) ───

async function ensureDateFile(date) {
  const fpath = localPath(date);
  if (fs.existsSync(fpath)) return;

  try {
    const buffer = await yadisk.downloadFile(remoteFilename(date));
    if (buffer) {
      fs.writeFileSync(fpath, buffer);
      console.log(`[Recovery] Файл восстановлен с Яндекс Диска: ${remoteFilename(date)}`);
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

  // Convert date from YYYY-MM-DD to DD.MM.YYYY; fallback to server date
  let date, time;
  if (dateRaw && timeRaw) {
    const p = String(dateRaw).split('-');
    date = p.length === 3 ? `${p[2]}.${p[1]}.${p[0]}` : dateRaw;
    time = timeRaw;
  } else {
    date = serverDateDDMMYYYY();
    const now = new Date();
    time = `${String(now.getHours()).padStart(2,'0')}:${String(now.getMinutes()).padStart(2,'0')}`;
  }

  try {
    // Restore from YaDisk if local file for this date is missing
    await ensureDateFile(date);

    const buffer = await appendRecord({ fio: driver.fio, plate: driver.plate, region: driver.region, date, time, box_count: count });

    // Upload to YaDisk in background
    yadisk.uploadFile(buffer, remoteFilename(date)).catch(err =>
      console.error('[YaDisk] Ошибка загрузки:', err.message)
    );

    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Manual send
app.post('/api/send', async (req, res) => {
  const date = serverDateDDMMYYYY();
  const fp = localPath(date);
  if (!fs.existsSync(fp)) {
    return res.status(400).json({ error: 'Нет данных за сегодня' });
  }
  try {
    const result = await mailer.sendReport(fp);
    fs.unlinkSync(fp);
    res.json({ success: true, message: `Отчёт отправлен (${result.count} записей, ${result.totalBoxes} коробов)` });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Reformat existing file on YaDisk: download → rewrite with current styles → upload back
app.post('/api/reformat', async (req, res) => {
  const date = req.body.date || serverDateDDMMYYYY(); // DD.MM.YYYY
  try {
    await ensureDateFile(date);
    const fpath = localPath(date);

    if (!fs.existsSync(fpath)) {
      return res.status(404).json({ error: `Файл за ${date} не найден ни локально, ни на YaDisk` });
    }

    // Read all existing rows
    const existingWb = XLSX.readFile(fpath);
    const existingWs = existingWb.Sheets[existingWb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(existingWs, { header: 1 });
    const dataRows = rows.slice(1);

    // Rewrite with current styling
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Данные');
    ws.addRow(['Дата', 'Время', 'Количество коробов', 'Регион', 'Номер ТС', 'Водитель']);
    ws.getColumn(1).width = 14; ws.getColumn(2).width = 10;
    ws.getColumn(3).width = 22; ws.getColumn(4).width = 10;
    ws.getColumn(5).width = 16; ws.getColumn(6).width = 32;
    styleHeaderRow(ws);
    dataRows.forEach(row => {
      const newRow = ws.addRow([row[0], row[1], row[2], row[3], row[4], row[5]]);
      styleDataRow(newRow, ws.rowCount);
    });

    const buffer = await wb.xlsx.writeBuffer();
    fs.writeFileSync(fpath, buffer);
    await yadisk.uploadFile(buffer, remoteFilename(date));

    res.json({ success: true, message: `Переформатировано ${dataRows.length} строк за ${date}` });
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
