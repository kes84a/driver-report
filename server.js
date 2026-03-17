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

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Данные');

  ws.columns = [
    { key: 'fio',       width: 30 },
    { key: 'date',      width: 14 },
    { key: 'time',      width: 10 },
    { key: 'box_count', width: 22 },
  ];

  // Header row added manually (avoids exceljs sheetNo bug with header property)
  ws.addRow(['ФИО', 'Дата', 'Время', 'Количество коробов']);
  styleHeaderRow(ws);

  // Load existing data rows (skip header)
  if (fs.existsSync(filePath)) {
    const existing = new ExcelJS.Workbook();
    await existing.xlsx.readFile(filePath);
    const existingWs = existing.getWorksheet(1);
    if (existingWs) {
      existingWs.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return;
        const newRow = ws.addRow([
          row.getCell(1).value,
          row.getCell(2).value,
          row.getCell(3).value,
          row.getCell(4).value,
        ]);
        styleDataRow(newRow, ws.rowCount);
      });
    }
  }

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

// Driver form submission
app.post('/api/submit', async (req, res) => {
  const { fio, date, time, box_count } = req.body;

  if (!fio || !date || !time || !box_count) {
    return res.status(400).json({ error: 'Заполните все поля' });
  }
  const count = parseInt(box_count);
  if (isNaN(count) || count < 1) {
    return res.status(400).json({ error: 'Некорректное количество коробов' });
  }

  try {
    await ensureTodayFile();
    const buffer = await appendRecord({ fio: fio.trim(), date, time, box_count: count });

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
