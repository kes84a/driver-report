'use strict';

require('dotenv').config();

const express = require('express');
const path    = require('path');
const fs      = require('fs');
const os      = require('os');
const XLSX    = require('xlsx');

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

// Append one record row to today.xlsx (create with headers if missing)
function appendRecord(record) {
  const filePath = todayPath();
  let rows;

  if (fs.existsSync(filePath)) {
    const wb = XLSX.readFile(filePath);
    rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1 });
  } else {
    rows = [['ФИО', 'Дата', 'Время', 'Количество коробов']];
  }

  rows.push([record.fio, record.date, record.time, record.box_count]);

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), 'Данные');
  XLSX.writeFile(wb, filePath);

  return Buffer.from(XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }));
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
    const buffer = appendRecord({ fio: fio.trim(), date, time, box_count: count });

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
