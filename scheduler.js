'use strict';

const cron   = require('node-cron');
const fs     = require('fs');
const path   = require('path');
const mailer = require('./mailer');
const yadisk = require('./yadisk');

const DATA_DIR = path.join(__dirname, 'data');

function serverDateDDMMYYYY() {
  const now = new Date();
  const dd   = String(now.getDate()).padStart(2, '0');
  const mm   = String(now.getMonth() + 1).padStart(2, '0');
  const yyyy = now.getFullYear();
  return `${dd}.${mm}.${yyyy}`;
}

function remoteFilename(date) {
  return `Отчёт_${date}.xlsx`;
}

function localPath(date) {
  return path.join(DATA_DIR, `report_${date}.xlsx`);
}

function startScheduler() {
  const sendTime = process.env.SEND_TIME || '18:00';
  const [h, m] = sendTime.split(':').map(Number);
  const hour   = isNaN(h) ? 18 : h;
  const minute = isNaN(m) ? 0  : m;

  cron.schedule(`${minute} ${hour} * * *`, async () => {
    const date   = serverDateDDMMYYYY();
    const fpath  = localPath(date);

    // Restore from YaDisk if missing locally (e.g. after cloud restart)
    if (!fs.existsSync(fpath)) {
      try {
        const buffer = await yadisk.downloadFile(remoteFilename(date));
        if (buffer) {
          fs.writeFileSync(fpath, buffer);
          console.log('[Планировщик] Файл восстановлен с Яндекс Диска');
        }
      } catch (err) {
        console.error('[Планировщик] Не удалось скачать файл:', err.message);
      }
    }

    if (!fs.existsSync(fpath)) {
      console.log('[Планировщик] Нет данных для отправки — пропуск');
      return;
    }

    console.log(`[Планировщик] Отправка отчёта в ${new Date().toLocaleString()}...`);
    try {
      await mailer.sendReport(fpath);
      fs.unlinkSync(fpath); // remove local file after sending (stays on YaDisk)
      console.log(`[Планировщик] Локальный файл удалён: ${fpath}`);
    } catch (err) {
      console.error('[Планировщик] Ошибка:', err.message);
    }
  });

  const hh = String(hour).padStart(2, '0');
  const mn = String(minute).padStart(2, '0');
  console.log(`Планировщик: ежедневная отправка в ${hh}:${mn}`);
}

module.exports = { startScheduler };
