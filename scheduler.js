'use strict';

const cron   = require('node-cron');
const fs     = require('fs');
const path   = require('path');
const mailer = require('./mailer');
const yadisk = require('./yadisk');

const DATA_DIR = path.join(__dirname, 'data');

function archiveToday() {
  const src = path.join(DATA_DIR, 'today.xlsx');
  if (!fs.existsSync(src)) return;

  const now = new Date();
  const dd = String(now.getDate()).padStart(2, '0');
  const mm = String(now.getMonth() + 1).padStart(2, '0');
  const yyyy = now.getFullYear();
  const dest = path.join(DATA_DIR, 'archive', `${yyyy}-${mm}-${dd}.xlsx`);

  fs.copyFileSync(src, dest);
  fs.unlinkSync(src);
  console.log(`[Архив] Сохранено: ${dest}`);
}

function startScheduler() {
  const sendTime = process.env.SEND_TIME || '18:00';
  const [h, m] = sendTime.split(':').map(Number);
  const hour   = isNaN(h) ? 18 : h;
  const minute = isNaN(m) ? 0  : m;

  cron.schedule(`${minute} ${hour} * * *`, async () => {
    const todayPath = path.join(DATA_DIR, 'today.xlsx');

    // Restore from YaDisk if missing locally (e.g. after cloud restart)
    if (!fs.existsSync(todayPath)) {
      try {
        const now = new Date();
        const dd = String(now.getDate()).padStart(2, '0');
        const mm = String(now.getMonth() + 1).padStart(2, '0');
        const yyyy = now.getFullYear();
        const filename = `Отчёт_${dd}.${mm}.${yyyy}.xlsx`;
        const buffer = await yadisk.downloadFile(filename);
        if (buffer) {
          fs.writeFileSync(todayPath, buffer);
          console.log('[Планировщик] Файл восстановлен с Яндекс Диска');
        }
      } catch (err) {
        console.error('[Планировщик] Не удалось скачать файл:', err.message);
      }
    }

    if (!fs.existsSync(todayPath)) {
      console.log('[Планировщик] Нет данных для отправки — пропуск');
      return;
    }

    console.log(`[Планировщик] Отправка отчёта в ${new Date().toLocaleString()}...`);
    try {
      await mailer.sendReport(todayPath);
      archiveToday();
    } catch (err) {
      console.error('[Планировщик] Ошибка:', err.message);
    }
  });

  const hh = String(hour).padStart(2, '0');
  const mn = String(minute).padStart(2, '0');
  console.log(`Планировщик: ежедневная отправка в ${hh}:${mn}`);
}

module.exports = { startScheduler, archiveToday };
