'use strict';

const nodemailer = require('nodemailer');
const XLSX = require('xlsx');
const fs = require('fs');

async function sendReport(filePath) {
  const required = ['SMTP_HOST', 'SMTP_PORT', 'SMTP_PASSWORD', 'EMAIL_SENDER', 'EMAIL_RECIPIENT'];
  const missing = required.filter(k => !process.env[k]);
  if (missing.length) throw new Error(`Не настроены параметры в .env: ${missing.join(', ')}`);

  const wb = XLSX.readFile(filePath);
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
  const dataRows = rows.slice(1); // skip header

  const now = new Date();
  const dd = String(now.getDate()).padStart(2, '0');
  const mm = String(now.getMonth() + 1).padStart(2, '0');
  const yyyy = now.getFullYear();
  const dateStr = `${dd}.${mm}.${yyyy}`;

  const totalBoxes = dataRows.reduce((s, r) => s + (parseInt(r[3]) || 0), 0);
  const driverLines = dataRows
    .map(r => `  • ${r[0]} — ${r[3]} коробов`)
    .join('\n');

  const text = [
    `Отчёт по водителям за ${dateStr}`,
    '',
    `Записей: ${dataRows.length}`,
    `Коробов итого: ${totalBoxes}`,
    '',
    'Водители:',
    driverLines || '  Нет данных',
  ].join('\n');

  const transporter = nodemailer.createTransport({
    host: process.env.SMTP_HOST,
    port: parseInt(process.env.SMTP_PORT),
    secure: parseInt(process.env.SMTP_PORT) === 465,
    auth: { user: process.env.EMAIL_SENDER, pass: process.env.SMTP_PASSWORD },
  });

  await transporter.sendMail({
    from: process.env.EMAIL_SENDER,
    to: process.env.EMAIL_RECIPIENT,
    subject: `Отчёт по водителям ${dateStr}`,
    text,
    attachments: [{
      filename: `Отчёт_водители_${dateStr}.xlsx`,
      content: fs.readFileSync(filePath),
    }],
  });

  console.log(`[Email] Отчёт за ${dateStr} → ${process.env.EMAIL_RECIPIENT} (${dataRows.length} записей)`);
  return { count: dataRows.length, totalBoxes };
}

module.exports = { sendReport };
