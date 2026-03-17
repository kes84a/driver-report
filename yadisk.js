'use strict';

const BASE = 'https://cloud-api.yandex.net/v1/disk/resources';

function token() {
  return process.env.YADISK_TOKEN;
}

function folder() {
  return process.env.YADISK_FOLDER || 'Отчёты водителей';
}

function authHeader() {
  return { Authorization: `OAuth ${token()}` };
}

// Create folder on YaDisk (ignore if already exists)
async function ensureFolder() {
  if (!token() || token() === 'your_yadisk_token_here') return;

  const res = await fetch(`${BASE}?path=${encodeURIComponent(folder())}`, {
    method: 'PUT',
    headers: authHeader(),
  });

  // 201 = created, 409 = already exists — both OK
  if (res.status !== 201 && res.status !== 409) {
    const body = await res.text();
    console.error('[YaDisk] Ошибка создания папки:', res.status, body);
  }
}

// Upload buffer as a file to YaDisk folder
async function uploadFile(buffer, filename) {
  if (!token() || token() === 'your_yadisk_token_here') {
    console.log('[YaDisk] Токен не задан — загрузка пропущена');
    return;
  }

  const remotePath = `${folder()}/${filename}`;

  // Step 1: get upload URL
  const urlRes = await fetch(
    `${BASE}/upload?path=${encodeURIComponent(remotePath)}&overwrite=true`,
    { headers: authHeader() }
  );

  if (!urlRes.ok) {
    const body = await urlRes.text();
    throw new Error(`YaDisk get-upload-url ${urlRes.status}: ${body}`);
  }

  const { href } = await urlRes.json();

  // Step 2: upload
  const uploadRes = await fetch(href, {
    method: 'PUT',
    body: buffer,
    headers: { 'Content-Type': 'application/octet-stream' },
  });

  if (!uploadRes.ok) {
    throw new Error(`YaDisk upload ${uploadRes.status}`);
  }

  console.log(`[YaDisk] Загружено: ${remotePath}`);
}

// Download a file from YaDisk folder, returns Buffer or null if not found
async function downloadFile(filename) {
  if (!token() || token() === 'your_yadisk_token_here') return null;

  const remotePath = `${folder()}/${filename}`;
  const res = await fetch(
    `${BASE}?path=${encodeURIComponent(remotePath)}&fields=file`,
    { headers: authHeader() }
  );

  if (!res.ok) return null;

  const { file: downloadUrl } = await res.json();
  if (!downloadUrl) return null;

  const fileRes = await fetch(downloadUrl);
  if (!fileRes.ok) return null;

  return Buffer.from(await fileRes.arrayBuffer());
}

module.exports = { ensureFolder, uploadFile, downloadFile };
