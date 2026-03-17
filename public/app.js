'use strict';

const form        = document.getElementById('form');
const errEl       = document.getElementById('err');
const okEl        = document.getElementById('ok');
const btn         = document.getElementById('btn');
const searchEl    = document.getElementById('search');
const suggestEl   = document.getElementById('suggestions');
const driverInfo  = document.getElementById('driver-info');
const driverName  = document.getElementById('driver-name');
const driverCar   = document.getElementById('driver-car');

let selectedDriver = null;
let searchTimer    = null;

// ── Driver search ─────────────────────────────────────────────────
searchEl.addEventListener('input', () => {
  clearTimeout(searchTimer);
  const q = searchEl.value.trim();
  selectedDriver = null;
  driverInfo.hidden = true;

  if (q.length < 1) { hideSuggestions(); return; }

  searchTimer = setTimeout(() => fetchDrivers(q), 200);
});

async function fetchDrivers(q) {
  try {
    const res  = await fetch('/api/driver?q=' + encodeURIComponent(q));
    const list = await res.json();

    if (!list.length) { hideSuggestions(); return; }

    suggestEl.innerHTML = '';
    list.forEach(d => {
      const item = document.createElement('div');
      item.className = 'suggest-item';
      item.innerHTML = '<strong>' + d.key + '</strong> — ' + d.fio +
                       '<span class="suggest-car">' + d.brand + ' ' + d.plate + '</span>';
      item.addEventListener('mousedown', e => { e.preventDefault(); selectDriver(d); });
      item.addEventListener('touchstart', e => { e.preventDefault(); selectDriver(d); });
      suggestEl.appendChild(item);
    });
    suggestEl.hidden = false;
  } catch { hideSuggestions(); }
}

function selectDriver(d) {
  selectedDriver = d;
  searchEl.value = d.key;
  hideSuggestions();
  driverName.textContent = d.fio;
  driverCar.textContent  = d.brand + ' · ' + d.plate;
  driverInfo.hidden = false;
  document.getElementById('box_count').focus();
}

function hideSuggestions() { suggestEl.hidden = true; }

document.addEventListener('click', e => {
  if (!e.target.closest('.field')) hideSuggestions();
});

// ── Form submit ───────────────────────────────────────────────────
form.addEventListener('submit', async e => {
  e.preventDefault();

  const box_count = document.getElementById('box_count').value;

  if (!selectedDriver) { showErr('Выберите водителя из списка'); return; }
  if (!box_count || +box_count < 1) { showErr('Введите количество коробов (не менее 1)'); return; }

  hideErr();
  btn.disabled    = true;
  btn.textContent = 'Сохранение…';

  try {
    const res  = await fetch('/api/submit', {
      method:  'POST',
      headers: { 'Content-Type': 'application/json' },
      body:    JSON.stringify({ driver_key: selectedDriver.key, box_count: +box_count }),
    });
    const data = await res.json();

    if (res.ok) {
      form.hidden = true;
      okEl.hidden = false;
      setTimeout(() => {
        form.hidden   = false;
        okEl.hidden   = true;
        form.reset();
        selectedDriver    = null;
        driverInfo.hidden = true;
        btn.disabled    = false;
        btn.textContent = 'Отправить';
      }, 3000);
    } else {
      showErr(data.error || 'Ошибка сервера');
      btn.disabled    = false;
      btn.textContent = 'Отправить';
    }
  } catch {
    showErr('Нет связи с сервером — проверьте подключение');
    btn.disabled    = false;
    btn.textContent = 'Отправить';
  }
});

function showErr(msg) { errEl.textContent = msg; errEl.hidden = false; }
function hideErr()     { errEl.hidden = true; }
