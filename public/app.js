'use strict';

// ── Set default date / time ──────────────────────────────────────
function setDefaults() {
  const now = new Date();
  document.getElementById('date').value = now.toISOString().split('T')[0];
  document.getElementById('time').value =
    String(now.getHours()).padStart(2, '0') + ':' +
    String(now.getMinutes()).padStart(2, '0');
}
setDefaults();

// ── Form submit ──────────────────────────────────────────────────
const form  = document.getElementById('form');
const errEl = document.getElementById('err');
const okEl  = document.getElementById('ok');
const btn   = document.getElementById('btn');

form.addEventListener('submit', async e => {
  e.preventDefault();

  const fio       = document.getElementById('fio').value.trim();
  const date      = document.getElementById('date').value;
  const time      = document.getElementById('time').value;
  const box_count = document.getElementById('box_count').value;

  if (!fio)        { showErr('Введите ФИО'); return; }
  if (!date)       { showErr('Укажите дату'); return; }
  if (!time)       { showErr('Укажите время'); return; }
  if (!box_count || +box_count < 1) { showErr('Введите количество коробов (не менее 1)'); return; }

  hideErr();
  btn.disabled = true;
  btn.textContent = 'Сохранение…';

  try {
    const res  = await fetch('/api/submit', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ fio, date, time, box_count: +box_count }),
    });
    const data = await res.json();

    if (res.ok) {
      form.hidden = true;
      okEl.hidden = false;
      setTimeout(() => {
        okEl.hidden  = false;
        form.hidden  = false;
        okEl.hidden  = true;
        form.reset();
        setDefaults();
        btn.disabled = false;
        btn.textContent = 'Отправить';
      }, 3000);
    } else {
      showErr(data.error || 'Ошибка сервера');
      btn.disabled = false;
      btn.textContent = 'Отправить';
    }
  } catch {
    showErr('Нет связи с сервером — проверьте подключение');
    btn.disabled = false;
    btn.textContent = 'Отправить';
  }
});

function showErr(msg) { errEl.textContent = msg; errEl.hidden = false; }
function hideErr()     { errEl.hidden = true; }
