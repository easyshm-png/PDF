const SS = SpreadsheetApp.getActive();
const MAX_ACC = 150; // макс. погрешность GPS, м

function sh(n, h) {
  let s = SS.getSheetByName(n);
  if (!s) { s = SS.insertSheet(n); s.appendRow(h); s.setFrozenRows(1); }
  return s;
}
const S = {
  obj: () => sh('Объекты', ['Название', 'Широта', 'Долгота', 'Радиус_м']),
  emp: () => sh('Сотрудники', ['Телефон', 'Имя', 'Устройство', 'Активен']),
  log: () => sh('Отметки', ['Время', 'Телефон', 'Имя', 'Объект', 'Тип', 'Широта', 'Долгота', 'Точность_м', 'Дистанция_м']),
  shift: () => sh('Смены', ['Дата', 'Имя', 'Объект', 'Приход', 'Уход', 'Часы']),
};

function setup() { Object.values(S).forEach(f => f()); }

const out = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
const norm = t => String(t).replace(/\D/g, '').slice(-10);

function dist(a, b, c, d) {
  const r = Math.PI / 180, x = (c - a) * r, y = (d - b) * r;
  const h = Math.sin(x / 2) ** 2 + Math.cos(a * r) * Math.cos(c * r) * Math.sin(y / 2) ** 2;
  return 2 * 6371000 * Math.asin(Math.sqrt(h));
}

function lastMark(phone) {
  const v = S.log().getDataRange().getValues();
  for (let i = v.length - 1; i > 0; i--) if (norm(v[i][1]) === phone) return v[i];
  return null;
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const p = JSON.parse(e.postData.contents);
    const phone = norm(p.phone);
    const emp = S.emp().getDataRange().getValues();
    const i = emp.findIndex((r, k) => k > 0 && norm(r[0]) === phone);
    if (i < 0) return out({ ok: false, msg: 'Номер не найден. Обратитесь к руководителю' });
    const [, name, dev, active] = emp[i];
    if (active === false || String(active).toLowerCase() === 'нет') return out({ ok: false, msg: 'Доступ отключён' });
    if (!dev) S.emp().getRange(i + 1, 3).setValue(p.device);
    else if (dev !== p.device) return out({ ok: false, msg: 'Номер привязан к другому телефону. Обратитесь к руководителю' });

    const last = lastMark(phone);
    const onShift = last && last[4] === 'Пришёл';
    if (p.action === 'status') return out({ ok: true, name, onShift, obj: onShift ? last[3] : '', since: onShift ? last[0] : '' });

    if (p.action === 'in' && onShift) return out({ ok: false, msg: 'Вы уже на смене: ' + last[3] });
    if (p.action === 'out' && !onShift) return out({ ok: false, msg: 'Нет открытой смены' });
    if (!(p.acc <= MAX_ACC)) return out({ ok: false, msg: `Слабый GPS (±${Math.round(p.acc)} м). Выйдите на открытое место и повторите` });

    let best = null;
    S.obj().getDataRange().getValues().slice(1).filter(r => r[0] && r[1] && r[2]).forEach(o => {
      const d = dist(p.lat, p.lon, +o[1], +o[2]), rad = +o[3] || 150;
      if (!best || d - rad < best.d - best.rad) best = { name: o[0], d, rad };
    });
    if (!best || best.d > best.rad) return out({ ok: false, msg: best ? `Вы не на объекте. Ближайший: ${best.name}, ${Math.round(best.d)} м` : 'Объекты не настроены' });
    if (p.action === 'out' && best.name !== last[3]) return out({ ok: false, msg: `Смена открыта на объекте ${last[3]}, а вы на ${best.name}` });

    const now = new Date();
    S.log().appendRow([now, "'" + phone, name, best.name, p.action === 'in' ? 'Пришёл' : 'Ушёл', p.lat, p.lon, Math.round(p.acc), Math.round(best.d)]);
    if (p.action === 'out') {
      const start = new Date(last[0]);
      S.shift().appendRow([Utilities.formatDate(start, Session.getScriptTimeZone(), 'yyyy-MM-dd'), name, best.name, start, now, Math.round((now - start) / 36e5 * 100) / 100]);
    }
    return out({ ok: true, name, onShift: p.action === 'in', obj: best.name, since: now });
  } catch (err) {
    return out({ ok: false, msg: 'Ошибка: ' + err.message });
  } finally {
    lock.releaseLock();
  }
}
