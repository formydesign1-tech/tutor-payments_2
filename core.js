/* ---------- CORE: расчёты (то же, что формулы в Excel) ---------- */
const Y0 = 2026, NM = 24;                       // месяцы: янв 2026 — дек 2027
const MNAMES = ['Январь','Февраль','Март','Апрель','Май','Июнь','Июль','Август','Сентябрь','Октябрь','Ноябрь','Декабрь'];
const DAYS = ['Пн','Вт','Ср','Чт','Пт','Сб','Вс'];
const PAID_WORDS = ['sent','paid','completed'];

function dayNum(y, m, d) { return Math.floor(Date.UTC(y, m, d) / 864e5); }       // m: 0..11, d=0 даёт последний день прошлого месяца
function wd(day) { return (((day + 3) % 7) + 7) % 7 + 1; }                        // Пн=1 … Вс=7
function parseISO(s) { const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(s || ''); return m ? dayNum(+m[1], +m[2] - 1, +m[3]) : null; }
const MONTHS = Array.from({ length: NM }, (_, i) => {
  const y = Y0 + Math.floor(i / 12), m = i % 12;
  return { y, m, ms: dayNum(y, m, 1), me: dayNum(y, m + 1, 0), name: MNAMES[m], key: y + '-' + String(m + 1).padStart(2, '0') };
});
function monthIdx(day) {                                                         // индекс месяца в MONTHS; -1 раньше начала, NM-1 позже конца
  if (day < MONTHS[0].ms) return -1;
  for (let i = 0; i < NM; i++) if (day <= MONTHS[i].me) return i;
  return NM - 1;
}

function safeNum(v) {
  if (v === null || v === undefined || v === '') return null;
  if (typeof v === 'number') return isFinite(v) ? v : null;
  const s = String(v).replace(/[\s ]/g, '').replace(',', '.');
  const n = parseFloat(s);
  return isNaN(n) ? null : n;
}

function parseProduct(pr) {
  pr = String(pr || '');
  const t = /Тип занятий:\s*(.*?)\s*,\s*Месяц:\s*([^)]*)\)/.exec(pr);
  const q = /\)\s*-\s*(\d+)\s*x\s*([\d.,]+)/.exec(pr);
  const type = t ? t[1].trim() : '';
  const month = t ? t[2].trim() : '';
  const sizeM = /Абонемент\s*(\d+)/i.exec(type);
  const size = sizeM ? +sizeM[1] : 1;
  if (!q) return { type, month, lessons: 0, perLesson: null };
  const qty = +q[1], unit = safeNum(q[2]);
  return { type, month, lessons: qty * size, perLesson: unit === null ? null : unit / size };
}

function scheduleFlags(s) {
  const t = (s.actual && String(s.actual).trim()) ? String(s.actual) : String(s.days || '');
  const low = t.toLowerCase();
  return DAYS.map(d => (low.indexOf(d.toLowerCase()) >= 0 ? 1 : 0));
}

function scheduleCheck(s) {
  const nd = scheduleFlags(s).reduce((a, b) => a + b, 0);
  const wk = safeNum(s.perWeek);
  const hasDays = (s.actual && String(s.actual).trim()) || (s.days && String(s.days).trim());
  if (!s.start || !hasDays || wk === null) return 'Заполните начало, дни и занятий в неделю';
  if (nd === 0) return 'Не распознаны дни (пишите Пн, Вт, Ср, Чт, Пт, Сб, Вс)';
  if (nd > wk) return 'В расписании ' + nd + ' дн., занятий в неделю ' + wk + ': впишите реальные дни в колонку «Дни, в которые ходит»';
  if (nd < wk) return 'Дней в расписании меньше, чем занятий в неделю';
  return '';
}

function lessonsIn(s, fl, ms, meEff) {
  const st = parseISO(s.start);
  if (st === null) return 0;
  const S = Math.max(ms, st);
  const nd = fl.reduce((a, b) => a + b, 0);
  const wk = safeNum(s.perWeek);
  if (nd > 0) {
    let cal = 0;
    fl.forEach((on, i) => {
      if (!on) return;
      const f = S + ((((i + 1) - wd(S)) % 7) + 7) % 7;
      if (f <= meEff) cal += Math.floor((meEff - f) / 7) + 1;
    });
    return Math.round(cal * ((wk !== null && wk < nd) ? wk / nd : 1));
  }
  return wk !== null ? Math.round(wk * Math.max(0, meEff - S + 1) / 7) : 0;
}

function isPaidRow(r) { return !!r.ok; }

function compute(students, payRows, repDay, toDate) {
  const repIdx = monthIdx(repDay);
  // --- платежи по ученикам
  const byName = new Map();
  const get = n => { if (!byName.has(n)) byName.set(n, { TL: 0, TP: 0, promo: '', pCnt: 0, listSum: 0, discSum: 0, firstPromo: '', n: 0 }); return byName.get(n); };
  const received = new Array(NM).fill(0), payCnt = new Array(NM).fill(0);
  payRows.forEach(r => {
    const g = get(r.n);
    const pp = parseProduct(r.pr);
    const am = safeNum(r.am) || 0;
    // скидка по промокоду (как «Скидка %» в Excel): считаем по всем платежам с промокодом
    if (r.pc) {
      const list = safeNum(r.pm) !== null ? safeNum(r.pm) : (pp.perLesson !== null ? pp.lessons * pp.perLesson : 0);
      const disc = safeNum(r.di) !== null ? safeNum(r.di) : (list ? list - am : 0);
      g.listSum += list || 0; g.discSum += disc || 0; g.pCnt += 1; if (!g.firstPromo) g.firstPromo = r.pc;
    }
    if (isPaidRow(r)) {
      g.TL += pp.lessons; g.TP += am; g.n += 1;
      const di = parseISO(String(r.c).slice(0, 10));
      if (di !== null) { const mi = monthIdx(di); if (mi >= 0 && di >= MONTHS[0].ms && di <= MONTHS[NM - 1].me) { received[mi] += am; payCnt[mi] += 1; } }
    }
  });
  // --- ученики
  const totals = {
    planL: new Array(NM).fill(0), planR: new Array(NM).fill(0), factL: new Array(NM).fill(0), factR: new Array(NM).fill(0),
    owedL: new Array(NM).fill(0), owedR: new Array(NM).fill(0), received, payCnt,
  };
  const rows = students.map(s => {
    const g = byName.get(s.name) || { TL: 0, TP: 0, promo: '', pCnt: 0, listSum: 0, discSum: 0, firstPromo: '', n: 0 };
    const active = (s.status || 'Активен') === 'Активен';
    const fl = scheduleFlags(s);
    let dpct = 0, promo = '';
    if (g.firstPromo) {
      const raw = g.listSum > 0 && g.discSum > 0 ? Math.round(g.discSum / g.listSum * 100) : 0;
      dpct = raw >= 1 ? raw : 0;
      promo = g.firstPromo + (g.listSum > 0 && g.discSum > 0 ? (raw >= 1 ? ' (-' + raw + '%)' : ' (-' + Math.round(g.discSum / g.pCnt) + ' ₽)') : '');
    }
    const price = safeNum(s.price) || 0;
    const exp = new Array(NM).fill(0), planR = new Array(NM).fill(0);
    for (let m = 0; m < NM; m++) {
      if (!active) continue;
      const mo = MONTHS[m];
      const meEff = toDate ? Math.min(mo.me, repDay) : mo.me;
      exp[m] = lessonsIn(s, fl, mo.ms, meEff);
      planR[m] = Math.round(exp[m] * price * (1 - dpct / 100));
    }
    // «водопад»: оплаченные занятия закрывают месяцы по порядку
    const cov = new Array(NM).fill(0), factR = new Array(NM).fill(0), owedR = new Array(NM).fill(0);
    let prior = 0, planSum = 0, factSum = 0;
    for (let m = 0; m < NM; m++) {
      cov[m] = Math.min(exp[m], Math.max(0, g.TL - prior));
      prior += exp[m];
      factR[m] = g.TL === 0 ? 0 : cov[m] / g.TL * g.TP;
      if (m <= repIdx && active && exp[m] > 0) owedR[m] = Math.round(planR[m] * Math.max(0, exp[m] - cov[m]) / exp[m]);
      totals.planL[m] += (m <= repIdx ? exp[m] : 0);
      totals.planR[m] += (m <= repIdx ? planR[m] : 0);
      totals.factL[m] += cov[m];
      totals.factR[m] += factR[m];
      totals.owedL[m] += (m <= repIdx && active ? Math.max(0, exp[m] - cov[m]) : 0);
      totals.owedR[m] += owedR[m];
      if (m <= repIdx) { planSum += planR[m]; }
      factSum += factR[m];
    }
    const repMonthFull = repIdx >= 0 && active ? lessonsIn(s, fl, MONTHS[repIdx].ms, MONTHS[repIdx].me) : 0;
    const repMonthDone = repIdx >= 0 && active ? lessonsIn(s, fl, MONTHS[repIdx].ms, Math.min(MONTHS[repIdx].me, repDay)) : 0;
    return {
      s, active, exp, planR, cov, factR, owedR, planSum, paid: g.TP, paidLessons: g.TL, balance: g.TP - planSum,
      owedTotal: owedR.reduce((a, b) => a + b, 0), credit: g.TP - factSum, promo, dpct, repMonthFull, repMonthDone, check: scheduleCheck(s), nPay: g.n,
    };
  });
  const sum = a => a.reduce((x, y) => x + y, 0);
  const upto = a => sum(a.slice(0, Math.max(0, repIdx + 1)));
  const recvAll = sum(received);
  const sumFactAll = sum(totals.factR);
  const factUpTo = upto(totals.factR);
  // платежи от учеников, которых нет в списке
  const known = new Set(students.map(s => s.name));
  const unknownPay = [...byName.keys()].filter(n => !known.has(n));
  return {
    repIdx, rows, totals,
    kpi: {
      received: recvAll, plan: upto(totals.planR), fact: factUpTo, owed: upto(totals.owedR),
      prepaid: recvAll - factUpTo, payments: sum(payCnt), unknownPay,
    },
  };
}

/* ---------- разбор выгрузки Tilda ---------- */
const TILDA_COLS = ['created','name','email','phone','checkbox','checkbox_2','checkbox_3','ma_name','ma_email','ma_id','product','amount','prodamount','tilda_orderid','tilda_order_currency','tilda_status','formid','formname','referer','ma_phone','paymentid','promocode','discount','subtotal'];

function normDate(v) {
  if (v === null || v === undefined || v === '') return '';
  if (v instanceof Date) { const p = n => String(n).padStart(2, '0'); return v.getFullYear() + '-' + p(v.getMonth() + 1) + '-' + p(v.getDate()) + ' ' + p(v.getHours()) + ':' + p(v.getMinutes()) + ':' + p(v.getSeconds()); }
  if (typeof v === 'number') { const ms = Math.round((v - 25569) * 864e5); const d = new Date(ms); const p = n => String(n).padStart(2, '0'); return d.getUTCFullYear() + '-' + p(d.getUTCMonth() + 1) + '-' + p(d.getUTCDate()) + ' ' + p(d.getUTCHours()) + ':' + p(d.getUTCMinutes()) + ':' + p(d.getUTCSeconds()); }
  const s = String(v).trim();
  let m = /^(\d{4}-\d{2}-\d{2})[ T](\d{2}:\d{2}(?::\d{2})?)/.exec(s);
  if (m) return m[1] + ' ' + (m[2].length === 5 ? m[2] + ':00' : m[2]);
  m = /^(\d{2})\.(\d{2})\.(\d{4})(?:[ T](\d{2}:\d{2}(?::\d{2})?))?/.exec(s);
  if (m) return m[3] + '-' + m[2] + '-' + m[1] + ' ' + (m[4] ? (m[4].length === 5 ? m[4] + ':00' : m[4]) : '00:00:00');
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s + ' 00:00:00';
  return s;
}

function rowsFromMatrix(matrix) {
  // matrix: массив строк (массивы ячеек). Находим шапку; если её нет, берём стандартный порядок колонок Tilda
  let hi = -1;
  for (let i = 0; i < Math.min(matrix.length, 15); i++) {
    const low = matrix[i].map(c => String(c === null || c === undefined ? '' : c).trim().toLowerCase());
    if (low.indexOf('created') >= 0 && low.indexOf('product') >= 0) { hi = i; break; }
  }
  let cols, body;
  if (hi >= 0) { cols = matrix[hi].map(c => String(c === null || c === undefined ? '' : c).trim().toLowerCase()); body = matrix.slice(hi + 1); }
  else { cols = TILDA_COLS; body = matrix; }
  const idx = {}; cols.forEach((c, i) => { if (c && !(c in idx)) idx[c] = i; });
  const need = ['created', 'name', 'product', 'amount', 'paymentid'];
  const missing = need.filter(k => !(k in idx));
  if (missing.length) return { error: 'В файле нет колонок: ' + missing.join(', ') };
  const out = [];
  let skipped = 0;
  body.forEach(r => {
    const g = k => { const v = r[idx[k]]; return (v === undefined || v === null) ? '' : v; };
    if (!String(g('created')).trim() && !String(g('name')).trim()) return;
    const name = String(g('name')).trim();
    if (!name) return;
    // платёж прошёл, только если заполнен paymentid; остальные строки (брошенные корзины, неудачные оплаты) пропускаем
    const pid = String(g('paymentid')).replace('custom.atolpay:', '').trim();
    if (!pid) { skipped++; return; }
    const created = normDate(g('created'));
    const pr = String(g('product'));
    const am = safeNum(g('amount'));
    const id = pid;
    const maName = String(g('ma_name')).trim();
    out.push({
      id, c: created, n: name, pa: (maName && maName !== name) ? maName : '', em: String(g('email')).trim(), ph: String(g('phone')).trim(),
      pr, am, pm: safeNum(g('prodamount')), di: safeNum(g('discount')), pc: String(g('promocode')).trim(), ok: 1,
    });
  });
  return { rows: out, skipped };
}

function parseCSV(text) {
  text = text.replace(/^﻿/, '');
  const first = text.split(/\r?\n/, 1)[0] || '';
  const cnt = ch => first.split(ch).length - 1;
  const delim = [';', '\t', ','].sort((a, b) => cnt(b) - cnt(a))[0];
  const rows = []; let row = [], cell = '', q = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (q) { if (c === '"') { if (text[i + 1] === '"') { cell += '"'; i++; } else q = false; } else cell += c; }
    else if (c === '"') q = true;
    else if (c === delim) { row.push(cell); cell = ''; }
    else if (c === '\n') { row.push(cell); rows.push(row); row = []; cell = ''; }
    else if (c === '\r') { /* skip */ }
    else cell += c;
  }
  if (cell !== '' || row.length) { row.push(cell); rows.push(row); }
  return rows;
}

/* ---------- импорт выгрузки в базу и проверка данных ---------- */
const str = (v, max) => String(v === null || v === undefined ? '' : v).slice(0, max || 500);
function sanitizeRow(r) {
  if (!r || typeof r !== 'object') return null;
  const n = str(r.n, 200).trim();
  if (!n) return null;
  const id = str(r.id, 200).trim();
  if (!id) return null;
  return {
    id, c: str(r.c, 40), n, pa: str(r.pa, 200), em: str(r.em, 200), ph: str(r.ph, 60), pr: str(r.pr, 1000),
    am: safeNum(r.am), pm: safeNum(r.pm), di: safeNum(r.di), pc: str(r.pc, 100), ok: r.ok ? 1 : 0,
  };
}
const STATUSES = ['Активен', 'Пауза', 'Завершил'];
function sanitizeStudent(s, newId) {
  if (!s || typeof s !== 'object') return null;
  const name = str(s.name, 200).trim();
  if (!name) return null;
  const start = /^\d{4}-\d{2}-\d{2}$/.test(str(s.start, 10)) ? str(s.start, 10) : '';
  const wk = safeNum(s.perWeek), pr = safeNum(s.price);
  return {
    id: str(s.id, 40) || newId(), name, exam: str(s.exam, 60), start, days: str(s.days, 200), actual: str(s.actual, 200),
    perWeek: wk === null ? '' : wk, price: pr === null ? 1300 : pr, status: STATUSES.includes(s.status) ? s.status : 'Активен',
    comment: str(s.comment, 500), email: str(s.email, 200), phone: str(s.phone, 60),
    tg: normNick(s.tg), tgName: str(s.tgName, 200).trim(),
  };
}

function applyImport(db, rows, newId) {
  const byId = new Map(db.payments.map((r, i) => [r.id, i]));
  let added = 0, updated = 0, same = 0;
  for (const raw of rows) {
    const r = sanitizeRow(raw); if (!r) continue;
    const i = byId.get(r.id);
    if (i === undefined) { db.payments.push(r); byId.set(r.id, db.payments.length - 1); added++; }
    else if (JSON.stringify(db.payments[i]) !== JSON.stringify(r)) { db.payments[i] = r; updated++; }
    else same++;
  }
  const have = new Set(db.students.map(s => s.name));
  const fresh = [];
  for (const raw of rows) {
    const r = sanitizeRow(raw); if (!r || have.has(r.n)) continue;
    have.add(r.n);
    const pp = parseProduct(r.pr);
    db.students.push({ id: newId(), name: r.n, exam: '', start: '', days: '', actual: '', perWeek: '', price: pp.perLesson ? Math.round(pp.perLesson) : 1300, status: 'Активен', comment: '', email: r.em, phone: r.ph });
    fresh.push(r.n);
  }
  return { added, updated, same, fresh, notPaid: rows.filter(r => r && !r.ok).length };
}

/* ---------- Telegram-ники и выгрузка должников для бота ---------- */
const normName = v => String(v === null || v === undefined ? '' : v).toLowerCase().replace(/ё/g, 'е').replace(/[\s ]+/g, ' ').trim();
function normNick(v) {                       // убираем @, ссылку t.me/, пробелы по краям
  let n = String(v === null || v === undefined ? '' : v).trim();
  n = n.replace(/^https?:\/\/(www\.)?(t\.me|telegram\.me)\//i, '').replace(/^(t\.me|telegram\.me)\//i, '').replace(/^@+/, '').replace(/\/+$/, '');
  return n.slice(0, 64);
}
const isNick = n => /^[A-Za-z][A-Za-z0-9_]{4,31}$/.test(n);   // правила Telegram: латиница, цифры, _; 5–32 символа

/** Список учеников (xlsx/csv с колонками «ФИО ученика», «Ник», «Имя») → ники и имена в карточки учеников. */
function matchRoster(students, matrix) {
  const hd = c => String(c === null || c === undefined ? '' : c).toLowerCase().replace(/[\s ]+/g, ' ').trim();
  let hi = -1, iN = -1, iNick = -1, iName = -1;
  for (let i = 0; i < Math.min(matrix.length, 20); i++) {
    const low = matrix[i].map(hd);
    const a = low.findIndex(x => x.startsWith('фио ученика') || x === 'ученик' || x === 'фио');
    const b = low.indexOf('ник');
    if (a >= 0 && b >= 0) { hi = i; iN = a; iNick = b; iName = low.indexOf('имя'); break; }
  }
  if (hi < 0) return { error: 'Не нашла колонки «ФИО ученика» и «Ник». Загрузи список учеников с такими заголовками.' };
  const byName = new Map();
  students.forEach(s => { const k = normName(s.name); if (!byName.has(k)) byName.set(k, []); byName.get(k).push(s); });
  const out = { changed: 0, same: 0, matched: 0, unmatched: [], noNick: [] };
  const seen = new Set();
  matrix.slice(hi + 1).forEach(r => {
    const full = String(r[iN] === undefined || r[iN] === null ? '' : r[iN]).trim();
    if (!full) return;
    const list = byName.get(normName(full));
    if (!list) { out.unmatched.push(full); return; }
    const nick = normNick(r[iNick]);
    const nm = iName >= 0 ? String(r[iName] === undefined || r[iName] === null ? '' : r[iName]).trim().slice(0, 200) : '';
    list.forEach(s => {
      seen.add(s.id); out.matched++;
      let ch = false;
      if (nick && s.tg !== nick) { s.tg = nick; ch = true; }
      if (nm && s.tgName !== nm) { s.tgName = nm; ch = true; }
      if (ch) out.changed++; else out.same++;
      if (!nick && !s.tg) out.noNick.push(s.name);
    });
  });
  out.notInFile = students.filter(s => !seen.has(s.id)).map(s => s.name);
  return out;
}

/** Таблица для бота: строка = ник + имя + месяц + сумма недоплаты. Только активные ученики с долгом. */
function debtRows(M) {
  const rows = [], skipped = [];
  M.rows.forEach(r => {
    if (!r.active) return;
    const items = [];
    for (let m = 0; m <= M.repIdx && m < NM; m++) if (r.owedR[m] > 0) items.push({ m, sum: r.owedR[m] });
    if (!items.length) return;
    const nick = normNick(r.s.tg);
    const total = items.reduce((a, x) => a + x.sum, 0);
    if (!nick) { skipped.push({ name: r.s.name, total, why: 'нет ника' }); return; }
    if (!isNick(nick)) { skipped.push({ name: r.s.name, total, why: 'ник «' + nick + '» не похож на Telegram-ник (нужна латиница, цифры и _, от 5 символов)' }); return; }
    items.forEach(x => rows.push({ nick, name: (r.s.tgName || r.s.name).trim(), month: MONTHS[x.m].name.toLowerCase(), sum: x.sum, student: r.s.name }));
  });
  return { rows, skipped, total: rows.reduce((a, x) => a + x.sum, 0) };
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = { matchRoster, debtRows, normNick, normName, isNick, compute, parseProduct, rowsFromMatrix, parseCSV, applyImport, sanitizeRow, sanitizeStudent, dayNum, monthIdx, MONTHS, scheduleCheck, normDate, safeNum, STATUSES };
}
