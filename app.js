/* ---------- состояние и обмен с сервером ---------- */
const $ = id => document.getElementById(id);
const rub = n => Math.round(n).toLocaleString('ru-RU') + '\u00a0₽';
const esc = s => String(s === null || s === undefined ? '' : s).replace(/[&<>"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const pad = n => String(n).padStart(2, '0');
const todayISO = () => { const d = new Date(); return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()); };
const fmtDate = iso => { const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(iso || ''); return m ? m[3] + '.' + m[2] + '.' + m[1] : (iso || ''); };
const uid = () => 's' + Math.random().toString(36).slice(2, 8);

const S = { students: [], payments: [], reminders: {}, settings: {}, debtAll: false, tab: 'summary', date: todayISO(), toDate: true, ready: false,
            filter: 'all', q: '', pf: 'all', pq: '', sortK: 'owedTotal', sortDir: -1, showAllPays: false, openChg: new Set() };
try { const m = localStorage.getItem('uo-mode'); if (m === '0') S.toDate = false; } catch (e) {}

let rosterDirty = false, rosterTimer = null, saving = 0;

function setStatus(text, err) { const el = $('status'); el.textContent = text || ''; el.className = 'status' + (err ? ' err' : ''); }

async function api(method, url, body) {
  const opt = { method, headers: {}, credentials: 'same-origin' };
  if (body !== undefined) { opt.headers['Content-Type'] = 'application/json'; opt.body = JSON.stringify(body); }
  const res = await fetch(url, opt);
  if (res.status === 401 && url !== '/api/login') { showLogin(); throw new Error('Нужно войти'); }
  let data = null; try { data = await res.json(); } catch (e) {}
  if (!res.ok) throw new Error((data && data.error) || ('Ошибка ' + res.status));
  return data;
}

async function loadState() {
  const d = await api('GET', '/api/state');
  S.students = d.students || []; S.payments = d.payments || []; S.reminders = d.reminders || {}; S.settings = d.settings || {}; S.ready = true;
}

async function saveRoster() {
  saving++; setStatus('Сохраняю…');
  try { await api('PUT', '/api/students', { students: S.students }); rosterDirty = false; setStatus('Сохранено'); }
  catch (e) { setStatus('Не удалось сохранить: ' + e.message, true); }
  finally { saving--; }
}
function scheduleRoster() { rosterDirty = true; clearTimeout(rosterTimer); setStatus('Есть несохранённые изменения…'); rosterTimer = setTimeout(saveRoster, 700); }
window.addEventListener('beforeunload', e => { if (rosterDirty) { e.preventDefault(); e.returnValue = ''; } });

/* ---------- вычисления для экрана ---------- */
const allPays = () => S.payments;
function model() { return compute(S.students, allPays(), parseISO(S.date) === null ? dayNum(2026, 8, 1) : parseISO(S.date), S.toDate); }

/* ---------- загрузка выгрузки ---------- */
async function readFile(file) {
  const name = (file.name || '').toLowerCase();
  if (/\.(xlsx|xls)$/.test(name)) {
    if (!window.XLSX) throw new Error('Не загрузилась библиотека для чтения Excel. Сохрани выгрузку как CSV и загрузи снова.');
    const wb = XLSX.read(await file.arrayBuffer(), { type: 'array', cellDates: false });
    const ws = wb.Sheets[wb.SheetNames[0]];
    return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '' });
  }
  const buf = await file.arrayBuffer();
  let text = new TextDecoder('utf-8').decode(buf);
  if (text.indexOf(String.fromCharCode(65533)) >= 0) text = new TextDecoder('windows-1251').decode(buf);
  return parseCSV(text);
}

async function handleFile(file) {
  if (!file) return;
  $('notice').innerHTML = '';
  try {
    const res = rowsFromMatrix(await readFile(file));
    if (res.error) throw new Error(res.error + '. Загрузи выгрузку заказов из Tilda целиком, вместе со строкой заголовков.');
    if (!res.rows.length) throw new Error('В файле нет строк с заполненным paymentid' + (res.skipped ? ' (пропущено строк без него: ' + res.skipped + ')' : '') + '. Берутся только прошедшие платежи.');
    if (rosterDirty) { clearTimeout(rosterTimer); await saveRoster(); }
    const m = await api('POST', '/api/payments/import', { rows: res.rows });
    await loadState();
    let html = '<div class="banner ok"><b>Выгрузка загружена</b><span>Добавлено новых платежей: ' + m.added + ', обновлено: ' + m.updated + ', уже были: ' + m.same + (res.skipped ? '. Пропущено строк без paymentid (оплата не прошла): ' + res.skipped : '') + '.</span>';
    if (m.fresh.length) html += '<span>Новые ученики (' + m.fresh.length + '): ' + m.fresh.map(esc).join(', ') + '. Заполни им начало занятий, дни и занятий в неделю на вкладке «Ученики», иначе план по ним не посчитается.</span>';
    html += '</div>';
    $('notice').innerHTML = html;
    if (m.fresh.length) S.tab = 'students';
    syncTabs(); render();
  } catch (e) {
    $('notice').innerHTML = '<div class="banner err"><b>Не получилось загрузить файл</b><span>' + esc(e && e.message ? e.message : e) + '</span></div>';
  }
}

/* ---------- вкладки ---------- */
function syncTabs() { document.querySelectorAll('[role=tab]').forEach(t => t.setAttribute('aria-selected', t.dataset.tab === S.tab ? 'true' : 'false')); }

function render() {
  const view = $('view');
  renderRemind();
  const n = S.students.length, p = allPays().length;
  $('asof').textContent = n || p ? ('Учеников: ' + n + ' · платежей: ' + p) : '';
  if (!S.ready) { view.innerHTML = '<div class="card"><span class="status">Загружаю данные…</span></div>'; return; }
  if (!n && !p) {
    view.innerHTML = '<div class="drop card"><h2>Пока нет данных</h2><div>Выгрузи заказы из Tilda (Excel или CSV) и перетащи файл сюда или нажми кнопку. Из платежей создастся список учеников.</div><button class="btn" type="button" id="empty-up">Загрузить выгрузку Tilda</button></div>';
    const b = $('empty-up'); if (b) b.onclick = () => $('file').click();
    return;
  }
  if (S.tab === 'summary') renderSummary(view); else if (S.tab === 'students') renderStudents(view); else renderPays(view);
}

/* ---------- Сводка ---------- */
const SER = [{ k: 'plan', n: 'План', c: 'var(--s-plan)' }, { k: 'fact', n: 'Факт', c: 'var(--s-fact)' }, { k: 'owed', n: 'Нужно доплатить', c: 'var(--s-owed)' }];

function niceStep(max) { const steps = [1e3, 2e3, 5e3, 1e4, 2e4, 5e4, 1e5, 2e5, 5e5]; return steps.find(s => max / s <= 6) || 1e6; }

function renderSummary(view) {
  const M = model(), K = M.kpi, T = M.totals, ri = M.repIdx;
  const monthsShown = [];
  for (let m = 0; m < NM; m++) {
    const past = m <= ri;
    if ((past && (T.planR[m] > 0 || T.received[m] > 0 || T.factR[m] > 0)) || (!past && T.received[m] > 0)) monthsShown.push(m);
  }
  const notes = [];
  const badSched = M.rows.filter(r => r.active && r.check);
  if (badSched.length) notes.push('Расписание не заполнено или противоречиво у ' + badSched.length + ' ' + (badSched.length === 1 ? 'ученика' : 'учеников') + ': ' + badSched.slice(0, 6).map(r => esc(r.s.name)).join(', ') + (badSched.length > 6 ? ' и др.' : '') + '. Исправь на вкладке «Ученики».');
  const unpaid = allPays().filter(r => !r.ok).length;
  if (unpaid) notes.push('Платежей без оплаты в Tilda: ' + unpaid + '. В суммы они не входят (вкладка «Платежи»).');
  if (K.unknownPay.length) notes.push('Есть платежи от учеников не из списка: ' + K.unknownPay.map(esc).join(', ') + '.');
  let h = '';
  if (notes.length) h += '<div class="banner warn">' + notes.map(t => '<span>' + t + '</span>').join('') + '</div>';
  h += '<section class="kpis">' + [
    ['Поступило оплат', rub(K.received), 'всего за все месяцы', ''],
    ['План по расписанию', rub(K.plan), ri >= 0 ? 'до месяца отчёта включительно' : '', 'var(--s-plan)'],
    ['Факт по месяцам', rub(K.fact), 'оплачено за эти месяцы', 'var(--s-fact)'],
    ['Нужно доплатить', rub(K.owed), 'оплата вперёд ещё ' + rub(Math.max(0, K.prepaid)), 'var(--s-owed)'],
  ].map(([l, v, n, c], i) => '<div class="kpi' + (i === 3 ? ' owed' : '') + '"><div class="l">' + (c ? '<span class="sw" style="background:' + c + '"></span>' : '') + l + '</div><div class="v">' + v + '</div><div class="n">' + n + '</div></div>').join('') + '</section>';

  h += debtCardHTML(M);
  const cm = monthsShown.filter(m => m <= ri).slice(-6);
  if (cm.length) {
    h += '<section class="card"><div class="chart-head"><h2>План, факт и доплата по месяцам, ₽</h2><div class="legend">' + SER.map(s => '<span><i style="background:' + s.c + '"></i>' + s.n + '</span>').join('') + '</div></div><div id="chart"></div>' +
      '<p class="note" style="margin:8px 0 0">План: сколько должно быть оплачено по расписанию. Факт: сколько из оплаченных денег пришлось на этот месяц (оплата вперёд закрывает следующие месяцы). Доплата: сумма недоплат учеников по неоплаченным занятиям.</p></section>';
  }
  // таблица по месяцам
  const sum = a => a.reduce((x, y) => x + y, 0);
  h += '<section class="card"><h2 style="margin-bottom:8px">По месяцам</h2><div class="scroll"><table><thead><tr><th>Месяц</th><th>Платежей</th><th>Поступило</th><th>План</th><th>Факт</th><th>Факт − план</th><th>Занятий оплачено</th><th>Нужно доплатить</th></tr></thead><tbody>';
  monthsShown.forEach(m => {
    const past = m <= ri, mo = MONTHS[m];
    h += '<tr><td>' + mo.name + ' ' + mo.y + '</td><td class="num">' + T.payCnt[m] + '</td><td class="num">' + rub(T.received[m]) + '</td>' +
      (past ? '<td class="num">' + rub(T.planR[m]) + '</td><td class="num">' + rub(T.factR[m]) + '</td><td class="num">' + rub(T.factR[m] - T.planR[m]) + '</td><td class="num">' + Math.round(T.factL[m]) + ' из ' + T.planL[m] + '</td><td class="num"><b>' + rub(T.owedR[m]) + '</b></td>'
            : '<td class="zero">—</td><td class="zero">—</td><td class="zero">—</td><td class="zero">—</td><td class="zero">—</td>') + '</tr>';
  });
  const pm = monthsShown.filter(m => m <= ri);
  h += '</tbody><tfoot><tr><td>Итого</td><td class="num">' + K.payments + '</td><td class="num">' + rub(K.received) + '</td><td class="num">' + rub(K.plan) + '</td><td class="num">' + rub(K.fact) + '</td><td class="num">' + rub(K.fact - K.plan) + '</td><td class="num">' + Math.round(sum(pm.map(m => T.factL[m]))) + ' из ' + sum(pm.map(m => T.planL[m])) + '</td><td class="num">' + rub(K.owed) + '</td></tr></tfoot></table></div></section>';

  // таблица учеников
  h += '<section class="card"><h2>Ученики</h2><div class="controls">' +
    [['all', 'Все'], ['owed', 'Нужно доплатить'], ['plus', 'Оплачено вперёд']].map(([f, n]) => '<button class="chip" data-f="' + f + '" aria-pressed="' + (S.filter === f) + '">' + n + '</button>').join('') +
    '<input type="search" id="q" placeholder="Найти по имени или нику" aria-label="Найти ученика" value="' + esc(S.q) + '"></div><div class="scroll" id="st"></div>' +
    '<p class="note" style="margin:10px 0 0">Баланс: оплачено минус план с начала занятий до месяца отчёта. Плюс значит оплата вперёд, минус значит долг. Колонки «Нужно доплатить» считаются по неоплаченным занятиям каждого месяца и могут не совпадать с балансом, если деньги внесены вперёд на другие месяцы.</p></section>';
  view.innerHTML = h;
  bindDebtCard();
  if (cm.length) drawChart(cm, T);
  view.querySelectorAll('.chip').forEach(c => c.onclick = () => { S.filter = c.dataset.f; view.querySelectorAll('.chip').forEach(x => x.setAttribute('aria-pressed', x === c)); drawStudents(M, monthsShown.filter(m => m <= ri && T.planR[m] > 0).slice(-3)); });
  const q = $('q'); q.addEventListener('input', e => { S.q = e.target.value.trim().toLowerCase(); drawStudents(M, monthsShown.filter(m => m <= ri && T.planR[m] > 0).slice(-3)); });
  drawStudents(M, monthsShown.filter(m => m <= ri && T.planR[m] > 0).slice(-3));
}

function drawChart(cm, T) {
  const W = 680, H = 280, L = 52, R = 12, Tp = 24, B = 34, iw = W - L - R, ih = H - Tp - B;
  const vals = cm.flatMap(m => [T.planR[m], T.factR[m], T.owedR[m]]);
  const step = niceStep(Math.max(1, ...vals));
  const max = Math.max(step, Math.ceil(Math.max(...vals) / step) * step);
  const y = v => Tp + ih - v / max * ih;
  let s = '<svg viewBox="0 0 ' + W + ' ' + H + '" role="img" aria-label="Столбчатая диаграмма: план, факт и доплата по месяцам">';
  for (let v = 0; v <= max; v += step) {
    s += '<line x1="' + L + '" x2="' + (W - R) + '" y1="' + y(v) + '" y2="' + y(v) + '" stroke="' + (v ? 'var(--line)' : 'var(--axis)') + '" stroke-width="1"/>';
    s += '<text x="' + (L - 8) + '" y="' + (y(v) + 4) + '" text-anchor="end" font-size="11" style="fill:var(--muted)">' + (v / 1000) + ' тыс</text>';
  }
  const gw = iw / cm.length, bw = Math.min(56, gw / 4.4), gap = 4;
  const keys = ['planR', 'factR', 'owedR'];
  cm.forEach((m, i) => {
    const gx = L + gw * i + gw / 2 - (3 * bw + 2 * gap) / 2;
    keys.forEach((k, j) => {
      const v = T[k][m], x = gx + j * (bw + gap), yy = y(v), h = Math.max(0, Tp + ih - yy), r = Math.min(4, h);
      s += '<path class="bar" data-m="' + m + '" data-s="' + j + '" d="M' + x + ' ' + (Tp + ih) + 'V' + (yy + r) + 'Q' + x + ' ' + yy + ' ' + (x + r) + ' ' + yy + 'H' + (x + bw - r) + 'Q' + (x + bw) + ' ' + yy + ' ' + (x + bw) + ' ' + (yy + r) + 'V' + (Tp + ih) + 'Z" fill="' + SER[j].c + '" tabindex="0"/>';
      s += '<text x="' + (x + bw / 2) + '" y="' + (yy - 5) + '" text-anchor="middle" font-size="11" style="fill:var(--ink)">' + Math.round(v / 1000) + '</text>';
    });
    s += '<text x="' + (L + gw * i + gw / 2) + '" y="' + (H - 10) + '" text-anchor="middle" font-size="12" style="fill:var(--ink)">' + MONTHS[m].name + '</text>';
  });
  $('chart').innerHTML = s + '</svg>';
  const tip = $('tip');
  document.querySelectorAll('#chart .bar').forEach(el => {
    const m = +el.dataset.m, j = +el.dataset.s;
    const show = e => {
      tip.textContent = MONTHS[m].name + ' · ' + SER[j].n + ': ' + rub(T[keys[j]][m]); tip.style.opacity = 1;
      const rc = el.getBoundingClientRect(); const x = (e.clientX || rc.left) + 12, yy = (e.clientY || rc.top) - 34;
      tip.style.left = Math.min(x, innerWidth - 250) + 'px'; tip.style.top = Math.max(4, yy) + 'px';
    };
    el.addEventListener('mousemove', show); el.addEventListener('mouseenter', show); el.addEventListener('click', show); el.addEventListener('focus', show);
    el.addEventListener('mouseleave', () => tip.style.opacity = 0); el.addEventListener('blur', () => tip.style.opacity = 0);
  });
}

function pill(b) { if (b < 0) return '<span class="pill bad">▼ долг ' + rub(-b) + '</span>'; if (b > 0) return '<span class="pill good">▲ вперёд ' + rub(b) + '</span>'; return '<span class="pill eq">= ровно</span>'; }

function drawStudents(M, dm) {
  const cell = v => '<td class="num' + (v ? '' : ' zero') + '">' + (v ? rub(v) : '—') + '</td>';
  const cols = [['name', 'Ученик', 'l'], ['tgName', 'Имя для бота', 'l'], ['start', 'Начало', ''], ['planSum', 'План', ''], ['paid', 'Оплачено', ''], ['balance', 'Баланс', '']].concat(dm.map(m => ['o' + m, MONTHS[m].name, ''])).concat([['owedTotal', 'Всего доплатить', '']]);
  const val = (r, k) => k === 'name' ? r.s.name.toLowerCase() : k === 'tgName' ? String(r.s.tgName || '').toLowerCase() : k === 'start' ? (r.s.start || '') : k === 'planSum' ? r.planSum : k[0] === 'o' && k !== 'owedTotal' ? r.owedR[+k.slice(1)] : r[k];
  let rows = M.rows.filter(r => (!S.q || (r.s.name + ' ' + (r.s.tgName || '') + ' ' + (r.s.tg || '')).toLowerCase().includes(S.q)) && (S.filter === 'all' || (S.filter === 'owed' && r.owedTotal > 0) || (S.filter === 'plus' && r.balance > 0)));
  rows.sort((a, b) => { const x = val(a, S.sortK), y = val(b, S.sortK); return (x > y ? 1 : x < y ? -1 : 0) * S.sortDir; });
  const head = '<thead><tr>' + cols.map(([k, n, c]) => '<th class="' + c + '" aria-sort="' + (S.sortK === k ? (S.sortDir > 0 ? 'ascending' : 'descending') : 'none') + '"><button data-k="' + k + '" style="all:unset;cursor:pointer">' + n + (S.sortK === k ? (S.sortDir > 0 ? ' ↑' : ' ↓') : '') + '</button></th>').join('') + '</tr></thead>';
  const body = rows.map(r => '<tr><td class="l">' + esc(r.s.name) + (r.active ? '' : ' <span class="pill eq">' + esc(r.s.status) + '</span>') + (r.check && r.active ? ' <span class="pill warn" title="' + esc(r.check) + '">расписание</span>' : '') + '</td><td class="l">' + esc(r.s.tgName || '') + '</td><td class="num">' + esc(fmtDate(r.s.start) || '—') + '</td>' + cell(r.planSum) + cell(r.paid) + '<td>' + pill(r.balance) + '</td>' + dm.map(m => cell(r.owedR[m])).join('') + '<td class="num' + (r.owedTotal ? '' : ' zero') + '">' + (r.owedTotal ? '<b>' + rub(r.owedTotal) + '</b>' : '—') + '</td></tr>').join('') || '<tr><td colspan="' + cols.length + '" style="text-align:center;color:var(--muted)">Никого не найдено</td></tr>';
  const sm = k => rows.reduce((a, r) => a + (k === 'planSum' ? r.planSum : k === 'paid' ? r.paid : k === 'balance' ? r.balance : k === 'owedTotal' ? r.owedTotal : r.owedR[+k.slice(1)]), 0);
  const foot = '<tfoot><tr><td class="l">Итого (' + rows.length + ')</td><td></td><td></td>' + ['planSum', 'paid', 'balance'].map(k => '<td class="num">' + rub(sm(k)) + '</td>').join('') + dm.map(m => '<td class="num">' + rub(sm('o' + m)) + '</td>').join('') + '<td class="num">' + rub(sm('owedTotal')) + '</td></tr></tfoot>';
  const el = $('st'); el.innerHTML = '<table>' + head + '<tbody>' + body + '</tbody>' + foot + '</table>';
  el.querySelectorAll('th button').forEach(b => b.onclick = () => { const k = b.dataset.k; if (S.sortK === k) S.sortDir *= -1; else { S.sortK = k; S.sortDir = (k === 'name' || k === 'tgName' || k === 'start') ? 1 : -1; } drawStudents(M, dm); });
}

/* ---------- выгрузка должников для бота ---------- */
const debtOpts = () => ({ reminders: S.reminders, onlyActual: !S.debtAll });
const sentDate = () => { const t = todayISO(); return S.date > t ? t : S.date; };    // напоминание не может быть «из будущего»

function renderRemind() {
  const el = $('remind'); if (!el) return;
  if (!S.ready || !S.students.length) { el.innerHTML = ''; return; }
  const D = debtRows(model(), { reminders: S.reminders, onlyActual: true });
  const n = D.students.length + D.skipped.length;
  if (!n) { el.innerHTML = ''; return; }
  const names = D.students.map(x => x.name).concat(D.skipped.map(x => x.name));
  const sum = D.students.reduce((a, x) => a + x.total, 0) + D.skipped.reduce((a, x) => a + x.total, 0);
  el.innerHTML = '<div class="banner warn"><b>Пора напомнить о долге: ' + n + ' ' + (n === 1 ? 'ученику' : 'ученикам') + ', на сумму ' + rub(sum) + '</b><span>' + names.slice(0, 8).map(esc).join(', ') + (names.length > 8 ? ' и ещё ' + (names.length - 8) : '') + '. После последнего напоминания у них было занятие.</span><span><button class="btn small" id="remgo" type="button">Открыть выгрузку</button></span></div>';
  $('remgo').onclick = () => { S.tab = 'summary'; syncTabs(); render(); const c = $('debtcard'); if (c) c.scrollIntoView({ behavior: 'smooth', block: 'start' }); };
}

function debtCardHTML(M) {
  const D = debtRows(M, debtOpts()), ri = M.repIdx;
  const who = new Set(D.rows.map(r => r.nick)).size;
  const dis = D.rows.length ? '' : ' disabled';
  let h = '<section class="card" id="debtcard"><div class="chart-head"><h2>Выгрузка должников для бота</h2><span class="btns"><button class="btn ghost small" id="debtdl2" type="button"' + dis + '>Скачать с датами занятий</button> <button class="btn small" id="debtdl" type="button"' + dis + '>Скачать Excel</button></span></div>';
  if (ri < 0) return h + '<p class="note" style="margin:0">Дата отчёта раньше начала учёта, долгов нет.</p></section>';
  h += '<p class="note" style="margin:0 0 8px">Долги на ' + esc(fmtDate(S.date)) + ' (' + (S.toDate ? 'занятия до даты' : 'месяц целиком') + '). ' +
    (S.debtAll ? 'Показаны все должники.' : 'Показаны только те, кому пора напомнить: после последнего напоминания у них было занятие (занятие считается прошедшим в конце дня). Кому ещё не напоминали, тот подходит сразу.') +
    ' Одна строка = ник, имя, месяц и сумма недоплаты за этот месяц. «Скачать Excel»: формат как в файле для бота (лист «Оплаты», колонки Ник, Имя, Месяц, Сумма). «С датами занятий» добавляет колонки «Занятий» и «Даты занятий»: это неоплаченные занятия месяца (оплата закрывает занятия по порядку, поэтому долг приходится на последние даты).</p>';
  if (D.rows.length) h += '<p style="margin:0 0 8px"><b>' + who + '</b> ' + (who === 1 ? 'получатель' : 'получателей') + ', строк в файле: <b>' + D.rows.length + '</b>, на сумму <b>' + rub(D.total) + '</b>.</p>';
  else h += '<p style="margin:0 0 8px">' + (D.quiet.length ? 'Сейчас напоминать некому: у всех должников после последнего напоминания не было занятий.' : 'Некому писать: в файл не попал ни один ученик с долгом.') + '</p>';
  if (D.skipped.length) h += '<div class="banner warn" style="margin:8px 0"><b>Не попадут в файл (' + D.skipped.length + '):</b>' + D.skipped.map(x => '<span>' + esc(x.name) + ' — долг ' + rub(x.total) + ', ' + esc(x.why) + '</span>').join('') + '<span>Исправь ник на вкладке «Ученики» и скачай файл снова.</span></div>';
  if (D.quiet.length) h += '<p class="note" style="margin:0 0 8px">Ждут следующего занятия (' + D.quiet.length + '): ' + D.quiet.map(x => esc(x.name) + ' (напоминали ' + esc(fmtDate(x.last)) + ')').join(', ') + '.</p>';
  h += '<div class="controls" style="margin:8px 0"><button class="btn small" id="debtmark" type="button"' + (D.students.length ? '' : ' disabled') + '>Отметить как отправленные (' + D.students.length + ')</button>' +
    '<label class="fld" style="flex-direction:row;align-items:center;gap:6px"><input type="checkbox" id="debtall"' + (S.debtAll ? ' checked' : '') + '> Показать всех должников</label>' +
    (Object.keys(S.reminders).length ? '<button class="btn ghost small" id="debtclear" type="button">Сбросить журнал напоминаний</button>' : '') + '</div>';
  if (D.rows.length) h += '<details><summary style="cursor:pointer">Что попадёт в файл</summary><div class="scroll"><table><thead><tr><th class="l">Ник</th><th class="l">Имя</th><th class="l">Месяц</th><th>Сумма</th><th>Занятий</th><th class="l">Даты занятий</th><th class="l">Напоминали</th></tr></thead><tbody>' +
    D.rows.map(r => '<tr><td class="l">' + esc(r.nick) + '</td><td class="l">' + esc(r.name) + '</td><td class="l">' + esc(r.month) + '</td><td class="num">' + rub(r.sum) + '</td><td class="num">' + r.lessons + '</td><td class="l">' + esc(r.dates) + '</td><td class="l">' + (r.last ? esc(fmtDate(r.last)) : '—') + '</td></tr>').join('') + '</tbody></table></div></details>';
  return h + manualHTML(D) + '</section>';
}
const tplText = () => (S.settings && typeof S.settings.template === 'string' && S.settings.template.trim()) ? S.settings.template : DEFAULT_TEMPLATE;

function manualHTML(D) {
  let h = '<div style="margin-top:16px;border-top:1px solid var(--line);padding-top:12px"><h2 style="margin-bottom:6px">Написать вручную</h2>' +
    '<p class="note" style="margin:0 0 8px">Для тех, кто не любит бота. Кнопка «Написать в Telegram» открывает чат с учеником, текст уже вставлен, остаётся нажать «Отправить». После отправки нажмите «Отправлено», и ученик уйдёт из списка до следующего занятия.</p>' +
    '<details><summary style="cursor:pointer">Текст сообщения (можно изменить)</summary><div style="margin:8px 0"><textarea id="tpl" rows="9" style="width:100%">' + esc(tplText()) + '</textarea>' +
    '<p class="note" style="margin:6px 0">Подстановки: {имя} (из «Имя для бота»), {ученик}, {сумма}, {занятия} (например «2 занятия»), {занятий} (только число), {даты}, {месяц}, {сегодня} (дата отчёта). Если у ученика нет дат занятий, строка с {даты} пропускается. Правки сохраняются сами.</p>' +
    '<button class="btn ghost small" id="tplreset" type="button">Вернуть стандартный текст</button></div></details>';
  if (!D.people.length) return h + '<p class="note" style="margin:8px 0 0">Сейчас писать некому.</p></div>';
  h += D.people.map(p => {
    const text = buildMessage(tplText(), p);
    return '<div class="mrow" data-id="' + esc(p.id) + '"><div class="mwho"><b>' + esc(p.name) + '</b><span class="note"> · ' + (p.nick ? '@' + esc(p.nick) : esc(p.why)) + ' · долг ' + rub(p.total) + '</span></div>' +
      '<div class="mtxt">' + esc(text) + '</div><div class="btns">' +
      (p.nick ? '<a class="btn small" target="_blank" rel="noopener" href="' + esc(tgLink(p.nick, text)) + '">Написать в Telegram</a>' : '') +
      '<button class="btn ghost small" data-copy type="button">Скопировать текст</button>' +
      '<button class="btn ghost small" data-sent type="button">Отправлено</button></div></div>';
  }).join('');
  return h + '</div>';
}

async function markSent(ids) {
  const D = debtRows(model(), debtOpts());
  const items = D.people.filter(p => ids.includes(p.id)).map(p => ({ id: p.id, date: sentDate(), sum: p.total }));
  if (!items.length) return;
  const r = await api('POST', '/api/reminders/mark', { items });
  S.reminders = r.reminders || {};
}

async function copyText(text) {
  try { await navigator.clipboard.writeText(text); return true; }
  catch (e) {
    const ta = document.createElement('textarea'); ta.value = text; ta.style.position = 'fixed'; ta.style.opacity = '0'; document.body.appendChild(ta); ta.select();
    let ok = false; try { ok = document.execCommand('copy'); } catch (x) {} ta.remove(); return ok;
  }
}

function bindManual() {
  const card = $('debtcard'); if (!card) return;
  const tpl = $('tpl');
  if (tpl) tpl.addEventListener('change', async () => {
    try { const r = await api('PUT', '/api/settings', { template: tpl.value }); S.settings = r.settings || {}; setStatus('Текст сохранён'); render(); }
    catch (e) { setStatus('Не удалось сохранить текст: ' + e.message, true); }
  });
  const rs = $('tplreset');
  if (rs) rs.onclick = async () => {
    try { const r = await api('PUT', '/api/settings', { template: '' }); S.settings = r.settings || {}; setStatus('Стандартный текст возвращён'); render(); }
    catch (e) { setStatus('Не удалось сохранить: ' + e.message, true); }
  };
  card.querySelectorAll('.mrow').forEach(row => {
    const id = row.dataset.id;
    const cp = row.querySelector('[data-copy]'), sn = row.querySelector('[data-sent]');
    if (cp) cp.onclick = async () => { const ok = await copyText(row.querySelector('.mtxt').textContent); setStatus(ok ? 'Текст скопирован' : 'Не получилось скопировать, выделите текст вручную', !ok); };
    if (sn) sn.onclick = async () => {
      try { await markSent([id]); setStatus('Отмечено как отправленное'); render(); }
      catch (e) { setStatus('Не удалось отметить: ' + e.message, true); }
    };
  });
}

function bindDebtCard() {
  const dl = (withDates) => () => {
    if (!window.XLSX) { setStatus('Не загрузилась библиотека для Excel. Обнови страницу.', true); return; }
    const D = debtRows(model(), debtOpts());
    const head = withDates ? ['Ник', 'Имя', 'Месяц', 'Сумма', 'Занятий', 'Даты занятий'] : ['Ник', 'Имя', 'Месяц', 'Сумма'];
    const aoa = [head].concat(D.rows.map(r => withDates ? [r.nick, r.name, r.month, r.sum, r.lessons, r.dates] : [r.nick, r.name, r.month, r.sum]));
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!cols'] = [{ wch: 22 }, { wch: 28 }, { wch: 12 }, { wch: 10 }].concat(withDates ? [{ wch: 9 }, { wch: 44 }] : []);
    const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'Оплаты');
    XLSX.writeFile(wb, 'Долги' + (withDates ? '_с_датами_' : '_') + S.date + '.xlsx');
    setStatus('Файл скачан: строк ' + D.rows.length);
  };
  const a = $('debtdl'), b = $('debtdl2');
  if (a) a.onclick = dl(false);
  if (b) b.onclick = dl(true);
  bindManual();
  const all = $('debtall'); if (all) all.onchange = () => { S.debtAll = all.checked; render(); };
  const mk = $('debtmark');
  if (mk) mk.onclick = async () => {
    const D = debtRows(model(), debtOpts());
    if (!D.students.length) return;
    if (!mk.dataset.sure) { mk.dataset.sure = '1'; mk.textContent = 'Точно? Эти ' + D.students.length + ' уйдут из списка до следующего занятия'; setTimeout(() => { if (mk.isConnected) { delete mk.dataset.sure; mk.textContent = 'Отметить как отправленные (' + D.students.length + ')'; } }, 4000); return; }
    try {
      const r = await api('POST', '/api/reminders/mark', { items: D.students.map(x => ({ id: x.id, date: sentDate(), sum: x.total })) });
      S.reminders = r.reminders || {};
      setStatus('Отмечено как отправленные: ' + r.count + '. Вернутся в список после следующего занятия.');
      render();
    } catch (e) { setStatus('Не удалось отметить: ' + e.message, true); }
  };
  const cl = $('debtclear');
  if (cl) cl.onclick = async () => {
    if (!cl.dataset.sure) { cl.dataset.sure = '1'; cl.textContent = 'Точно сбросить?'; setTimeout(() => { if (cl.isConnected) { delete cl.dataset.sure; cl.textContent = 'Сбросить журнал напоминаний'; } }, 3000); return; }
    try { const r = await api('POST', '/api/reminders/clear', {}); S.reminders = r.reminders || {}; setStatus('Журнал напоминаний очищен'); render(); }
    catch (e) { setStatus('Не удалось сбросить: ' + e.message, true); }
  };
}

async function handleNickFile(file) {
  if (!file) return;
  $('notice').innerHTML = '';
  try {
    const res = matchRoster(S.students, await readFile(file));
    if (res.error) throw new Error(res.error);
    if (res.changed) { saving++; try { await api('PUT', '/api/students', { students: S.students }); setStatus('Сохранено'); } finally { saving--; } }
    let html = '<div class="banner ' + (res.unmatched.length ? 'warn' : 'ok') + '"><b>Ники загружены</b><span>Найдено учеников: ' + res.matched + ', обновлено: ' + res.changed + ', без изменений: ' + res.same + '.</span>';
    if (res.unmatched.length) html += '<span>Нет в базе (' + res.unmatched.length + '), пропущены: ' + res.unmatched.map(esc).join(', ') + '. Имя в списке должно совпадать с ФИО ученика на вкладке «Ученики».</span>';
    if (res.noNick.length) html += '<span>Ника нет у (' + res.noNick.length + '): ' + res.noNick.map(esc).join(', ') + '. Им бот не напишет.</span>';
    $('notice').innerHTML = html + '</div>';
    render();
  } catch (e) {
    $('notice').innerHTML = '<div class="banner err"><b>Не получилось загрузить список</b><span>' + esc(e && e.message ? e.message : e) + '</span></div>';
  }
}

function timelineHTML(s) {
  const lines = describeSchedule(s);
  return '<b>Расписание по периодам</b>' + (lines.length ? lines.map(l => '<div>' + esc(l) + '</div>').join('') : '<div>не заполнено (укажите начало, дни и занятий в неделю выше)</div>');
}
function chgRowHTML(c, si, ci) {
  const f = (k, cls, ph, type) => '<input class="' + cls + '" data-cf="' + k + '" ' + (type ? 'type="' + type + '" ' : '') + 'value="' + esc(c[k]) + '" placeholder="' + ph + '" aria-label="' + ph + '">';
  return '<div class="chg-row" data-si="' + si + '" data-ci="' + ci + '" data-kind="' + esc(c.kind) + '">' +
    '<select data-cf="kind" aria-label="Что меняется"><option value="off"' + (c.kind === 'off' ? ' selected' : '') + '>Не ходит</option><option value="sched"' + (c.kind === 'sched' ? ' selected' : '') + '>Другое расписание</option></select>' +
    '<label class="fld">с <input type="date" data-cf="from" value="' + esc(c.from) + '"></label>' +
    '<label class="fld">по (необязательно) <input type="date" data-cf="to" value="' + esc(c.to) + '"></label>' +
    '<span class="sched-only">' + f('days', 'w-act', 'Дни: Пн, Ср, Пт') + f('perWeek', 'w-num', 'В неделю', 'number') + f('price', 'w-num', 'Цена, ₽', 'number') + '</span>' +
    f('note', 'w-comm', 'Комментарий') +
    '<button class="btn ghost small" data-delchg type="button">Убрать</button>' +
    '<div class="chk" data-cchk>' + esc(changeCheck(c)) + '</div></div>';
}
function chgEditorHTML(s, i) {
  const list = s.changes || [];
  return '<tr class="chgrow" data-i="' + i + '"><td colspan="14"><div class="chg"><div class="note" style="margin:0 0 8px">Изменения в расписании «' + esc(s.name) + '». Расписание выше остаётся основным: до первого изменения и после периода с датой «по» ученик занимается по нему. ' +
    '«Не ходит»: занятий в этот период нет и платить за них не нужно. «Другое расписание»: впишите дни, в которые ученик ходит теперь (например «Пн, Ср, Пт»), число занятий в неделю посчитается само (2 дня значит 2 раза в неделю, 3 дня значит 3); цену заполняйте, только если она меняется. Без даты «по» изменение действует, пока не начнётся следующее.</div>' +
    '<div class="tl" data-tl>' + timelineHTML(s) + '</div>' +
    list.map((c, ci) => chgRowHTML(c, i, ci)).join('') +
    '<button class="btn small" data-addchg type="button">+ Добавить изменение</button></div></td></tr>';
}

function exportStudents() {
  if (!window.XLSX) { setStatus('Не загрузилась библиотека для Excel. Обнови страницу.', true); return; }
  const X = rosterExport(model(), S.payments);
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([X.head].concat(X.rows));
  ws['!cols'] = X.head.map((h, k) => ({ wch: Math.min(40, Math.max(10, h.length * 0.8, ...X.rows.map(r => String(r[k] === undefined ? '' : r[k]).length))) }));
  XLSX.utils.book_append_sheet(wb, ws, 'Ученики');
  if (X.chRows.length) {
    const w2 = XLSX.utils.aoa_to_sheet([X.chHead].concat(X.chRows));
    w2['!cols'] = X.chHead.map(h => ({ wch: Math.max(14, h.length) }));
    XLSX.utils.book_append_sheet(wb, w2, 'Изменения расписания');
  }
  XLSX.writeFile(wb, 'Ученики_' + todayISO() + '.xlsx');
  setStatus('Список скачан: учеников ' + X.rows.length);
}

/* ---------- Ученики (редактирование) ---------- */
function renderStudents(view) {
  const M = model(); const byName = new Map(M.rows.map(r => [r.s.id, r]));
  let h = '<section class="card"><div class="chart-head"><h2>Список учеников и расписание</h2><span class="btns"><button class="btn ghost small" id="exportst" type="button">Скачать список (Excel)</button> <button class="btn ghost small" id="nickbtn" type="button">Загрузить ники из списка</button> <button class="btn small" id="addst" type="button">+ Добавить ученика</button></span></div>' +
    '<p class="note" style="margin:0 0 10px">Правки сохраняются сами. План платежей считается по дате начала, дням недели и числу занятий в неделю. Если ученик ходит не во все дни из расписания, впиши реальные дни в колонку «Дни, в которые ходит».</p><div class="scroll"><table id="rt"><thead><tr>' +
    ['Ученик', 'Ник в Telegram', 'Имя для бота', 'Экзамен', 'Начало', 'Дни и время', 'Дни, в которые ходит', 'В неделю', 'Цена занятия, ₽', 'Статус', 'Скидка', 'Комментарий', 'Проверка расписания', 'Изменения', ''].map((n, i) => '<th class="l">' + n + '</th>').join('') + '</tr></thead><tbody>';
  S.students.forEach((s, i) => {
    const r = byName.get(s.id);
    h += '<tr data-i="' + i + '"><td class="l"><input class="w-name" data-f="name" value="' + esc(s.name) + '" aria-label="Ученик"></td>' +
      '<td class="l"><input class="w-act" data-f="tg" value="' + esc(s.tg) + '" placeholder="без @" aria-label="Ник в Telegram"></td>' +
      '<td class="l"><input class="w-name" data-f="tgName" value="' + esc(s.tgName) + '" placeholder="как обращаться" aria-label="Имя для бота"></td>' +
      '<td class="l"><input class="w-act" data-f="exam" value="' + esc(s.exam) + '" aria-label="Экзамен"></td>' +
      '<td class="l"><input type="date" data-f="start" value="' + esc(s.start) + '" aria-label="Начало занятий"></td>' +
      '<td class="l"><input class="w-days" data-f="days" value="' + esc(s.days) + '" placeholder="Пн 17:00, Ср 17:00" aria-label="Дни и время"></td>' +
      '<td class="l"><input class="w-act" data-f="actual" value="' + esc(s.actual) + '" placeholder="Чт и Сб" aria-label="Дни, в которые ходит"></td>' +
      '<td class="l"><input class="w-num" type="number" min="0" max="7" data-f="perWeek" value="' + esc(s.perWeek) + '" aria-label="Занятий в неделю"></td>' +
      '<td class="l"><input class="w-num" type="number" min="0" step="100" data-f="price" value="' + esc(s.price) + '" aria-label="Цена занятия"></td>' +
      '<td class="l"><select data-f="status" aria-label="Статус">' + ['Активен', 'Пауза', 'Завершил'].map(o => '<option' + (o === (s.status || 'Активен') ? ' selected' : '') + '>' + o + '</option>').join('') + '</select></td>' +
      '<td class="l" style="color:var(--ink2)">' + esc(r ? r.promo : '') + '</td>' +
      '<td class="l"><input class="w-comm" data-f="comment" value="' + esc(s.comment) + '" aria-label="Комментарий"></td>' +
      '<td class="chk" data-chk>' + esc(scheduleCheck(s)) + '</td>' +
      '<td class="l"><button class="btn ghost small" data-chg type="button" aria-expanded="' + S.openChg.has(s.id) + '">' + ((s.changes || []).length ? 'Изменения · ' + s.changes.length : 'Изменить') + '</button></td>' +
      '<td class="l"><button class="btn ghost small" data-del type="button" aria-label="Удалить ученика">Удалить</button></td></tr>';
    if (S.openChg.has(s.id)) h += chgEditorHTML(s, i);
  });
  h += '</tbody></table></div></section>';
  view.innerHTML = h;
  $('nickbtn').onclick = () => $('nickfile').click();
  $('exportst').onclick = exportStudents;
  $('addst').onclick = () => { S.students.push({ id: uid(), name: 'Новый ученик', exam: '', start: '', days: '', actual: '', perWeek: '', price: 1300, status: 'Активен', comment: '', tg: '', tgName: '', changes: [] }); scheduleRoster(); renderStudents(view); const rows = view.querySelectorAll('#rt tbody tr'); const last = rows[rows.length - 1]; if (last) { const inp = last.querySelector('input'); inp.focus(); inp.select(); } };
  const tb = view.querySelector('#rt tbody');
  tb.addEventListener('change', e => {
    const el = e.target;
    if (el.dataset.cf) {
      const row = el.closest('[data-ci]'), st = S.students[+row.dataset.si], c = st && st.changes[+row.dataset.ci]; if (!c) return;
      const k = el.dataset.cf; let v = el.value;
      if (k === 'perWeek' || k === 'price') v = v === '' ? '' : Number(v);
      c[k] = v; if (k === 'kind') row.dataset.kind = v;
      row.querySelector('[data-cchk]').textContent = changeCheck(c);
      const tl = row.closest('.chg').querySelector('[data-tl]'); if (tl) tl.innerHTML = timelineHTML(st);
      scheduleRoster(); return;
    }
    const tr = el.closest('tr'); if (!tr || !el.dataset.f) return;
    const s = S.students[+tr.dataset.i]; const f = el.dataset.f;
    let v = el.value;
    if (f === 'perWeek' || f === 'price') v = v === '' ? '' : Number(v);
    if (f === 'name') { v = String(v).trim(); if (!v) { el.value = s.name; return; } }
    if (f === 'tg') { v = normNick(v); el.value = v; }
    s[f] = v; tr.querySelector('[data-chk]').textContent = scheduleCheck(s); scheduleRoster();
  });
  tb.addEventListener('click', e => {
    const tg = e.target.closest('[data-chg]');
    if (tg) { const st = S.students[+tg.closest('tr').dataset.i]; if (S.openChg.has(st.id)) S.openChg.delete(st.id); else S.openChg.add(st.id); renderStudents(view); return; }
    const ad = e.target.closest('[data-addchg]');
    if (ad) { const st = S.students[+ad.closest('tr').dataset.i]; st.changes = st.changes || []; st.changes.push({ id: uid(), kind: 'off', from: S.date || todayISO(), to: '', days: '', perWeek: '', price: '', note: '' }); scheduleRoster(); renderStudents(view); return; }
    const dc = e.target.closest('[data-delchg]');
    if (dc) { const row = dc.closest('[data-ci]'), st = S.students[+row.dataset.si]; st.changes.splice(+row.dataset.ci, 1); scheduleRoster(); renderStudents(view); return; }
    const b = e.target.closest('[data-del]'); if (!b) return;
    const tr = b.closest('tr'), i = +tr.dataset.i;
    if (b.dataset.sure) { S.students.splice(i, 1); scheduleRoster(); renderStudents(view); return; }
    b.dataset.sure = '1'; b.textContent = 'Точно удалить?'; setTimeout(() => { if (b.isConnected) { delete b.dataset.sure; b.textContent = 'Удалить'; } }, 3000);
  });
}

/* ---------- Платежи ---------- */
function renderPays(view) {
  const known = new Set(S.students.map(s => s.name));
  const rows = S.payments.map(r => ({ r })).sort((a, b) => a.r.c < b.r.c ? 1 : -1);
  const cnt = { all: rows.length, paid: rows.filter(x => x.r.ok).length, unpaid: rows.filter(x => !x.r.ok).length, unk: rows.filter(x => !known.has(x.r.n)).length };
  let f = rows.filter(x => (!S.pq || (x.r.n + ' ' + x.r.pa).toLowerCase().includes(S.pq)) && (S.pf === 'all' || (S.pf === 'paid' && x.r.ok) || (S.pf === 'unpaid' && !x.r.ok) || (S.pf === 'unk' && !known.has(x.r.n))));
  const total = f.length; if (!S.showAllPays) f = f.slice(0, 150);
  let h = '<section class="card"><h2>Платежи</h2><div class="controls">' + [['all', 'Все'], ['paid', 'Оплачено'], ['unpaid', 'Без оплаты'], ['unk', 'Не из списка']].map(([k, n]) => '<button class="chip" data-pf="' + k + '" aria-pressed="' + (S.pf === k) + '">' + n + ' · ' + cnt[k === 'all' ? 'all' : k] + '</button>').join('') +
    '<input type="search" id="pq" placeholder="Найти по ученику" aria-label="Найти по ученику" value="' + esc(S.pq) + '"></div><div class="scroll"><table><thead><tr><th class="l">Дата</th><th class="l">Ученик</th><th class="l">Плательщик</th><th class="l">Тип занятий</th><th class="l">Месяц</th><th>Занятий</th><th>Сумма</th><th class="l">Промокод</th><th class="l">Статус</th><th></th></tr></thead><tbody>';
  h += f.map(x => { const r = x.r, pp = parseProduct(r.pr); return '<tr data-id="' + esc(r.id) + '"><td class="l num">' + esc(fmtDate(r.c)) + ' ' + esc((r.c || '').slice(11, 16)) + '</td><td class="l">' + esc(r.n) + (known.has(r.n) ? '' : ' <span class="pill warn">не в списке</span>') + '</td><td class="l" style="color:var(--ink2)">' + esc(r.pa) + '</td><td class="l">' + esc(pp.type || '—') + '</td><td class="l">' + esc(pp.month || '—') + '</td><td class="num">' + (pp.lessons || '<span class="pill warn">не распознано</span>') + '</td><td class="num">' + rub(safeNum(r.am) || 0) + '</td><td class="l">' + esc(r.pc) + '</td><td class="l">' + (r.ok ? '<span class="pill good">оплачено</span>' : '<span class="pill warn">без оплаты</span>') + '</td><td><button class="btn ghost small" data-del type="button" aria-label="Удалить платёж">Удалить</button></td></tr>'; }).join('') || '<tr><td colspan="10" style="text-align:center;color:var(--muted)">Ничего не найдено</td></tr>';
  h += '</tbody></table></div>' + (total > f.length ? '<p class="note" style="margin-top:10px">Показано ' + f.length + ' из ' + total + '. <button class="btn ghost small" id="allp" type="button">Показать все</button></p>' : '') +
    '<p class="note" style="margin-top:10px">Повторная загрузка той же выгрузки платежи не дублирует: платёж узнаётся по номеру (paymentid). Из выгрузки берутся только строки с заполненным paymentid.</p></section>';
  view.innerHTML = h;
  view.querySelectorAll('[data-pf]').forEach(c => c.onclick = () => { S.pf = c.dataset.pf; S.showAllPays = false; renderPays(view); });
  const pq = $('pq'); pq.addEventListener('change', e => { S.pq = e.target.value.trim().toLowerCase(); renderPays(view); });
  pq.addEventListener('keydown', e => { if (e.key === 'Enter') { S.pq = pq.value.trim().toLowerCase(); renderPays(view); } });
  const ap = $('allp'); if (ap) ap.onclick = () => { S.showAllPays = true; renderPays(view); };
  view.querySelectorAll('[data-del]').forEach(b => b.onclick = async () => {
    if (!b.dataset.sure) { b.dataset.sure = '1'; b.textContent = 'Точно удалить?'; setTimeout(() => { if (b.isConnected) { delete b.dataset.sure; b.textContent = 'Удалить'; } }, 3000); return; }
    const tr = b.closest('tr'), id = tr.dataset.id;
    try { await api('POST', '/api/payments/delete', { id }); S.payments = S.payments.filter(r => r.id !== id); setStatus('Платёж удалён'); }
    catch (e) { setStatus('Не удалось удалить: ' + e.message, true); }
    renderPays(view);
  });
}

/* ---------- вход, резервные копии, запуск ---------- */
function showLogin() { $('app').hidden = true; $('login').hidden = false; $('pw').focus(); }
async function showApp() {
  $('login').hidden = true; $('app').hidden = false;
  render();
  try { await loadState(); } catch (e) { setStatus('Не удалось загрузить данные: ' + e.message, true); S.ready = true; }
  render();
}
$('loginform').addEventListener('submit', async e => {
  e.preventDefault(); $('loginerr').textContent = '';
  try { await api('POST', '/api/login', { password: $('pw').value }); $('pw').value = ''; await showApp(); }
  catch (err) { $('loginerr').textContent = err.message; }
});
$('logout').onclick = async () => { try { await api('POST', '/api/logout', {}); } catch (e) {} S.students = []; S.payments = []; S.reminders = {}; S.settings = {}; S.ready = false; showLogin(); };
$('restorebtn').onclick = () => $('restorefile').click();
$('restorefile').addEventListener('change', async e => {
  const f = e.target.files[0]; e.target.value = ''; if (!f) return;
  try {
    const data = JSON.parse(await f.text());
    const n = (data.students || []).length, p = (data.payments || []).length;
    $('notice').innerHTML = '<div class="banner warn"><b>Заменить все данные из файла?</b><span>В файле ' + n + ' учеников и ' + p + ' платежей. Текущие данные на сайте будут заменены.</span><span><button class="btn small" id="okrestore" type="button">Заменить</button> <button class="btn ghost small" id="norestore" type="button">Отмена</button></span></div>';
    $('norestore').onclick = () => { $('notice').innerHTML = ''; };
    $('okrestore').onclick = async () => {
      try { const r = await api('POST', '/api/import', data); await loadState(); $('notice').innerHTML = '<div class="banner ok"><b>Данные восстановлены</b><span>Учеников: ' + r.students + ', платежей: ' + r.payments + '.</span></div>'; render(); }
      catch (err) { $('notice').innerHTML = '<div class="banner err"><b>Не получилось восстановить</b><span>' + esc(err.message) + '</span></div>'; }
    };
  } catch (err) { $('notice').innerHTML = '<div class="banner err"><b>Это не файл резервной копии</b><span>' + esc(err.message) + '</span></div>'; }
});

document.querySelectorAll('[role=tab]').forEach(t => t.onclick = () => { S.tab = t.dataset.tab; syncTabs(); render(); });
$('repdate').value = S.date;
$('repdate').addEventListener('change', e => { S.date = e.target.value || todayISO(); if (!e.target.value) e.target.value = S.date; render(); });
$('mode').value = S.toDate ? '1' : '0';
$('mode').addEventListener('change', e => { S.toDate = e.target.value === '1'; try { localStorage.setItem('uo-mode', e.target.value); } catch (x) {} render(); });
$('nickfile').addEventListener('change', async e => { const f = e.target.files[0]; e.target.value = ''; await handleNickFile(f); });
$('upbtn').onclick = () => $('file').click();
$('file').addEventListener('change', async e => { const f = e.target.files[0]; e.target.value = ''; await handleFile(f); });
['dragenter', 'dragover'].forEach(ev => document.addEventListener(ev, e => { if (e.dataTransfer && [...e.dataTransfer.types].includes('Files')) { e.preventDefault(); document.body.classList.add('dragging'); } }));
['dragleave', 'drop'].forEach(ev => document.addEventListener(ev, e => { if (ev === 'dragleave' && e.relatedTarget) return; document.body.classList.remove('dragging'); }));
document.addEventListener('drop', async e => { if (e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files.length && !$('app').hidden) { e.preventDefault(); await handleFile(e.dataTransfer.files[0]); } });

(async () => { try { await api('GET', '/api/session'); await showApp(); } catch (e) { showLogin(); } })();
