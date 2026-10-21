'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');
const core = require('./core.js');

const sample = JSON.parse(fs.readFileSync(path.join(__dirname, 'sample-data.json'), 'utf8'));
const REPORT_DAY = core.dayNum(2026, 9, 10);        // 10 октября 2026

test('итоги совпадают с расчётом в Excel (режим «занятия до даты»)', () => {
  const r = core.compute(sample.students, sample.payments, REPORT_DAY, true);
  assert.equal(r.kpi.received, 355910);
  assert.equal(r.kpi.plan, 331510);
  assert.equal(Math.round(r.kpi.fact), 289154);
  assert.equal(r.kpi.owed, 43610);
  assert.equal(Math.round(r.kpi.prepaid), 66756);
  assert.equal(r.kpi.payments, 79);
  assert.equal(r.totals.owedR[8], 5900);            // сентябрь
  assert.equal(r.totals.owedR[9], 37710);           // октябрь
});

test('режим «месяц целиком» считает октябрь полностью', () => {
  const r = core.compute(sample.students, sample.payments, REPORT_DAY, false);
  assert.equal(r.totals.planL[9], 291);
});

test('разбор названия товара из Tilda', () => {
  const p = core.parseProduct('Оплата уроков (Тип занятий: Абонемент 8 занятий , Месяц: Октябрь 2026) - 1x10400 = 10400');
  assert.equal(p.lessons, 8);
  assert.equal(p.perLesson, 1300);
  assert.equal(p.month, 'Октябрь 2026');
  const q = core.parseProduct('Оплата уроков (Тип занятий: Разовое занятие , Месяц: Сентябрь 2026) - 3x1300 = 3900');
  assert.equal(q.lessons, 3);
  assert.equal(core.parseProduct('что-то другое').lessons, 0);
});

test('safeNum понимает суммы из Tilda', () => {
  assert.equal(core.safeNum('5200.0000'), 5200);
  assert.equal(core.safeNum('1 300,50'), 1300.5);
  assert.equal(core.safeNum(''), null);
});

test('CSV из Tilda: точка с запятой, BOM, статус и дедупликация', () => {
  const text = fs.readFileSync(path.join(__dirname, 'sample-tilda.csv'), 'utf8');
  const parsed = core.rowsFromMatrix(core.parseCSV(text));
  assert.ok(!parsed.error);
  assert.equal(parsed.rows.length, 12);
  assert.ok(parsed.rows.every(r => r.id && r.n && r.ok === 1));
  const db = { students: [], payments: [] };
  let n = 0;
  const first = core.applyImport(db, parsed.rows, () => 'n' + (++n));
  assert.equal(first.added, 12);
  const again = core.applyImport(db, parsed.rows, () => 'n' + (++n));
  assert.equal(again.added, 0);
  assert.equal(again.same, 12);
  assert.equal(db.payments.length, 12);
  assert.ok(db.students.length > 0);
});

test('берутся только строки с заполненным paymentid', () => {
  const header = ['created', 'name', 'product', 'amount', 'tilda_status', 'paymentid'];
  const prod = 'Оплата уроков (Тип занятий: Разовое занятие , Месяц: Октябрь 2026) - 1x1300 = 1300';
  const res = core.rowsFromMatrix([
    header,
    ['2026-10-01 10:00:00', 'Ученик А', prod, '1300', 'sent', 'custom.atolpay:123'],
    ['2026-10-01 11:00:00', 'Ученик Б', prod, '1300', 'sent', ''],          // нет paymentid: оплата не прошла
    ['2026-10-01 12:00:00', 'Ученик В', prod, '1300', '', 'custom.atolpay:456'],  // статус пустой, но paymentid есть: оплачено
  ]);
  assert.equal(res.rows.length, 2);
  assert.equal(res.skipped, 1);
  assert.deepEqual(res.rows.map(r => r.id), ['123', '456']);
  assert.ok(res.rows.every(r => r.ok === 1));
});

test('без колонки paymentid файл не принимается', () => {
  const res = core.rowsFromMatrix([['created', 'name', 'product', 'amount'], ['2026-10-01', 'А', 'x', '1']]);
  assert.match(res.error, /paymentid/);
});

test('проверка расписания подсказывает, что не так', () => {
  assert.match(core.scheduleCheck({ start: '', days: '', perWeek: '' }), /Заполните/);
  assert.match(core.scheduleCheck({ start: '2026-09-01', days: 'Пн, Ср, Пт', perWeek: 2 }), /впишите реальные дни/);
  assert.equal(core.scheduleCheck({ start: '2026-09-01', days: 'Пн, Ср', perWeek: 2 }), '');
});

test('оплата вперёд закрывает следующие месяцы', () => {
  const s = [{ id: 'a', name: 'А', start: '2026-09-01', days: 'Пн', actual: '', perWeek: 1, price: 1000, status: 'Активен' }];
  const pay = [{ id: '1', c: '2026-09-01 10:00:00', n: 'А', pr: 'Оплата уроков (Тип занятий: Абонемент 12 занятий , Месяц: Сентябрь 2026) - 1x12000 = 12000', am: 12000, pm: 12000, di: null, pc: '', ok: 1 }];
  // понедельников: сентябрь 4, октябрь 4, ноябрь 5 — оплачено 12 занятий
  const oct = core.compute(s, pay, core.dayNum(2026, 9, 31), false);
  assert.equal(oct.kpi.owed, 0);
  assert.equal(oct.kpi.received, 12000);
  const nov = core.compute(s, pay, core.dayNum(2026, 10, 30), false);
  assert.equal(nov.kpi.owed, 1000);                                  // 13-е занятие не оплачено
});

/* ---------- ники и выгрузка должников для бота ---------- */
test('normNick убирает @ и ссылку, isNick проверяет правила Telegram', () => {
  assert.equal(core.normNick(' @tgg_111 '), 'tgg_111');
  assert.equal(core.normNick('https://t.me/martyunina'), 'martyunina');
  assert.equal(core.normNick(null), '');
  assert.ok(core.isNick('tgg_111'));
  assert.ok(!core.isNick('Наталья'));
  assert.ok(!core.isNick('🎩'));
  assert.ok(!core.isNick('abc'));
});

test('matchRoster переносит ники и имена из списка учеников по ФИО', () => {
  const st = [
    { id: '1', name: 'Татьяна Миронова' }, { id: '2', name: 'Исмаил' }, { id: '3', name: 'Без ника' }, { id: '4', name: 'Не в файле' },
  ];
  const m = [
    ['Ученики'], [],
    ['№', 'ФИО ученика', 'Телефон', 'Ник', 'Имя'],
    [1, 'татьяна  миронова', '+7', '@exp_onyou', 'Татьяна Миронова'],
    [2, 'Исмаил', '+7', 'isma16160', 'Исмаил Гасымов'],
    [3, 'Без ника', '+7', '', ''],
    [4, 'Лишний', '+7', 'extra_nick', 'Лишний'],
  ];
  const r = core.matchRoster(st, m);
  assert.equal(r.matched, 3);
  assert.equal(r.changed, 2);
  assert.deepEqual(r.unmatched, ['Лишний']);
  assert.deepEqual(r.noNick, ['Без ника']);
  assert.deepEqual(r.notInFile, ['Не в файле']);
  assert.equal(st[0].tg, 'exp_onyou');
  assert.equal(st[1].tgName, 'Исмаил Гасымов');
  assert.equal(core.matchRoster(st, m).changed, 0);                 // повторная загрузка ничего не меняет
  assert.match(core.matchRoster(st, [['a', 'b']]).error, /ФИО ученика/);
});

test('sanitizeStudent сохраняет ник и имя для бота', () => {
  const s = core.sanitizeStudent({ name: 'А', tg: '@nick_abc', tgName: ' Имя ' }, () => 'x');
  assert.equal(s.tg, 'nick_abc');
  assert.equal(s.tgName, 'Имя');
  assert.equal(core.sanitizeStudent({ name: 'Б' }, () => 'y').tg, '');
});

test('debtRows: только активные должники, по строке на месяц, без ника — в пропущенных', () => {
  const mk = (id, name, tg, status) => ({ id, name, start: '2026-09-01', days: 'Пн', actual: '', perWeek: 1, price: 1000, status: status || 'Активен', tg: tg || '', tgName: '' });
  const st = [mk('a', 'Аня', 'anya_tg1'), mk('b', 'Боря', ''), mk('c', 'Вика', 'Вика'), mk('d', 'Глеб', 'gleb_tg1', 'Пауза'), mk('e', 'Дима', 'dima_tg1')];
  st[0].tgName = 'Анна';
  const pay = [{ id: '1', c: '2026-09-01 10:00:00', n: 'Дима', pr: 'Оплата уроков (Тип занятий: Абонемент 12 занятий , Месяц: Сентябрь 2026) - 1x12000 = 12000', am: 12000, pm: 12000, di: null, pc: '', ok: 1 }];
  const M = core.compute(st, pay, core.dayNum(2026, 9, 31), false);   // конец октября: сентябрь 4 + октябрь 4 урока
  const D = core.debtRows(M);
  assert.deepEqual(D.rows.map(r => [r.nick, r.name, r.month, r.sum]), [['anya_tg1', 'Анна', 'сентябрь', 4000], ['anya_tg1', 'Анна', 'октябрь', 4000]]);
  assert.equal(D.total, 8000);
  assert.deepEqual(D.skipped.map(x => x.name), ['Боря', 'Вика']);    // без ника и «не-ник»; Глеб на паузе, Дима оплатил
});

test('debtRows на реальных данных: сумма строк = «Нужно доплатить» активных учеников с ником', () => {
  const st = sample.students.map((s, i) => Object.assign({}, s, { tg: 'student_' + String(i).padStart(3, '0') }));
  const M = core.compute(st, sample.payments, REPORT_DAY, true);
  const D = core.debtRows(M);
  assert.equal(D.skipped.length, 0);
  assert.equal(D.total, M.kpi.owed);
  assert.ok(D.rows.every(r => r.sum > 0));
});

/* ---------- изменения в расписании ученика ---------- */
const OCT = { ms: core.dayNum(2026, 9, 1), me: core.dayNum(2026, 10, 0) };
const baseSt = (changes) => ({ id: 'a', name: 'А', start: '2026-09-01', days: 'Пн, Ср', actual: '', perWeek: 2, price: 1000, status: 'Активен', changes: changes || [] });
const octPlan = st => { const segs = core.buildSegments(st, ['Пн', 'Вт', 'Ср', 'Чт', 'Пт', 'Сб', 'Вс'].map(d => (st.days.includes(d) ? 1 : 0))); return core.monthPlan(st, segs, OCT.ms, OCT.me); };

test('без изменений октябрь: 4 понедельника + 4 среды', () => {
  assert.deepEqual(octPlan(baseSt()), { lessons: 8, gross: 8000 });
});

test('с 15 октября ходит 1 раз в неделю (только среда)', () => {
  const r = octPlan(baseSt([{ kind: 'sched', from: '2026-10-15', to: '', days: 'Ср', perWeek: 1, price: '' }]));
  assert.equal(r.lessons, 6);                                   // до 14.10: Пн 5, 12, Ср 7, 14; после: Ср 21, 28
  assert.equal(r.gross, 6000);
});

test('не ходит с даты по дату, потом возвращается к обычному расписанию', () => {
  const r = octPlan(baseSt([{ kind: 'off', from: '2026-10-15', to: '2026-10-27' }]));
  assert.equal(r.lessons, 5);                                   // 4 до перерыва + Ср 28.10
});

test('не ходит с даты без даты окончания: занятий больше нет', () => {
  assert.equal(octPlan(baseSt([{ kind: 'off', from: '2026-10-15' }])).lessons, 4);
});

test('следующее изменение закрывает предыдущее', () => {
  const r = octPlan(baseSt([{ kind: 'off', from: '2026-10-15' }, { kind: 'sched', from: '2026-10-22', days: 'Ср' }]));
  assert.equal(r.lessons, 5);                                   // 4 + Ср 28.10
});

test('изменение цены с даты считается по новой цене', () => {
  const r = octPlan(baseSt([{ kind: 'sched', from: '2026-10-15', price: 1500 }]));
  assert.equal(r.lessons, 8);
  assert.equal(r.gross, 4 * 1000 + 4 * 1500);                   // 4 занятия до 15.10 и 4 после
});

test('изменения до начала занятий и мусорные даты не ломают расчёт', () => {
  assert.equal(octPlan(baseSt([{ kind: 'off', from: '' }, { kind: 'off', from: 'вчера' }])).lessons, 8);
  assert.equal(octPlan(baseSt([{ kind: 'off', from: '2026-01-01', to: '2026-01-31' }])).lessons, 8);
});

test('compute учитывает изменения: план и долг за октябрь', () => {
  const st = [baseSt([{ kind: 'off', from: '2026-10-15' }])];
  const M = core.compute(st, [], core.dayNum(2026, 9, 31), false);
  assert.equal(M.rows[0].exp[9], 4);
  assert.equal(M.rows[0].planR[9], 4000);
  assert.equal(M.rows[0].owedR[9], 4000);
  assert.equal(core.compute([baseSt()], [], core.dayNum(2026, 9, 31), false).rows[0].exp[9], 8);
});

test('sanitizeStudent чистит изменения; describeChange и changeCheck', () => {
  const s = core.sanitizeStudent({ name: 'А', changes: [{ kind: 'sched', from: '2026-10-15', days: 'Ср', perWeek: '1' }, { kind: 'off', from: '2026-11-01', to: '2026-11-10', note: 'болеет' }, null, 5] }, () => 'id1');
  assert.equal(s.changes.length, 2);
  assert.equal(s.changes[0].perWeek, 1);
  assert.equal(s.changes[0].id, 'id1');
  assert.equal(core.describeChange(s.changes[1]), 'с 01.11.2026 по 10.11.2026: не ходит (болеет)');
  assert.equal(core.describeChange(s.changes[0]), 'с 15.10.2026: Ср, 1 зан. в неделю');
  assert.equal(core.changeCheck(s.changes[0]), '');
  assert.match(core.changeCheck({ kind: 'sched', from: '2026-10-15' }), /Укажите/);
  assert.match(core.changeCheck({ kind: 'off', from: '2026-10-15', to: '2026-10-01' }), /раньше/);
  assert.match(core.changeCheck({ kind: 'sched', from: '', days: 'Ср' }), /дату/);
  assert.equal(core.sanitizeStudent({ name: 'Б' }, () => 'x').changes.length, 0);
});

test('rosterExport: все столбцы, плательщик из последнего платежа, лист изменений', () => {
  const st = [Object.assign(baseSt([{ kind: 'off', from: '2026-10-15', note: 'отпуск' }]), { tg: 'abcde', tgName: 'Аня', phone: '+7 1', comment: 'к' })];
  const pay = [{ id: '1', c: '2026-09-01 10:00:00', n: 'А', pa: 'Мама Ани', pr: 'Оплата уроков (Тип занятий: Абонемент 12 занятий , Месяц: Сентябрь 2026) - 1x12000 = 12000', am: 12000, pm: 12000, di: null, pc: '', ok: 1 }];
  const M = core.compute(st, pay, core.dayNum(2026, 9, 31), false);
  const X = core.rosterExport(M, pay);
  const row = X.rows[0], col = n => row[X.head.indexOf(n)];
  assert.equal(col('ФИО плательщика (если отличается)'), 'Мама Ани');
  assert.equal(col('Ник в Telegram'), 'abcde');
  assert.equal(col('Начало занятий'), '01.09.2026');
  assert.equal(col('Обычно занятий в месяц'), 8);
  assert.equal(col('Оплачено, ₽'), 12000);
  assert.match(col('Изменения в расписании'), /не ходит \(отпуск\)/);
  assert.ok(X.rows.every(r => r.length === X.head.length));
  assert.deepEqual(X.chRows, [['А', 'Не ходит', '15.10.2026', '', '', '', '', 'отпуск']]);
});

/* ---------- долги с количеством и датами занятий ---------- */
test('lessonDays: даты занятий совпадают с количеством и учитывают изменения', () => {
  const st = baseSt([{ kind: 'off', from: '2026-10-15', to: '2026-10-27' }]);
  const fl = [1, 0, 1, 0, 0, 0, 0];
  const segs = core.buildSegments(st, fl);
  const days = core.lessonDays(st, segs, OCT.ms, OCT.me);
  assert.deepEqual(days.map(core.dayDM), ['05.10', '07.10', '12.10', '14.10', '28.10']);
  assert.equal(days.length, core.monthPlan(st, segs, OCT.ms, OCT.me).lessons);
});

test('debtRows: количество и даты неоплаченных занятий — последние занятия месяца', () => {
  // Пн и Ср, 1000 ₽; оплачено 14 занятий: сентябрь (9) и 5 занятий октября; октябрь по 31.10 = 8 занятий, не оплачены 3 последних
  const st = [Object.assign(baseSt(), { tg: 'anya_tg1' })];
  const pay = [{ id: '1', c: '2026-09-01 10:00:00', n: 'А', pr: 'Оплата уроков (Тип занятий: Абонемент 14 занятий , Месяц: Сентябрь 2026) - 1x14000 = 14000', am: 14000, pm: 14000, di: null, pc: '', ok: 1 }];
  const M = core.compute(st, pay, core.dayNum(2026, 9, 31), false);
  const D = core.debtRows(M);
  assert.equal(M.rows[0].exp[8], 9);                              // сентябрь: Пн 7, 14, 21, 28 и Ср 2, 9, 16, 23, 30
  assert.equal(D.rows.length, 1);
  assert.equal(D.rows[0].month, 'октябрь');
  assert.equal(D.rows[0].lessons, 3);
  assert.equal(D.rows[0].sum, 3000);
  assert.equal(D.rows[0].dates, '21.10, 26.10, 28.10');
});

/* ---------- напоминания: только те, у кого было занятие после последнего напоминания ---------- */
test('debtRows с журналом: без нового занятия не напоминаем, после занятия снова в списке', () => {
  // Пн и Ср, 1000 ₽, оплат нет; октябрь: Пн 5, 12, 19, 26; Ср 7, 14, 21, 28
  const st = [Object.assign(baseSt(), { id: 'a', tg: 'anya_tg1' }), Object.assign(baseSt(), { id: 'b', name: 'Б', tg: 'boris_tg1' })];
  const run = (day, rem) => core.debtRows(core.compute(st, [], core.dayNum(2026, 9, day), true), { reminders: rem, onlyActual: true });
  assert.equal(run(14, {}).students.length, 2);                                   // раньше не писали: оба подходят
  const rem = { a: { date: '2026-10-14', sum: 1, at: '' }, b: { date: '2026-10-12', sum: 1, at: '' } };
  const d14 = run(14, rem);                                                       // у Ани после 14.10 занятий нет, у Бори была Ср 14.10
  assert.deepEqual(d14.students.map(x => x.id), ['b']);
  assert.deepEqual(d14.quiet.map(x => x.name), ['А']);
  const d15 = run(15, rem);                                                       // 15.10 (Чт) занятий нет ни у кого
  assert.deepEqual(d15.students.map(x => x.id), ['b']);
  const d19 = run(19, Object.assign({}, rem, { b: { date: '2026-10-14', sum: 1, at: '' } }));   // 19.10 Пн: занятие у обоих
  assert.deepEqual(d19.students.map(x => x.id).sort(), ['a', 'b']);
  const all = core.debtRows(core.compute(st, [], core.dayNum(2026, 9, 14), true), { reminders: rem });
  assert.equal(all.students.length, 2);                                           // без фильтра видны все должники
});

test('lessonsAfter и sanitizeReminders', () => {
  const st = baseSt();
  assert.equal(core.lessonsAfter(st, core.dayNum(2026, 9, 14), core.dayNum(2026, 9, 18)), 0);
  assert.equal(core.lessonsAfter(st, core.dayNum(2026, 9, 14), core.dayNum(2026, 9, 19)), 1);
  assert.equal(core.lessonsAfter(st, core.dayNum(2026, 9, 19), core.dayNum(2026, 9, 14)), 0);
  const r = core.sanitizeReminders({ a: { date: '2026-10-14', sum: '5200' }, b: { date: 'вчера' }, c: null });
  assert.deepEqual(Object.keys(r), ['a']);
  assert.equal(r.a.sum, 5200);
  assert.deepEqual(core.sanitizeReminders([1, 2]), {});
});

/* ---------- сообщение для ручной отправки ---------- */
test('buildMessage: текст как у бота, подстановки, ссылка в Telegram', () => {
  const p = { greet: 'Татьяна Миронова', name: 'Татьяна Миронова', today: '21.10.2026', total: 1300, lessons: 1, parts: [{ month: 'октябрь', lessons: 1, dates: '21.10' }] };
  const m = core.buildMessage('', p);
  assert.match(m, /^Здравствуйте! 😊\nНа сегодняшний день \(21\.10\.2026\) вы не оплатили 1 занятие на сумму 1 300 ₽\.\nДаты занятий: 21\.10\.\n/);
  assert.match(m, /Спасибо!$/);
  const two = core.buildMessage('{сумма} / {занятий} / {даты}', { greet: 'А', name: 'А', total: 12000, lessons: 5, parts: [{ month: 'сентябрь', lessons: 2, dates: '22.09, 29.09' }, { month: 'октябрь', lessons: 3, dates: '06.10, 13.10, 20.10' }] });
  assert.equal(two, '12 000 ₽ / 5 / сентябрь: 22.09, 29.09; октябрь: 06.10, 13.10, 20.10');
  const noDates = core.buildMessage('Долг: {сумма} ({даты})\nСтрока {даты}\nКонец', { greet: 'А', name: 'А', total: 500, lessons: 1, parts: [{ month: 'октябрь', lessons: 1, dates: '' }] });
  assert.equal(noDates, 'Долг: 500 ₽\nКонец');
  const link = core.tgLink('abcde', 'Привет\nмир & ещё');
  assert.equal(link, 'https://t.me/abcde?text=' + encodeURIComponent('Привет\nмир & ещё'));
  assert.ok(!link.includes(' ') && !link.includes('\n'));
});

test('debtRows.people: все должники с ником или причиной, суммы и даты по месяцам', () => {
  const mk = (id, name, tg) => ({ id, name, start: '2026-09-01', days: 'Пн', actual: '', perWeek: 1, price: 1000, status: 'Активен', tg, tgName: '' });
  const st = [mk('a', 'Аня', 'anya_tg1'), mk('b', 'Боря', ''), mk('c', 'Вика', 'Вика')];
  const D = core.debtRows(core.compute(st, [], core.dayNum(2026, 9, 31), false));
  assert.deepEqual(D.people.map(x => x.id), ['a', 'b', 'c']);
  assert.equal(D.people[0].nick, 'anya_tg1');
  assert.equal(D.people[1].nick, ''); assert.equal(D.people[1].why, 'нет ника');
  assert.match(D.people[2].why, /не похож/);
  assert.equal(D.people[0].lessons, 8);                            // сентябрь 4 + октябрь 4, оплат нет
  assert.deepEqual(D.people[0].parts.map(x => x.month), ['сентябрь', 'октябрь']);
  assert.equal(D.people[0].parts[1].dates, '05.10, 12.10, 19.10, 26.10');
  assert.equal(D.people[0].total, 8000);
});

test('lessonsWord: склонение', () => {
  assert.deepEqual([1, 2, 4, 5, 11, 12, 14, 21, 22, 25, 111].map(core.lessonsWord), ['1 занятие', '2 занятия', '4 занятия', '5 занятий', '11 занятий', '12 занятий', '14 занятий', '21 занятие', '22 занятия', '25 занятий', '111 занятий']);
});

test('debtRows.people: дата отчёта для сообщения', () => {
  const st = [{ id: 'a', name: 'А', start: '2026-09-01', days: 'Пн', actual: '', perWeek: 1, price: 1000, status: 'Активен', tg: 'anya_tg1', tgName: '' }];
  const D = core.debtRows(core.compute(st, [], core.dayNum(2026, 9, 21), true));
  assert.equal(D.people[0].today, '21.10.2026');
  assert.match(core.buildMessage('', D.people[0]), /^Здравствуйте! 😊\nНа сегодняшний день \(21\.10\.2026\) вы не оплатили 7 занятий на сумму 7 000 ₽\./);
});
