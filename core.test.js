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
