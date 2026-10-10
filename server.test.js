'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const os = require('os');
const path = require('path');
const { createApp } = require('./server.js');

const sample = JSON.parse(fs.readFileSync(path.join(__dirname, 'sample-data.json'), 'utf8'));

async function start() {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'tp-'));
  const { app } = createApp({ password: 'secret-pass', dataDir: dir });
  const server = await new Promise(r => { const s = app.listen(0, () => r(s)); });
  const base = 'http://127.0.0.1:' + server.address().port;
  return { base, dir, close: () => server.close() };
}
const json = (method, body, cookie) => ({ method, headers: { 'Content-Type': 'application/json', ...(cookie ? { Cookie: cookie } : {}) }, body: body === undefined ? undefined : JSON.stringify(body) });

test('без входа API закрыт, с паролем открывается', async () => {
  const t = await start();
  try {
    assert.equal((await fetch(t.base + '/api/state')).status, 401);
    const bad = await fetch(t.base + '/api/login', json('POST', { password: 'nope' }));
    assert.equal(bad.status, 401);
    const ok = await fetch(t.base + '/api/login', json('POST', { password: 'secret-pass' }));
    assert.equal(ok.status, 200);
    const cookie = ok.headers.get('set-cookie').split(';')[0];
    assert.match(ok.headers.get('set-cookie'), /HttpOnly/);
    const st = await fetch(t.base + '/api/state', { headers: { Cookie: cookie } });
    assert.equal(st.status, 200);
    assert.deepEqual((await st.json()).students, []);
    assert.equal((await fetch(t.base + '/healthz')).status, 200);
  } finally { t.close(); }
});

test('импорт платежей, повторный импорт без дублей, удаление, резервная копия', async () => {
  const t = await start();
  try {
    const login = await fetch(t.base + '/api/login', json('POST', { password: 'secret-pass' }));
    const cookie = login.headers.get('set-cookie').split(';')[0];
    const rows = sample.payments.slice(0, 5);
    const a = await (await fetch(t.base + '/api/payments/import', json('POST', { rows }, cookie))).json();
    assert.equal(a.added, 5);
    const b = await (await fetch(t.base + '/api/payments/import', json('POST', { rows }, cookie))).json();
    assert.equal(b.added, 0); assert.equal(b.same, 5);
    let st = await (await fetch(t.base + '/api/state', { headers: { Cookie: cookie } })).json();
    assert.equal(st.payments.length, 5);
    assert.ok(st.students.length > 0);
    await fetch(t.base + '/api/payments/delete', json('POST', { id: rows[0].id }, cookie));
    st = await (await fetch(t.base + '/api/state', { headers: { Cookie: cookie } })).json();
    assert.equal(st.payments.length, 4);
    // восстановление из копии и сохранение на диск
    const imp = await fetch(t.base + '/api/import', json('POST', sample, cookie));
    assert.equal((await imp.json()).payments, 79);
    const saved = JSON.parse(fs.readFileSync(path.join(t.dir, 'db.json'), 'utf8'));
    assert.equal(saved.students.length, 30);
    const exp = await (await fetch(t.base + '/api/export', { headers: { Cookie: cookie } })).json();
    assert.equal(exp.payments.length, 79);
  } finally { t.close(); }
});

test('нельзя записать без JSON и с чужой cookie', async () => {
  const t = await start();
  try {
    const login = await fetch(t.base + '/api/login', json('POST', { password: 'secret-pass' }));
    const cookie = login.headers.get('set-cookie').split(';')[0];
    const r = await fetch(t.base + '/api/students', { method: 'PUT', headers: { 'Content-Type': 'text/plain', Cookie: cookie }, body: 'x' });
    assert.equal(r.status, 415);
    const forged = await fetch(t.base + '/api/students', json('PUT', { students: [] }, 'tp_session=9999999999999.abc'));
    assert.equal(forged.status, 401);
  } finally { t.close(); }
});
