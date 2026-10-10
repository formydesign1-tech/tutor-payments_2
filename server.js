'use strict';
const crypto = require('crypto');
const path = require('path');
const express = require('express');
const core = require('./lib/core.js');
const { JsonStore } = require('./lib/store.js');

const COOKIE = 'tp_session';
const SESSION_DAYS = 30;

const sha = s => crypto.createHash('sha256').update(String(s)).digest();
const newId = () => 's' + crypto.randomBytes(4).toString('hex');

function createApp({ password, dataDir, secret }) {
  if (!password) throw new Error('Не задан APP_PASSWORD');
  const sessionSecret = secret || sha('session:' + password).toString('hex');
  const store = new JsonStore(path.join(dataDir, 'db.json')).load();

  const sign = v => crypto.createHmac('sha256', sessionSecret).update(v).digest('base64url');
  const makeToken = () => { const exp = String(Date.now() + SESSION_DAYS * 864e5); return exp + '.' + sign(exp); };
  const checkToken = t => {
    if (!t || typeof t !== 'string') return false;
    const [exp, mac] = t.split('.');
    if (!exp || !mac || Number(exp) < Date.now()) return false;
    const good = Buffer.from(sign(exp)), got = Buffer.from(mac);
    return good.length === got.length && crypto.timingSafeEqual(good, got);
  };
  const readCookie = req => {
    const raw = req.headers.cookie || '';
    for (const part of raw.split(';')) {
      const i = part.indexOf('=');
      if (i > 0 && part.slice(0, i).trim() === COOKIE) return decodeURIComponent(part.slice(i + 1).trim());
    }
    return '';
  };

  const app = express();
  app.disable('x-powered-by');
  app.set('trust proxy', 1);                 // за прокси Railway определяем https по заголовку

  app.use((req, res, next) => {
    res.set({
      'X-Content-Type-Options': 'nosniff',
      'Referrer-Policy': 'no-referrer',
      'X-Frame-Options': 'DENY',
      'Content-Security-Policy': "default-src 'self'; script-src 'self'; style-src 'self' 'unsafe-inline' https://fonts.googleapis.com; font-src https://fonts.gstatic.com; img-src 'self' data:; connect-src 'self'; frame-ancestors 'none'; base-uri 'none'; form-action 'self'",
    });
    next();
  });

  app.get('/healthz', (req, res) => res.type('text').send('ok'));

  // --- вход
  const attempts = new Map();
  app.post('/api/login', express.json({ limit: '4kb' }), (req, res) => {
    const ip = req.ip || 'x', now = Date.now();
    const a = attempts.get(ip) || { n: 0, reset: now + 10 * 60 * 1000 };
    if (now > a.reset) { a.n = 0; a.reset = now + 10 * 60 * 1000; }
    if (a.n >= 10) return res.status(429).json({ error: 'Слишком много попыток. Подожди несколько минут.' });
    const given = req.body && typeof req.body.password === 'string' ? req.body.password : '';
    const ok = crypto.timingSafeEqual(sha(given), sha(password));
    if (!ok) { a.n += 1; attempts.set(ip, a); return res.status(401).json({ error: 'Неверный пароль' }); }
    attempts.delete(ip);
    res.cookie(COOKIE, makeToken(), { httpOnly: true, sameSite: 'lax', secure: req.secure, maxAge: SESSION_DAYS * 864e5, path: '/' });
    res.json({ ok: true });
  });
  app.post('/api/logout', (req, res) => { res.clearCookie(COOKIE, { path: '/' }); res.json({ ok: true }); });

  const needAuth = (req, res, next) => (checkToken(readCookie(req)) ? next() : res.status(401).json({ error: 'Нужно войти' }));
  const needJson = (req, res, next) => (req.is('application/json') ? next() : res.status(415).json({ error: 'Ожидается JSON' }));
  const noStore = (req, res, next) => { res.set('Cache-Control', 'no-store'); next(); };
  const bodyParser = express.json({ limit: '8mb' });

  app.use('/api', noStore);
  app.get('/api/session', needAuth, (req, res) => res.json({ ok: true }));

  app.get('/api/state', needAuth, (req, res) => res.json(store.snapshot()));

  app.put('/api/students', needAuth, needJson, bodyParser, async (req, res) => {
    const list = req.body && req.body.students;
    if (!Array.isArray(list) || list.length > 1000) return res.status(400).json({ error: 'Неверный список учеников' });
    const students = list.map(s => core.sanitizeStudent(s, newId)).filter(Boolean);
    await store.update(d => { d.students = students; });
    res.json({ ok: true, count: students.length });
  });

  app.post('/api/payments/import', needAuth, needJson, bodyParser, async (req, res) => {
    const rows = req.body && req.body.rows;
    if (!Array.isArray(rows) || !rows.length || rows.length > 20000) return res.status(400).json({ error: 'Нет платежей для загрузки' });
    const stats = await store.update(d => core.applyImport(d, rows, newId));
    res.json(stats);
  });

  app.post('/api/payments/delete', needAuth, needJson, bodyParser, async (req, res) => {
    const id = req.body && typeof req.body.id === 'string' ? req.body.id : '';
    if (!id) return res.status(400).json({ error: 'Не указан платёж' });
    const removed = await store.update(d => { const n = d.payments.length; d.payments = d.payments.filter(p => p.id !== id); return n - d.payments.length; });
    res.json({ ok: true, removed });
  });

  app.get('/api/export', needAuth, (req, res) => {
    const day = new Date().toISOString().slice(0, 10);
    res.set('Content-Disposition', `attachment; filename="uchet-oplat-${day}.json"`);
    res.json({ students: store.data.students, payments: store.data.payments });
  });

  app.post('/api/import', needAuth, needJson, bodyParser, async (req, res) => {
    const b = req.body || {};
    if (!Array.isArray(b.students) || !Array.isArray(b.payments)) return res.status(400).json({ error: 'В файле нет списков students и payments' });
    const students = b.students.map(s => core.sanitizeStudent(s, newId)).filter(Boolean);
    const payments = []; const seen = new Set();
    for (const p of b.payments) { const r = core.sanitizeRow(p); if (r && !seen.has(r.id)) { seen.add(r.id); payments.push(r); } }
    await store.update(d => { d.students = students; d.payments = payments; });
    res.json({ ok: true, students: students.length, payments: payments.length });
  });

  // --- страница и скрипты
  app.get('/core.js', (req, res) => res.sendFile(path.join(__dirname, 'lib', 'core.js')));
  app.get('/vendor/xlsx.js', (req, res) => res.sendFile(require.resolve('xlsx/dist/xlsx.full.min.js')));
  app.use(express.static(path.join(__dirname, 'public'), { maxAge: '5m' }));

  app.use((err, req, res, next) => {            // eslint-disable-line no-unused-vars
    if (err && err.type === 'entity.too.large') return res.status(413).json({ error: 'Файл слишком большой' });
    if (err && err.type === 'entity.parse.failed') return res.status(400).json({ error: 'Неверный JSON' });
    console.error(err);
    res.status(500).json({ error: 'Ошибка сервера' });
  });

  return { app, store };
}

module.exports = { createApp };

if (require.main === module) {
  const password = process.env.APP_PASSWORD;
  if (!password) { console.error('Задай переменную окружения APP_PASSWORD (пароль для входа на сайт).'); process.exit(1); }
  const dataDir = process.env.DATA_DIR || path.join(__dirname, 'data');
  const { app } = createApp({ password, dataDir, secret: process.env.SESSION_SECRET });
  const port = Number(process.env.PORT) || 3000;
  app.listen(port, () => console.log(`Учёт оплат: http://localhost:${port} (данные: ${dataDir})`));
}
