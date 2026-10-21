'use strict';
const fs = require('fs');
const path = require('path');

/**
 * Хранилище в одном JSON-файле. Данных мало (десятки учеников, сотни платежей),
 * поэтому база не нужна. Запись атомарная: пишем во временный файл и переименовываем.
 * Все изменения идут по очереди, так что два запроса не затирают друг друга.
 */
class JsonStore {
  constructor(file) {
    this.file = file;
    this.data = { students: [], payments: [], reminders: {}, settings: {}, rev: 0 };
    this.queue = Promise.resolve();
  }

  load() {
    fs.mkdirSync(path.dirname(this.file), { recursive: true });
    if (fs.existsSync(this.file)) {
      const parsed = JSON.parse(fs.readFileSync(this.file, 'utf8'));
      this.data = {
        students: Array.isArray(parsed.students) ? parsed.students : [],
        payments: Array.isArray(parsed.payments) ? parsed.payments : [],
        reminders: parsed.reminders && typeof parsed.reminders === 'object' ? parsed.reminders : {},
        settings: parsed.settings && typeof parsed.settings === 'object' ? parsed.settings : {},
        rev: Number(parsed.rev) || 0,
      };
    }
    return this;
  }

  snapshot() {
    return JSON.parse(JSON.stringify(this.data));
  }

  /** fn получает рабочую копию данных, меняет её и возвращает результат; после этого копия сохраняется. */
  update(fn) {
    const run = async () => {
      const draft = this.snapshot();
      const result = await fn(draft);
      draft.rev = (draft.rev || 0) + 1;
      const tmp = this.file + '.tmp';
      await fs.promises.writeFile(tmp, JSON.stringify(draft));
      await fs.promises.rename(tmp, this.file);
      this.data = draft;
      return result;
    };
    const next = this.queue.then(run, run);
    this.queue = next.catch(() => {});
    return next;
  }
}

module.exports = { JsonStore };
