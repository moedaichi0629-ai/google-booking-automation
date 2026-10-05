'use strict';

const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const GAS_FILE_PATH = path.join(__dirname, '..', 'gas', 'Code.gs');

// gas/Code.gs はGoogle Apps Script用のファイル（module.exports等は持たない）。
// このファイル自体は一切変更せず、Node上の仮想環境（vm）でそのまま読み込み、
// SpreadsheetApp・CalendarApp・GmailApp などのGAS固有オブジェクトを
// テスト用の簡易な偽物（フェイク）に差し替えて実行する。
// これにより、本番コードと全く同じ実装をテストできる。

/**
 * Utilities.formatDate の簡易版（Code.gs で使用する yyyy/MM/dd/HH/mm/M/d に対応）。
 */
function formatDate(date, timeZone, pattern) {
  const parts = {};
  new Intl.DateTimeFormat('en-US', {
    timeZone,
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    hourCycle: 'h23',
  })
    .formatToParts(date)
    .forEach((p) => {
      parts[p.type] = p.value;
    });
  return pattern
    .replace('yyyy', parts.year)
    .replace('MM', parts.month)
    .replace('dd', parts.day)
    .replace('HH', parts.hour)
    .replace('mm', parts.minute)
    .replace(/M/g, String(Number(parts.month)))
    .replace(/d/g, String(Number(parts.day)));
}

/**
 * Utilities.parseDate の簡易版（'yyyy-MM-dd HH:mm' と Asia/Tokyo のみ対応）。
 * 形式が合わない場合は、実機と同様に例外を投げる。
 */
function parseDate(text, timeZone, pattern) {
  if (pattern !== 'yyyy-MM-dd HH:mm' || timeZone !== 'Asia/Tokyo') {
    throw new Error('テスト用parseDateが未対応の形式です: ' + pattern + ' / ' + timeZone);
  }
  const m = String(text).match(/^(\d{4})-(\d{2})-(\d{2}) (\d{1,2}):(\d{2})$/);
  if (!m) {
    throw new Error('Unparseable date: "' + text + '"');
  }
  return new Date(`${m[1]}-${m[2]}-${m[3]}T${m[4].padStart(2, '0')}:${m[5]}:00+09:00`);
}

/**
 * 2次元配列で値を保持するシートのフェイク。
 */
class FakeSheet {
  constructor(rows) {
    this.rows = rows.map((r) => r.slice());
    this.activeRow = 2;
  }

  getRange(row, column, numRows = 1, numColumns = 1) {
    const sheet = this;
    return {
      getValues() {
        const out = [];
        for (let r = 0; r < numRows; r++) {
          const src = sheet.rows[row - 1 + r] || [];
          const line = [];
          for (let c = 0; c < numColumns; c++) {
            const v = src[column - 1 + c];
            line.push(v === undefined ? '' : v);
          }
          out.push(line);
        }
        return out;
      },
      getValue() {
        return this.getValues()[0][0];
      },
      setValue(value) {
        while (sheet.rows.length < row) {
          sheet.rows.push([]);
        }
        const line = sheet.rows[row - 1];
        while (line.length < column) {
          line.push('');
        }
        line[column - 1] = value;
        return this;
      },
      getRow: () => row,
    };
  }

  getActiveCell() {
    return this.getRange(this.activeRow, 1);
  }
}

/**
 * GAS環境のフェイクを作り、gas/Code.gs を読み込んだコンテキストを返す。
 *
 * options:
 *   now               : 実行日時（new Date() の引数なし呼び出しがこの日時になる）
 *   rows              : 回答シートの内容（1行目はヘッダー）
 *   failCalendar      : true にするとカレンダー登録が失敗する
 *   failEmail         : true にするとメール送信が失敗する
 */
function loadGasLogic(options = {}) {
  const fixedNow = new Date(options.now || '2026-10-05T09:00:00+09:00').getTime();

  class FakeDate extends Date {
    constructor(...args) {
      if (args.length === 0) {
        super(fixedNow);
      } else {
        super(...args);
      }
    }

    static now() {
      return fixedNow;
    }
  }

  const sheet = new FakeSheet(options.rows || []);
  const createdEvents = [];
  const sentEmails = [];
  const alerts = [];
  const failures = {
    calendar: Boolean(options.failCalendar),
    email: Boolean(options.failEmail),
  };

  const sandbox = {
    Date: FakeDate,
    String,
    Object,
    Error,
    isNaN,
    SpreadsheetApp: {
      getActiveSheet: () => sheet,
      getUi: () => ({
        alert: (message) => alerts.push(message),
        createMenu: () => ({ addItem() { return this; }, addToUi() {} }),
      }),
    },
    CalendarApp: {
      getDefaultCalendar: () => ({
        createEvent(title, start, end, opts) {
          if (failures.calendar) {
            throw new Error('カレンダー登録テスト用のエラー');
          }
          const id = 'event-' + (createdEvents.length + 1) + '@google.com';
          createdEvents.push({ id, title, start, end, description: opts && opts.description });
          return { getId: () => id };
        },
      }),
    },
    GmailApp: {
      sendEmail(to, subject, body) {
        if (failures.email) {
          throw new Error('メール送信テスト用のエラー');
        }
        sentEmails.push({ to, subject, body });
      },
    },
    Utilities: { formatDate, parseDate },
  };

  const context = vm.createContext(sandbox);
  vm.runInContext(fs.readFileSync(GAS_FILE_PATH, 'utf8'), context, { filename: 'gas/Code.gs' });

  return { gas: context, sheet, createdEvents, sentEmails, alerts, failures };
}

module.exports = { loadGasLogic };
