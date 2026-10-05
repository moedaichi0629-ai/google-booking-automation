'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { loadGasLogic } = require('../test-support/loadGasLogic');

const HEADER = ['タイムスタンプ', 'お名前', 'メールアドレス', '予約日', '予約時間', '相談内容', '処理ステータス', 'CalendarイベントID', 'メール送信日時', 'エラー内容'];

// サンプルデータはすべて架空の情報（example.com）。実行日時は 2026-10-05 09:00（日本時間）
function reservationRow(overrides = {}) {
  const base = ['2026/10/05 08:55:00', '山田 花子', 'demo-yamada@example.com', '2026-10-10', '10:00', 'ホームページ制作の相談', '', '', '', ''];
  Object.keys(overrides).forEach((index) => {
    base[index] = overrides[index];
  });
  return base;
}

// 列番号（0始まり）
const STATUS = 6;
const EVENT_ID = 7;
const MAIL_SENT_AT = 8;
const ERROR = 9;

function submitForm(env, row = 2) {
  env.gas.onFormSubmit({ range: { getSheet: () => env.sheet, getRow: () => row } });
}

function rerunSelected(env, row = 2) {
  env.sheet.activeRow = row;
  env.gas.reprocessSelectedReservation();
}

test('新規予約：カレンダー登録・メール送信を行い、G〜J列に完了を記録する', () => {
  const env = loadGasLogic({ rows: [HEADER, reservationRow()] });

  submitForm(env);

  const row = env.sheet.rows[1];
  assert.equal(env.createdEvents.length, 1);
  assert.equal(env.createdEvents[0].title, '山田 花子様 オンライン相談');
  assert.equal(env.sentEmails.length, 1);
  assert.equal(env.sentEmails[0].to, 'demo-yamada@example.com');
  assert.match(env.sentEmails[0].body, /10月10日 10:00より/);
  assert.equal(row[STATUS], '完了');
  assert.equal(row[EVENT_ID], 'event-1@google.com');
  assert.notEqual(row[MAIL_SENT_AT], '');
  assert.equal(row[ERROR], '');
});

test('完了済み予約：再実行してもカレンダー登録・メール送信を行わない', () => {
  const env = loadGasLogic({ rows: [HEADER, reservationRow()] });
  submitForm(env);
  const before = env.sheet.rows[1].slice();

  rerunSelected(env);

  assert.equal(env.createdEvents.length, 1, 'カレンダーの二重登録は発生しない');
  assert.equal(env.sentEmails.length, 1, '予約完了メールは再送信されない');
  assert.deepEqual(env.sheet.rows[1], before, 'G〜J列の記録も変わらない');
  assert.match(env.alerts[0], /すでに処理が完了しています/);
});

test('完了済み予約：フォーム送信時トリガーが同じ行で再度動いても二重処理しない', () => {
  const env = loadGasLogic({ rows: [HEADER, reservationRow()] });
  submitForm(env);
  submitForm(env);

  assert.equal(env.createdEvents.length, 1);
  assert.equal(env.sentEmails.length, 1);
});

test('メール送信エラーの予約：再実行でメールだけ送信し、カレンダーは二重登録しない', () => {
  const env = loadGasLogic({ rows: [HEADER, reservationRow()], failEmail: true });
  submitForm(env);

  let row = env.sheet.rows[1];
  assert.equal(row[STATUS], 'Calendar登録済み・メール未送信');
  assert.equal(row[EVENT_ID], 'event-1@google.com');
  assert.equal(row[MAIL_SENT_AT], '');
  assert.match(row[ERROR], /メール送信に失敗しました/);

  // 原因を解消して再実行
  env.failures.email = false;
  rerunSelected(env);

  row = env.sheet.rows[1];
  assert.equal(env.createdEvents.length, 1, 'カレンダーは最初の1件のまま');
  assert.equal(env.sentEmails.length, 1);
  assert.equal(row[STATUS], '完了');
  assert.notEqual(row[MAIL_SENT_AT], '');
  assert.equal(row[ERROR], '');
  assert.match(env.alerts[0], /再実行が完了しました/);

  // 完了後にもう一度再実行しても何もしない
  rerunSelected(env);
  assert.equal(env.createdEvents.length, 1);
  assert.equal(env.sentEmails.length, 1);
});

test('カレンダー登録エラーの予約：再実行でカレンダー登録・メール送信を行える', () => {
  const env = loadGasLogic({ rows: [HEADER, reservationRow()], failCalendar: true });
  submitForm(env);

  let row = env.sheet.rows[1];
  assert.equal(row[STATUS], 'エラー');
  assert.equal(row[EVENT_ID], '');
  assert.match(row[ERROR], /Calendar登録に失敗しました/);
  assert.equal(env.sentEmails.length, 0);

  env.failures.calendar = false;
  rerunSelected(env);

  row = env.sheet.rows[1];
  assert.equal(env.createdEvents.length, 1);
  assert.equal(env.sentEmails.length, 1);
  assert.equal(row[STATUS], '完了');
});

test('入力不足エラーの予約：入力を補ってから再実行すると処理される', () => {
  const env = loadGasLogic({ rows: [HEADER, reservationRow({ 5: '' })] });
  submitForm(env);

  assert.equal(env.sheet.rows[1][STATUS], 'エラー');
  assert.match(env.sheet.rows[1][ERROR], /入力値が不足しています/);
  assert.equal(env.createdEvents.length, 0);

  env.sheet.rows[1][5] = 'ホームページ制作の相談';
  rerunSelected(env);

  assert.equal(env.sheet.rows[1][STATUS], '完了');
  assert.equal(env.createdEvents.length, 1);
  assert.equal(env.sentEmails.length, 1);
});

test('過去日時の予約はエラーとして記録し、カレンダー登録・メール送信を行わない', () => {
  const env = loadGasLogic({ rows: [HEADER, reservationRow({ 3: '2026-10-01' })] });
  submitForm(env);

  assert.equal(env.sheet.rows[1][STATUS], 'エラー');
  assert.match(env.sheet.rows[1][ERROR], /予約日時が過去/);
  assert.equal(env.createdEvents.length, 0);
  assert.equal(env.sentEmails.length, 0);
});

test('予約時間の形式が不正な場合はエラーとして記録する', () => {
  const env = loadGasLogic({ rows: [HEADER, reservationRow({ 4: '10時' })] });
  submitForm(env);

  assert.equal(env.sheet.rows[1][STATUS], 'エラー');
  assert.match(env.sheet.rows[1][ERROR], /予約日時の解析に失敗しました/);
});

test('同じ日時の別々の予約は、それぞれ登録・送信される', () => {
  const env = loadGasLogic({
    rows: [HEADER, reservationRow(), reservationRow({ 1: '佐藤 一郎', 2: 'demo-sato@example.com' })],
  });
  submitForm(env, 2);
  submitForm(env, 3);

  assert.equal(env.createdEvents.length, 2);
  assert.deepEqual(env.sentEmails.map((m) => m.to), ['demo-yamada@example.com', 'demo-sato@example.com']);
});

test('再実行メニュー：ヘッダー行・空行を選択した場合は処理せずアラートを表示する', () => {
  const env = loadGasLogic({ rows: [HEADER, reservationRow({ 1: '' })] });

  rerunSelected(env, 1);
  rerunSelected(env, 2);

  assert.match(env.alerts[0], /ヘッダー行が選択されています/);
  assert.match(env.alerts[1], /予約データがありません/);
  assert.equal(env.createdEvents.length, 0);
  assert.equal(env.sentEmails.length, 0);
});
