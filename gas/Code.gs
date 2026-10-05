/**
 * オンライン相談 予約受付自動化ツール - Phase 2: 予約自動処理
 * 確定仕様: docs/02_requirements.md
 */

// 列番号（A=1 ... J=10）
var COLUMNS = {
  TIMESTAMP: 1,
  NAME: 2,
  EMAIL: 3,
  RESERVATION_DATE: 4,
  RESERVATION_TIME: 5,
  CONSULTATION: 6,
  STATUS: 7,
  EVENT_ID: 8,
  MAIL_SENT_AT: 9,
  ERROR: 10
};
var TOTAL_COLUMNS = 10;
var FIRST_DATA_ROW = 2; // 1行目はヘッダー

// 処理ステータス（G列）
var STATUS = {
  PROCESSING: '処理中',
  COMPLETED: '完了',
  ERROR: 'エラー',
  CALENDAR_ONLY: 'Calendar登録済み・メール未送信'
};

var TIMEZONE = 'Asia/Tokyo';
var RESERVATION_DURATION_MINUTES = 60;
var EMAIL_SUBJECT = '【オンライン相談】ご予約ありがとうございます';

/**
 * インストール型「フォーム送信時」トリガーから呼び出す関数。
 * トリガー設定: Apps Scriptエディタ > トリガー > 追加
 *   イベントのソース: スプレッドシートから / イベントの種類: フォーム送信時 / 実行する関数: onFormSubmit
 */
function onFormSubmit(e) {
  var sheet = e.range.getSheet();
  var row = e.range.getRow();
  processReservationRow(sheet, row);
}

/**
 * スプレッドシートを開いたときにカスタムメニューを追加する（シンプルトリガー）。
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('予約管理')
    .addItem('選択した予約を再実行', 'reprocessSelectedReservation')
    .addToUi();
}

/**
 * カスタムメニュー「選択した予約を再実行」から呼び出す関数。
 * 現在選択されている行を対象に、通常処理と同じ処理関数を実行する。
 */
function reprocessSelectedReservation() {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSheet();
  var row = sheet.getActiveCell().getRow();

  if (row < FIRST_DATA_ROW) {
    ui.alert('ヘッダー行が選択されています。予約データの行を選択してください。');
    return;
  }

  var name = sheet.getRange(row, COLUMNS.NAME).getValue();
  if (!name) {
    ui.alert('選択した行に予約データがありません。');
    return;
  }

  var processed = processReservationRow(sheet, row);
  if (!processed) {
    ui.alert(row + '行目はすでに処理が完了しています（予約完了メール送信済み）。カレンダー登録・メール送信は行いませんでした。');
    return;
  }
  ui.alert(row + '行目の再実行が完了しました。処理結果はG〜J列をご確認ください。');
}

/**
 * 予約1件分の自動処理（フォーム送信時・手動再実行時で共通利用）。
 *
 * I列（メール送信日時）が記録済みの行は正常完了済みとみなし、何もせずに false を返す。
 * I列はメール送信に成功したときだけ記録されるため、エラーで終わった行
 * （I列が空）は再実行で処理を続けられる。それ以外は処理を行い true を返す。
 */
function processReservationRow(sheet, row) {
  var values = sheet.getRange(row, 1, 1, TOTAL_COLUMNS).getValues()[0];
  if (isCompletedRow(values)) {
    return false;
  }

  var name = values[COLUMNS.NAME - 1];
  var email = values[COLUMNS.EMAIL - 1];
  var dateValue = values[COLUMNS.RESERVATION_DATE - 1];
  var timeValue = values[COLUMNS.RESERVATION_TIME - 1];
  var consultation = values[COLUMNS.CONSULTATION - 1];
  var existingEventId = values[COLUMNS.EVENT_ID - 1];

  if (!name || !email || !dateValue || !timeValue || !consultation) {
    recordError(sheet, row, '入力値が不足しています（お名前・メールアドレス・予約日・予約時間・相談内容のいずれかが空です）');
    return true;
  }

  var startDateTime;
  try {
    startDateTime = buildReservationDateTime(dateValue, timeValue);
  } catch (err) {
    recordError(sheet, row, '予約日時の解析に失敗しました: ' + err.message);
    return true;
  }
  var endDateTime = new Date(startDateTime.getTime() + RESERVATION_DURATION_MINUTES * 60 * 1000);

  var eventId = existingEventId ? String(existingEventId) : '';

  if (!eventId) {
    // 新規作成時のみ過去日時判定を行う（既にCalendar登録済みの再実行は対象外）
    if (startDateTime.getTime() <= new Date().getTime()) {
      recordError(sheet, row, '予約日時が過去のため、Calendar登録・メール送信を行いません');
      return true;
    }

    setStatus(sheet, row, STATUS.PROCESSING);

    try {
      eventId = createCalendarEvent(name, email, consultation, startDateTime, endDateTime);
    } catch (err) {
      recordError(sheet, row, 'Calendar登録に失敗しました: ' + err.message);
      return true;
    }
    // メール送信より先にイベントIDを保存する
    sheet.getRange(row, COLUMNS.EVENT_ID).setValue(eventId);
  } else {
    setStatus(sheet, row, STATUS.PROCESSING);
  }

  try {
    sendConfirmationEmail(email, name, startDateTime);
  } catch (err) {
    sheet.getRange(row, COLUMNS.STATUS).setValue(STATUS.CALENDAR_ONLY);
    sheet.getRange(row, COLUMNS.MAIL_SENT_AT).setValue('');
    sheet.getRange(row, COLUMNS.ERROR).setValue('メール送信に失敗しました: ' + err.message);
    return true;
  }

  sheet.getRange(row, COLUMNS.STATUS).setValue(STATUS.COMPLETED);
  sheet.getRange(row, COLUMNS.MAIL_SENT_AT).setValue(new Date());
  sheet.getRange(row, COLUMNS.ERROR).setValue('');
  return true;
}

/**
 * 予約行が正常完了済みかどうかを判定する。
 * I列（メール送信日時）はメール送信に成功したときだけ記録され、
 * メール送信に失敗したときは空欄に戻されるため、この列を完了の判定に使う。
 */
function isCompletedRow(values) {
  var mailSentAt = values[COLUMNS.MAIL_SENT_AT - 1];
  return mailSentAt !== '' && mailSentAt !== null && mailSentAt !== undefined;
}

/**
 * D列（予約日）とE列（予約時間）からAsia/Tokyoの予約開始日時を生成する。
 */
function buildReservationDateTime(dateValue, timeValue) {
  var dateStr;
  if (Object.prototype.toString.call(dateValue) === '[object Date]') {
    dateStr = Utilities.formatDate(dateValue, TIMEZONE, 'yyyy-MM-dd');
  } else {
    dateStr = String(dateValue).trim();
  }
  var timeStr = String(timeValue).trim();

  var parsed = Utilities.parseDate(dateStr + ' ' + timeStr, TIMEZONE, 'yyyy-MM-dd HH:mm');
  if (isNaN(parsed.getTime())) {
    throw new Error('日付または時間の形式が不正です');
  }
  return parsed;
}

/**
 * デフォルトCalendarへ予定を登録し、イベントIDを返す。
 */
function createCalendarEvent(name, email, consultation, startDateTime, endDateTime) {
  var calendar = CalendarApp.getDefaultCalendar();
  var title = name + '様 オンライン相談';
  var description = '相談内容：\n' + consultation + '\n\nメールアドレス：\n' + email;

  var event = calendar.createEvent(title, startDateTime, endDateTime, {
    description: description
  });
  return event.getId();
}

/**
 * 予約完了メールを送信する。
 */
function sendConfirmationEmail(email, name, startDateTime) {
  var formattedDateTime = Utilities.formatDate(startDateTime, TIMEZONE, 'M月d日 HH:mm');
  var body = name + '様\n\n' +
    'ご予約ありがとうございます。\n' +
    formattedDateTime + 'より、オンライン相談のご予約を承りました。\n\n' +
    'よろしくお願いいたします。';

  GmailApp.sendEmail(email, EMAIL_SUBJECT, body);
}

function setStatus(sheet, row, status) {
  sheet.getRange(row, COLUMNS.STATUS).setValue(status);
}

function recordError(sheet, row, message) {
  sheet.getRange(row, COLUMNS.STATUS).setValue(STATUS.ERROR);
  sheet.getRange(row, COLUMNS.ERROR).setValue(message);
}
