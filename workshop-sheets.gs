/**
 * workshop.html の申込を Google スプレッドシートへ 1申込1行で追記する受信スクリプト。
 * スプレッドシートの「拡張機能 > Apps Script」に貼り付け、Webアプリとしてデプロイする。
 * 設定手順は README.md を参照。
 */

const SHEET_NAME = '申込一覧';
const MAX_LEN = 2000; // 1項目あたりの最大文字数（極端に長い送信を切り詰める）

// [シートの列名, 受信データのキー]
const COLUMNS = [
  ['受付日時', null],
  ['申込ID', 'submissionId'],
  ['スレッズ名', 'name'],
  ['メールアドレス', 'email'],
  ['コース', 'level'],
  ['希望AI', 'ai'],
  ['希望日程', 'dates'],
  ['AI経験', 'experience'],
  ['支払い方法', 'payment'],
  ['参加費', 'price'],
  ['紹介者', 'referrer'],
  ['作りたいもの', 'idea'],
  ['質問・要望', 'note'],
  ['送信日時(端末)', 'submittedAt'],
];

function doPost(e) {
  const lock = LockService.getScriptLock();
  let locked = false;
  try {
    lock.waitLock(10000); // 同時送信で行が混ざったり重複チェックをすり抜けたりしないよう1件ずつ処理
    locked = true;
    const d = JSON.parse(e.postData.contents);
    if (!d.submissionId || !d.name || !d.email) { // フォーム側と同じく必須項目のみ確認
      return json_({ ok: false, error: 'invalid' });
    }

    const sheet = getSheet_();
    // 同じ申込IDがすでにあれば追記しない（再送信・二重クリック対策）
    const dup = sheet.getRange('B:B').createTextFinder(String(d.submissionId)).matchEntireCell(true).findNext();
    if (dup) return json_({ ok: true, duplicate: true });

    sheet.appendRow(COLUMNS.map(([, key]) => key ? clean_(d[key]) : new Date()));
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    if (locked) lock.releaseLock();
  }
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(COLUMNS.map(([label]) => label));
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// 文字列化して長さを制限。= + - で始まる値は数式として実行されないよう先頭に ' を付ける
function clean_(v) {
  const s = String(v == null ? '' : v).slice(0, MAX_LEN);
  return /^[=+\-]/.test(s) ? "'" + s : s;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
