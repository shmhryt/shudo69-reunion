/**
 * 鏡開き協賛 意向表明フォーム → スプレッドシート受信用
 *
 * 使い方
 * 1. 回答を貯めたいスプレッドシートを開く（例：shudo69alumni@gmail.com のドライブに新規作成）
 * 2. 拡張機能 → Apps Script を開き、このファイルの中身を丸ごと貼り付けて保存
 * 3. デプロイ → 新しいデプロイ → 種類「ウェブアプリ」
 *      次のユーザーとして実行：自分
 *      アクセスできるユーザー：全員
 *    → 承認して、表示された「ウェブアプリのURL」（…/exec）をコピー
 * 4. kagami.html の ENDPOINT にそのURLを貼る
 *
 * シート「協賛意向」が無ければ自動で作成し、1行目に見出しを入れます。
 */

const SHEET_NAME = '協賛意向';
const NOTIFY_TO = ''; // 例：'shudo69alumni@gmail.com'（空なら通知メールなし）

const COLUMNS = [
  ['受付日時', null],
  ['協賛区分', 'type'],
  ['企業・団体名', 'org'],
  ['ご担当者名／お名前', 'name'],
  ['卒業回', 'kai'],
  ['メール', 'email'],
  ['電話', 'tel'],
  ['希望枠数', 'units'],
  ['最近挑戦していること', 'challenge'],
  ['掲示用のお名前', 'display'],
  ['ご質問・ご要望', 'memo'],
  ['同意', 'agree'],
  ['抽選結果', null],   // 実行委員会の記入用
  ['備考（委員会）', null] // 実行委員会の記入用
];

function doPost(e) {
  const p = (e && e.parameter) || {};
  if (p.website) return json_({ ok: true }); // bot 対策（隠し項目に入力があれば捨てる）

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheet_();
    const row = COLUMNS.map(([label, key]) => {
      if (label === '受付日時') return new Date();
      if (!key) return '';
      return clean_(p[key]);
    });
    sheet.appendRow(row);

    if (NOTIFY_TO) {
      MailApp.sendEmail(
        NOTIFY_TO,
        '【鏡開き協賛】意向表明が届きました：' + (p.org || p.name || ''),
        COLUMNS.filter(c => c[1]).map(([label, key]) => label + '：' + (p[key] || '')).join('\n')
      );
    }
    return json_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(COLUMNS.map(c => c[0]));
    sheet.getRange(1, 1, 1, COLUMNS.length).setFontWeight('bold').setBackground('#721e27').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// 先頭が = + - @ の入力を数式として実行させない
function clean_(v) {
  v = String(v == null ? '' : v).slice(0, 2000);
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
