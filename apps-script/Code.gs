/**
 * 家計簿アプリの「自動で共有」の受け口（Google スプレッドシートの Apps Script）
 *
 * つかい方
 *   1. 家計シート（スプレッドシート）を開き、「拡張機能 → Apps Script」
 *   2. 最初からある中身を全部消して、このコードを貼り付ける（アプリの［コードをコピー］なら合言葉が入った状態でコピーされる）
 *   3. 下の KEY を、アプリに入れる合言葉と同じにする
 *   4. 保存 →「デプロイ → 新しいデプロイ」→ 種類「ウェブアプリ」
 *      実行するユーザー：自分 ／ アクセスできるユーザー：全員 → デプロイ → アクセスを承認
 *   5. 出てきた「ウェブアプリの URL」をアプリに貼る
 *
 * データは、このスプレッドシートの隠しシート「_kakeibo_app」に入る。
 * 「@OnlyCurrentDoc」で、このスプレッドシート以外にはさわれない権限にしている。
 *
 * @OnlyCurrentDoc
 */
const KEY = 'ここに合言葉';
const SHEET = '_kakeibo_app';
const CHUNK = 40000;   // 1つのセルに入る文字数（上限 50,000）より少なく

// 読み書き（アプリからの通常の通信。フォームで送られたときは payload に入っている）
function doPost(e) {
  let req = {};
  try {
    const raw = e && e.parameter && e.parameter.payload ? e.parameter.payload : e.postData.contents;
    req = JSON.parse(raw);
  } catch (err) { return out_({ ok: false, error: 'bad' }); }
  return out_(handle_(req));
}

// ブラウザで URL を開いたときの確認用（データは出さない）と、
// ブラウザが別サイトへの通信を止めたときの予備（JSONP：callback を付けて読む）
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.callback) {
    if (!/^[A-Za-z0-9_]{1,64}$/.test(p.callback)) return out_({ ok: false, error: 'bad' });
    const res = p.action ? handle_({ action: p.action, key: p.key }) : { ok: true, app: 'kakeibo' };
    return ContentService.createTextOutput(p.callback + '(' + JSON.stringify(res) + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return out_({ ok: true, app: 'kakeibo' });
}

function handle_(req) {
  if (!KEY || KEY === 'ここに合言葉' || req.key !== KEY) return { ok: false, error: 'key' };
  if (req.action === 'get') return { ok: true, doc: read_() };
  if (req.action === 'put') {
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const cur = read_();
      // 相手が先に書き込んでいたら、上書きせずに今の内容を返す
      if (cur && req.baseRev !== undefined && req.baseRev !== null && cur.rev !== req.baseRev) return { ok: false, error: 'conflict', doc: cur };
      const doc = { rev: (cur ? cur.rev : 0) + 1, at: new Date().toISOString(), state: req.state };
      write_(doc);
      return { ok: true, rev: doc.rev, at: doc.at };
    } finally {
      lock.releaseLock();
    }
  }
  return { ok: false, error: 'action' };
}

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET);
  if (!sh) { sh = ss.insertSheet(SHEET); sh.hideSheet(); }
  return sh;
}

// セルの先頭に「k」を付けて、数字や数式として読まれないようにしている
function read_() {
  const sh = sheet_();
  const n = sh.getLastRow();
  if (!n) return null;
  const txt = sh.getRange(1, 1, n, 1).getValues().map(r => String(r[0]).slice(1)).join('');
  return txt ? JSON.parse(txt) : null;
}

function write_(doc) {
  const sh = sheet_();
  const txt = JSON.stringify(doc);
  const rows = [];
  for (let i = 0; i < txt.length; i += CHUNK) rows.push(['k' + txt.slice(i, i + CHUNK)]);
  sh.clearContents();
  if (rows.length) sh.getRange(1, 1, rows.length, 1).setValues(rows);
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
