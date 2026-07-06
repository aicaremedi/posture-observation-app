/**
 * 姿勢観察支援アプリ ─ Google Apps Script バックエンド
 *
 * 機能:
 *  - doPost: アプリからのデータ受信（画像/動画/骨格JSON/測定値）
 *  - Google Driveへの自動保存（利用者別・カテゴリ別フォルダ）
 *  - 測定値のスプレッドシート自動追記
 *  - cleanupOldVideos: 3ヶ月より古い動画の自動削除（トリガーで毎日実行）
 *
 * デプロイ方法は docs/GAS設定手順.md を参照。
 */

// ルートフォルダ名（Google Driveのマイドライブ直下に自動作成される）
const ROOT_FOLDER_NAME = '姿勢観察アプリデータ';

// 動画の保存期間（日数）。これより古いWebM動画は cleanupOldVideos で削除される
const VIDEO_RETENTION_DAYS = 90;

/**
 * アプリからのPOSTリクエストを受け取るエントリポイント
 */
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);

    let result;
    switch (payload.action) {
      case 'saveFile':
        result = saveFile(payload);
        break;
      case 'saveMeasurement':
        result = saveMeasurement(payload);
        break;
      case 'saveAnnotation':
        result = saveAnnotation(payload);
        break;
      case 'listUsers':
        result = listUsers();
        break;
      case 'addUser':
        result = addUser(payload);
        break;
      case 'listRecords':
        result = listRecords(payload);
        break;
      case 'getFileData':
        result = getFileData(payload);
        break;
      default:
        result = { status: 'error', message: '不明なaction: ' + payload.action };
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: String(err)
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * GETリクエスト:
 *  - ?page=report&user=名前&key=トークン → 評価表ページ（家族・ケアマネ閲覧用）
 *  - それ以外 → 動作確認用のステータスJSON
 */
function doGet(e) {
  if (e && e.parameter && e.parameter.page === 'report') {
    return renderReport(e.parameter.user, e.parameter.key);
  }
  return ContentService.createTextOutput(JSON.stringify({
    status: 'ok',
    app: '姿勢観察支援アプリ バックエンド',
    time: new Date().toISOString()
  })).setMimeType(ContentService.MimeType.JSON);
}

// ===== 利用者管理 =====

function getUsersSheet() {
  const root = getOrCreateRootFolder();
  const ss = getOrCreateSpreadsheet(root, '利用者一覧');
  const sheet = ss.getSheets()[0];
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['利用者名', 'トークン', '作成日']);
    sheet.getRange(1, 1, 1, 3).setFontWeight('bold');
  }
  return sheet;
}

function listUsers() {
  const sheet = getUsersSheet();
  const last = sheet.getLastRow();
  const users = [];
  if (last >= 2) {
    const values = sheet.getRange(2, 1, last - 1, 2).getValues();
    values.forEach(function(row) {
      if (row[0]) users.push({ name: String(row[0]), token: String(row[1]) });
    });
  }
  return { status: 'ok', users: users };
}

function addUser(payload) {
  const name = String(payload.name || '').trim();
  if (!name) return { status: 'error', message: '利用者名が空です' };

  const existing = listUsers().users;
  for (var i = 0; i < existing.length; i++) {
    if (existing[i].name === name) return { status: 'ok', user: existing[i], existed: true };
  }

  const token = Utilities.getUuid().replace(/-/g, '');
  getUsersSheet().appendRow([name, token, new Date()]);
  getOrCreateUserFolder(name);
  return { status: 'ok', user: { name: name, token: token } };
}

// ===== クラウド記録一覧・取得 =====

function listRecords(payload) {
  const user = String(payload.user || '').trim();
  if (!user) return { status: 'error', message: '利用者名が空です' };

  const root = getOrCreateRootFolder();
  const it = root.getFoldersByName(user);
  if (!it.hasNext()) return { status: 'ok', records: [] };
  const userFolder = it.next();

  const categories = ['姿勢観察', '動作観察', '簡易検査', 'セラピスト助言'];
  const records = [];
  categories.forEach(function(cat) {
    const cit = userFolder.getFoldersByName(cat);
    if (!cit.hasNext()) return;
    const files = cit.next().getFiles();
    while (files.hasNext()) {
      const f = files.next();
      records.push({
        id: f.getId(),
        name: f.getName(),
        category: cat,
        mimeType: f.getMimeType(),
        date: f.getDateCreated().toISOString(),
        size: f.getSize(),
        url: f.getUrl()
      });
    }
  });
  records.sort(function(a, b) { return a.date < b.date ? 1 : -1; });
  return { status: 'ok', records: records.slice(0, 100) };
}

function getFileData(payload) {
  const file = DriveApp.getFileById(String(payload.fileId));
  if (!isInsideRootFolder(file)) {
    return { status: 'error', message: 'アクセスできないファイルです' };
  }
  if (file.getSize() > 25 * 1024 * 1024) {
    return { status: 'error', message: 'ファイルが大きすぎます。Google Driveで直接開いてください。', url: file.getUrl() };
  }
  const blob = file.getBlob();
  return {
    status: 'ok',
    base64: Utilities.base64Encode(blob.getBytes()),
    mimeType: blob.getContentType(),
    name: file.getName()
  };
}

/** アプリのルートフォルダ配下のファイルかを確認（他ファイルへのアクセスを防ぐ） */
function isInsideRootFolder(file) {
  var parents = file.getParents();
  var current = parents.hasNext() ? parents.next() : null;
  var depth = 0;
  while (current && depth < 10) {
    if (current.getName() === ROOT_FOLDER_NAME) return true;
    var p = current.getParents();
    current = p.hasNext() ? p.next() : null;
    depth++;
  }
  return false;
}

// ===== 評価表ページ（家族・ケアマネ閲覧用） =====

function renderReport(user, key) {
  user = String(user || '');
  key = String(key || '');

  // トークン検証
  var valid = false;
  const users = listUsers().users;
  for (var i = 0; i < users.length; i++) {
    if (users[i].name === user && users[i].token === key && key.length > 10) { valid = true; break; }
  }
  if (!valid) {
    return HtmlService.createHtmlOutput(
      '<div style="font-family:sans-serif;text-align:center;padding:48px;">' +
      '<h3>リンクが無効です</h3><p>QRコードを再度読み込むか、施設にお問い合わせください。</p></div>');
  }

  const root = getOrCreateRootFolder();
  const uit = root.getFoldersByName(user);
  const userFolder = uit.hasNext() ? uit.next() : null;

  // 最新の写真（棒人間付きスクショ・助言画像）最大6枚
  var imgHtml = '';
  if (userFolder) {
    const recs = listRecords({ user: user }).records
      .filter(function(r) { return r.mimeType === 'image/png'; })
      .slice(0, 6);
    recs.forEach(function(r) {
      try {
        const f = DriveApp.getFileById(r.id);
        const b64 = Utilities.base64Encode(f.getBlob().getBytes());
        const d = new Date(r.date);
        const dateStr = d.getFullYear() + '/' + (d.getMonth() + 1) + '/' + d.getDate();
        imgHtml +=
          '<div class="photo"><img src="data:image/png;base64,' + b64 + '">' +
          '<div class="cap">' + escapeHtml(r.category) + ' ─ ' + dateStr + '</div></div>';
      } catch (err) { /* 1枚の失敗で全体を止めない */ }
    });
  }
  if (!imgHtml) imgHtml = '<p class="empty">まだ写真がありません</p>';

  // セラピスト助言（最新10件）
  var adviceHtml = '';
  const adviceRows = userFolder ? readSheetRows(userFolder, '助言記録', 10) : [];
  adviceRows.forEach(function(row) {
    adviceHtml += '<tr><td>' + escapeHtml(row[0]) + '</td><td>' + escapeHtml(row[2]) + '</td><td>' + escapeHtml(row[3]) + '</td></tr>';
  });
  if (!adviceHtml) adviceHtml = '<tr><td colspan="3" class="empty">まだ助言がありません</td></tr>';

  // 測定値（最新10件）
  var measureHtml = '';
  const measureRows = userFolder ? readSheetRows(userFolder, '測定データ', 10) : [];
  measureRows.forEach(function(row) {
    measureHtml += '<tr><td>' + escapeHtml(row[0]) + '</td><td>' + escapeHtml(row[1]) + '</td><td>' + escapeHtml(row[2]) + '</td><td>' + escapeHtml(row[3]) + '</td></tr>';
  });
  if (!measureHtml) measureHtml = '<tr><td colspan="4" class="empty">まだ測定値がありません</td></tr>';

  const now = new Date();
  const html =
    '<!DOCTYPE html><html lang="ja"><head><meta charset="UTF-8">' +
    '<meta name="viewport" content="width=device-width, initial-scale=1.0">' +
    '<title>' + escapeHtml(user) + 'さんの評価表</title><style>' +
    'body{font-family:"Hiragino Kaku Gothic ProN","Noto Sans JP",sans-serif;margin:0;background:#f9fafb;color:#1f2937;}' +
    '.header{background:linear-gradient(135deg,#1e3a5f,#2563eb);color:#fff;padding:24px 16px;text-align:center;}' +
    '.header h1{font-size:1.3rem;margin:0 0 4px;}.header p{margin:0;opacity:.8;font-size:.85rem;}' +
    '.wrap{max-width:720px;margin:0 auto;padding:16px;}' +
    '.card{background:#fff;border:1px solid #e5e7eb;border-radius:12px;padding:16px;margin-bottom:16px;}' +
    '.card h2{font-size:1.05rem;color:#2563eb;border-bottom:2px solid #dbeafe;padding-bottom:6px;margin:0 0 12px;}' +
    '.photos{display:grid;grid-template-columns:1fr 1fr;gap:10px;}' +
    '.photo img{width:100%;border-radius:8px;display:block;}' +
    '.cap{font-size:.75rem;color:#4b5563;margin-top:4px;text-align:center;}' +
    'table{width:100%;border-collapse:collapse;font-size:.85rem;}' +
    'th,td{border:1px solid #e5e7eb;padding:6px 8px;text-align:left;}th{background:#f3f4f6;}' +
    '.empty{color:#9ca3af;text-align:center;}' +
    '.footer{text-align:center;font-size:.75rem;color:#9ca3af;padding:16px;}' +
    '@media(min-width:600px){.photos{grid-template-columns:1fr 1fr 1fr;}}' +
    '</style></head><body>' +
    '<div class="header"><h1>' + escapeHtml(user) + ' さんの評価表</h1><p>姿勢観察支援アプリ ─ ' +
    now.getFullYear() + '/' + (now.getMonth() + 1) + '/' + now.getDate() + ' 時点</p></div>' +
    '<div class="wrap">' +
    '<div class="card"><h2>📷 最近の姿勢記録</h2><div class="photos">' + imgHtml + '</div></div>' +
    '<div class="card"><h2>💬 セラピストからの助言</h2><table><tr><th>日時</th><th>テンプレート</th><th>コメント</th></tr>' + adviceHtml + '</table></div>' +
    '<div class="card"><h2>📊 測定値の推移</h2><table><tr><th>日時</th><th>項目</th><th>値</th><th>補足</th></tr>' + measureHtml + '</table></div>' +
    '</div><div class="footer">このページはQRコードを知っている方のみ閲覧できます</div>' +
    '</body></html>';

  return HtmlService.createHtmlOutput(html)
    .setTitle(user + 'さんの評価表')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

/** 利用者フォルダ内のスプレッドシートから最新maxRows行を取得（新しい順） */
function readSheetRows(userFolder, name, maxRows) {
  const it = userFolder.getFilesByName(name);
  if (!it.hasNext()) return [];
  const ss = SpreadsheetApp.open(it.next());
  const sheet = ss.getSheets()[0];
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const start = Math.max(2, last - maxRows + 1);
  return sheet.getRange(start, 1, last - start + 1, sheet.getLastColumn()).getDisplayValues().reverse();
}

function escapeHtml(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

/**
 * 画像・動画・骨格JSONをGoogle Driveに保存
 */
function saveFile(payload) {
  const userFolder = getOrCreateUserFolder(payload.user);
  const category = categoryFromTemplate(payload.template);
  const categoryFolder = getOrCreateFolder(userFolder, category);

  // メインファイル（画像 or 動画）
  const bytes = Utilities.base64Decode(payload.dataBase64);
  const blob = Utilities.newBlob(bytes, payload.mimeType, payload.filename);
  const file = categoryFolder.createFile(blob);

  // 骨格データJSON（画像撮影時に同時保存）
  let skeletonUrl = null;
  if (payload.skeleton) {
    const jsonName = payload.filename.replace(/\.(png|webm)$/, '') + '_skeleton.json';
    const jsonBlob = Utilities.newBlob(
      JSON.stringify(payload.skeleton, null, 2), 'application/json', jsonName);
    const jsonFile = categoryFolder.createFile(jsonBlob);
    skeletonUrl = jsonFile.getUrl();
  }

  return {
    status: 'ok',
    fileUrl: file.getUrl(),
    skeletonUrl: skeletonUrl
  };
}

/**
 * 測定値（体重など）を利用者のスプレッドシートに追記
 */
function saveMeasurement(payload) {
  const userFolder = getOrCreateUserFolder(payload.user);
  const ss = getOrCreateSpreadsheet(userFolder, '測定データ');
  const sheet = ss.getSheets()[0];

  // 初回はヘッダー行を作成
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['日時', '項目', '値', '補足']);
    sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
  }

  sheet.appendRow([
    new Date(payload.timestamp),
    payload.item,
    payload.value,
    payload.extra || ''
  ]);

  return { status: 'ok' };
}

/**
 * セラピスト助言（コメント＋マーカー描画済み画像）を保存
 * - 画像: 利用者フォルダ内「セラピスト助言」に保存
 * - コメント: 「助言記録」スプレッドシートに追記
 */
function saveAnnotation(payload) {
  const userFolder = getOrCreateUserFolder(payload.user);

  // マーカー描画済み画像（あれば）
  let imageUrl = '';
  if (payload.dataBase64) {
    const annoFolder = getOrCreateFolder(userFolder, 'セラピスト助言');
    const bytes = Utilities.base64Decode(payload.dataBase64);
    const blob = Utilities.newBlob(bytes, 'image/png', payload.filename);
    const file = annoFolder.createFile(blob);
    imageUrl = file.getUrl();
  }

  // コメントをスプレッドシートに追記
  const ss = getOrCreateSpreadsheet(userFolder, '助言記録');
  const sheet = ss.getSheets()[0];
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['助言日時', '対象記録の撮影日時', 'テンプレート', 'コメント', '画像URL']);
    sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
  }
  sheet.appendRow([
    new Date(payload.timestamp),
    payload.recordTimestamp ? new Date(payload.recordTimestamp) : '',
    payload.template || '',
    payload.comment || '',
    imageUrl
  ]);

  return { status: 'ok', imageUrl: imageUrl };
}

/**
 * 3ヶ月より古い動画（WebM）を削除する。
 * ※時間主導型トリガーで1日1回実行するよう設定する（手順書参照）
 */
function cleanupOldVideos() {
  const root = getOrCreateRootFolder();
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - VIDEO_RETENTION_DAYS);

  let deleted = 0;
  deleted += cleanupFolderRecursive(root, cutoff);
  Logger.log('削除した動画: ' + deleted + '件');
  return deleted;
}

function cleanupFolderRecursive(folder, cutoff) {
  let deleted = 0;

  const files = folder.getFilesByType('video/webm');
  while (files.hasNext()) {
    const file = files.next();
    if (file.getDateCreated() < cutoff) {
      file.setTrashed(true);
      deleted++;
    }
  }

  const subFolders = folder.getFolders();
  while (subFolders.hasNext()) {
    deleted += cleanupFolderRecursive(subFolders.next(), cutoff);
  }
  return deleted;
}

// ===== フォルダ操作ヘルパー =====

function getOrCreateRootFolder() {
  return getOrCreateFolderIn(DriveApp.getRootFolder(), ROOT_FOLDER_NAME);
}

function getOrCreateUserFolder(userName) {
  const root = getOrCreateRootFolder();
  return getOrCreateFolder(root, userName || '未設定');
}

function getOrCreateFolder(parent, name) {
  return getOrCreateFolderIn(parent, name);
}

function getOrCreateFolderIn(parent, name) {
  const it = parent.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return parent.createFolder(name);
}

function getOrCreateSpreadsheet(folder, name) {
  const it = folder.getFilesByName(name);
  if (it.hasNext()) {
    return SpreadsheetApp.open(it.next());
  }
  const ss = SpreadsheetApp.create(name);
  const file = DriveApp.getFileById(ss.getId());
  file.moveTo(folder);
  return ss;
}

/**
 * テンプレートコードからカテゴリ名を判定
 */
function categoryFromTemplate(code) {
  if (!code) return 'その他';
  if (code.indexOf('A') === 0) return '姿勢観察';
  if (code.indexOf('B') === 0) return '動作観察';
  if (code.indexOf('C') === 0) return '簡易検査';
  return 'その他';
}
