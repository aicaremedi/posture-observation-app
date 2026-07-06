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
 * 動作確認用（ブラウザでGAS URLを開くと表示される）
 */
function doGet() {
  return ContentService.createTextOutput(JSON.stringify({
    status: 'ok',
    app: '姿勢観察支援アプリ バックエンド',
    time: new Date().toISOString()
  })).setMimeType(ContentService.MimeType.JSON);
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
