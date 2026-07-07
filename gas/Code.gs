/**
 * 姿勢観察支援アプリ ─ Google Apps Script バックエンド v2
 *
 * アカウント体系:
 *   システムアカウント（AI Care Medi）→ 法人 → 事業所 → 利用者
 *
 * 機能:
 *  - 事業所ログイン認証（事業所ID + パスワード → トークン発行）
 *  - システム管理API（法人・事業所の発行/停止/パスワード再発行）
 *  - Google Driveへの自動保存（法人/事業所/利用者/カテゴリ別フォルダ）
 *  - 評価表ページ配信（?page=report&fid=..&user=..&key=..）
 *  - cleanupOldVideos: 90日より古い動画の自動削除（日次トリガー）
 *
 * 管理DB: マイドライブ >「姿勢観察アプリデータ」>「_システム管理」スプレッドシート
 *   - 法人シート / 事業所シート / 設定シート（システム管理者パスワード）
 */

const ROOT_FOLDER_NAME = '姿勢観察アプリデータ';
const VIDEO_RETENTION_DAYS = 90;
const SYSTEM_SS_NAME = '_システム管理';
const TOKEN_SALT = 'posture-app-token-v2';

// ===== エントリポイント =====

function doPost(e) {
  try {
    const p = JSON.parse(e.postData.contents);
    let r;
    switch (p.action) {
      // 認証不要
      case 'adminInit': r = adminInit(p); break;
      case 'adminLogin': r = adminLogin(p); break;
      case 'login': r = facilityLogin(p); break;
      // システム管理者専用
      case 'adminOverview': r = withAdmin(p, adminOverview); break;
      case 'adminListCorps': r = withAdmin(p, adminListCorps); break;
      case 'adminAddCorp': r = withAdmin(p, adminAddCorp); break;
      case 'adminUpdateCorp': r = withAdmin(p, adminUpdateCorp); break;
      case 'adminSetCorpStatus': r = withAdmin(p, adminSetCorpStatus); break;
      case 'adminListFacilities': r = withAdmin(p, adminListFacilities); break;
      case 'adminAddFacility': r = withAdmin(p, adminAddFacility); break;
      case 'adminSetFacilityStatus': r = withAdmin(p, adminSetFacilityStatus); break;
      case 'adminResetPassword': r = withAdmin(p, adminResetPassword); break;
      case 'adminResetCorpPassword': r = withAdmin(p, adminResetCorpPassword); break;
      case 'adminSaveSettings': r = withAdmin(p, adminSaveSettings); break;
      case 'adminGenerateInvoices': r = withAdmin(p, adminGenerateInvoices); break;
      case 'adminSetInvoiceStatus': r = withAdmin(p, adminSetInvoiceStatus); break;
      case 'adminDeleteInvoice': r = withAdmin(p, adminDeleteInvoice); break;
      // 法人ポータル（要法人ログイン）
      case 'corpLogin': r = corpLogin(p); break;
      case 'corpOverview': r = withCorp(p, corpOverview); break;
      case 'corpListUsers': r = withCorp(p, corpListUsers); break;
      case 'corpListRecords': r = withCorp(p, corpListRecords); break;
      case 'corpGetFileData': r = withCorp(p, corpGetFileData); break;
      case 'corpResetFacilityPassword': r = withCorp(p, corpResetFacilityPassword); break;
      // 事業所（要ログイン）
      case 'listUsers': r = withFacility(p, apiListUsers); break;
      case 'addUser': r = withFacility(p, apiAddUser); break;
      case 'listRecords': r = withFacility(p, apiListRecords); break;
      case 'getFileData': r = withFacility(p, apiGetFileData); break;
      case 'saveFile': r = withFacility(p, apiSaveFile); break;
      case 'saveMeasurement': r = withFacility(p, apiSaveMeasurement); break;
      case 'saveAnnotation': r = withFacility(p, apiSaveAnnotation); break;
      default: r = { status: 'error', message: '不明なaction: ' + p.action };
    }
    return jsonOut(r);
  } catch (err) {
    return jsonOut({ status: 'error', message: String(err) });
  }
}

function doGet(e) {
  if (e && e.parameter && e.parameter.page === 'report') {
    return renderReport(e.parameter.fid, e.parameter.user, e.parameter.key);
  }
  return jsonOut({
    status: 'ok',
    app: '姿勢観察支援アプリ バックエンド v2',
    time: new Date().toISOString()
  });
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ===== システム管理DB =====
// 高速化: 1回の実行内ではスプレッドシート・各シートの内容をキャッシュして
// Drive検索/シートアクセスを最小限にする

var _cache = { ss: null, db: null };
const DB_CACHE_KEY = 'systemDb_v2';
const DB_CACHE_SEC = 300; // 5分（更新系の操作時は即時無効化する）
const TAX_RATE = 0.10; // 消費税率

function getSystemSS() {
  if (_cache.ss) return _cache.ss;

  // スプレッドシートIDを記憶してDrive検索を省略（2回目以降は直接オープン）
  const props = PropertiesService.getScriptProperties();
  const savedId = props.getProperty('systemSSId');
  if (savedId) {
    try {
      _cache.ss = SpreadsheetApp.openById(savedId);
      return _cache.ss;
    } catch (e) { /* 削除された場合は下の作成パスへ */ }
  }

  const root = getOrCreateRootFolder();
  const ss = getOrCreateSpreadsheet(root, SYSTEM_SS_NAME);
  ensureSheet(ss, '法人', ['法人ID', '法人名', '状態', '作成日']);
  const fSheet = ensureSheet(ss, '事業所',
    ['事業所ID', '事業所名', '法人ID', 'パスワードハッシュ', 'ソルト', '状態', '作成日', 'パスワード']);
  // 既存シートに「パスワード」列がない場合は追加（後方互換）
  if (String(fSheet.getRange(1, 8).getValue()) !== 'パスワード') {
    fSheet.getRange(1, 8).setValue('パスワード').setFontWeight('bold');
  }
  ensureSheet(ss, '設定', ['キー', '値']);
  props.setProperty('systemSSId', ss.getId());
  _cache.ss = ss;
  return ss;
}

/**
 * 管理DB（設定/法人/事業所）をまとめて読み込む。
 * CacheServiceで5分間キャッシュされるため、読み取り系のリクエストは
 * スプレッドシートを開かずに応答できる（大幅な高速化）。
 */
function loadDb() {
  if (_cache.db) return _cache.db;

  const cached = CacheService.getScriptCache().get(DB_CACHE_KEY);
  if (cached) {
    _cache.db = JSON.parse(cached);
    return _cache.db;
  }

  const ss = getSystemSS();
  ensureExtendedStructure(ss);
  const settings = {};
  sheetRows(ss.getSheetByName('設定')).forEach(function(r) {
    settings[String(r[0])] = String(r[1]);
  });
  const db = {
    settings: settings,
    corps: sheetRows(ss.getSheetByName('法人')).map(function(r) { return r.map(String); }),
    facilities: sheetRows(ss.getSheetByName('事業所')).map(function(r) { return r.map(String); }),
    invoices: sheetRows(ss.getSheetByName('請求')).map(function(r) { return r.map(String); })
  };
  try {
    CacheService.getScriptCache().put(DB_CACHE_KEY, JSON.stringify(db), DB_CACHE_SEC);
  } catch (e) { /* キャッシュ容量超過時は都度読み込みにフォールバック */ }
  _cache.db = db;
  return db;
}

function invalidateDb() {
  _cache.db = null;
  CacheService.getScriptCache().remove(DB_CACHE_KEY);
}

/** 法人シートの拡張列と請求シートを保証（キャッシュ再構築時のみ実行される） */
function ensureExtendedStructure(ss) {
  const corpSheet = ss.getSheetByName('法人');
  if (corpSheet && String(corpSheet.getRange(1, 5).getValue()) !== '担当者') {
    corpSheet.getRange(1, 5, 1, 6)
      .setValues([['担当者', 'メール', '電話', '住所', '単価税抜', '備考']])
      .setFontWeight('bold');
  }
  // 保存単位列（'事業所' = 事業所ごとにフォルダ分割 / '法人' = 法人でまとめて共有）
  if (corpSheet && String(corpSheet.getRange(1, 11).getValue()) !== '保存単位') {
    corpSheet.getRange(1, 11).setValue('保存単位').setFontWeight('bold');
  }
  // 法人ポータルのログイン情報列
  if (corpSheet && String(corpSheet.getRange(1, 12).getValue()) !== 'PWハッシュ') {
    corpSheet.getRange(1, 12, 1, 3)
      .setValues([['PWハッシュ', 'PWソルト', 'パスワード']])
      .setFontWeight('bold');
  }
  ensureSheet(ss, '請求',
    ['請求ID', '法人ID', '対象月', '事業所数', '単価税抜', '税抜額', '消費税', '税込額', '状態', '発行日', '支払期限']);
}

function corpRows() { return loadDb().corps; }
function facilityRows() { return loadDb().facilities; }

function ensureSheet(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  return sheet;
}

function sheetRows(sheet) {
  const last = sheet.getLastRow();
  if (last < 2) return [];
  return sheet.getRange(2, 1, last - 1, sheet.getLastColumn()).getValues();
}

function getSetting(key) {
  const settings = loadDb().settings;
  return settings.hasOwnProperty(key) ? settings[key] : null;
}

function setSetting(key, value) {
  const sheet = getSystemSS().getSheetByName('設定');
  const rows = sheetRows(sheet);
  invalidateDb();
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0]) === key) { sheet.getRange(i + 2, 2).setValue(value); return; }
  }
  sheet.appendRow([key, value]);
}

// ===== ハッシュ・トークン =====

function sha256hex(text) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8)
    .map(function(b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}

function randomToken() { return Utilities.getUuid().replace(/-/g, ''); }

function randomPassword() {
  // 紛らわしい文字（l,1,o,0,i）を除いた8文字
  const chars = 'abcdefghjkmnpqrstuvwxyz23456789';
  var pw = '';
  for (var i = 0; i < 8; i++) pw += chars.charAt(Math.floor(Math.random() * chars.length));
  return pw;
}

function facilityTokenFor(facilityId, passwordHash) {
  return sha256hex(facilityId + ':' + passwordHash + ':' + TOKEN_SALT);
}

// ===== システム管理者認証 =====

function adminInit(p) {
  if (getSetting('adminHash')) return { status: 'error', message: 'システムパスワードは設定済みです' };
  const pw = String(p.password || '');
  if (pw.length < 8) return { status: 'error', message: 'パスワードは8文字以上にしてください' };
  const salt = randomToken();
  setSetting('adminSalt', salt);
  setSetting('adminHash', sha256hex(salt + pw));
  return { status: 'ok', adminToken: adminTokenNow() };
}

function adminLogin(p) {
  const hash = getSetting('adminHash');
  const salt = getSetting('adminSalt');
  if (!hash) return { status: 'error', code: 'NOINIT', message: '初期設定が必要です' };
  if (sha256hex(salt + String(p.password || '')) !== hash) {
    return { status: 'error', message: 'パスワードが違います' };
  }
  return { status: 'ok', adminToken: adminTokenNow() };
}

function adminTokenNow() {
  return sha256hex('admin:' + getSetting('adminHash') + ':' + TOKEN_SALT);
}

function withAdmin(p, fn) {
  if (!getSetting('adminHash') || String(p.adminToken || '') !== adminTokenNow()) {
    return { status: 'error', code: 'AUTH', message: '管理者認証エラー。再ログインしてください' };
  }
  return fn(p);
}

// ===== 法人・事業所管理（システム管理者） =====

function adminOverview(p) {
  const db = loadDb();
  const billSettings = {};
  ['bill_company', 'bill_address', 'bill_tel', 'bill_bank', 'bill_invoiceNo', 'bill_dueDays', 'bill_note']
    .forEach(function(k) { billSettings[k] = db.settings[k] || ''; });
  return {
    status: 'ok',
    corps: adminListCorps(p).corps,
    facilities: adminListFacilities(p).facilities,
    invoices: adminListInvoices(),
    billSettings: billSettings
  };
}

function adminListCorps(p) {
  const rows = corpRows();
  return {
    status: 'ok',
    corps: rows.map(function(r) {
      return {
        corpId: String(r[0]),
        name: String(r[1]),
        active: String(r[2]) === '有効',
        contact: String(r[4] || ''),
        email: String(r[5] || ''),
        tel: String(r[6] || ''),
        address: String(r[7] || ''),
        unitPrice: Number(r[8] || 0),
        note: String(r[9] || ''),
        storageScope: String(r[10] || '事業所'),
        portalPassword: String(r[13] || '')
      };
    })
  };
}

function adminUpdateCorp(p) {
  const sheet = getSystemSS().getSheetByName('法人');
  const rows = sheetRows(sheet);
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0]) === String(p.corpId)) {
      sheet.getRange(i + 2, 5, 1, 6).setValues([[
        String(p.contact || ''), String(p.email || ''), String(p.tel || ''),
        String(p.address || ''), Number(p.unitPrice || 0), String(p.note || '')
      ]]);
      sheet.getRange(i + 2, 11).setValue(p.storageScope === '法人' ? '法人' : '事業所');
      invalidateDb();
      return { status: 'ok' };
    }
  }
  return { status: 'error', message: '法人が見つかりません' };
}

function adminSetCorpStatus(p) {
  const sheet = getSystemSS().getSheetByName('法人');
  const rows = sheetRows(sheet);
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0]) === String(p.corpId)) {
      sheet.getRange(i + 2, 3).setValue(p.active ? '有効' : '停止');
      invalidateDb();
      return { status: 'ok' };
    }
  }
  return { status: 'error', message: '法人が見つかりません' };
}

// ===== 請求元設定 =====

function adminSaveSettings(p) {
  const s = p.settings || {};
  ['bill_company', 'bill_address', 'bill_tel', 'bill_bank', 'bill_invoiceNo', 'bill_dueDays', 'bill_note']
    .forEach(function(k) {
      if (s.hasOwnProperty(k)) setSetting(k, String(s[k]));
    });
  return { status: 'ok' };
}

// ===== 請求管理 =====

function adminListInvoices() {
  const corps = {};
  corpRows().forEach(function(r) { corps[String(r[0])] = String(r[1]); });
  return loadDb().invoices.map(function(r) {
    return {
      invoiceId: String(r[0]), corpId: String(r[1]), corpName: corps[String(r[1])] || '',
      month: String(r[2]), count: Number(r[3] || 0), unitPrice: Number(r[4] || 0),
      net: Number(r[5] || 0), tax: Number(r[6] || 0), total: Number(r[7] || 0),
      state: String(r[8] || ''), issuedAt: String(r[9] || ''), dueDate: String(r[10] || '')
    };
  });
}

/** 対象月（YYYY-MM）の請求書を有効法人に対して一括生成 */
function adminGenerateInvoices(p) {
  const month = String(p.month || '');
  if (!/^\d{4}-\d{2}$/.test(month)) return { status: 'error', message: '対象月の形式が不正です（YYYY-MM）' };

  const existing = {};
  loadDb().invoices.forEach(function(r) { existing[String(r[1]) + ':' + String(r[2])] = true; });

  // 法人ごとの有効事業所数
  const activeCount = {};
  facilityRows().forEach(function(r) {
    if (String(r[5]) === '有効') {
      const cid = String(r[2]);
      activeCount[cid] = (activeCount[cid] || 0) + 1;
    }
  });

  const dueDays = Number(getSetting('bill_dueDays') || 30);
  const sheet = getSystemSS().getSheetByName('請求');
  const today = new Date();
  const due = new Date(today.getTime() + dueDays * 24 * 60 * 60 * 1000);
  const fmt = function(d) {
    return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2);
  };

  var created = 0, skippedExisting = 0, skippedNoTarget = 0;
  corpRows().forEach(function(r) {
    const corpId = String(r[0]);
    if (String(r[2]) !== '有効') { skippedNoTarget++; return; }
    if (existing[corpId + ':' + month]) { skippedExisting++; return; }
    const unitPrice = Number(r[8] || 0);
    const count = activeCount[corpId] || 0;
    if (unitPrice <= 0 || count <= 0) { skippedNoTarget++; return; }

    const net = unitPrice * count;
    const tax = Math.floor(net * TAX_RATE);
    sheet.appendRow([
      'INV-' + month + '-' + corpId, corpId, month, count, unitPrice,
      net, tax, net + tax, '発行済', fmt(today), fmt(due)
    ]);
    created++;
  });
  invalidateDb();
  return { status: 'ok', created: created, skippedExisting: skippedExisting, skippedNoTarget: skippedNoTarget };
}

function adminSetInvoiceStatus(p) {
  const sheet = getSystemSS().getSheetByName('請求');
  const rows = sheetRows(sheet);
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0]) === String(p.invoiceId)) {
      sheet.getRange(i + 2, 9).setValue(String(p.state || '発行済'));
      invalidateDb();
      return { status: 'ok' };
    }
  }
  return { status: 'error', message: '請求書が見つかりません' };
}

function adminDeleteInvoice(p) {
  const sheet = getSystemSS().getSheetByName('請求');
  const rows = sheetRows(sheet);
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0]) === String(p.invoiceId)) {
      sheet.deleteRow(i + 2);
      invalidateDb();
      return { status: 'ok' };
    }
  }
  return { status: 'error', message: '請求書が見つかりません' };
}

function adminAddCorp(p) {
  const name = String(p.name || '').trim();
  if (!name) return { status: 'error', message: '法人名が空です' };
  const sheet = getSystemSS().getSheetByName('法人');
  const rows = sheetRows(sheet);
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][1]) === name) return { status: 'error', message: '同名の法人が既に存在します' };
  }
  var maxNum = 0;
  rows.forEach(function(r) {
    const m = String(r[0]).match(/^C(\d+)$/);
    if (m) maxNum = Math.max(maxNum, parseInt(m[1], 10));
  });
  const corpId = 'C' + ('000' + (maxNum + 1)).slice(-3);
  // 法人ポータルのログイン情報も同時に発行
  const password = randomPassword();
  const salt = randomToken();
  sheet.appendRow([corpId, name, '有効', new Date(),
    '', '', '', '', 0, '', '事業所',
    sha256hex(salt + password), salt, password]);
  invalidateDb();
  getOrCreateFolderIn(getOrCreateRootFolder(), name);
  return { status: 'ok', corp: { corpId: corpId, name: name, portalPassword: password } };
}

/** 法人ポータルのパスワード発行/再発行 */
function adminResetCorpPassword(p) {
  const sheet = getSystemSS().getSheetByName('法人');
  const rows = sheetRows(sheet);
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0]) === String(p.corpId)) {
      const password = randomPassword();
      const salt = randomToken();
      sheet.getRange(i + 2, 12, 1, 3).setValues([[sha256hex(salt + password), salt, password]]);
      invalidateDb();
      return { status: 'ok', corpId: String(p.corpId), password: password };
    }
  }
  return { status: 'error', message: '法人が見つかりません' };
}

function adminListFacilities(p) {
  const rows = facilityRows();
  const corps = {};
  corpRows().forEach(function(r) { corps[String(r[0])] = String(r[1]); });
  var list = rows.map(function(r) {
    return {
      facilityId: String(r[0]),
      name: String(r[1]),
      corpId: String(r[2]),
      corpName: corps[String(r[2])] || '',
      active: String(r[5]) === '有効',
      password: String(r[7] || '')
    };
  });
  if (p.corpId) list = list.filter(function(f) { return f.corpId === String(p.corpId); });
  return { status: 'ok', facilities: list };
}

function adminAddFacility(p) {
  const corpId = String(p.corpId || '').trim();
  const name = String(p.name || '').trim();
  if (!corpId || !name) return { status: 'error', message: '法人IDと事業所名を指定してください' };

  const corpRow = findCorp(corpId);
  if (!corpRow) return { status: 'error', message: '法人が見つかりません: ' + corpId };

  const sheet = getSystemSS().getSheetByName('事業所');
  const rows = sheetRows(sheet);
  const existingIds = {};
  rows.forEach(function(r) { existingIds[String(r[0])] = true; });

  var facilityId;
  do { facilityId = 'F' + ('0000' + Math.floor(Math.random() * 10000)).slice(-4); }
  while (existingIds[facilityId]);

  const password = randomPassword();
  const salt = randomToken();
  sheet.appendRow([facilityId, name, corpId, sha256hex(salt + password), salt, '有効', new Date(), password]);
  invalidateDb();

  // 法人フォルダ/事業所フォルダを自動生成
  const corpFolder = getOrCreateFolderIn(getOrCreateRootFolder(), String(corpRow[1]));
  getOrCreateFolderIn(corpFolder, name);

  return {
    status: 'ok',
    facility: { facilityId: facilityId, password: password, name: name, corpId: corpId, corpName: String(corpRow[1]) }
  };
}

function adminSetFacilityStatus(p) {
  const sheet = getSystemSS().getSheetByName('事業所');
  const rows = sheetRows(sheet);
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0]) === String(p.facilityId)) {
      sheet.getRange(i + 2, 6).setValue(p.active ? '有効' : '停止');
      invalidateDb();
      return { status: 'ok' };
    }
  }
  return { status: 'error', message: '事業所が見つかりません' };
}

function adminResetPassword(p) {
  const sheet = getSystemSS().getSheetByName('事業所');
  const rows = sheetRows(sheet);
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0]) === String(p.facilityId)) {
      const password = randomPassword();
      const salt = randomToken();
      sheet.getRange(i + 2, 4).setValue(sha256hex(salt + password));
      sheet.getRange(i + 2, 5).setValue(salt);
      sheet.getRange(i + 2, 8).setValue(password);
      invalidateDb();
      return { status: 'ok', facilityId: String(p.facilityId), password: password };
    }
  }
  return { status: 'error', message: '事業所が見つかりません' };
}

function findCorp(corpId) {
  const rows = corpRows();
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0]) === corpId) return rows[i];
  }
  return null;
}

function findFacility(facilityId) {
  const rows = facilityRows();
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0]) === String(facilityId)) return rows[i];
  }
  return null;
}

// ===== 法人ポータル（本部向け・閲覧＋最小限の管理操作） =====

function corpTokenFor(corpId, passwordHash) {
  return sha256hex('corp:' + corpId + ':' + passwordHash + ':' + TOKEN_SALT);
}

function corpLogin(p) {
  const row = findCorp(String(p.corpId || '').trim().toUpperCase());
  if (!row || !String(row[11] || '')) {
    return { status: 'error', message: '法人IDまたはパスワードが違います' };
  }
  if (String(row[2]) !== '有効') {
    return { status: 'error', message: 'このアカウントは停止中です。販売元にお問い合わせください' };
  }
  if (sha256hex(String(row[12]) + String(p.password || '')) !== String(row[11])) {
    return { status: 'error', message: '法人IDまたはパスワードが違います' };
  }
  return {
    status: 'ok',
    session: {
      corpId: String(row[0]),
      corpName: String(row[1]),
      token: corpTokenFor(String(row[0]), String(row[11]))
    }
  };
}

function withCorp(p, fn) {
  const row = findCorp(String(p.corpId || ''));
  if (!row || String(row[2]) !== '有効' || !String(row[11] || '') ||
      String(p.token || '') !== corpTokenFor(String(row[0]), String(row[11]))) {
    return { status: 'error', code: 'AUTH', message: '認証エラー。再ログインしてください' };
  }
  return fn(p, { corpRow: row, corpId: String(row[0]), corpName: String(row[1]) });
}

/** 法人配下の事業所のデータフォルダ（保存単位を考慮） */
function corpDataFolderFor(corpRow, facilityRow) {
  const corpFolder = getOrCreateFolderIn(getOrCreateRootFolder(), String(corpRow[1]));
  const scope = String(corpRow[10] || '事業所');
  return (scope === '法人') ? corpFolder : getOrCreateFolderIn(corpFolder, String(facilityRow[1]));
}

/** 事業所が自法人のものであることを確認して行を返す */
function corpOwnFacility(ctx, facilityId) {
  const row = findFacility(String(facilityId || ''));
  if (!row || String(row[2]) !== ctx.corpId) return null;
  return row;
}

function corpOverview(p, ctx) {
  const facilities = facilityRows()
    .filter(function(r) { return String(r[2]) === ctx.corpId; })
    .map(function(r) {
      return { facilityId: String(r[0]), name: String(r[1]), active: String(r[5]) === '有効' };
    });
  const invoices = adminListInvoices().filter(function(i) { return i.corpId === ctx.corpId; });
  const db = loadDb();
  const billSettings = {};
  ['bill_company', 'bill_address', 'bill_tel', 'bill_bank', 'bill_invoiceNo', 'bill_note']
    .forEach(function(k) { billSettings[k] = db.settings[k] || ''; });
  const c = ctx.corpRow;
  return {
    status: 'ok',
    corp: {
      corpId: ctx.corpId, name: ctx.corpName,
      contact: String(c[4] || ''), address: String(c[7] || ''),
      unitPrice: Number(c[8] || 0), storageScope: String(c[10] || '事業所')
    },
    facilities: facilities,
    invoices: invoices,
    billSettings: billSettings
  };
}

function corpListUsers(p, ctx) {
  const fRow = corpOwnFacility(ctx, p.facilityId);
  if (!fRow) return { status: 'error', message: '事業所が見つかりません' };
  const folder = corpDataFolderFor(ctx.corpRow, fRow);
  const users = [];
  sheetRows(getUsersSheetIn(folder)).forEach(function(r) {
    if (r[0]) users.push({ name: String(r[0]), token: String(r[1]) });
  });
  return { status: 'ok', users: users };
}

function corpListRecords(p, ctx) {
  const fRow = corpOwnFacility(ctx, p.facilityId);
  if (!fRow) return { status: 'error', message: '事業所が見つかりません' };
  const folder = corpDataFolderFor(ctx.corpRow, fRow);
  return { status: 'ok', records: listRecordsIn(folder, String(p.user || '')) };
}

function corpGetFileData(p, ctx) {
  const fRow = corpOwnFacility(ctx, p.facilityId);
  if (!fRow) return { status: 'error', message: '事業所が見つかりません' };
  const folder = corpDataFolderFor(ctx.corpRow, fRow);
  const file = DriveApp.getFileById(String(p.fileId));
  if (!isInsideFolder(file, folder)) {
    return { status: 'error', message: 'アクセスできないファイルです' };
  }
  if (file.getSize() > 25 * 1024 * 1024) {
    return { status: 'error', message: 'ファイルが大きすぎます', url: file.getUrl() };
  }
  const blob = file.getBlob();
  return {
    status: 'ok',
    base64: Utilities.base64Encode(blob.getBytes()),
    mimeType: blob.getContentType(),
    name: file.getName()
  };
}

/** 現場がパスワードを忘れた際に本部から再発行（自法人の事業所のみ） */
function corpResetFacilityPassword(p, ctx) {
  const fRow = corpOwnFacility(ctx, p.facilityId);
  if (!fRow) return { status: 'error', message: '事業所が見つかりません' };
  return adminResetPassword({ facilityId: p.facilityId });
}

// ===== 事業所ログイン =====

function facilityLogin(p) {
  const row = findFacility(p.facilityId);
  if (!row) return { status: 'error', message: '事業所IDまたはパスワードが違います' };
  if (String(row[5]) !== '有効') return { status: 'error', message: 'このアカウントは停止中です。管理者にお問い合わせください' };
  const hash = sha256hex(String(row[4]) + String(p.password || ''));
  if (hash !== String(row[3])) return { status: 'error', message: '事業所IDまたはパスワードが違います' };

  const corp = findCorp(String(row[2]));
  return {
    status: 'ok',
    session: {
      facilityId: String(row[0]),
      facilityName: String(row[1]),
      corpName: corp ? String(corp[1]) : '',
      token: facilityTokenFor(String(row[0]), String(row[3]))
    }
  };
}

function withFacility(p, fn) {
  const row = findFacility(p.facilityId);
  if (!row || String(row[5]) !== '有効' ||
      String(p.token || '') !== facilityTokenFor(String(row[0]), String(row[3]))) {
    return { status: 'error', code: 'AUTH', message: '認証エラー。再ログインしてください' };
  }
  const corp = findCorp(String(row[2]));
  const corpFolder = getOrCreateFolderIn(getOrCreateRootFolder(), corp ? String(corp[1]) : '不明法人');
  // 保存単位: '法人' なら法人フォルダを共有（同一法人の全事業所で利用者・記録を共有）
  const scope = corp ? String(corp[10] || '事業所') : '事業所';
  const facilityFolder = (scope === '法人') ? corpFolder : getOrCreateFolderIn(corpFolder, String(row[1]));
  return fn(p, { folder: facilityFolder, facilityId: String(row[0]), facilityName: String(row[1]) });
}

// ===== 利用者管理（事業所スコープ） =====

function getUsersSheetIn(facilityFolder) {
  const ss = getOrCreateSpreadsheet(facilityFolder, '利用者一覧');
  const sheet = ss.getSheets()[0];
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['利用者名', 'トークン', '作成日']);
    sheet.getRange(1, 1, 1, 3).setFontWeight('bold');
  }
  return sheet;
}

function apiListUsers(p, ctx) {
  const rows = sheetRows(getUsersSheetIn(ctx.folder));
  const users = [];
  rows.forEach(function(r) {
    if (r[0]) users.push({ name: String(r[0]), token: String(r[1]) });
  });
  return { status: 'ok', users: users };
}

function apiAddUser(p, ctx) {
  const name = String(p.name || '').trim();
  if (!name) return { status: 'error', message: '利用者名が空です' };

  const existing = apiListUsers(p, ctx).users;
  for (var i = 0; i < existing.length; i++) {
    if (existing[i].name === name) return { status: 'ok', user: existing[i], existed: true };
  }

  const token = randomToken();
  getUsersSheetIn(ctx.folder).appendRow([name, token, new Date()]);
  getOrCreateFolderIn(ctx.folder, name);
  return { status: 'ok', user: { name: name, token: token } };
}

// ===== 記録の保存・取得（事業所スコープ） =====

function apiSaveFile(p, ctx) {
  const userFolder = getOrCreateFolderIn(ctx.folder, String(p.user || '未設定'));
  const categoryFolder = getOrCreateFolderIn(userFolder, categoryFromTemplate(p.template));

  const bytes = Utilities.base64Decode(p.dataBase64);
  const blob = Utilities.newBlob(bytes, p.mimeType, p.filename);
  const file = categoryFolder.createFile(blob);

  var skeletonUrl = null;
  if (p.skeleton) {
    const jsonName = p.filename.replace(/\.(png|webm)$/, '') + '_skeleton.json';
    const jsonBlob = Utilities.newBlob(JSON.stringify(p.skeleton, null, 2), 'application/json', jsonName);
    skeletonUrl = categoryFolder.createFile(jsonBlob).getUrl();
  }
  return { status: 'ok', fileUrl: file.getUrl(), skeletonUrl: skeletonUrl };
}

function apiSaveMeasurement(p, ctx) {
  const userFolder = getOrCreateFolderIn(ctx.folder, String(p.user || '未設定'));
  const ss = getOrCreateSpreadsheet(userFolder, '測定データ');
  const sheet = ss.getSheets()[0];
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['日時', '項目', '値', '補足']);
    sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
  }
  sheet.appendRow([new Date(p.timestamp), p.item, p.value, p.extra || '']);
  return { status: 'ok' };
}

function apiSaveAnnotation(p, ctx) {
  const userFolder = getOrCreateFolderIn(ctx.folder, String(p.user || '未設定'));

  var imageUrl = '';
  if (p.dataBase64) {
    const annoFolder = getOrCreateFolderIn(userFolder, 'セラピスト助言');
    const bytes = Utilities.base64Decode(p.dataBase64);
    imageUrl = annoFolder.createFile(Utilities.newBlob(bytes, 'image/png', p.filename)).getUrl();
  }

  const ss = getOrCreateSpreadsheet(userFolder, '助言記録');
  const sheet = ss.getSheets()[0];
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['助言日時', '対象記録の撮影日時', 'テンプレート', 'コメント', '画像URL']);
    sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
  }
  sheet.appendRow([
    new Date(p.timestamp),
    p.recordTimestamp ? new Date(p.recordTimestamp) : '',
    p.template || '', p.comment || '', imageUrl
  ]);
  return { status: 'ok', imageUrl: imageUrl };
}

function apiListRecords(p, ctx) {
  return { status: 'ok', records: listRecordsIn(ctx.folder, String(p.user || '')) };
}

function listRecordsIn(facilityFolder, user) {
  if (!user) return [];
  const uit = facilityFolder.getFoldersByName(user);
  if (!uit.hasNext()) return [];
  const userFolder = uit.next();

  const categories = ['姿勢観察', '動作観察', '簡易検査', 'セラピスト助言'];
  const records = [];
  categories.forEach(function(cat) {
    const cit = userFolder.getFoldersByName(cat);
    if (!cit.hasNext()) return;
    const files = cit.next().getFiles();
    while (files.hasNext()) {
      const f = files.next();
      records.push({
        id: f.getId(), name: f.getName(), category: cat,
        mimeType: f.getMimeType(), date: f.getDateCreated().toISOString(),
        size: f.getSize(), url: f.getUrl()
      });
    }
  });
  records.sort(function(a, b) { return a.date < b.date ? 1 : -1; });
  return records.slice(0, 100);
}

function apiGetFileData(p, ctx) {
  const file = DriveApp.getFileById(String(p.fileId));
  if (!isInsideFolder(file, ctx.folder)) {
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

/** ファイルが指定フォルダ配下にあるか（他事業所のデータへのアクセスを防ぐ） */
function isInsideFolder(file, targetFolder) {
  const targetId = targetFolder.getId();
  var parents = file.getParents();
  var current = parents.hasNext() ? parents.next() : null;
  var depth = 0;
  while (current && depth < 10) {
    if (current.getId() === targetId) return true;
    var p = current.getParents();
    current = p.hasNext() ? p.next() : null;
    depth++;
  }
  return false;
}

// ===== 評価表ページ（家族・ケアマネ閲覧用） =====

function renderReport(fid, user, key) {
  fid = String(fid || '');
  user = String(user || '');
  key = String(key || '');

  const invalid = HtmlService.createHtmlOutput(
    '<div style="font-family:sans-serif;text-align:center;padding:48px;">' +
    '<h3>リンクが無効です</h3><p>QRコードを再度読み込むか、施設にお問い合わせください。</p></div>');

  const row = findFacility(fid);
  if (!row || String(row[5]) !== '有効') return invalid;
  const corp = findCorp(String(row[2]));
  const corpFolder = getOrCreateFolderIn(getOrCreateRootFolder(), corp ? String(corp[1]) : '不明法人');
  const scope = corp ? String(corp[10] || '事業所') : '事業所';
  const facilityFolder = (scope === '法人') ? corpFolder : getOrCreateFolderIn(corpFolder, String(row[1]));

  // 利用者トークン検証
  var valid = false;
  sheetRows(getUsersSheetIn(facilityFolder)).forEach(function(r) {
    if (String(r[0]) === user && String(r[1]) === key && key.length > 10) valid = true;
  });
  if (!valid) return invalid;

  const uit = facilityFolder.getFoldersByName(user);
  const userFolder = uit.hasNext() ? uit.next() : null;

  // 最新の写真 最大6枚
  var imgHtml = '';
  listRecordsIn(facilityFolder, user)
    .filter(function(r) { return r.mimeType === 'image/png'; })
    .slice(0, 6)
    .forEach(function(r) {
      try {
        const f = DriveApp.getFileById(r.id);
        const b64 = Utilities.base64Encode(f.getBlob().getBytes());
        const d = new Date(r.date);
        const dateStr = d.getFullYear() + '/' + (d.getMonth() + 1) + '/' + d.getDate();
        imgHtml += '<div class="photo"><img src="data:image/png;base64,' + b64 + '">' +
          '<div class="cap">' + escapeHtml(r.category) + ' ─ ' + dateStr + '</div></div>';
      } catch (err) { /* 1枚の失敗で全体を止めない */ }
    });
  if (!imgHtml) imgHtml = '<p class="empty">まだ写真がありません</p>';

  var adviceHtml = '';
  (userFolder ? readSheetRows(userFolder, '助言記録', 10) : []).forEach(function(r) {
    adviceHtml += '<tr><td>' + escapeHtml(r[0]) + '</td><td>' + escapeHtml(r[2]) + '</td><td>' + escapeHtml(r[3]) + '</td></tr>';
  });
  if (!adviceHtml) adviceHtml = '<tr><td colspan="3" class="empty">まだ助言がありません</td></tr>';

  var measureHtml = '';
  (userFolder ? readSheetRows(userFolder, '測定データ', 10) : []).forEach(function(r) {
    measureHtml += '<tr><td>' + escapeHtml(r[0]) + '</td><td>' + escapeHtml(r[1]) + '</td><td>' + escapeHtml(r[2]) + '</td><td>' + escapeHtml(r[3]) + '</td></tr>';
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
    '<div class="header"><h1>' + escapeHtml(user) + ' さんの評価表</h1><p>' +
    escapeHtml(String(row[1])) + ' ─ ' +
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

// ===== 動画クリーンアップ（日次トリガー） =====

function cleanupOldVideos() {
  const root = getOrCreateRootFolder();
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - VIDEO_RETENTION_DAYS);
  const deleted = cleanupFolderRecursive(root, cutoff);
  Logger.log('削除した動画: ' + deleted + '件');
  return deleted;
}

function cleanupFolderRecursive(folder, cutoff) {
  var deleted = 0;
  const files = folder.getFilesByType('video/webm');
  while (files.hasNext()) {
    const file = files.next();
    if (file.getDateCreated() < cutoff) { file.setTrashed(true); deleted++; }
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

function getOrCreateFolderIn(parent, name) {
  const it = parent.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return parent.createFolder(name);
}

function getOrCreateSpreadsheet(folder, name) {
  const it = folder.getFilesByName(name);
  if (it.hasNext()) return SpreadsheetApp.open(it.next());
  const ss = SpreadsheetApp.create(name);
  DriveApp.getFileById(ss.getId()).moveTo(folder);
  return ss;
}

function categoryFromTemplate(code) {
  if (!code) return 'その他';
  if (code.indexOf('A') === 0) return '姿勢観察';
  if (code.indexOf('B') === 0) return '動作観察';
  if (code.indexOf('C') === 0) return '簡易検査';
  return 'その他';
}
