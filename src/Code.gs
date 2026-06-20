/**
 * Manabase（Padlet風 振り返りボードアプリ） - サーバー（Google Apps Script）
 *
 * 外部APIは一切使いません。組み込みの SpreadsheetApp / DriveApp のみで動作します。
 * 1デプロイ＝1クラス。シートと写真フォルダは初回アクセス時に自動作成されます。
 */

var PROP_SPREADSHEET_ID = 'SPREADSHEET_ID';
var PROP_FOLDER_ID = 'PHOTO_FOLDER_ID';
var PROP_TEACHER_HASH = 'TEACHER_HASH';
var PROP_TEACHER_SALT = 'TEACHER_SALT';
var PROP_TEACHER_NAME = 'TEACHER_NAME';
var PROP_SCHEMA_VERSION = 'SCHEMA_VERSION';
// SHEET_DEFS を変更したら必ずこの版数を上げる（次回アクセス時に1回だけ移行が走る）
var SCHEMA_VERSION = '8';
// クライアント(Index.html)の APP_BUILD と必ず一致させること。
// デプロイ更新忘れ（古いコードが動いている状態）を検知するために使う。
var APP_BUILD = '9';

var SHEET_STUDENTS = 'Students';
var SHEET_BOARDS = 'Boards';
var SHEET_SECTIONS = 'Sections';
var SHEET_REFLECTIONS = 'Reflections';
var SHEET_COMMENTS = 'Comments';
var SHEET_LIKES = 'Likes';

var SHEET_DEFS = {};
SHEET_DEFS[SHEET_STUDENTS] = ['number', 'name', 'salt', 'passwordHash', 'createdAt'];
// 末尾の archived は後から追加した列（true で児童のボード一覧から非表示。データは保持）
SHEET_DEFS[SHEET_BOARDS] = ['boardId', 'subject', 'unit', 'date', 'title', 'createdAt', 'archived'];
// 末尾の color は後から追加した列（セクションの色分け用）
SHEET_DEFS[SHEET_SECTIONS] = ['sectionId', 'boardId', 'name', 'sortOrder', 'createdAt', 'color'];
// 末尾の sectionId / mediaType / updatedAt / pinned / title / link は後から追加した列（既存データは移行で保持）
// link はリンクプレビュー情報(JSON文字列) {url,title,image,desc}
SHEET_DEFS[SHEET_REFLECTIONS] = ['reflectionId', 'boardId', 'studentName', 'text', 'photoUrl', 'photoFileId', 'color', 'sortOrder', 'createdAt', 'sectionId', 'mediaType', 'updatedAt', 'pinned', 'title', 'link'];
SHEET_DEFS[SHEET_COMMENTS] = ['commentId', 'reflectionId', 'author', 'text', 'createdAt'];
// 末尾の type は後から追加した列（リアクションの種類。空＝❤）
SHEET_DEFS[SHEET_LIKES] = ['reflectionId', 'studentName', 'createdAt', 'type'];

// 使えるリアクション（先頭が既定＝❤）
var REACTIONS = ['❤', '👍', '😲', '🤔', '😢'];

// ============================ エントリ ============================

function doGet() {
  ensureInit_();
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Manabase')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** 任意：エディタから手動実行しても初期化できます（通常は不要）。 */
function setup() {
  ensureInit_();
  Logger.log('初期化完了: ' + getSpreadsheet_().getUrl());
}

// ============================ 初期化 ============================

var _ssCache = null; // 1リクエスト内でのキャッシュ

/** スプレッドシート・フォルダ・各シートを必要に応じて自動作成します。 */
function ensureInit_() {
  if (_ssCache) return _ssCache;
  var props = PropertiesService.getScriptProperties();

  var ss = openStoredSpreadsheet_(props);
  if (!ss) {
    // 初回アクセスでサーバー処理が同時に複数走ると、各々が「IDが無い」と判断して
    // それぞれスプレッドシートを新規作成してしまう（＝重複）。排他ロックで防ぐ。
    var lock = LockService.getScriptLock();
    try { lock.waitLock(20000); } catch (e) {}
    try {
      ss = openStoredSpreadsheet_(props); // ロック取得後に再確認（ダブルチェック）
      if (!ss) {
        ss = SpreadsheetApp.create('Manabase データ');
        props.setProperty(PROP_SPREADSHEET_ID, ss.getId());
      }
      if (!props.getProperty(PROP_FOLDER_ID)) {
        var folder0 = DriveApp.createFolder('Manabase 写真');
        props.setProperty(PROP_FOLDER_ID, folder0.getId());
      }
    } finally {
      try { lock.releaseLock(); } catch (e) {}
    }
  }

  // スキーマ版数が一致していれば、毎回のヘッダー再検証（全シートの読み込み）を省いて高速化する
  if (props.getProperty(PROP_SCHEMA_VERSION) === SCHEMA_VERSION) {
    _ssCache = ss;
    return ss;
  }

  // 各シートを保証（ヘッダーが旧版・不一致なら作り直す）。版数が変わった時のみ実行。
  Object.keys(SHEET_DEFS).forEach(function (name) {
    var def = SHEET_DEFS[name];
    var sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      sheet.getRange(1, 1, 1, def.length).setValues([def]);
      return;
    }
    // 既存シートのヘッダー行を検証。旧スキーマ（例: 旧 Boards の classCode 列・
    // unit 列なし）のまま残っていると appendRow の列ズレで不具合になるため、
    // 1行目（データ未投入時）またはヘッダー不一致を検知したら定義どおりに補正する。
    var lastRow = sheet.getLastRow();
    var width = Math.max(def.length, sheet.getLastColumn());
    var header = sheet.getRange(1, 1, 1, width).getValues()[0];

    var matches = true;
    for (var k = 0; k < def.length; k++) {
      if (String(header[k] || '') !== def[k]) { matches = false; break; }
    }
    if (matches) return;

    // 既存の各列が定義と矛盾しないか（＝列を末尾に追加しただけ等の安全な変更か）を判定。
    // 既存の非空ヘッダーが def の同じ位置と一致していれば、データを壊さず見出しだけ補正する。
    var safeRelabel = true;
    for (var c = 0; c < width; c++) {
      var hv = String(header[c] || '');
      if (hv !== '' && hv !== (def[c] || '')) { safeRelabel = false; break; }
    }

    if (safeRelabel || lastRow <= 1) {
      // データはそのまま、見出し行だけを定義どおりに上書き（新しい列は既存行で空のまま）
      if (lastRow >= 1 && width > def.length) sheet.getRange(1, 1, 1, width).clearContent();
      sheet.getRange(1, 1, 1, def.length).setValues([def]);
    } else {
      // 列の意味が食い違う旧スキーマ（例: 旧 Boards の classCode）は破壊せず退避し、新規作成
      var backup = name + '_旧_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
      sheet.setName(backup);
      var fresh = ss.insertSheet(name);
      fresh.getRange(1, 1, 1, def.length).setValues([def]);
    }
  });
  // 自動生成された空の "シート1"/"Sheet1" を削除
  ['シート1', 'Sheet1'].forEach(function (n) {
    var s = ss.getSheetByName(n);
    if (s && ss.getSheets().length > 1) { try { ss.deleteSheet(s); } catch (e) {} }
  });

  if (!props.getProperty(PROP_FOLDER_ID)) {
    var folder = DriveApp.createFolder('Manabase 写真');
    props.setProperty(PROP_FOLDER_ID, folder.getId());
  }
  // 移行完了。次回以降は上のゲートでヘッダー再検証をスキップする。
  props.setProperty(PROP_SCHEMA_VERSION, SCHEMA_VERSION);
  _ssCache = ss;
  return ss;
}

/** 記録済みIDからスプレッドシートを開く。無ければ null。 */
function openStoredSpreadsheet_(props) {
  var id = props.getProperty(PROP_SPREADSHEET_ID);
  if (!id) return null;
  try { return SpreadsheetApp.openById(id); } catch (e) { return null; }
}

function getSpreadsheet_() {
  return ensureInit_();
}

/**
 * 「今アプリが実際に使っているスプレッドシート」を確認する関数。
 * エディタからこの関数を実行し、ログに出るURLのファイルを残してください。
 * （同名のもう一方は使われていないので、ゴミ箱に入れて構いません）
 */
function getActiveSpreadsheetUrl() {
  var ss = openStoredSpreadsheet_(PropertiesService.getScriptProperties());
  var url = ss ? ss.getUrl() : '(まだ作成されていません)';
  Logger.log('使用中のスプレッドシート: ' + url);
  return url;
}
function getSheet_(name) {
  var sheet = getSpreadsheet_().getSheetByName(name);
  if (!sheet) {
    // 通常 ensureInit_ で作成済みだが、手動削除など想定外時に分かりやすく失敗させる
    throw new Error('シート「' + name + '」が見つかりません。データの初期化に失敗している可能性があります。');
  }
  return sheet;
}
function getPhotoFolder_() {
  ensureInit_();
  return DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty(PROP_FOLDER_ID));
}

// ============================ 認証ヘルパー ============================

function sha256_(str) {
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, str, Utilities.Charset.UTF_8);
  return raw.map(function (b) {
    b = (b < 0) ? b + 256 : b;
    var s = b.toString(16);
    return s.length === 1 ? '0' + s : s;
  }).join('');
}
function newSalt_() { return Utilities.getUuid(); }

// ============================ シート読み取りユーティリティ ============================

/** シートをオブジェクト配列で取得します。 */
function readSheet_(name) {
  var sheet = getSheet_(name);
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  var headers = values[0];
  var out = [];
  for (var i = 1; i < values.length; i++) {
    if (values[i].join('') === '') continue;
    var obj = {};
    for (var j = 0; j < headers.length; j++) obj[headers[j]] = values[i][j];
    obj._row = i + 1;
    out.push(obj);
  }
  return out;
}

function genId_(p) { return p + '_' + Utilities.getUuid().slice(0, 8); }

function deleteRowsWhere_(name, key, value) {
  var sheet = getSheet_(name);
  var values = sheet.getDataRange().getValues();
  var headers = values[0];
  var col = headers.indexOf(key);
  if (col < 0) return;
  for (var i = values.length - 1; i >= 1; i--) {
    if (values[i][col] === value) sheet.deleteRow(i + 1);
  }
}

// ============================ ログイン / 名簿 ============================

/** 起動時情報：名簿（パスワード設定済みか）と先生パスワード設定状況。 */
function getLoginInfo() {
  var students = readSheet_(SHEET_STUDENTS).map(function (s) {
    return { number: s.number, name: s.name, hasPassword: !!s.passwordHash };
  });
  students.sort(function (a, b) { return (a.number || 0) - (b.number || 0); });
  var teacherSet = !!PropertiesService.getScriptProperties().getProperty(PROP_TEACHER_HASH);
  return { students: students, teacherSet: teacherSet, build: APP_BUILD };
}

/** デプロイ状態の診断用。クライアントの APP_BUILD と一致していれば最新。 */
function getServerInfo() {
  // 外部リクエスト権限（UrlFetchApp）が使えるか実際に試す
  var urlFetch = false, urlFetchError = '';
  try {
    UrlFetchApp.fetch('https://www.google.com/generate_204', { muteHttpExceptions: true, followRedirects: false });
    urlFetch = true;
  } catch (e) {
    urlFetchError = String((e && e.message) || e);
  }
  return { build: APP_BUILD, schema: SCHEMA_VERSION, urlFetch: urlFetch, urlFetchError: urlFetchError };
}

/** 生徒ログイン。初回（パスワード未設定）はここで設定します。 */
function studentLogin(name, password) {
  password = String(password || '');
  if (password.length < 1) throw new Error('パスワードを入力してください。');
  var sheet = getSheet_(SHEET_STUDENTS);
  var rows = readSheet_(SHEET_STUDENTS);
  var st = rows.filter(function (r) { return r.name === name; })[0];
  if (!st) throw new Error('名簿に見つかりません。');

  var headers = SHEET_DEFS[SHEET_STUDENTS];
  var saltCol = headers.indexOf('salt') + 1;
  var hashCol = headers.indexOf('passwordHash') + 1;

  if (!st.passwordHash) {
    // 初回ログイン：パスワード設定
    var salt = newSalt_();
    sheet.getRange(st._row, saltCol).setValue(salt);
    sheet.getRange(st._row, hashCol).setValue(sha256_(salt + password));
    return { ok: true, name: name, firstTime: true };
  }
  if (sha256_(st.salt + password) === st.passwordHash) {
    return { ok: true, name: name, firstTime: false };
  }
  throw new Error('パスワードが違います。');
}

/** 先生ログイン。初回はここでパスワードを設定します。 */
function teacherLogin(password) {
  password = String(password || '');
  if (password.length < 1) throw new Error('パスワードを入力してください。');
  var props = PropertiesService.getScriptProperties();
  var hash = props.getProperty(PROP_TEACHER_HASH);
  if (!hash) {
    var salt = newSalt_();
    props.setProperty(PROP_TEACHER_SALT, salt);
    props.setProperty(PROP_TEACHER_HASH, sha256_(salt + password));
    return { ok: true, firstTime: true, name: getTeacherName_() };
  }
  var salt2 = props.getProperty(PROP_TEACHER_SALT);
  if (sha256_(salt2 + password) === hash) return { ok: true, firstTime: false, name: getTeacherName_() };
  throw new Error('パスワードが違います。');
}

/** 先生の表示名（未設定なら「先生」）。 */
function getTeacherName_() {
  return PropertiesService.getScriptProperties().getProperty(PROP_TEACHER_NAME) || '先生';
}

/** 先生の設定取得（名前）。 */
function getTeacherSettings(teacherPassword) {
  if (!isTeacher_(teacherPassword)) throw new Error('先生のみ操作できます。');
  return { name: getTeacherName_() };
}

/** 先生の表示名を変更。 */
function setTeacherName(name, teacherPassword) {
  if (!isTeacher_(teacherPassword)) throw new Error('先生のみ操作できます。');
  name = String(name || '').trim();
  if (!name) throw new Error('名前を入力してください。');
  PropertiesService.getScriptProperties().setProperty(PROP_TEACHER_NAME, name);
  return { name: name };
}

/** 先生のパスワードを変更（現在のパスワードで本人確認）。 */
function setTeacherPassword(newPassword, teacherPassword) {
  if (!isTeacher_(teacherPassword)) throw new Error('現在のパスワードが正しくありません。');
  newPassword = String(newPassword || '');
  if (newPassword.length < 1) throw new Error('新しいパスワードを入力してください。');
  var props = PropertiesService.getScriptProperties();
  var salt = newSalt_();
  props.setProperty(PROP_TEACHER_SALT, salt);
  props.setProperty(PROP_TEACHER_HASH, sha256_(salt + newPassword));
  return { ok: true };
}

/** 生徒の本人確認（投稿・いいね等の操作前に呼ぶ簡易チェック）。 */
function verifyStudent_(name, password) {
  var st = readSheet_(SHEET_STUDENTS).filter(function (r) { return r.name === name; })[0];
  if (!st || !st.passwordHash) return false;
  return sha256_(st.salt + password) === st.passwordHash;
}

// --- 先生による名簿管理 ---

function getStudents() {
  var students = readSheet_(SHEET_STUDENTS).map(function (s) {
    return { number: s.number, name: s.name, hasPassword: !!s.passwordHash };
  });
  students.sort(function (a, b) { return (a.number || 0) - (b.number || 0); });
  return students;
}

function addStudent(number, name) {
  name = String(name || '').trim();
  if (!name) throw new Error('名前を入力してください。');
  getSheet_(SHEET_STUDENTS).appendRow([number || nextNumber_(), name, '', '', new Date()]);
  return getStudents();
}

function nextNumber_() {
  var max = 0;
  readSheet_(SHEET_STUDENTS).forEach(function (s) { if (Number(s.number) > max) max = Number(s.number); });
  return max + 1;
}

/**
 * Excel などからの一括貼り付け。タブ / カンマ / 空白区切り、1行1名。
 * 「1<TAB>山田太郎」「山田太郎<TAB>1」「山田太郎」いずれも可。
 */
function importStudents(text) {
  var lines = String(text || '').split(/\r?\n/);
  var sheet = getSheet_(SHEET_STUDENTS);
  var auto = nextNumber_();
  var rows = [];
  lines.forEach(function (line) {
    line = line.replace(/　/g, ' ').trim();
    if (!line) return;
    var parts = line.split(/[\t,、]+/).map(function (p) { return p.trim(); }).filter(String);
    if (parts.length === 1) parts = parts[0].split(/\s+/);
    var num = '', name = '';
    if (parts.length >= 2) {
      if (/^\d+$/.test(parts[0])) { num = Number(parts[0]); name = parts.slice(1).join(' '); }
      else if (/^\d+$/.test(parts[parts.length - 1])) { num = Number(parts[parts.length - 1]); name = parts.slice(0, -1).join(' '); }
      else { name = parts.join(' '); }
    } else {
      name = parts[0];
    }
    if (!name) return;
    if (num === '') num = auto++;
    rows.push([num, name, '', '', new Date()]);
  });
  if (rows.length) sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 5).setValues(rows);
  return getStudents();
}

function removeStudent(name) {
  deleteRowsWhere_(SHEET_STUDENTS, 'name', name);
  return getStudents();
}

/** 生徒のパスワードをリセット（次回ログイン時に再設定）。 */
function resetStudentPassword(name) {
  var sheet = getSheet_(SHEET_STUDENTS);
  var st = readSheet_(SHEET_STUDENTS).filter(function (r) { return r.name === name; })[0];
  if (!st) return getStudents();
  var headers = SHEET_DEFS[SHEET_STUDENTS];
  sheet.getRange(st._row, headers.indexOf('salt') + 1).setValue('');
  sheet.getRange(st._row, headers.indexOf('passwordHash') + 1).setValue('');
  return getStudents();
}

// ============================ ボード ============================

/**
 * ボード一覧。includeArchived=true で非表示（アーカイブ）ボードも含めます。
 * 児童側（boardList）は false で呼ぶため、非表示ボードは一覧に出ません。
 */
function getBoards(includeArchived) {
  var boards = readSheet_(SHEET_BOARDS).map(rowToBoard_);
  if (!includeArchived) boards = boards.filter(function (b) { return !b.archived; });
  // 未読バッジ用に各ボードの投稿数を付与（Reflections を1回読むだけ）
  var counts = {};
  readSheet_(SHEET_REFLECTIONS).forEach(function (r) { counts[r.boardId] = (counts[r.boardId] || 0) + 1; });
  boards.forEach(function (b) { b.cardCount = counts[b.boardId] || 0; });
  boards.sort(function (a, b) { return new Date(b.createdAt) - new Date(a.createdAt); });
  return boards;
}

function rowToBoard_(r) {
  return {
    boardId: r.boardId,
    subject: r.subject,
    unit: r.unit,
    date: fmtDate_(r.date),
    title: r.title,
    createdAt: toMs_(r.createdAt),
    archived: r.archived === true || r.archived === 'true' || r.archived === 1
  };
}

/** ボードの非表示（アーカイブ）切り替え（先生のみ）。データは消さず一覧から隠すだけ。 */
function setBoardArchived(boardId, archived, teacherPassword) {
  if (!isTeacher_(teacherPassword)) throw new Error('この操作は先生のみ可能です。');
  var sheet = getSheet_(SHEET_BOARDS);
  var b = readSheet_(SHEET_BOARDS).filter(function (r) { return r.boardId === boardId; })[0];
  if (!b) throw new Error('ボードが見つかりません。');
  sheet.getRange(b._row, SHEET_DEFS[SHEET_BOARDS].indexOf('archived') + 1).setValue(!!archived);
  return true;
}
function fmtDate_(d) {
  if (d instanceof Date) return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd');
  if (d == null) return '';
  return String(d);
}
/**
 * 日時をクライアントへ返すための数値(エポックms)に変換します。
 * google.script.run は Date オブジェクトを含む戻り値をうまくシリアライズできず
 * null になる場合があるため、必ず数値に変換してから返します。
 */
function toMs_(d) {
  if (d instanceof Date) return d.getTime();
  if (d == null || d === '') return null;
  var t = new Date(d).getTime();
  return isNaN(t) ? null : t;
}

function getBoard(boardId) {
  var b = readSheet_(SHEET_BOARDS).filter(function (r) { return r.boardId === boardId; })[0];
  return b ? rowToBoard_(b) : null;
}

function createBoard(subject, unit, date, title) {
  subject = String(subject || '').trim();
  unit = String(unit || '').trim();
  if (!subject) throw new Error('教科を選択してください。');
  if (!unit) throw new Error('単元名を入力してください。');
  date = String(date || '').trim();
  if (!date) date = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  title = String(title || '').trim();
  if (!title) title = unit;
  var id = genId_('b');
  // 列順は SHEET_DEFS[SHEET_BOARDS] と一致させること（boardId, subject, unit, date, title, createdAt, archived）
  getSheet_(SHEET_BOARDS).appendRow([id, subject, unit, "'" + date, title, new Date(), false]);
  var created = getBoard(id);
  if (!created) throw new Error('ボードの作成に失敗しました。もう一度お試しください。');
  // 既定セクションを1つ用意しておく（最初から投稿できるように）
  getSheet_(SHEET_SECTIONS).appendRow([genId_('s'), id, 'みんなの投稿', 1, new Date(), '']);
  return created;
}

function deleteBoard(boardId) {
  // 紐づく振り返り・コメント・いいね・写真・セクションを削除
  var refs = readSheet_(SHEET_REFLECTIONS).filter(function (r) { return r.boardId === boardId; });
  refs.forEach(function (r) {
    if (r.photoFileId) { try { DriveApp.getFileById(r.photoFileId).setTrashed(true); } catch (e) {} }
    deleteRowsWhere_(SHEET_COMMENTS, 'reflectionId', r.reflectionId);
    deleteRowsWhere_(SHEET_LIKES, 'reflectionId', r.reflectionId);
  });
  deleteRowsWhere_(SHEET_REFLECTIONS, 'boardId', boardId);
  deleteRowsWhere_(SHEET_SECTIONS, 'boardId', boardId);
  deleteRowsWhere_(SHEET_BOARDS, 'boardId', boardId);
  return true;
}

// ============================ セクション ============================

function getSections(boardId) {
  var secs = readSheet_(SHEET_SECTIONS)
    .filter(function (s) { return s.boardId === boardId; })
    .map(function (s) { return { sectionId: s.sectionId, boardId: s.boardId, name: s.name, sortOrder: Number(s.sortOrder) || 0, color: s.color || '' }; });
  secs.sort(function (a, b) { return a.sortOrder - b.sortOrder; });
  return secs;
}

/** セクション作成（先生のみ）。color は任意。 */
function createSection(boardId, name, teacherPassword, color) {
  if (!isTeacher_(teacherPassword)) throw new Error('セクションの操作は先生のみ可能です。');
  name = String(name || '').trim();
  if (!name) throw new Error('セクション名を入力してください。');
  if (!getBoard(boardId)) throw new Error('ボードが見つかりません。');
  var maxOrder = 0;
  getSections(boardId).forEach(function (s) { if (s.sortOrder > maxOrder) maxOrder = s.sortOrder; });
  getSheet_(SHEET_SECTIONS).appendRow([genId_('s'), boardId, name, maxOrder + 1, new Date(), String(color || '')]);
  return getSections(boardId);
}

/** セクションの色を変更（先生のみ）。 */
function setSectionColor(sectionId, color, teacherPassword) {
  if (!isTeacher_(teacherPassword)) throw new Error('セクションの操作は先生のみ可能です。');
  var sheet = getSheet_(SHEET_SECTIONS);
  var s = readSheet_(SHEET_SECTIONS).filter(function (x) { return x.sectionId === sectionId; })[0];
  if (!s) throw new Error('セクションが見つかりません。');
  sheet.getRange(s._row, SHEET_DEFS[SHEET_SECTIONS].indexOf('color') + 1).setValue(String(color || ''));
  return getSections(s.boardId);
}

/** セクション改名（先生のみ）。 */
function renameSection(sectionId, name, teacherPassword) {
  if (!isTeacher_(teacherPassword)) throw new Error('セクションの操作は先生のみ可能です。');
  name = String(name || '').trim();
  if (!name) throw new Error('セクション名を入力してください。');
  var sheet = getSheet_(SHEET_SECTIONS);
  var s = readSheet_(SHEET_SECTIONS).filter(function (x) { return x.sectionId === sectionId; })[0];
  if (!s) throw new Error('セクションが見つかりません。');
  sheet.getRange(s._row, SHEET_DEFS[SHEET_SECTIONS].indexOf('name') + 1).setValue(name);
  return getSections(s.boardId);
}

/** セクション削除（先生のみ）。中の投稿は未分類（既定列）に移します。 */
function deleteSection(sectionId, teacherPassword) {
  if (!isTeacher_(teacherPassword)) throw new Error('セクションの操作は先生のみ可能です。');
  var s = readSheet_(SHEET_SECTIONS).filter(function (x) { return x.sectionId === sectionId; })[0];
  if (!s) return true;
  var boardId = s.boardId;
  // このセクションの投稿は sectionId を空にして未分類へ（投稿自体は消さない）
  var sheet = getSheet_(SHEET_REFLECTIONS);
  var secCol = SHEET_DEFS[SHEET_REFLECTIONS].indexOf('sectionId') + 1;
  readSheet_(SHEET_REFLECTIONS).forEach(function (r) {
    if (r.sectionId === sectionId) sheet.getRange(r._row, secCol).setValue('');
  });
  deleteRowsWhere_(SHEET_SECTIONS, 'sectionId', sectionId);
  return getSections(boardId);
}

/** ボードに最低1つセクションがあることを保証し、既定セクションIDを返す。 */
function ensureDefaultSection_(boardId) {
  var secs = getSections(boardId);
  if (secs.length) return secs[0].sectionId;
  var id = genId_('s');
  getSheet_(SHEET_SECTIONS).appendRow([id, boardId, 'みんなの投稿', 1, new Date(), '']);
  return id;
}

// ============================ 振り返り（カード） ============================

/** ボードのカード一覧（いいね・コメント込み）。currentName で自分のいいね判定。 */
function getBoardData(boardId, currentName) {
  var board = getBoard(boardId);
  if (!board) throw new Error('ボードが見つかりません。');

  var refs = readSheet_(SHEET_REFLECTIONS).filter(function (r) { return r.boardId === boardId; });

  // リアクション集計（種類別）。type 空は ❤ とみなす。
  var likes = readSheet_(SHEET_LIKES);
  var likeCount = {}, reactByRef = {}, myReactByRef = {};
  likes.forEach(function (l) {
    var t = l.type || '❤';
    likeCount[l.reflectionId] = (likeCount[l.reflectionId] || 0) + 1;
    var m = reactByRef[l.reflectionId] = reactByRef[l.reflectionId] || {};
    m[t] = (m[t] || 0) + 1;
    if (l.studentName === currentName) {
      var mm = myReactByRef[l.reflectionId] = myReactByRef[l.reflectionId] || {};
      mm[t] = true;
    }
  });

  // コメント集計
  var comments = readSheet_(SHEET_COMMENTS);
  var byRef = {};
  comments.forEach(function (c) {
    (byRef[c.reflectionId] = byRef[c.reflectionId] || []).push({
      commentId: c.commentId, author: c.author, text: c.text, createdAt: toMs_(c.createdAt)
    });
  });
  Object.keys(byRef).forEach(function (k) {
    byRef[k].sort(function (a, b) { return new Date(a.createdAt) - new Date(b.createdAt); });
  });

  var cards = refs.map(function (r) {
    return {
      reflectionId: r.reflectionId,
      sectionId: r.sectionId || '',
      studentName: r.studentName,
      text: r.text,
      photoUrl: r.photoUrl,
      mediaType: r.mediaType || (r.photoUrl ? 'image' : ''),
      color: r.color,
      title: r.title || '',
      sortOrder: Number(r.sortOrder) || 0,
      createdAt: toMs_(r.createdAt),
      updatedAt: toMs_(r.updatedAt),
      pinned: r.pinned === true || r.pinned === 'true' || r.pinned === 1,
      link: parseLink_(r.link),
      likeCount: likeCount[r.reflectionId] || 0,
      reactions: reactByRef[r.reflectionId] || {},
      myReactions: myReactByRef[r.reflectionId] || {},
      comments: byRef[r.reflectionId] || []
    };
  });
  cards.sort(cardCompare_);

  return { board: board, sections: getSections(boardId), cards: cards };
}

/** ピン留め優先 → 並び順 → 作成日時 の順で比較。 */
function cardCompare_(a, b) {
  if (!!a.pinned !== !!b.pinned) return a.pinned ? -1 : 1;
  if (a.sortOrder !== b.sortOrder) return a.sortOrder - b.sortOrder;
  return (a.createdAt || 0) - (b.createdAt || 0);
}

/**
 * リアルタイム更新用の軽量シグネチャ。
 * これが前回と変われば、クライアントは getBoardData を取り直して再描画する。
 */
function getBoardSignature(boardId) {
  var refCount = 0, maxMs = 0;
  var refIds = {};
  readSheet_(SHEET_REFLECTIONS).forEach(function (r) {
    if (r.boardId !== boardId) return;
    refCount++;
    refIds[r.reflectionId] = true;
    var u = toMs_(r.updatedAt) || toMs_(r.createdAt) || 0;
    if (u > maxMs) maxMs = u;
  });
  var secCount = readSheet_(SHEET_SECTIONS).filter(function (s) { return s.boardId === boardId; }).length;
  // いいね・コメントの増減も拾うため総数を含める（このボード分に限定）
  var likeCount = readSheet_(SHEET_LIKES).filter(function (l) { return refIds[l.reflectionId]; }).length;
  var comCount = readSheet_(SHEET_COMMENTS).filter(function (c) { return refIds[c.reflectionId]; }).length;
  return refCount + '|' + maxMs + '|' + secCount + '|' + likeCount + '|' + comCount;
}

/**
 * 投稿。media は { data(base64), mimeType, filename, kind:'image'|'video' } または null。
 */
function postReflection(boardId, sectionId, studentName, password, title, text, color, media, link, teacherPassword) {
  // 先生は自分の表示名で投稿できる。それ以外は生徒の本人確認。
  var author;
  if (teacherPassword && isTeacher_(teacherPassword)) {
    author = getTeacherName_();
  } else {
    if (!verifyStudent_(studentName, password)) throw new Error('ログイン情報が正しくありません。');
    author = studentName;
  }
  title = String(title || '').trim();
  text = String(text || '').trim();
  var linkObj = sanitizeLink_(link);
  if (!title && !text && !media && !linkObj) throw new Error('タイトル・本文・写真・動画・リンクのいずれかを入力してください。');
  var board = getBoard(boardId);
  if (!board) throw new Error('ボードが見つかりません。');

  if (!sectionId) sectionId = ensureDefaultSection_(boardId);

  var url = '', fileId = '', mediaType = '';
  if (media && media.data) {
    var saved = saveMedia_(media, board);
    url = saved.url; fileId = saved.fileId; mediaType = saved.mediaType;
  }

  // 末尾に並べる
  var maxOrder = 0;
  readSheet_(SHEET_REFLECTIONS).forEach(function (r) {
    if (r.boardId === boardId && Number(r.sortOrder) > maxOrder) maxOrder = Number(r.sortOrder);
  });

  var id = genId_('r');
  var now = new Date();
  // 列順は SHEET_DEFS[SHEET_REFLECTIONS] と一致させること
  getSheet_(SHEET_REFLECTIONS).appendRow([
    id, boardId, author, text, url, fileId, color || '#fff7c0', maxOrder + 1, now, sectionId, mediaType, now, false, title,
    linkObj ? JSON.stringify(linkObj) : ''
  ]);
  // 速度重視：作成したカード1枚だけを返す（クライアントは部分描画する）
  return {
    card: {
      reflectionId: id, sectionId: sectionId, studentName: author, title: title, text: text,
      photoUrl: url, mediaType: mediaType, color: color || '#fff7c0', sortOrder: maxOrder + 1,
      createdAt: now.getTime(), updatedAt: now.getTime(), pinned: false,
      link: linkObj, reactions: {}, myReactions: {},
      likeCount: 0, likedByMe: false, comments: []
    }
  };
}

/** クライアントから来たリンク情報を安全なオブジェクトに整形（無効なら null）。 */
function sanitizeLink_(link) {
  if (!link) return null;
  var url = normalizeUrl_(typeof link === 'string' ? link : link.url);
  if (!url) return null;
  var s = (typeof link === 'object') ? link : {};
  return {
    url: url,
    title: String(s.title || '').slice(0, 300),
    image: String(s.image || '').slice(0, 1000),
    desc: String(s.desc || '').slice(0, 500),
    site: String(s.site || hostOf_(url)).slice(0, 200)
  };
}

/**
 * 投稿の編集。本人 または 先生のみ。
 * media を渡せば差し替え、removeMedia=true なら添付を削除、どちらも無ければ本文/色のみ更新。
 */
function editReflection(reflectionId, studentName, password, teacherPassword, title, text, color, media, removeMedia, link, removeLink) {
  var sheet = getSheet_(SHEET_REFLECTIONS);
  var r = readSheet_(SHEET_REFLECTIONS).filter(function (x) { return x.reflectionId === reflectionId; })[0];
  if (!r) throw new Error('投稿が見つかりません。');
  var allowed = isTeacher_(teacherPassword) || (r.studentName === studentName && verifyStudent_(studentName, password));
  if (!allowed) throw new Error('編集する権限がありません。');

  var def = SHEET_DEFS[SHEET_REFLECTIONS];
  title = String(title || '').trim();
  text = String(text || '').trim();
  var linkObj = sanitizeLink_(link);

  // 変更後に添付が残るか（先に検証し、空投稿になるなら何も書き換えない）
  var willHaveMedia = (media && media.data) ? true : (removeMedia ? false : !!r.photoFileId);
  var willHaveLink = linkObj ? true : (removeLink ? false : !!r.link);
  if (!title && !text && !willHaveMedia && !willHaveLink) throw new Error('タイトル・本文・写真・動画・リンクのいずれかは必要です。');

  sheet.getRange(r._row, def.indexOf('title') + 1).setValue(title);
  sheet.getRange(r._row, def.indexOf('text') + 1).setValue(text);
  if (color) sheet.getRange(r._row, def.indexOf('color') + 1).setValue(color);

  var linkCol = def.indexOf('link') + 1;
  if (linkObj) sheet.getRange(r._row, linkCol).setValue(JSON.stringify(linkObj));
  else if (removeLink) sheet.getRange(r._row, linkCol).setValue('');

  if (media && media.data) {
    if (r.photoFileId) { try { DriveApp.getFileById(r.photoFileId).setTrashed(true); } catch (e) {} }
    var board = getBoard(r.boardId) || { subject: 'その他', unit: '' };
    var saved = saveMedia_(media, board);
    sheet.getRange(r._row, def.indexOf('photoUrl') + 1).setValue(saved.url);
    sheet.getRange(r._row, def.indexOf('photoFileId') + 1).setValue(saved.fileId);
    sheet.getRange(r._row, def.indexOf('mediaType') + 1).setValue(saved.mediaType);
  } else if (removeMedia) {
    if (r.photoFileId) { try { DriveApp.getFileById(r.photoFileId).setTrashed(true); } catch (e) {} }
    sheet.getRange(r._row, def.indexOf('photoUrl') + 1).setValue('');
    sheet.getRange(r._row, def.indexOf('photoFileId') + 1).setValue('');
    sheet.getRange(r._row, def.indexOf('mediaType') + 1).setValue('');
  }

  var now = new Date();
  sheet.getRange(r._row, def.indexOf('updatedAt') + 1).setValue(now);
  // 速度重視：変更後の値だけ返す（クライアントは該当カードを差し替える）
  var newUrl = (media && media.data) ? readCell_(sheet, r._row, def, 'photoUrl') : (removeMedia ? '' : r.photoUrl);
  var newType = (media && media.data) ? readCell_(sheet, r._row, def, 'mediaType') : (removeMedia ? '' : (r.mediaType || (r.photoUrl ? 'image' : '')));
  var newLink = linkObj ? linkObj : (removeLink ? null : parseLink_(r.link));
  return {
    update: {
      reflectionId: reflectionId, title: title, text: text, color: color || r.color,
      photoUrl: newUrl, mediaType: newType, link: newLink, updatedAt: now.getTime()
    }
  };
}

/** セルの link(JSON文字列) を安全にオブジェクト化（無ければ null）。 */
function parseLink_(v) {
  if (!v) return null;
  try { var o = JSON.parse(v); return (o && o.url) ? o : null; } catch (e) { return null; }
}

function readCell_(sheet, row, def, key) {
  return sheet.getRange(row, def.indexOf(key) + 1).getValue();
}

/** ピン留めの切り替え。本人 または 先生。 */
function togglePin(reflectionId, studentName, password, teacherPassword) {
  var sheet = getSheet_(SHEET_REFLECTIONS);
  var r = readSheet_(SHEET_REFLECTIONS).filter(function (x) { return x.reflectionId === reflectionId; })[0];
  if (!r) throw new Error('投稿が見つかりません。');
  var allowed = isTeacher_(teacherPassword) || (r.studentName === studentName && verifyStudent_(studentName, password));
  if (!allowed) throw new Error('ピン留めする権限がありません。');
  var def = SHEET_DEFS[SHEET_REFLECTIONS];
  var pinned = !(r.pinned === true || r.pinned === 'true' || r.pinned === 1);
  sheet.getRange(r._row, def.indexOf('pinned') + 1).setValue(pinned);
  sheet.getRange(r._row, def.indexOf('updatedAt') + 1).setValue(new Date());
  return { pinned: pinned };
}

/** 削除：本人 または 先生。 */
function deleteReflection(reflectionId, studentName, password, teacherPassword) {
  var r = readSheet_(SHEET_REFLECTIONS).filter(function (x) { return x.reflectionId === reflectionId; })[0];
  if (!r) return true;
  var allowed = isTeacher_(teacherPassword) || (r.studentName === studentName && verifyStudent_(studentName, password));
  if (!allowed) throw new Error('削除する権限がありません。');
  if (r.photoFileId) { try { DriveApp.getFileById(r.photoFileId).setTrashed(true); } catch (e) {} }
  deleteRowsWhere_(SHEET_COMMENTS, 'reflectionId', reflectionId);
  deleteRowsWhere_(SHEET_LIKES, 'reflectionId', reflectionId);
  deleteRowsWhere_(SHEET_REFLECTIONS, 'reflectionId', reflectionId);
  return true;
}

/** ドラッグ＆ドロップ後：並び順とセクション移動を保存。items=[{id, sectionId}] 。 */
function updateLayout(boardId, items) {
  var sheet = getSheet_(SHEET_REFLECTIONS);
  var def = SHEET_DEFS[SHEET_REFLECTIONS];
  var orderCol = def.indexOf('sortOrder') + 1;
  var secCol = def.indexOf('sectionId') + 1;
  var rowById = {};
  readSheet_(SHEET_REFLECTIONS).forEach(function (r) { rowById[r.reflectionId] = r._row; });
  (items || []).forEach(function (it, idx) {
    var row = rowById[it.id];
    if (!row) return;
    sheet.getRange(row, orderCol).setValue(idx + 1);
    sheet.getRange(row, secCol).setValue(it.sectionId || '');
  });
  return true;
}

function isTeacher_(password) {
  if (!password) return false;
  var props = PropertiesService.getScriptProperties();
  var hash = props.getProperty(PROP_TEACHER_HASH);
  var salt = props.getProperty(PROP_TEACHER_SALT);
  return !!hash && sha256_(salt + password) === hash;
}

/** ドラッグ＆ドロップ後の並び順を保存。 */
function updateOrder(boardId, orderedIds) {
  var sheet = getSheet_(SHEET_REFLECTIONS);
  var rows = readSheet_(SHEET_REFLECTIONS);
  var orderCol = SHEET_DEFS[SHEET_REFLECTIONS].indexOf('sortOrder') + 1;
  var rowById = {};
  rows.forEach(function (r) { rowById[r.reflectionId] = r._row; });
  orderedIds.forEach(function (id, idx) {
    if (rowById[id]) sheet.getRange(rowById[id], orderCol).setValue(idx + 1);
  });
  return true;
}

// --- リアクション（複数種類） ---
function toggleReaction(reflectionId, studentName, password, type) {
  if (!verifyStudent_(studentName, password)) throw new Error('ログイン情報が正しくありません。');
  type = String(type || '❤');
  if (REACTIONS.indexOf(type) < 0) type = '❤';
  var sheet = getSheet_(SHEET_LIKES);
  var existing = readSheet_(SHEET_LIKES).filter(function (l) {
    return l.reflectionId === reflectionId && l.studentName === studentName && (l.type || '❤') === type;
  })[0];
  if (existing) {
    sheet.deleteRow(existing._row);
  } else {
    sheet.appendRow([reflectionId, studentName, new Date(), type]);
  }
  // 更新後の種類別集計と自分の反応を返す
  var reactions = {}, mine = {}, total = 0;
  readSheet_(SHEET_LIKES).forEach(function (l) {
    if (l.reflectionId !== reflectionId) return;
    var t = l.type || '❤';
    reactions[t] = (reactions[t] || 0) + 1;
    total++;
    if (l.studentName === studentName) mine[t] = true;
  });
  return { reactions: reactions, myReactions: mine, count: total };
}

// --- コメント ---
function addComment(reflectionId, author, password, text) {
  if (!verifyStudent_(author, password)) throw new Error('ログイン情報が正しくありません。');
  text = String(text || '').trim();
  if (!text) throw new Error('コメントを入力してください。');
  var id = genId_('c');
  getSheet_(SHEET_COMMENTS).appendRow([id, reflectionId, author, text, new Date()]);
  return readSheet_(SHEET_COMMENTS)
    .filter(function (c) { return c.reflectionId === reflectionId; })
    .sort(function (a, b) { return new Date(a.createdAt) - new Date(b.createdAt); })
    .map(function (c) { return { commentId: c.commentId, author: c.author, text: c.text, createdAt: toMs_(c.createdAt) }; });
}

// ============================ リンクプレビュー（OGP） ============================

/**
 * URL を読み込み、OGP/メタ情報からプレビュー（タイトル・画像・説明）を作る。
 * クライアントの「プレビュー取得」および投稿保存時に使う。失敗しても URL だけは返す。
 */
function fetchLinkPreview(url) {
  url = normalizeUrl_(url);
  if (!url) throw new Error('URL を入力してください。');
  var info = { url: url, title: '', image: '', desc: '', site: hostOf_(url) };
  try {
    var res = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      followRedirects: true,
      validateHttpsCertificates: true,
      headers: { 'User-Agent': 'Mozilla/5.0 (compatible; ManabaseBot/1.0)' }
    });
    if (res.getResponseCode() >= 400) return info;
    var ct = String(res.getHeaders()['Content-Type'] || res.getHeaders()['content-type'] || '');
    if (ct && ct.indexOf('text/html') < 0 && ct.indexOf('application/xhtml') < 0) return info;
    var html = res.getContentText();
    // 先頭の <head> 付近だけ見れば十分（重い解析を避ける）
    html = html.slice(0, 200000);
    info.title = metaContent_(html, 'og:title') || titleTag_(html) || '';
    info.image = absUrl_(url, metaContent_(html, 'og:image') || metaContent_(html, 'twitter:image') || '');
    info.desc = metaContent_(html, 'og:description') || metaName_(html, 'description') || '';
    if (!info.title) info.title = info.site;
  } catch (e) {
    // ネットワークエラー等は URL だけのプレビューにフォールバック
  }
  return info;
}

function normalizeUrl_(url) {
  url = String(url || '').trim();
  if (!url) return '';
  if (!/^https?:\/\//i.test(url)) url = 'https://' + url;
  if (!/^https?:\/\/[^\s]+\.[^\s]+/i.test(url)) return '';
  return url;
}
function hostOf_(url) {
  var m = String(url).match(/^https?:\/\/([^\/?#]+)/i);
  return m ? m[1].replace(/^www\./, '') : '';
}
function metaContent_(html, prop) {
  // <meta property="og:xxx" content="..."> （property/name どちらでも、属性順も両対応）
  var re = new RegExp('<meta[^>]+(?:property|name)\\s*=\\s*["\']' + escRe_(prop) + '["\'][^>]*>', 'i');
  var tag = html.match(re);
  if (!tag) return '';
  var c = tag[0].match(/content\s*=\s*["\']([\s\S]*?)["\']/i);
  return c ? decodeEntities_(c[1].trim()) : '';
}
function metaName_(html, name) { return metaContent_(html, name); }
function titleTag_(html) {
  var m = html.match(/<title[^>]*>([\s\S]*?)<\/title>/i);
  return m ? decodeEntities_(m[1].trim()) : '';
}
function absUrl_(base, u) {
  u = String(u || '').trim();
  if (!u) return '';
  if (/^https?:\/\//i.test(u)) return u;
  if (/^\/\//.test(u)) return 'https:' + u;
  var m = base.match(/^(https?:\/\/[^\/]+)(\/[^?#]*)?/i);
  if (!m) return u;
  if (u.charAt(0) === '/') return m[1] + u;
  var dir = (m[2] || '/').replace(/[^\/]*$/, '');
  return m[1] + dir + u;
}
function escRe_(s) { return String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }
function decodeEntities_(s) {
  return String(s)
    .replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&#x27;/gi, "'")
    .replace(/&nbsp;/g, ' ');
}

// ============================ メディア（写真・動画） ============================

/** 写真または動画を Drive に保存し、表示用URLと種別を返す。 */
function saveMedia_(media, board) {
  var folder = getPhotoFolder_();
  var subName = (board.subject || 'その他') + '_' + (board.unit || '');
  var sub = getOrCreateSubfolder_(folder, subName);
  // MIME は "video/webm;codecs=vp9" のように余分が付くことがあるので主要部だけ使う
  var mime = String(media.mimeType || 'image/jpeg').split(';')[0].trim() || 'image/jpeg';
  var isVideo = (media.kind === 'video') || /^video\//.test(mime);
  var isPdf = (media.kind === 'pdf') || mime === 'application/pdf';
  var bytes = decodeBase64_(media.data);
  var defName = (isPdf ? 'file_' : isVideo ? 'video_' : 'photo_') + Date.now() + (isPdf ? '.pdf' : isVideo ? '.mp4' : '.jpg');
  var blob = Utilities.newBlob(bytes, mime, media.filename || defName);
  var file = sub.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  var id = file.getId();
  if (isVideo || isPdf) {
    // 動画・PDFは preview（iframe）で表示する
    return { fileId: id, url: 'https://drive.google.com/file/d/' + id + '/preview', mediaType: isPdf ? 'pdf' : 'video' };
  }
  return { fileId: id, url: 'https://drive.google.com/thumbnail?id=' + id + '&sz=w1000', mediaType: 'image' };
}

/**
 * base64 を頑丈にデコードする。
 * - "data:...;base64," のデータURL接頭辞が混ざっていても除去
 * - 改行・空白を除去
 * - URLセーフ（- _）形式にも対応
 */
function decodeBase64_(data) {
  data = String(data || '');
  var ci = data.indexOf('base64,');
  if (ci >= 0) data = data.slice(ci + 7);
  data = data.replace(/\s+/g, '');
  if (!data) throw new Error('メディアのデータが空です。もう一度撮影／選択してください。');
  if (data.indexOf('-') >= 0 || data.indexOf('_') >= 0) {
    return Utilities.base64DecodeWebSafe(data);
  }
  return Utilities.base64Decode(data);
}
function getOrCreateSubfolder_(parent, name) {
  var it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

// ============================ 出力（先生画面） ============================

/** ボード全体の出力データ（印刷用）。 */
function exportBoard(boardId) {
  return getBoardData(boardId, null);
}

/** 児童別の出力データ：その児童の全ボードにわたる振り返りを時系列で。 */
function exportStudent(studentName) {
  var boards = {};
  readSheet_(SHEET_BOARDS).forEach(function (b) { boards[b.boardId] = rowToBoard_(b); });
  var refs = readSheet_(SHEET_REFLECTIONS)
    .filter(function (r) { return r.studentName === studentName; })
    .sort(function (a, b) { return new Date(a.createdAt) - new Date(b.createdAt); })
    .map(function (r) {
      var bd = boards[r.boardId] || {};
      return {
        boardTitle: bd.title || '(削除済みボード)',
        subject: bd.subject || '', unit: bd.unit || '', date: bd.date || '',
        title: r.title || '', text: r.text, photoUrl: r.photoUrl,
        mediaType: r.mediaType || (r.photoUrl ? 'image' : ''),
        link: parseLink_(r.link),
        color: r.color, createdAt: toMs_(r.createdAt)
      };
    });
  return { studentName: studentName, reflections: refs };
}
