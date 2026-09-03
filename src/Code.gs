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
var SCHEMA_VERSION = '9';
// クライアント(Index.html)の APP_BUILD と必ず一致させること。
// デプロイ更新忘れ（古いコードが動いている状態）を検知するために使う。
var APP_BUILD = '26';

var SHEET_STUDENTS = 'Students';
var SHEET_BOARDS = 'Boards';
var SHEET_SECTIONS = 'Sections';
var SHEET_REFLECTIONS = 'Reflections';
var SHEET_COMMENTS = 'Comments';
var SHEET_LIKES = 'Likes';
var SHEET_CLASSES = 'Classes';

var SHEET_DEFS = {};
SHEET_DEFS[SHEET_CLASSES] = ['classId', 'name', 'sortOrder', 'createdAt'];
// 末尾の classId は後から追加した列（どのクラスの生徒か）
SHEET_DEFS[SHEET_STUDENTS] = ['number', 'name', 'salt', 'passwordHash', 'createdAt', 'classId'];
// 末尾の archived / classId は後から追加した列
SHEET_DEFS[SHEET_BOARDS] = ['boardId', 'subject', 'unit', 'date', 'title', 'createdAt', 'archived', 'classId'];
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
var _propsCache = null; // ScriptProperties は1リクエストにつき1回だけまとめて読む（毎回の往復を減らす）

function props_() {
  if (!_propsCache) _propsCache = PropertiesService.getScriptProperties().getProperties() || {};
  return _propsCache;
}
function getProp_(key) { var v = props_()[key]; return (v == null) ? null : v; }
function setProp_(key, value) {
  PropertiesService.getScriptProperties().setProperty(key, value);
  if (_propsCache) _propsCache[key] = value;
}

/** スプレッドシート・フォルダ・各シートを必要に応じて自動作成します。 */
function ensureInit_() {
  if (_ssCache) return _ssCache;

  var ss = openStoredSpreadsheet_();
  if (!ss) {
    // 初回アクセスでサーバー処理が同時に複数走ると、各々が「IDが無い」と判断して
    // それぞれスプレッドシートを新規作成してしまう（＝重複）。排他ロックで防ぐ。
    var lock = LockService.getScriptLock();
    try { lock.waitLock(20000); } catch (e) {}
    try {
      _propsCache = null; // ロック待ちの間に他の実行が設定した可能性があるので読み直す
      ss = openStoredSpreadsheet_(); // ロック取得後に再確認（ダブルチェック）
      if (!ss) {
        ss = SpreadsheetApp.create('Manabase データ');
        setProp_(PROP_SPREADSHEET_ID, ss.getId());
      }
      if (!getProp_(PROP_FOLDER_ID)) {
        var folder0 = DriveApp.createFolder('Manabase 写真');
        setProp_(PROP_FOLDER_ID, folder0.getId());
      }
    } finally {
      try { lock.releaseLock(); } catch (e) {}
    }
  }

  // スキーマ版数が一致していれば、毎回のヘッダー再検証（全シートの読み込み）を省いて高速化する
  if (getProp_(PROP_SCHEMA_VERSION) === SCHEMA_VERSION) {
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

  if (!getProp_(PROP_FOLDER_ID)) {
    var folder = DriveApp.createFolder('Manabase 写真');
    setProp_(PROP_FOLDER_ID, folder.getId());
  }
  // 旧データ（クラスなし）を既定クラスに割り当てる移行（失敗してもバージョンは進める）
  _ssCache = ss; // 以降のヘルパーが getSpreadsheet_ を使うため先にキャッシュ
  try { ensureLegacyClass_(); } catch (e) { /* 何度でも安全に再実行できるため握りつぶす */ }
  // 移行完了。次回以降は上のゲートでヘッダー再検証をスキップする。
  setProp_(PROP_SCHEMA_VERSION, SCHEMA_VERSION);
  _ssCache = ss;
  return ss;
}

/**
 * クラス導入前のデータ（classId 空）を、既定クラスにまとめて割り当てる。
 * 既存の生徒・ボードがあるのにクラスが無い場合のみ「1組」を作って割り当てる。
 */
function ensureLegacyClass_() {
  var classSheet = getSheet_(SHEET_CLASSES);
  var classes = readSheet_(SHEET_CLASSES);
  var students = readSheet_(SHEET_STUDENTS);
  var boards = readSheet_(SHEET_BOARDS);

  var orphanStudents = students.filter(function (s) { return !s.classId; });
  var orphanBoards = boards.filter(function (b) { return !b.classId; });
  if (!orphanStudents.length && !orphanBoards.length) return;

  // 割り当て先クラス：既存があれば先頭、無ければ「1組」を作成
  var targetId;
  if (classes.length) {
    targetId = classes[0].classId;
  } else {
    targetId = genId_('cls');
    classSheet.appendRow([targetId, '1組', 1, new Date()]);
  }

  var sSheet = getSheet_(SHEET_STUDENTS);
  var sCol = SHEET_DEFS[SHEET_STUDENTS].indexOf('classId') + 1;
  orphanStudents.forEach(function (s) { sSheet.getRange(s._row, sCol).setValue(targetId); });

  var bSheet = getSheet_(SHEET_BOARDS);
  var bCol = SHEET_DEFS[SHEET_BOARDS].indexOf('classId') + 1;
  orphanBoards.forEach(function (b) { bSheet.getRange(b._row, bCol).setValue(targetId); });
}

/** 記録済みIDからスプレッドシートを開く。無ければ null。 */
function openStoredSpreadsheet_() {
  var id = getProp_(PROP_SPREADSHEET_ID);
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
  var ss = openStoredSpreadsheet_();
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
  return DriveApp.getFolderById(getProp_(PROP_FOLDER_ID));
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
  return deleteRowsWhereIn_(name, key, [value]);
}

/**
 * key 列の値が values のいずれかに一致する行をまとめて削除する。
 * シートは1回だけ読み、下から連続区間ごとに deleteRows するので、
 * 1行ずつ deleteRow するより大幅に速い（投稿数が多いボードの削除など）。
 */
function deleteRowsWhereIn_(name, key, values) {
  var want = {};
  (values || []).forEach(function (v) { if (v != null && v !== '') want[String(v)] = true; });
  if (!Object.keys(want).length) return 0;
  var sheet = getSheet_(name);
  return withLock_(function () {
    var data = sheet.getDataRange().getValues();
    var col = data.length ? data[0].indexOf(key) : -1;
    if (col < 0) return 0;
    var rows = [];
    for (var i = 1; i < data.length; i++) {
      if (want[String(data[i][col])]) rows.push(i + 1);
    }
    rows.sort(function (a, b) { return b - a; });
    var k = 0;
    while (k < rows.length) {
      var end = rows[k], start = end;
      while (k + 1 < rows.length && rows[k + 1] === start - 1) { k++; start = rows[k]; }
      sheet.deleteRows(start, end - start + 1);
      k++;
    }
    return rows.length;
  });
}

/**
 * 行番号に依存する書き込み（行削除・列まとめ書き）を直列化する。
 * ロックが取れなくても処理自体は行う（止まるより動くことを優先）。
 */
function withLock_(fn) {
  var lock = LockService.getScriptLock();
  var locked = false;
  try { locked = lock.tryLock(10000); } catch (e) { locked = false; }
  try { return fn(); }
  finally { if (locked) { try { lock.releaseLock(); } catch (e) {} } }
}

// ============================ 小さなキャッシュユーティリティ ============================

/** JSON を gzip + base64 で CacheService に保存（100KB 上限を超えるものは諦める）。 */
function cachePutJson_(key, obj, ttlSec) {
  try {
    var json = JSON.stringify(obj);
    var gz = Utilities.gzip(Utilities.newBlob(json, 'application/octet-stream'));
    var s = Utilities.base64Encode(gz.getBytes());
    if (s.length > 100000) return false;
    CacheService.getScriptCache().put(key, s, ttlSec);
    return true;
  } catch (e) { return false; }
}
function cacheGetJson_(key) {
  try {
    var s = CacheService.getScriptCache().get(key);
    if (!s) return null;
    var blob = Utilities.newBlob(Utilities.base64Decode(s), 'application/x-gzip');
    return JSON.parse(Utilities.ungzip(blob).getDataAsString());
  } catch (e) { return null; }
}
function cacheRemove_(key) {
  try { CacheService.getScriptCache().remove(key); } catch (e) {}
}

// ============================ ログイン / 名簿 ============================

/** 起動時情報：クラス一覧＋クラス別名簿と先生パスワード設定状況。 */
var CACHE_LOGIN_INFO = 'login_info_v1';
function getLoginInfo() {
  // 授業の始めに全員が一斉に開くので、名簿・クラス一覧は短時間キャッシュする
  // （名簿・クラス・パスワード設定が変わる操作では clearLoginCache_ で即無効化）
  var cached = cacheGetJson_(CACHE_LOGIN_INFO);
  if (cached && cached.build === APP_BUILD) return cached;
  var byClass = {};
  readSheet_(SHEET_STUDENTS).forEach(function (s) {
    var k = String(s.classId || '');
    (byClass[k] = byClass[k] || []).push({ number: s.number, name: asText_(s.name), hasPassword: !!s.passwordHash });
  });
  Object.keys(byClass).forEach(function (k) {
    byClass[k].sort(function (a, b) { return (a.number || 0) - (b.number || 0); });
  });
  var teacherSet = !!getProp_(PROP_TEACHER_HASH);
  var info = { classes: getClasses(), studentsByClass: byClass, teacherSet: teacherSet, build: APP_BUILD };
  cachePutJson_(CACHE_LOGIN_INFO, info, 120);
  return info;
}
function clearLoginCache_() { cacheRemove_(CACHE_LOGIN_INFO); }

// ============================ クラス ============================

function getClasses() {
  var classes = readSheet_(SHEET_CLASSES).map(function (c) {
    return { classId: c.classId, name: c.name, sortOrder: Number(c.sortOrder) || 0 };
  });
  classes.sort(function (a, b) { return a.sortOrder - b.sortOrder; });
  return classes;
}

function createClass(name, teacherPassword) {
  if (!isTeacher_(teacherPassword)) throw new Error('クラスの操作は先生のみ可能です。');
  name = String(name || '').trim();
  if (!name) throw new Error('クラス名を入力してください。');
  var maxOrder = 0;
  getClasses().forEach(function (c) { if (c.sortOrder > maxOrder) maxOrder = c.sortOrder; });
  getSheet_(SHEET_CLASSES).appendRow([genId_('cls'), name, maxOrder + 1, new Date()]);
  clearLoginCache_();
  return getClasses();
}

function renameClass(classId, name, teacherPassword) {
  if (!isTeacher_(teacherPassword)) throw new Error('クラスの操作は先生のみ可能です。');
  name = String(name || '').trim();
  if (!name) throw new Error('クラス名を入力してください。');
  var sheet = getSheet_(SHEET_CLASSES);
  var c = readSheet_(SHEET_CLASSES).filter(function (x) { return x.classId === classId; })[0];
  if (!c) throw new Error('クラスが見つかりません。');
  sheet.getRange(c._row, SHEET_DEFS[SHEET_CLASSES].indexOf('name') + 1).setValue(name);
  clearLoginCache_();
  return getClasses();
}

/** クラス削除（先生のみ）。生徒・ボードが残っている場合は削除できない。 */
function deleteClass(classId, teacherPassword) {
  if (!isTeacher_(teacherPassword)) throw new Error('クラスの操作は先生のみ可能です。');
  var hasStudents = readSheet_(SHEET_STUDENTS).some(function (s) { return String(s.classId || '') === String(classId); });
  var hasBoards = readSheet_(SHEET_BOARDS).some(function (b) { return String(b.classId || '') === String(classId); });
  if (hasStudents || hasBoards) throw new Error('このクラスには生徒またはボードが残っています。先に移動か削除をしてください。');
  deleteRowsWhere_(SHEET_CLASSES, 'classId', classId);
  clearLoginCache_();
  return getClasses();
}

/** デプロイ状態の診断用。クライアントの APP_BUILD と一致していれば最新。 */
function getServerInfo() {
  return { build: APP_BUILD, schema: SCHEMA_VERSION };
}

/** 生徒ログイン。初回（パスワード未設定）はここで設定します。 */
function studentLogin(classId, name, password) {
  password = String(password || '');
  if (password.length < 1) throw new Error('パスワードを入力してください。');
  var sheet = getSheet_(SHEET_STUDENTS);
  var rows = readSheet_(SHEET_STUDENTS);
  var st = rows.filter(function (r) { return r.name === name && String(r.classId || '') === String(classId || ''); })[0];
  if (!st) throw new Error('名簿に見つかりません。');

  var headers = SHEET_DEFS[SHEET_STUDENTS];
  var saltCol = headers.indexOf('salt') + 1;
  var hashCol = headers.indexOf('passwordHash') + 1;

  if (!st.passwordHash) {
    // 初回ログイン：パスワード設定
    var salt = newSalt_();
    sheet.getRange(st._row, saltCol, 1, 2).setValues([[salt, sha256_(salt + password)]]);
    clearLoginCache_();
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
  var hash = getProp_(PROP_TEACHER_HASH);
  if (!hash) {
    var salt = newSalt_();
    setProp_(PROP_TEACHER_SALT, salt);
    setProp_(PROP_TEACHER_HASH, sha256_(salt + password));
    clearLoginCache_();
    return { ok: true, firstTime: true, name: getTeacherName_() };
  }
  var salt2 = getProp_(PROP_TEACHER_SALT);
  if (sha256_(salt2 + password) === hash) return { ok: true, firstTime: false, name: getTeacherName_() };
  throw new Error('パスワードが違います。');
}

/** 先生の表示名（未設定なら「先生」）。 */
function getTeacherName_() {
  return getProp_(PROP_TEACHER_NAME) || '先生';
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
  setProp_(PROP_TEACHER_NAME, name);
  return { name: name };
}

/** 先生のパスワードを変更（現在のパスワードで本人確認）。 */
function setTeacherPassword(newPassword, teacherPassword) {
  if (!isTeacher_(teacherPassword)) throw new Error('現在のパスワードが正しくありません。');
  newPassword = String(newPassword || '');
  if (newPassword.length < 1) throw new Error('新しいパスワードを入力してください。');
  var salt = newSalt_();
  setProp_(PROP_TEACHER_SALT, salt);
  setProp_(PROP_TEACHER_HASH, sha256_(salt + newPassword));
  return { ok: true };
}

/** 生徒の本人確認（投稿・いいね等の操作前に呼ぶ簡易チェック）。classId は任意（あれば厳密化）。 */
function verifyStudent_(name, password, classId) {
  var st = readSheet_(SHEET_STUDENTS).filter(function (r) {
    if (r.name !== name) return false;
    if (classId != null && classId !== '' && String(r.classId || '') !== String(classId)) return false;
    return true;
  })[0];
  if (!st || !st.passwordHash) return false;
  return sha256_(st.salt + password) === st.passwordHash;
}

// --- 先生による名簿管理 ---

function getStudents(classId, teacherPassword) {
  requireTeacher_(teacherPassword, '名簿の閲覧');
  var students = readSheet_(SHEET_STUDENTS)
    .filter(function (s) { return String(s.classId || '') === String(classId || ''); })
    .map(function (s) { return { number: s.number, name: s.name, hasPassword: !!s.passwordHash }; });
  students.sort(function (a, b) { return (a.number || 0) - (b.number || 0); });
  return students;
}

function addStudent(number, name, classId, teacherPassword) {
  requireTeacher_(teacherPassword, '名簿の操作');
  name = String(name || '').trim();
  if (!name) throw new Error('名前を入力してください。');
  if (!classId) throw new Error('クラスを選択してください。');
  // 名前はログインの識別子なので、同じクラス内での重複は登録しない
  if (classNames_(classId)[name]) throw new Error('「' + name + '」はこのクラスにすでに登録されています。');
  var num = Number(number);
  getSheet_(SHEET_STUDENTS).appendRow([(num > 0 ? num : nextNumber_(classId)), name, '', '', new Date(), classId]);
  clearLoginCache_();
  return getStudents(classId, teacherPassword);
}

/** クラス内の登録済み名前の集合（重複登録の防止用）。 */
function classNames_(classId) {
  var names = {};
  readSheet_(SHEET_STUDENTS).forEach(function (s) {
    if (String(s.classId || '') === String(classId || '')) names[asText_(s.name)] = true;
  });
  return names;
}

/** 出席番号の変更（先生のみ）。名簿順の並べ替えに使われる。 */
function setStudentNumber(name, classId, number, teacherPassword) {
  requireTeacher_(teacherPassword, '名簿の操作');
  var num = Number(number);
  if (!(num > 0)) throw new Error('出席番号は1以上の数字で入力してください。');
  var sheet = getSheet_(SHEET_STUDENTS);
  var st = readSheet_(SHEET_STUDENTS).filter(function (r) {
    return asText_(r.name) === name && String(r.classId || '') === String(classId || '');
  })[0];
  if (!st) throw new Error('生徒が見つかりません。');
  sheet.getRange(st._row, SHEET_DEFS[SHEET_STUDENTS].indexOf('number') + 1).setValue(num);
  clearLoginCache_();
  return getStudents(classId, teacherPassword);
}

function nextNumber_(classId) {
  var max = 0;
  readSheet_(SHEET_STUDENTS).forEach(function (s) {
    if (String(s.classId || '') !== String(classId || '')) return;
    if (Number(s.number) > max) max = Number(s.number);
  });
  return max + 1;
}

/**
 * Excel などからの一括貼り付け。タブ / カンマ / 空白区切り、1行1名。
 * 「1<TAB>山田太郎」「山田太郎<TAB>1」「山田太郎」いずれも可。
 */
function importStudents(text, classId, teacherPassword) {
  requireTeacher_(teacherPassword, '名簿の操作');
  if (!classId) throw new Error('クラスを選択してください。');
  var lines = String(text || '').split(/\r?\n/);
  var sheet = getSheet_(SHEET_STUDENTS);
  var auto = nextNumber_(classId);
  var existing = classNames_(classId);
  var skipped = 0;
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
    // 同じクラスにすでにいる名前（二重貼り付けなど）は飛ばす
    if (existing[name]) { skipped++; return; }
    existing[name] = true;
    if (num === '') num = auto++;
    rows.push([num, name, '', '', new Date(), classId]);
  });
  if (rows.length) sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  clearLoginCache_();
  return { students: getStudents(classId, teacherPassword), added: rows.length, skipped: skipped };
}

function removeStudent(name, classId, teacherPassword) {
  requireTeacher_(teacherPassword, '名簿の操作');
  var sheet = getSheet_(SHEET_STUDENTS);
  withLock_(function () {
    var values = sheet.getDataRange().getValues();
    var headers = values[0];
    var nCol = headers.indexOf('name'), cCol = headers.indexOf('classId');
    for (var i = values.length - 1; i >= 1; i--) {
      if (asText_(values[i][nCol]) === name && String(values[i][cCol] || '') === String(classId || '')) sheet.deleteRow(i + 1);
    }
  });
  clearLoginCache_();
  return getStudents(classId, teacherPassword);
}

/** 生徒のパスワードをリセット（次回ログイン時に再設定）。 */
function resetStudentPassword(name, classId, teacherPassword) {
  requireTeacher_(teacherPassword, '名簿の操作');
  var sheet = getSheet_(SHEET_STUDENTS);
  var st = readSheet_(SHEET_STUDENTS).filter(function (r) {
    return r.name === name && String(r.classId || '') === String(classId || '');
  })[0];
  if (!st) return getStudents(classId, teacherPassword);
  var headers = SHEET_DEFS[SHEET_STUDENTS];
  // salt / passwordHash は隣り合う列なので1回で書く
  sheet.getRange(st._row, headers.indexOf('salt') + 1, 1, 2).setValues([['', '']]);
  clearLoginCache_();
  return getStudents(classId, teacherPassword);
}

// ============================ ボード ============================

/**
 * ボード一覧。includeArchived=true で非表示（アーカイブ）ボードも含めます。
 * 児童側（boardList）は false で呼ぶため、非表示ボードは一覧に出ません。
 */
function getBoards(includeArchived, classId) {
  var boards = readSheet_(SHEET_BOARDS).map(rowToBoard_);
  if (classId != null && classId !== '') boards = boards.filter(function (b) { return String(b.classId || '') === String(classId); });
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
    subject: asText_(r.subject),
    unit: asText_(r.unit),
    date: fmtDate_(r.date),
    title: asText_(r.title),
    createdAt: toMs_(r.createdAt),
    archived: r.archived === true || r.archived === 'true' || r.archived === 1,
    classId: r.classId || ''
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

/**
 * テキスト用の値を必ず文字列にして返す。
 * 「5/22」等のスラッシュ入り文字を Sheets が日付に変換してしまった場合でも、
 * Date オブジェクトのまま返すと google.script.run のシリアライズが壊れて
 * 読み込みエラーになるため、ここで安全に文字列へ戻す。
 */
function asText_(v) {
  if (v == null) return '';
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Tokyo', 'M/d');
  return String(v);
}

function getBoard(boardId) {
  var b = readSheet_(SHEET_BOARDS).filter(function (r) { return r.boardId === boardId; })[0];
  return b ? rowToBoard_(b) : null;
}

function createBoard(subject, unit, date, title, classId, teacherPassword) {
  requireTeacher_(teacherPassword, 'ボードの作成');
  subject = String(subject || '').trim();
  unit = String(unit || '').trim();
  if (!subject) throw new Error('教科を選択してください。');
  if (!unit) throw new Error('単元名を入力してください。');
  if (!classId) throw new Error('クラスを選択してください。');
  date = String(date || '').trim();
  if (!date) date = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  title = String(title || '').trim();
  if (!title) title = unit;
  var id = genId_('b');
  // 列順は SHEET_DEFS[SHEET_BOARDS] と一致させること。unit/title も日付化を防ぐためテキスト保存。
  getSheet_(SHEET_BOARDS).appendRow([id, subject, "'" + unit, "'" + date, "'" + title, new Date(), false, classId]);
  var created = getBoard(id);
  if (!created) throw new Error('ボードの作成に失敗しました。もう一度お試しください。');
  // セクションは作らない（0セクションから開始。先生が「＋セクション」で追加する）
  return created;
}

/**
 * ボードをコピー（先生のみ）。セクション構成（名前・色・並び）と「先生名義の投稿」を複製する。
 * 児童の投稿・コメント・いいねはコピーしない。targetClassId で別クラスにも複製可。
 */
function copyBoard(boardId, targetClassId, teacherPassword) {
  if (!isTeacher_(teacherPassword)) throw new Error('この操作は先生のみ可能です。');
  if (!targetClassId) throw new Error('コピー先のクラスを選んでください。');
  var src = readSheet_(SHEET_BOARDS).filter(function (r) { return r.boardId === boardId; })[0];
  if (!src) throw new Error('コピー元のボードが見つかりません。');

  var sameClass = String(src.classId || '') === String(targetClassId);
  var title = asText_(src.title);
  if (sameClass) title = title + ' のコピー';   // 同クラスは名前で区別

  var newId = genId_('b');
  var today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  // 列順は SHEET_DEFS[SHEET_BOARDS] と一致させること
  getSheet_(SHEET_BOARDS).appendRow([newId, asText_(src.subject), "'" + asText_(src.unit), "'" + today, "'" + title, new Date(), false, targetClassId]);

  // セクションを複製し、旧→新の対応表を作る
  var secSheet = getSheet_(SHEET_SECTIONS);
  var secMap = {};
  getSections(boardId).forEach(function (s) {
    var nid = genId_('s');
    secMap[s.sectionId] = nid;
    secSheet.appendRow([nid, newId, "'" + s.name, s.sortOrder, new Date(), s.color || '']);
  });

  // 先生名義の投稿だけを複製（児童の投稿は複製しない）
  var teacherName = getTeacherName_();
  var refSheet = getSheet_(SHEET_REFLECTIONS);
  var defR = SHEET_DEFS[SHEET_REFLECTIONS];
  readSheet_(SHEET_REFLECTIONS).forEach(function (r) {
    if (r.boardId !== boardId) return;
    if (asText_(r.studentName) !== teacherName) return;
    var newSec = r.sectionId ? (secMap[r.sectionId] || '') : '';
    var pinned = (r.pinned === true || r.pinned === 'true' || r.pinned === 1);
    // 写真/動画は同じ表示URLを使い回す（photoFileId は空にして、元ボード削除時の巻き込み消去を防ぐ）
    var row = [];
    row[defR.indexOf('reflectionId')] = genId_('r');
    row[defR.indexOf('boardId')] = newId;
    row[defR.indexOf('studentName')] = teacherName;
    row[defR.indexOf('text')] = "'" + asText_(r.text);
    row[defR.indexOf('photoUrl')] = r.photoUrl || '';
    row[defR.indexOf('photoFileId')] = '';
    row[defR.indexOf('color')] = r.color || '#fff7c0';
    row[defR.indexOf('sortOrder')] = Number(r.sortOrder) || 1;
    row[defR.indexOf('createdAt')] = new Date();
    row[defR.indexOf('sectionId')] = newSec;
    row[defR.indexOf('mediaType')] = r.mediaType || '';
    row[defR.indexOf('updatedAt')] = new Date();
    row[defR.indexOf('pinned')] = pinned;
    row[defR.indexOf('title')] = "'" + asText_(r.title);
    row[defR.indexOf('link')] = r.link || '';
    refSheet.appendRow(row);
  });
  return getBoard(newId);
}

function deleteBoard(boardId, teacherPassword) {
  requireTeacher_(teacherPassword, 'ボードの削除');
  // 紐づく振り返り・コメント・いいね・写真・セクションを削除（各シートは1回ずつまとめて処理）
  var refs = readSheet_(SHEET_REFLECTIONS).filter(function (r) { return r.boardId === boardId; });
  var refIds = refs.map(function (r) { return r.reflectionId; });
  refs.forEach(function (r) {
    if (r.photoFileId) { try { DriveApp.getFileById(r.photoFileId).setTrashed(true); } catch (e) {} }
  });
  deleteRowsWhereIn_(SHEET_COMMENTS, 'reflectionId', refIds);
  deleteRowsWhereIn_(SHEET_LIKES, 'reflectionId', refIds);
  deleteRowsWhere_(SHEET_REFLECTIONS, 'boardId', boardId);
  deleteRowsWhere_(SHEET_SECTIONS, 'boardId', boardId);
  deleteRowsWhere_(SHEET_BOARDS, 'boardId', boardId);
  clearSig_(boardId);
  return true;
}

// ============================ セクション ============================

function getSections(boardId) {
  var secs = readSheet_(SHEET_SECTIONS)
    .filter(function (s) { return s.boardId === boardId; })
    .map(function (s) { return { sectionId: s.sectionId, boardId: s.boardId, name: asText_(s.name), sortOrder: Number(s.sortOrder) || 0, color: s.color || '' }; });
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
  // 先頭に ' を付けて「テキスト」として保存（5/22 等が日付に変換されるのを防ぐ）
  getSheet_(SHEET_SECTIONS).appendRow([genId_('s'), boardId, "'" + name, maxOrder + 1, new Date(), String(color || '')]);
  clearSig_(boardId);
  return getSections(boardId);
}

/** セクションの色を変更（先生のみ）。 */
function setSectionColor(sectionId, color, teacherPassword) {
  if (!isTeacher_(teacherPassword)) throw new Error('セクションの操作は先生のみ可能です。');
  var sheet = getSheet_(SHEET_SECTIONS);
  var s = readSheet_(SHEET_SECTIONS).filter(function (x) { return x.sectionId === sectionId; })[0];
  if (!s) throw new Error('セクションが見つかりません。');
  sheet.getRange(s._row, SHEET_DEFS[SHEET_SECTIONS].indexOf('color') + 1).setValue(String(color || ''));
  clearSig_(s.boardId);
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
  sheet.getRange(s._row, SHEET_DEFS[SHEET_SECTIONS].indexOf('name') + 1).setValue("'" + name);
  clearSig_(s.boardId);
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
  clearSig_(boardId);
  return getSections(boardId);
}

/**
 * セクションの並び順を保存（先生のみ）。orderedIds は左から右の順のセクションID。
 * 渡されなかったセクションは末尾に元の順で残す。
 */
function reorderSections(boardId, orderedIds, teacherPassword) {
  if (!isTeacher_(teacherPassword)) throw new Error('セクションの操作は先生のみ可能です。');
  var sheet = getSheet_(SHEET_SECTIONS);
  var orderCol = SHEET_DEFS[SHEET_SECTIONS].indexOf('sortOrder') + 1;
  var rowById = {};
  readSheet_(SHEET_SECTIONS).forEach(function (s) {
    if (s.boardId === boardId) rowById[s.sectionId] = s._row;
  });
  var idx = 0;
  var placed = {};
  (orderedIds || []).forEach(function (id) {
    if (!rowById[id] || placed[id]) return;
    placed[id] = true;
    idx++;
    sheet.getRange(rowById[id], orderCol).setValue(idx);
  });
  // 並び替え対象に含まれなかったセクション（別端末での追加など）は末尾へ
  getSections(boardId).forEach(function (s) {
    if (placed[s.sectionId] || !rowById[s.sectionId]) return;
    idx++;
    sheet.getRange(rowById[s.sectionId], orderCol).setValue(idx);
  });
  clearSig_(boardId);
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

/**
 * ボードのカード一覧（いいね・コメント込み）。currentName で自分のリアクション判定。
 *
 * 【速度の考え方】
 *   ・全員分に共通の部分（board / sections / roster / cards）は、シグネチャをキーにして
 *     CacheService に gzip 圧縮して置く（loadBoardBundle_ が作る）。
 *   ・getBoardSignature の計算時に同じデータを読むので、その場でキャッシュを温めておく。
 *     → 変更を検知した30人が一斉に getBoardData を呼んでも、シートの読み直しはほぼ起きない。
 *   ・「自分のリアクション」だけを返す直前に付け足す（personalizeBundle_）。
 */
function getBoardData(boardId, currentName) {
  var sig = getBoardSignature(boardId);
  if (sig === SIG_GONE) throw new Error('ボードが見つかりません。');
  var bundle = cacheGetJson_(bundleKey_(boardId, sig));
  if (!bundle) {
    bundle = loadBoardBundle_(boardId);
    if (!bundle) throw new Error('ボードが見つかりません。');
    cachePutJson_(bundleKey_(boardId, bundle.sig), bundle, BUNDLE_TTL_SEC);
  }
  return personalizeBundle_(bundle, currentName);
}

var SIG_TTL_SEC = 8;       // シグネチャのキャッシュ秒数（書き込み時は clearSig_ で即無効化）
var BUNDLE_TTL_SEC = 120;  // ボードデータのキャッシュ秒数（キーにシグネチャを含むので古い版は自然に使われなくなる）
var SIG_GONE = 'gone';
function sigKey_(boardId) { return 'sig_' + boardId; }
function bundleKey_(boardId, sig) { return 'bd_' + boardId + '_' + hashString_(sig); }

/** 共通データから、ログイン中の人向けの応答を作る（reactors は名前一覧なので外へは出さない）。 */
function personalizeBundle_(bundle, currentName) {
  var cards = bundle.cards.map(function (c) {
    var mine = {};
    if (currentName && c.reactors) {
      Object.keys(c.reactors).forEach(function (t) {
        if (c.reactors[t].indexOf(currentName) >= 0) mine[t] = true;
      });
    }
    var out = {};
    Object.keys(c).forEach(function (k) { if (k !== 'reactors') out[k] = c[k]; });
    out.myReactions = mine;
    return out;
  });
  return { board: bundle.board, sections: bundle.sections, roster: bundle.roster, cards: cards, sig: bundle.sig };
}

/**
 * ボードの全データを読み、共通部分とシグネチャをまとめて返す（無ければ null）。
 * シグネチャは「投稿数 | 最終更新 | セクション | 並び・所属 | リアクション | コメント」を
 * ハッシュ化したもので、どれかが変われば必ず値が変わる。
 */
function loadBoardBundle_(boardId) {
  var board = getBoard(boardId);
  if (!board) return null;

  var refs = readSheet_(SHEET_REFLECTIONS).filter(function (r) { return r.boardId === boardId; });
  var refIds = {};
  refs.forEach(function (r) { refIds[r.reflectionId] = true; });

  // リアクション集計（種類別）。type 空は ❤ とみなす。誰が押したかは reactors に持つ（本人判定用）。
  var likeCount = {}, reactByRef = {}, reactorsByRef = {}, likeKeys = [];
  readSheet_(SHEET_LIKES).forEach(function (l) {
    if (!refIds[l.reflectionId]) return;
    var t = l.type || '❤';
    var who = asText_(l.studentName);
    likeCount[l.reflectionId] = (likeCount[l.reflectionId] || 0) + 1;
    var m = reactByRef[l.reflectionId] = reactByRef[l.reflectionId] || {};
    m[t] = (m[t] || 0) + 1;
    var rr = reactorsByRef[l.reflectionId] = reactorsByRef[l.reflectionId] || {};
    (rr[t] = rr[t] || []).push(who);
    likeKeys.push(l.reflectionId + ':' + who + ':' + t);
  });

  // コメント集計
  var byRef = {}, comKeys = [];
  readSheet_(SHEET_COMMENTS).forEach(function (c) {
    if (!refIds[c.reflectionId]) return;
    (byRef[c.reflectionId] = byRef[c.reflectionId] || []).push({
      commentId: c.commentId, author: asText_(c.author), text: asText_(c.text), createdAt: toMs_(c.createdAt)
    });
    comKeys.push(String(c.commentId));
  });
  Object.keys(byRef).forEach(function (k) {
    byRef[k].sort(function (a, b) { return (a.createdAt || 0) - (b.createdAt || 0); });
  });

  var maxMs = 0, layoutKeys = [];
  var cards = refs.map(function (r) {
    var created = toMs_(r.createdAt), updated = toMs_(r.updatedAt);
    var u = updated || created || 0;
    if (u > maxMs) maxMs = u;
    var sortOrder = Number(r.sortOrder) || 0, sectionId = r.sectionId || '';
    layoutKeys.push(r.reflectionId + ':' + sectionId + ':' + sortOrder);
    return {
      reflectionId: r.reflectionId,
      sectionId: sectionId,
      studentName: asText_(r.studentName),
      text: asText_(r.text),
      photoUrl: r.photoUrl,
      mediaType: r.mediaType || (r.photoUrl ? 'image' : ''),
      color: r.color,
      title: asText_(r.title),
      sortOrder: sortOrder,
      createdAt: created,
      updatedAt: updated,
      pinned: r.pinned === true || r.pinned === 'true' || r.pinned === 1,
      link: parseLink_(r.link),
      likeCount: likeCount[r.reflectionId] || 0,
      reactions: reactByRef[r.reflectionId] || {},
      reactors: reactorsByRef[r.reflectionId] || {},
      comments: byRef[r.reflectionId] || []
    };
  });
  cards.sort(cardCompare_);

  var sections = getSections(boardId);
  var secKeys = sections.map(function (s) { return s.sectionId + ':' + s.name + ':' + (s.color || '') + ':' + s.sortOrder; });

  var sig = [
    refs.length, maxMs,
    sections.length + '.' + hashString_(secKeys.sort().join('|')),
    hashString_(layoutKeys.sort().join('|')),
    hashString_(likeKeys.sort().join('|')),
    hashString_(comKeys.sort().join('|'))
  ].join('|');

  return { board: board, sections: sections, roster: getBoardRoster_(boardId, board), cards: cards, sig: sig };
}

/**
 * ボードのクラスの名簿（出席番号順の並べ替えに使う）。
 * 名前と出席番号だけを返す（パスワード等は返さない）。
 */
function getBoardRoster_(boardId, board) {
  board = board || getBoard(boardId);
  if (!board) return [];
  var classId = String(board.classId || '');
  var roster = readSheet_(SHEET_STUDENTS)
    .filter(function (s) { return String(s.classId || '') === classId; })
    .map(function (s) { return { number: Number(s.number) || 0, name: asText_(s.name) }; });
  roster.sort(function (a, b) { return a.number - b.number; });
  return roster;
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
 * 多数の児童が同時にポーリングしても重い計算を共有できるよう、短時間キャッシュする。
 * 書き込み時に clearSig_ で無効化するので、変更は即座に反映される。
 */
function getBoardSignature(boardId) {
  var cache = CacheService.getScriptCache();
  var key = sigKey_(boardId);
  var hit = cache.get(key);
  if (hit != null) return hit;
  // 同時に何台もミスしたときは、1台だけが計算して残りはその結果を使う
  var lock = LockService.getScriptLock();
  var locked = false;
  try { locked = lock.tryLock(4000); } catch (e) { locked = false; }
  try {
    if (locked) { hit = cache.get(key); if (hit != null) return hit; }
    var bundle = loadBoardBundle_(boardId);
    var sig = bundle ? bundle.sig : SIG_GONE;
    try { cache.put(key, sig, SIG_TTL_SEC); } catch (e) {}
    // 直後に来る getBoardData のために、同じデータをキャッシュしておく
    if (bundle) cachePutJson_(bundleKey_(boardId, sig), bundle, BUNDLE_TTL_SEC);
    return sig;
  } finally {
    if (locked) { try { lock.releaseLock(); } catch (e) {} }
  }
}
/** 文字列を短い数値文字列にまとめる（シグネチャ・キャッシュキー用の簡易ハッシュ）。 */
function hashString_(str) {
  var h = 0;
  str = String(str || '');
  for (var i = 0; i < str.length; i++) { h = (h * 31 + str.charCodeAt(i)) | 0; }
  return String(h);
}

/** ボードのシグネチャ・キャッシュを無効化（書き込み後に呼ぶと即反映される）。 */
function clearSig_(boardId) {
  if (!boardId) return;
  cacheRemove_(sigKey_(boardId));
}

/**
 * 投稿。media は { data(base64), mimeType, filename, kind:'image'|'video' } または null。
 */
function postReflection(boardId, sectionId, studentName, password, title, text, color, media, link, teacherPassword, classId) {
  // 先生は自分の表示名で投稿できる。それ以外は生徒の本人確認。
  var author;
  if (teacherPassword && isTeacher_(teacherPassword)) {
    author = getTeacherName_();
  } else {
    if (!verifyStudent_(studentName, password, classId)) throw new Error('ログイン情報が正しくありません。');
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
    id, boardId, author, "'" + text, url, fileId, color || '#fff7c0', maxOrder + 1, now, sectionId, mediaType, now, false, "'" + title,
    linkObj ? JSON.stringify(linkObj) : ''
  ]);
  clearSig_(boardId);
  // 速度重視：作成したカード1枚だけを返す（クライアントは部分描画する）
  return {
    card: {
      reflectionId: id, sectionId: sectionId, studentName: author, title: title, text: text,
      photoUrl: url, mediaType: mediaType, color: color || '#fff7c0', sortOrder: maxOrder + 1,
      createdAt: now.getTime(), updatedAt: now.getTime(), pinned: false,
      link: linkObj, reactions: {}, myReactions: {},
      likeCount: 0, comments: []
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
    favicon: String(s.favicon || '').slice(0, 1000),
    desc: String(s.desc || '').slice(0, 500),
    site: String(s.site || hostOf_(url)).slice(0, 200)
  };
}

/**
 * 投稿の編集。本人 または 先生のみ。
 * media を渡せば差し替え、removeMedia=true なら添付を削除、どちらも無ければ本文/色のみ更新。
 */
function editReflection(reflectionId, studentName, password, teacherPassword, title, text, color, media, removeMedia, link, removeLink, classId) {
  var sheet = getSheet_(SHEET_REFLECTIONS);
  var r = readSheet_(SHEET_REFLECTIONS).filter(function (x) { return x.reflectionId === reflectionId; })[0];
  if (!r) throw new Error('投稿が見つかりません。');
  var allowed = isTeacher_(teacherPassword) || (r.studentName === studentName && verifyStudent_(studentName, password, classId));
  if (!allowed) throw new Error('編集する権限がありません。');

  var def = SHEET_DEFS[SHEET_REFLECTIONS];
  title = String(title || '').trim();
  text = String(text || '').trim();
  var linkObj = sanitizeLink_(link);

  // 変更後に添付が残るか（先に検証し、空投稿になるなら何も書き換えない）
  var willHaveMedia = (media && media.data) ? true : (removeMedia ? false : !!r.photoFileId);
  var willHaveLink = linkObj ? true : (removeLink ? false : !!r.link);
  if (!title && !text && !willHaveMedia && !willHaveLink) throw new Error('タイトル・本文・写真・動画・リンクのいずれかは必要です。');

  sheet.getRange(r._row, def.indexOf('title') + 1).setValue("'" + title);
  sheet.getRange(r._row, def.indexOf('text') + 1).setValue("'" + text);
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
  clearSig_(r.boardId);
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
function togglePin(reflectionId, studentName, password, teacherPassword, classId) {
  var sheet = getSheet_(SHEET_REFLECTIONS);
  var r = readSheet_(SHEET_REFLECTIONS).filter(function (x) { return x.reflectionId === reflectionId; })[0];
  if (!r) throw new Error('投稿が見つかりません。');
  var allowed = isTeacher_(teacherPassword) || (r.studentName === studentName && verifyStudent_(studentName, password, classId));
  if (!allowed) throw new Error('ピン留めする権限がありません。');
  var def = SHEET_DEFS[SHEET_REFLECTIONS];
  var pinned = !(r.pinned === true || r.pinned === 'true' || r.pinned === 1);
  sheet.getRange(r._row, def.indexOf('pinned') + 1).setValue(pinned);
  sheet.getRange(r._row, def.indexOf('updatedAt') + 1).setValue(new Date());
  clearSig_(r.boardId);
  return { pinned: pinned };
}

/** 削除：本人 または 先生。 */
function deleteReflection(reflectionId, studentName, password, teacherPassword, classId) {
  var r = readSheet_(SHEET_REFLECTIONS).filter(function (x) { return x.reflectionId === reflectionId; })[0];
  if (!r) return true;
  var allowed = isTeacher_(teacherPassword) || (r.studentName === studentName && verifyStudent_(studentName, password, classId));
  if (!allowed) throw new Error('削除する権限がありません。');
  if (r.photoFileId) { try { DriveApp.getFileById(r.photoFileId).setTrashed(true); } catch (e) {} }
  deleteRowsWhere_(SHEET_COMMENTS, 'reflectionId', reflectionId);
  deleteRowsWhere_(SHEET_LIKES, 'reflectionId', reflectionId);
  deleteRowsWhere_(SHEET_REFLECTIONS, 'reflectionId', reflectionId);
  clearSig_(r.boardId);
  return true;
}

/**
 * ドラッグ＆ドロップ後：並び順とセクション移動を保存。items=[{id, sectionId}]（画面の並び順）。
 * ログイン中の生徒か先生のみ。カード1枚ずつセルを書くと枚数×2回の往復になるので、
 * sortOrder 列と sectionId 列をそれぞれ1回のまとめ書きで更新する。
 */
function updateLayout(boardId, items, studentName, password, teacherPassword, classId) {
  if (!(isTeacher_(teacherPassword) || verifyStudent_(studentName, password, classId))) {
    throw new Error('ログイン情報が正しくありません。');
  }
  var sheet = getSheet_(SHEET_REFLECTIONS);
  var def = SHEET_DEFS[SHEET_REFLECTIONS];
  var idIdx = def.indexOf('reflectionId'), boardIdx = def.indexOf('boardId');
  var orderIdx = def.indexOf('sortOrder'), secIdx = def.indexOf('sectionId');
  var want = {};
  (items || []).forEach(function (it, idx) { if (it && it.id && !want[it.id]) want[it.id] = { order: idx + 1, sec: String(it.sectionId || '') }; });
  withLock_(function () {
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return;
    var n = lastRow - 1;
    var width = Math.max(orderIdx, secIdx, idIdx, boardIdx) + 1;
    var data = sheet.getRange(2, 1, n, width).getValues();
    var orders = [], secs = [], changedOrder = false, changedSec = false;
    for (var i = 0; i < n; i++) {
      var row = data[i], w = want[String(row[idIdx])];
      var o = row[orderIdx], sc = row[secIdx];
      if (w && String(row[boardIdx]) === String(boardId)) {
        if (Number(o) !== w.order) { o = w.order; changedOrder = true; }
        if (String(sc || '') !== w.sec) { sc = w.sec; changedSec = true; }
      }
      orders.push([o]); secs.push([sc]);
    }
    if (changedOrder) sheet.getRange(2, orderIdx + 1, n, 1).setValues(orders);
    if (changedSec) sheet.getRange(2, secIdx + 1, n, 1).setValues(secs);
  });
  clearSig_(boardId);
  return true;
}

function isTeacher_(password) {
  if (!password) return false;
  var hash = getProp_(PROP_TEACHER_HASH);
  var salt = getProp_(PROP_TEACHER_SALT);
  return !!hash && sha256_(salt + password) === hash;
}
function requireTeacher_(password, what) {
  if (!isTeacher_(password)) throw new Error((what || 'この操作') + 'は先生のみ可能です。先生用ログインをし直してください。');
}

// --- リアクション（複数種類） ---
function toggleReaction(reflectionId, studentName, password, type, classId, boardId) {
  if (!verifyStudent_(studentName, password, classId)) throw new Error('ログイン情報が正しくありません。');
  type = String(type || '❤');
  if (REACTIONS.indexOf(type) < 0) type = '❤';
  var sheet = getSheet_(SHEET_LIKES);
  // シートは1回だけ読み、更新後の集計はその結果から組み立てる（読み直さない）
  var mineOnCard = readSheet_(SHEET_LIKES).filter(function (l) { return l.reflectionId === reflectionId; });
  var existing = mineOnCard.filter(function (l) {
    return l.studentName === studentName && (l.type || '❤') === type;
  })[0];
  if (existing) {
    withLock_(function () { sheet.deleteRow(existing._row); });
    mineOnCard = mineOnCard.filter(function (l) { return l !== existing; });
  } else {
    sheet.appendRow([reflectionId, studentName, new Date(), type]);
    mineOnCard.push({ reflectionId: reflectionId, studentName: studentName, type: type });
  }
  clearSig_(boardId);
  var reactions = {}, mine = {}, total = 0;
  mineOnCard.forEach(function (l) {
    var t = l.type || '❤';
    reactions[t] = (reactions[t] || 0) + 1;
    total++;
    if (l.studentName === studentName) mine[t] = true;
  });
  return { reactions: reactions, myReactions: mine, count: total };
}

// --- コメント ---
function addComment(reflectionId, author, password, text, classId, boardId) {
  if (!verifyStudent_(author, password, classId)) throw new Error('ログイン情報が正しくありません。');
  text = String(text || '').trim();
  if (!text) throw new Error('コメントを入力してください。');
  var id = genId_('c');
  var now = new Date();
  // 既存分は追記の前に1回だけ読み、新しいコメントを末尾に足して返す（読み直さない）
  var list = readSheet_(SHEET_COMMENTS)
    .filter(function (c) { return c.reflectionId === reflectionId; })
    .map(function (c) { return { commentId: c.commentId, author: asText_(c.author), text: asText_(c.text), createdAt: toMs_(c.createdAt) }; });
  getSheet_(SHEET_COMMENTS).appendRow([id, reflectionId, author, "'" + text, now]);
  clearSig_(boardId);
  list.push({ commentId: id, author: author, text: text, createdAt: now.getTime() });
  list.sort(function (a, b) { return (a.createdAt || 0) - (b.createdAt || 0); });
  return list;
}

// ============================ リンク（プレビューはクライアント側で生成） ============================
// ※ 外部リクエスト権限(UrlFetchApp)を使わない方針。プレビュー情報はクライアントが作って渡す。

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
function exportBoard(boardId, teacherPassword) {
  requireTeacher_(teacherPassword, '出力');
  return getBoardData(boardId, null);
}

/** 児童別の出力データ：その児童の（クラス内）全ボードにわたる振り返りを時系列で。 */
function exportStudent(studentName, classId, teacherPassword) {
  requireTeacher_(teacherPassword, '出力');
  var boards = {};
  readSheet_(SHEET_BOARDS).forEach(function (b) { boards[b.boardId] = rowToBoard_(b); });
  var refs = readSheet_(SHEET_REFLECTIONS)
    .filter(function (r) {
      if (r.studentName !== studentName) return false;
      if (classId) { var bd = boards[r.boardId]; if (!bd || String(bd.classId || '') !== String(classId)) return false; }
      return true;
    })
    .sort(function (a, b) { return new Date(a.createdAt) - new Date(b.createdAt); })
    .map(function (r) {
      var bd = boards[r.boardId] || {};
      return {
        boardTitle: bd.title || '(削除済みボード)',
        subject: bd.subject || '', unit: bd.unit || '', date: bd.date || '',
        title: asText_(r.title), text: asText_(r.text), photoUrl: r.photoUrl,
        mediaType: r.mediaType || (r.photoUrl ? 'image' : ''),
        link: parseLink_(r.link),
        color: r.color, createdAt: toMs_(r.createdAt)
      };
    });
  return { studentName: studentName, reflections: refs };
}
