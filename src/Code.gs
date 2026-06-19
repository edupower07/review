/**
 * 振り返りボード（Padlet風アプリ） - サーバー（Google Apps Script）
 *
 * 外部APIは一切使いません。組み込みの SpreadsheetApp / DriveApp のみで動作します。
 * 1デプロイ＝1クラス。シートと写真フォルダは初回アクセス時に自動作成されます。
 */

var PROP_SPREADSHEET_ID = 'SPREADSHEET_ID';
var PROP_FOLDER_ID = 'PHOTO_FOLDER_ID';
var PROP_TEACHER_HASH = 'TEACHER_HASH';
var PROP_TEACHER_SALT = 'TEACHER_SALT';

var SHEET_STUDENTS = 'Students';
var SHEET_BOARDS = 'Boards';
var SHEET_REFLECTIONS = 'Reflections';
var SHEET_COMMENTS = 'Comments';
var SHEET_LIKES = 'Likes';

var SHEET_DEFS = {};
SHEET_DEFS[SHEET_STUDENTS] = ['number', 'name', 'salt', 'passwordHash', 'createdAt'];
SHEET_DEFS[SHEET_BOARDS] = ['boardId', 'subject', 'unit', 'date', 'title', 'createdAt'];
SHEET_DEFS[SHEET_REFLECTIONS] = ['reflectionId', 'boardId', 'studentName', 'text', 'photoUrl', 'photoFileId', 'color', 'sortOrder', 'createdAt'];
SHEET_DEFS[SHEET_COMMENTS] = ['commentId', 'reflectionId', 'author', 'text', 'createdAt'];
SHEET_DEFS[SHEET_LIKES] = ['reflectionId', 'studentName', 'createdAt'];

// ============================ エントリ ============================

function doGet() {
  ensureInit_();
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('振り返りボード')
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

  var ss;
  var ssId = props.getProperty(PROP_SPREADSHEET_ID);
  if (ssId) {
    try { ss = SpreadsheetApp.openById(ssId); } catch (e) { ss = null; }
  }
  if (!ss) {
    ss = SpreadsheetApp.create('振り返りボード データ');
    props.setProperty(PROP_SPREADSHEET_ID, ss.getId());
  }

  // 各シートを保証（ヘッダーが旧版・不一致なら作り直す）
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
    if (!matches) {
      if (lastRow <= 1) {
        // データがまだ無いので安全にヘッダーを置き換える
        if (lastRow >= 1) sheet.getRange(1, 1, 1, width).clearContent();
        sheet.getRange(1, 1, 1, def.length).setValues([def]);
      } else {
        // データがある旧スキーマのシートは破壊せず退避し、新しい空シートを作る
        var backup = name + '_旧_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
        sheet.setName(backup);
        var fresh = ss.insertSheet(name);
        fresh.getRange(1, 1, 1, def.length).setValues([def]);
      }
    }
  });
  // 自動生成された空の "シート1"/"Sheet1" を削除
  ['シート1', 'Sheet1'].forEach(function (n) {
    var s = ss.getSheetByName(n);
    if (s && ss.getSheets().length > 1) { try { ss.deleteSheet(s); } catch (e) {} }
  });

  if (!props.getProperty(PROP_FOLDER_ID)) {
    var folder = DriveApp.createFolder('振り返りボード 写真');
    props.setProperty(PROP_FOLDER_ID, folder.getId());
  }
  _ssCache = ss;
  return ss;
}

function getSpreadsheet_() {
  return ensureInit_();
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
  return { students: students, teacherSet: teacherSet };
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
    return { ok: true, firstTime: true };
  }
  var salt2 = props.getProperty(PROP_TEACHER_SALT);
  if (sha256_(salt2 + password) === hash) return { ok: true, firstTime: false };
  throw new Error('パスワードが違います。');
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

function getBoards() {
  var boards = readSheet_(SHEET_BOARDS).map(rowToBoard_);
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
    createdAt: r.createdAt
  };
}
function fmtDate_(d) {
  if (d instanceof Date) return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd');
  if (d == null) return '';
  return String(d);
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
  // 列順は SHEET_DEFS[SHEET_BOARDS] と一致させること（boardId, subject, unit, date, title, createdAt）
  getSheet_(SHEET_BOARDS).appendRow([id, subject, unit, "'" + date, title, new Date()]);
  var created = getBoard(id);
  if (!created) throw new Error('ボードの作成に失敗しました。もう一度お試しください。');
  return created;
}

function deleteBoard(boardId) {
  // 紐づく振り返り・コメント・いいね・写真を削除
  var refs = readSheet_(SHEET_REFLECTIONS).filter(function (r) { return r.boardId === boardId; });
  refs.forEach(function (r) {
    if (r.photoFileId) { try { DriveApp.getFileById(r.photoFileId).setTrashed(true); } catch (e) {} }
    deleteRowsWhere_(SHEET_COMMENTS, 'reflectionId', r.reflectionId);
    deleteRowsWhere_(SHEET_LIKES, 'reflectionId', r.reflectionId);
  });
  deleteRowsWhere_(SHEET_REFLECTIONS, 'boardId', boardId);
  deleteRowsWhere_(SHEET_BOARDS, 'boardId', boardId);
  return true;
}

// ============================ 振り返り（カード） ============================

/** ボードのカード一覧（いいね・コメント込み）。currentName で自分のいいね判定。 */
function getBoardData(boardId, currentName) {
  var board = getBoard(boardId);
  if (!board) throw new Error('ボードが見つかりません。');

  var refs = readSheet_(SHEET_REFLECTIONS).filter(function (r) { return r.boardId === boardId; });

  // いいね集計
  var likes = readSheet_(SHEET_LIKES);
  var likeCount = {}, likedByMe = {};
  likes.forEach(function (l) {
    likeCount[l.reflectionId] = (likeCount[l.reflectionId] || 0) + 1;
    if (l.studentName === currentName) likedByMe[l.reflectionId] = true;
  });

  // コメント集計
  var comments = readSheet_(SHEET_COMMENTS);
  var byRef = {};
  comments.forEach(function (c) {
    (byRef[c.reflectionId] = byRef[c.reflectionId] || []).push({
      commentId: c.commentId, author: c.author, text: c.text, createdAt: c.createdAt
    });
  });
  Object.keys(byRef).forEach(function (k) {
    byRef[k].sort(function (a, b) { return new Date(a.createdAt) - new Date(b.createdAt); });
  });

  var cards = refs.map(function (r) {
    return {
      reflectionId: r.reflectionId,
      studentName: r.studentName,
      text: r.text,
      photoUrl: r.photoUrl,
      color: r.color,
      sortOrder: Number(r.sortOrder) || 0,
      createdAt: r.createdAt,
      likeCount: likeCount[r.reflectionId] || 0,
      likedByMe: !!likedByMe[r.reflectionId],
      comments: byRef[r.reflectionId] || []
    };
  });
  cards.sort(function (a, b) {
    if (a.sortOrder !== b.sortOrder) return a.sortOrder - b.sortOrder;
    return new Date(a.createdAt) - new Date(b.createdAt);
  });

  return { board: board, cards: cards };
}

function postReflection(boardId, studentName, password, text, color, photo) {
  if (!verifyStudent_(studentName, password)) throw new Error('ログイン情報が正しくありません。');
  text = String(text || '').trim();
  if (!text && !photo) throw new Error('テキストか写真を入力してください。');
  var board = getBoard(boardId);
  if (!board) throw new Error('ボードが見つかりません。');

  var photoUrl = '', photoFileId = '';
  if (photo && photo.data) {
    var saved = savePhoto_(photo, board);
    photoUrl = saved.url; photoFileId = saved.fileId;
  }

  // 末尾に並べる
  var maxOrder = 0;
  readSheet_(SHEET_REFLECTIONS).forEach(function (r) {
    if (r.boardId === boardId && Number(r.sortOrder) > maxOrder) maxOrder = Number(r.sortOrder);
  });

  var id = genId_('r');
  getSheet_(SHEET_REFLECTIONS).appendRow([
    id, boardId, studentName, text, photoUrl, photoFileId, color || '#fff7c0', maxOrder + 1, new Date()
  ]);
  return getBoardData(boardId, studentName);
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

// --- いいね ---
function toggleLike(reflectionId, studentName, password) {
  if (!verifyStudent_(studentName, password)) throw new Error('ログイン情報が正しくありません。');
  var sheet = getSheet_(SHEET_LIKES);
  var existing = readSheet_(SHEET_LIKES).filter(function (l) {
    return l.reflectionId === reflectionId && l.studentName === studentName;
  })[0];
  if (existing) {
    sheet.deleteRow(existing._row);
  } else {
    sheet.appendRow([reflectionId, studentName, new Date()]);
  }
  var count = readSheet_(SHEET_LIKES).filter(function (l) { return l.reflectionId === reflectionId; }).length;
  return { count: count, liked: !existing };
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
    .map(function (c) { return { commentId: c.commentId, author: c.author, text: c.text, createdAt: c.createdAt }; });
}

// ============================ 写真 ============================

function savePhoto_(photo, board) {
  var folder = getPhotoFolder_();
  var subName = board.subject + '_' + board.unit;
  var sub = getOrCreateSubfolder_(folder, subName);
  var bytes = Utilities.base64Decode(photo.data);
  var blob = Utilities.newBlob(bytes, photo.mimeType || 'image/jpeg', photo.filename || ('photo_' + Date.now() + '.jpg'));
  var file = sub.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { fileId: file.getId(), url: 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1000' };
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
        text: r.text, photoUrl: r.photoUrl, color: r.color, createdAt: r.createdAt
      };
    });
  return { studentName: studentName, reflections: refs };
}
