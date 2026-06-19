/**
 * クライアント（google.script.run）から呼び出すサーバー API。
 */

// ---------- クラス / 名簿 ----------

/** 先生用パスコードの照合。 */
function verifyTeacher(passcode) {
  var saved = PropertiesService.getScriptProperties().getProperty(PROP_TEACHER_PASSCODE);
  return String(passcode) === String(saved);
}

/** 全クラスを取得します（先生画面用）。 */
function getClasses() {
  var sheet = getSheet_(SHEET_CLASSES);
  var values = sheet.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < values.length; i++) {
    if (!values[i][0]) continue;
    out.push({ classCode: values[i][0], className: values[i][1] });
  }
  return out;
}

/** クラスを作成し、生成したクラスコードを返します。 */
function createClass(className) {
  className = String(className || '').trim();
  if (!className) throw new Error('クラス名を入力してください。');
  var sheet = getSheet_(SHEET_CLASSES);

  // 重複しないコードを生成
  var existing = {};
  getClasses().forEach(function (c) { existing[c.classCode] = true; });
  var code;
  do { code = genClassCode_(); } while (existing[code]);

  sheet.appendRow([code, className, new Date()]);
  return { classCode: code, className: className };
}

/** クラスを削除します（所属生徒・ボード・振り返りも削除）。 */
function deleteClass(classCode) {
  deleteRowsWhere_(getSheet_(SHEET_CLASSES), 0, classCode);
  deleteRowsWhere_(getSheet_(SHEET_STUDENTS), 0, classCode);
  // クラスに紐づくボードを取得し、各ボードの振り返りも削除
  var boards = getBoards(classCode);
  boards.forEach(function (b) { deleteBoard(b.boardId); });
  return true;
}

/** クラスの生徒一覧を取得します（出席番号順）。 */
function getStudents(classCode) {
  var sheet = getSheet_(SHEET_STUDENTS);
  var values = sheet.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < values.length; i++) {
    if (values[i][0] === classCode) {
      out.push({ number: values[i][1], name: values[i][2] });
    }
  }
  out.sort(function (a, b) { return (a.number || 0) - (b.number || 0); });
  return out;
}

/** 生徒を1人追加します。 */
function addStudent(classCode, number, name) {
  name = String(name || '').trim();
  if (!name) throw new Error('名前を入力してください。');
  getSheet_(SHEET_STUDENTS).appendRow([classCode, number || '', name, new Date()]);
  return getStudents(classCode);
}

/**
 * 名簿を一括登録します。1行1名。
 * 「1,山田太郎」「1 山田太郎」「山田太郎」いずれの形式も可。
 */
function importStudents(classCode, text) {
  var lines = String(text || '').split(/\r?\n/);
  var sheet = getSheet_(SHEET_STUDENTS);
  var rows = [];
  lines.forEach(function (line) {
    line = line.trim();
    if (!line) return;
    var m = line.match(/^(\d+)[\s,、\t]+(.+)$/);
    if (m) {
      rows.push([classCode, Number(m[1]), m[2].trim(), new Date()]);
    } else {
      rows.push([classCode, '', line, new Date()]);
    }
  });
  if (rows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 4).setValues(rows);
  }
  return getStudents(classCode);
}

/** 生徒を削除します。 */
function removeStudent(classCode, name) {
  var sheet = getSheet_(SHEET_STUDENTS);
  var values = sheet.getDataRange().getValues();
  for (var i = values.length - 1; i >= 1; i--) {
    if (values[i][0] === classCode && values[i][2] === name) {
      sheet.deleteRow(i + 1);
    }
  }
  return getStudents(classCode);
}

// ---------- ボード ----------

/** クラスコードが存在するか確認し、クラス情報を返します（生徒の入口）。 */
function getClassInfo(classCode) {
  classCode = String(classCode || '').trim().toUpperCase();
  var found = getClasses().filter(function (c) { return c.classCode === classCode; })[0];
  if (!found) return null;
  return { classCode: found.classCode, className: found.className };
}

/** クラスのボード一覧を取得します（新しい順）。 */
function getBoards(classCode) {
  var sheet = getSheet_(SHEET_BOARDS);
  var values = sheet.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < values.length; i++) {
    if (values[i][0] && values[i][1] === classCode) {
      out.push(rowToBoard_(values[i]));
    }
  }
  out.sort(function (a, b) { return new Date(b.createdAt) - new Date(a.createdAt); });
  return out;
}

function rowToBoard_(r) {
  return {
    boardId: r[0],
    classCode: r[1],
    subject: r[2],
    date: r[3] instanceof Date ? Utilities.formatDate(r[3], 'Asia/Tokyo', 'yyyy-MM-dd') : r[3],
    title: r[4],
    createdAt: r[5]
  };
}

/** 1件のボード情報を取得します。 */
function getBoard(boardId) {
  var sheet = getSheet_(SHEET_BOARDS);
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (values[i][0] === boardId) return rowToBoard_(values[i]);
  }
  return null;
}

/** ボードを作成します。 */
function createBoard(classCode, subject, date, title) {
  subject = String(subject || '').trim();
  if (!subject) throw new Error('教科を選択してください。');
  if (!date) throw new Error('日付を選択してください。');
  var boardId = genId_('b');
  if (!title) title = subject + ' ' + date;
  getSheet_(SHEET_BOARDS).appendRow([boardId, classCode, subject, date, title, new Date()]);
  return getBoard(boardId);
}

/** ボードを削除します（紐づく振り返りも削除）。 */
function deleteBoard(boardId) {
  deleteRowsWhere_(getSheet_(SHEET_BOARDS), 0, boardId);
  deleteRowsWhere_(getSheet_(SHEET_REFLECTIONS), 1, boardId);
  return true;
}

// ---------- 振り返り（カード） ----------

/** ボードの振り返りを取得します（古い順 = 投稿順）。 */
function getReflections(boardId) {
  var sheet = getSheet_(SHEET_REFLECTIONS);
  var values = sheet.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < values.length; i++) {
    if (values[i][1] === boardId) {
      out.push({
        reflectionId: values[i][0],
        boardId: values[i][1],
        studentName: values[i][2],
        text: values[i][3],
        photoUrl: values[i][4],
        photoFileId: values[i][5],
        color: values[i][6],
        createdAt: values[i][7]
      });
    }
  }
  out.sort(function (a, b) { return new Date(a.createdAt) - new Date(b.createdAt); });
  return out;
}

/**
 * 振り返りを投稿します。
 * photo は { data: base64文字列, mimeType, filename } または null。
 */
function postReflection(boardId, studentName, text, color, photo) {
  text = String(text || '').trim();
  if (!text && !photo) throw new Error('テキストか写真のどちらかを入力してください。');

  var board = getBoard(boardId);
  if (!board) throw new Error('ボードが見つかりません。');

  var photoUrl = '';
  var photoFileId = '';
  if (photo && photo.data) {
    var saved = savePhoto_(photo, board);
    photoUrl = saved.url;
    photoFileId = saved.fileId;
  }

  var id = genId_('r');
  getSheet_(SHEET_REFLECTIONS).appendRow([
    id, boardId, studentName || '名無し', text, photoUrl, photoFileId, color || '#fff7c0', new Date()
  ]);

  return {
    reflectionId: id,
    boardId: boardId,
    studentName: studentName || '名無し',
    text: text,
    photoUrl: photoUrl,
    photoFileId: photoFileId,
    color: color || '#fff7c0',
    createdAt: new Date()
  };
}

/** 振り返りを削除します（写真も削除）。 */
function deleteReflection(reflectionId) {
  var sheet = getSheet_(SHEET_REFLECTIONS);
  var values = sheet.getDataRange().getValues();
  for (var i = values.length - 1; i >= 1; i--) {
    if (values[i][0] === reflectionId) {
      var fileId = values[i][5];
      if (fileId) {
        try { DriveApp.getFileById(fileId).setTrashed(true); } catch (err) {}
      }
      sheet.deleteRow(i + 1);
    }
  }
  return true;
}

/** 写真を Drive に保存し、表示用 URL を返します。 */
function savePhoto_(photo, board) {
  var folder = getPhotoFolder_();
  // 教科/日付ごとにサブフォルダを分けて整理
  var subName = board.subject + '_' + board.date;
  var sub = getOrCreateSubfolder_(folder, subName);

  var bytes = Utilities.base64Decode(photo.data);
  var mime = photo.mimeType || 'image/jpeg';
  var name = photo.filename || ('photo_' + new Date().getTime() + '.jpg');
  var blob = Utilities.newBlob(bytes, mime, name);
  var file = sub.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return {
    fileId: file.getId(),
    url: 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1000'
  };
}

function getOrCreateSubfolder_(parent, name) {
  var it = parent.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return parent.createFolder(name);
}

// ---------- 汎用 ----------

/** 指定列が値に一致する行をすべて削除します。 */
function deleteRowsWhere_(sheet, colIndex, value) {
  var values = sheet.getDataRange().getValues();
  for (var i = values.length - 1; i >= 1; i--) {
    if (values[i][colIndex] === value) {
      sheet.deleteRow(i + 1);
    }
  }
}
