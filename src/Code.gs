/**
 * 振り返りボード（Padlet風アプリ） - メインエントリ / ルーティング
 *
 * - スプレッドシートをデータベース、Google ドライブを写真置き場として使用します。
 * - 初回は setup() を一度だけ実行して、スプレッドシートとドライブフォルダを作成します。
 */

// Script Properties のキー
var PROP_SPREADSHEET_ID = 'SPREADSHEET_ID';
var PROP_FOLDER_ID = 'PHOTO_FOLDER_ID';
var PROP_TEACHER_PASSCODE = 'TEACHER_PASSCODE';

// シート名
var SHEET_CLASSES = 'Classes';
var SHEET_STUDENTS = 'Students';
var SHEET_BOARDS = 'Boards';
var SHEET_REFLECTIONS = 'Reflections';

/**
 * 初回セットアップ。GAS エディタから一度だけ手動実行してください。
 * - 専用スプレッドシートを作成
 * - 写真保存用ドライブフォルダを作成
 * - 先生用パスコードを初期化（変更可能）
 */
function setup() {
  var props = PropertiesService.getScriptProperties();

  // スプレッドシート作成
  if (!props.getProperty(PROP_SPREADSHEET_ID)) {
    var ss = SpreadsheetApp.create('振り返りボード データ');
    initSheets_(ss);
    props.setProperty(PROP_SPREADSHEET_ID, ss.getId());
    Logger.log('スプレッドシートを作成しました: ' + ss.getUrl());
  }

  // 写真フォルダ作成
  if (!props.getProperty(PROP_FOLDER_ID)) {
    var folder = DriveApp.createFolder('振り返りボード 写真');
    props.setProperty(PROP_FOLDER_ID, folder.getId());
    Logger.log('写真フォルダを作成しました: ' + folder.getUrl());
  }

  // 先生用パスコード初期化
  if (!props.getProperty(PROP_TEACHER_PASSCODE)) {
    props.setProperty(PROP_TEACHER_PASSCODE, '1234');
    Logger.log('先生用パスコードの初期値: 1234 （setTeacherPasscode で変更できます）');
  }

  Logger.log('セットアップ完了。Web アプリとしてデプロイしてください。');
}

/** 先生用パスコードを変更します。 */
function setTeacherPasscode(newPasscode) {
  PropertiesService.getScriptProperties().setProperty(PROP_TEACHER_PASSCODE, String(newPasscode));
}

/** 各シートのヘッダーを初期化します。 */
function initSheets_(ss) {
  var first = ss.getSheets()[0];
  first.setName(SHEET_CLASSES);
  first.getRange(1, 1, 1, 3).setValues([['classCode', 'className', 'createdAt']]);

  var students = ss.insertSheet(SHEET_STUDENTS);
  students.getRange(1, 1, 1, 4).setValues([['classCode', 'number', 'name', 'createdAt']]);

  var boards = ss.insertSheet(SHEET_BOARDS);
  boards.getRange(1, 1, 1, 6).setValues([['boardId', 'classCode', 'subject', 'date', 'title', 'createdAt']]);

  var reflections = ss.insertSheet(SHEET_REFLECTIONS);
  reflections.getRange(1, 1, 1, 8).setValues([
    ['reflectionId', 'boardId', 'studentName', 'text', 'photoUrl', 'photoFileId', 'color', 'createdAt']
  ]);
}

/** Web アプリのエントリポイント。 */
function doGet(e) {
  var template = HtmlService.createTemplateFromFile('index');
  return template.evaluate()
    .setTitle('振り返りボード')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setFaviconUrl('https://ssl.gstatic.com/docs/script/images/favicon.ico')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** HTML 内で他の HTML ファイルを読み込むためのヘルパー。 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ---- 内部ヘルパー ----

function getSpreadsheet_() {
  var id = PropertiesService.getScriptProperties().getProperty(PROP_SPREADSHEET_ID);
  if (!id) throw new Error('未セットアップです。GAS エディタから setup() を実行してください。');
  return SpreadsheetApp.openById(id);
}

function getSheet_(name) {
  return getSpreadsheet_().getSheetByName(name);
}

function getPhotoFolder_() {
  var id = PropertiesService.getScriptProperties().getProperty(PROP_FOLDER_ID);
  if (!id) throw new Error('未セットアップです。GAS エディタから setup() を実行してください。');
  return DriveApp.getFolderById(id);
}

/** ランダムな ID を生成します。 */
function genId_(prefix) {
  return prefix + '_' + Utilities.getUuid().slice(0, 8);
}

/** 4桁のクラスコード（英数字）を生成します。 */
function genClassCode_() {
  var chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  var code = '';
  for (var i = 0; i < 5; i++) {
    code += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return code;
}
