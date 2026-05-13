function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setTitle('NiceNotes3 · 会議資料ワークスペース');
}

/**
 * ホーム画面の並び・フォルダ閲覧モード等（JSON を ScriptProperties に保存）。
 * @return {{ folderPdfMode: Object, folderKind: Object, fileLogicalPages: Object, layout: { childrenOrder: Object, fileOrder: Object, fileDue: Object } }}
 */
function nn_getWorkspaceMeta_() {
  const def = {
    folderPdfMode: {},
    folderKind: {},
    fileLogicalPages: {},
    layout: { childrenOrder: {}, fileOrder: {}, fileDue: {} },
  };
  try {
    const s = PropertiesService.getScriptProperties().getProperty('NN_WORKSPACE_META');
    if (!s) return def;
    const o = JSON.parse(s);
    if (!o.folderPdfMode || typeof o.folderPdfMode !== 'object') o.folderPdfMode = {};
    if (!o.folderKind || typeof o.folderKind !== 'object') o.folderKind = {};
    if (!o.fileLogicalPages || typeof o.fileLogicalPages !== 'object') o.fileLogicalPages = {};
    if (!o.layout || typeof o.layout !== 'object') o.layout = {};
    if (!o.layout.childrenOrder) o.layout.childrenOrder = {};
    if (!o.layout.fileOrder) o.layout.fileOrder = {};
    if (!o.layout.fileDue) o.layout.fileDue = {};
    return o;
  } catch (e) {
    return def;
  }
}

/**
 * @param {string} jsonString JSON 文字列（全体を上書き）
 * @return {{ success: boolean, error?: string }}
 */
function saveWorkspaceMeta(jsonString) {
  try {
    JSON.parse(jsonString);
    PropertiesService.getScriptProperties().setProperty('NN_WORKSPACE_META', jsonString);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * フォルダを別の親フォルダへ移動（Drive 上の階層が変わる）。
 * @param {string} folderId
 * @param {string} newParentFolderId
 * @return {{ success: boolean, error?: string }}
 */
function nn_moveFolderToParent(folderId, newParentFolderId) {
  if (!folderId || !newParentFolderId || folderId === newParentFolderId) {
    return { success: false, error: '無効な移動先です。' };
  }
  try {
    DriveApp.getFolderById(folderId).moveTo(DriveApp.getFolderById(newParentFolderId));
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function getFileData(fileId) {
  try {
    const file = DriveApp.getFileById(fileId);
    return Utilities.base64Encode(file.getBlob().getBytes());
  } catch (e) {
    return { error: e.toString() };
  }
}

function saveAnnotation(fileName, jsonData) {
  try {
    const folder = DriveApp.getRootFolder();
    const targetName = fileName + '_rev.json';
    const files = folder.getFilesByName(targetName);
    if (files.hasNext()) {
      files.next().setContent(jsonData); 
    } else {
      folder.createFile(targetName, jsonData, MimeType.PLAIN_TEXT); 
    }
    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function loadAnnotation(fileName) {
  try {
    const files = DriveApp.getRootFolder().getFilesByName(fileName + '_rev.json');
    if (files.hasNext()) return { success: true, data: files.next().getBlob().getDataAsString() };
    return { success: true, data: null };
  } catch (e) { return { success: false, error: e.toString() }; }
}

// 【新規追加】ストローク配列を受け取りGoogle APIへ送る関数
function recognizeSentence(allStrokes) {
  const url = "https://www.google.com.hk/inputtools/request?ime=handwriting&app=mobilesearch&cs=1&oe=UTF-8";
  const payload = {
    "options": "enable_pre_space",
    "requests": [{
      "writing_guide": { "writing_area_width": 1000, "writing_area_height": 1000 },
      "ink": allStrokes,
      "language": "ja"
    }]
  };

  try {
    const response = UrlFetchApp.fetch(url, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload)
    });
    const result = JSON.parse(response.getContentText());
    if (result[0] === "SUCCESS") {
      return result[1][0][1][0];
    }
  } catch (e) {
    return null;
  }
  return null;
}

// =============================================================================
// 職員会議ワークスペース（Drive ツリー・資料同期・手書き共有）
// google.script.run で index.html から呼び出す。ScriptProperties キーは元アプリ互換。
// =============================================================================

var NN_USER_MAIN_FOLDER_IDS_KEY = 'NN_USER_MAIN_FOLDER_IDS';

function nn_getScriptOwnerEmail_() {
  try {
    const f = DriveApp.getFileById(ScriptApp.getScriptId());
    return String(f.getOwner().getEmail() || '')
      .toLowerCase()
      .trim();
  } catch (e) {
    return '';
  }
}

function nn_getActiveUserEmail_() {
  try {
    return String(Session.getActiveUser().getEmail() || '')
      .toLowerCase()
      .trim();
  } catch (e) {
    return '';
  }
}

function nn_parseUserFolderMap_(props) {
  const raw = props.getProperty(NN_USER_MAIN_FOLDER_IDS_KEY);
  if (!raw) return {};
  try {
    const o = JSON.parse(raw);
    return o && typeof o === 'object' ? o : {};
  } catch (e) {
    return {};
  }
}

function nn_setUserFolderMap_(props, map) {
  props.setProperty(NN_USER_MAIN_FOLDER_IDS_KEY, JSON.stringify(map));
}

function nn_sameEmail_(a, b) {
  return !!(a && b && a === b);
}

/**
 * ログインユーザーごとの MyNiceNotes ルート。非オーナーは共有の MAIN_FOLDER_ID を使わない。
 */
function nn_getMainFolderIdForCurrentUser_(props) {
  const active = nn_getActiveUserEmail_();
  const map = nn_parseUserFolderMap_(props);
  if (active && map[active]) return map[active];
  const legacy = props.getProperty('MAIN_FOLDER_ID');
  const owner = nn_getScriptOwnerEmail_();
  if (legacy && nn_sameEmail_(owner, active)) return legacy;
  if (legacy && !active) return legacy;
  return null;
}

function nn_setMainFolderIdForCurrentUser_(props, folderId) {
  const active = nn_getActiveUserEmail_();
  if (!active) {
    props.setProperty('MAIN_FOLDER_ID', folderId);
    return;
  }
  const map = nn_parseUserFolderMap_(props);
  map[active] = folderId;
  nn_setUserFolderMap_(props, map);
  const owner = nn_getScriptOwnerEmail_();
  if (nn_sameEmail_(owner, active)) props.setProperty('MAIN_FOLDER_ID', folderId);
}

function initializeApp() {
  const props = PropertiesService.getScriptProperties();
  const existing = nn_getMainFolderIdForCurrentUser_(props);
  if (existing) {
    try {
      DriveApp.getFolderById(existing);
      return { success: true, mainFolderId: existing };
    } catch (e) {
      /* フォルダ欠損時は再作成 */
    }
  }
  try {
    const active = nn_getActiveUserEmail_();
    const owner = nn_getScriptOwnerEmail_();
    let parentFolder;
    if (nn_sameEmail_(owner, active)) {
      const scriptId = ScriptApp.getScriptId();
      const parents = DriveApp.getFileById(scriptId).getParents();
      parentFolder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
    } else {
      parentFolder = DriveApp.getRootFolder();
    }
    const newFolder = parentFolder.createFolder('MyNiceNotes');
    const mainFolderId = newFolder.getId();
    nn_setMainFolderIdForCurrentUser_(props, mainFolderId);
    return { success: true, mainFolderId: mainFolderId };
  } catch (e2) {
    return { success: false, error: '初期化に失敗しました: ' + e2.toString() };
  }
}

/**
 * クイック入力（音声メモ等）で Blob を Drive のワークスペース配下に保存する。
 * @param {string} base64 Base64（データ URL プレフィックスなし）
 * @param {string} mimeType MIME
 * @param {string} fileName ファイル名
 * @return {{ success: boolean, url?: string, fileId?: string, error?: string }}
 */
function nnSaveCaptureToDrive(base64, mimeType, fileName) {
  try {
    const props = PropertiesService.getScriptProperties();
    const mainFolderId = nn_getMainFolderIdForCurrentUser_(props);
    if (!mainFolderId) {
      return { success: false, error: '管理フォルダが未初期化です。ホームを開いて初期化してください。' };
    }
    const main = DriveApp.getFolderById(mainFolderId);
    let subIter = main.getFoldersByName('MyNiceNotes_Captures');
    if (!subIter.hasNext()) {
      subIter = main.getFoldersByName('NiceNotes_Captures');
    }
    const sub = subIter.hasNext() ? subIter.next() : main.createFolder('MyNiceNotes_Captures');
    const bytes = Utilities.base64Decode(base64);
    const blob = Utilities.newBlob(bytes, mimeType || 'application/octet-stream', fileName || 'capture');
    const file = sub.createFile(blob);
    return { success: true, url: file.getUrl(), fileId: file.getId() };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function extractIdFromUrl(input) {
  const match = String(input || '').match(/[-\w]{25,}/);
  return match ? match[0] : String(input || '');
}

function registerFolder(inputData, isOrg) {
  const id = extractIdFromUrl(inputData);
  try {
    DriveApp.getFolderById(id);
    const props = PropertiesService.getScriptProperties();
    let folders = JSON.parse(props.getProperty('REGISTERED_FOLDERS') || '[]');
    if (folders.indexOf(id) === -1) {
      folders.push(id);
      props.setProperty('REGISTERED_FOLDERS', JSON.stringify(folders));
    }
    return { success: true, folderId: id };
  } catch (e) {
    if (isOrg) {
      return {
        success: false,
        error:
          '❌ 組織のフォルダにアクセスできません。\n\n【確認事項】\n該当フォルダの共有設定で「リンクを知っている全員」に権限が付与されているか、またはあなたのアカウントにアクセス権があるか、管理者に確認を促してください。',
      };
    }
    return {
      success: false,
      error: '❌ フォルダが見つかりません。URLまたはIDが正しいか確認してください。',
    };
  }
}

function importPdf(inputData) {
  const id = extractIdFromUrl(inputData);
  try {
    const props = PropertiesService.getScriptProperties();
    const mainFolderId = nn_getMainFolderIdForCurrentUser_(props);
    if (!mainFolderId) {
      return { success: false, error: '管理フォルダが未初期化です。' };
    }
    const mainFolder = DriveApp.getFolderById(mainFolderId);

    const originalFile = DriveApp.getFileById(id);
    if (originalFile.getMimeType() !== MimeType.PDF) {
      return { success: false, error: '指定されたファイルはPDFではありません。' };
    }

    const copyName = '[編集用] ' + originalFile.getName();
    const copyFile = originalFile.makeCopy(copyName, mainFolder);

    return { success: true, newFileId: copyFile.getId(), newFileName: copyFile.getName() };
  } catch (e) {
    return {
      success: false,
      error: '❌ ファイルの取り込みに失敗しました。権限やURLを確認してください。\n詳細: ' + e.toString(),
    };
  }
}

function setFolderColor(folderId, colorCode) {
  const props = PropertiesService.getScriptProperties();
  let colors = JSON.parse(props.getProperty('FOLDER_COLORS') || '{}');
  colors[folderId] = colorCode;
  props.setProperty('FOLDER_COLORS', JSON.stringify(colors));

  const cacheFileName = 'app_folder_tree_cache.json';
  const files = DriveApp.getRootFolder().getFilesByName(cacheFileName);
  if (files.hasNext()) {
    try {
      const cacheFile = files.next();
      const data = JSON.parse(cacheFile.getBlob().getDataAsString());
      data.folderColors = colors;
      cacheFile.setContent(JSON.stringify(data));
    } catch (e) {
      /* ignore */
    }
  }
  return { success: true };
}

function getFolderTree(forceRefresh) {
  const cacheFileName = 'app_folder_tree_cache.json';
  const rootFolder = DriveApp.getRootFolder();
  let cacheFile = null;
  const files = rootFolder.getFilesByName(cacheFileName);
  if (files.hasNext()) cacheFile = files.next();

  if (!forceRefresh && cacheFile) {
    try {
      const cached = JSON.parse(cacheFile.getBlob().getDataAsString());
      cached.workspaceMeta = nn_getWorkspaceMeta_();
      return cached;
    } catch (e) {
      /* fall through */
    }
  }

  const props = PropertiesService.getScriptProperties();
  let roots = [];
  const mainId = nn_getMainFolderIdForCurrentUser_(props);
  if (mainId) roots.push(mainId);
  const registeredStr = props.getProperty('REGISTERED_FOLDERS');
  if (registeredStr) roots = roots.concat(JSON.parse(registeredStr));
  roots = Array.from(new Set(roots));

  const resultFolders = [];
  const resultFiles = [];
  const folderColors = JSON.parse(props.getProperty('FOLDER_COLORS') || '{}');

  function scan(folder, parentId) {
    const currentId = folder.getId();
    resultFolders.push({ id: currentId, name: folder.getName(), parentId: parentId });

    const fIter = folder.getFiles();
    while (fIter.hasNext()) {
      const f = fIter.next();
      const mime = f.getMimeType();
      if (mime === MimeType.PDF || mime === MimeType.JPEG || mime === MimeType.PNG) {
        resultFiles.push({
          id: f.getId(),
          name: f.getName(),
          folderId: currentId,
          updated: f.getLastUpdated().getTime(),
          created: f.getDateCreated().getTime(),
        });
      }
    }
    const dIter = folder.getFolders();
    while (dIter.hasNext()) scan(dIter.next(), currentId);
  }

  for (let r = 0; r < roots.length; r++) {
    try {
      scan(DriveApp.getFolderById(roots[r]), null);
    } catch (e) {
      /* skip inaccessible root */
    }
  }

  const result = {
    folders: resultFolders,
    files: resultFiles,
    folderColors: folderColors,
    lastUpdated: new Date().getTime(),
    workspaceMeta: nn_getWorkspaceMeta_(),
  };
  const jsonString = JSON.stringify(result);

  if (cacheFile) cacheFile.setContent(jsonString);
  else rootFolder.createFile(cacheFileName, jsonString, MimeType.PLAIN_TEXT);

  return result;
}

function updateMeetingState(stateJson) {
  try {
    CacheService.getScriptCache().put('shared_meeting_state', stateJson, 300);
    return true;
  } catch (e) {
    return false;
  }
}

function getMeetingState() {
  try {
    return CacheService.getScriptCache().get('shared_meeting_state');
  } catch (e) {
    return null;
  }
}

// =============================================================================
// NiceNotes (Phase 1-2: Schema, Initialization & API)
// -----------------------------------------------------------------------------
// Things 風タスク管理機能。既存の PDF ビューア機能と共存させるため、すべての
// 識別子に `nn_` / `NN_` プレフィックスを付ける。フロント (index.html) からは
// google.script.run.withSuccessHandler(cb).nn_xxx(args) で呼び出す前提。
// 仕様は docs/SCHEMA.md と docs/API_SPEC.md を単一ソースとする。
//
// スタンドアロン GAS のため排他は LockService.getScriptLock() を使用する。
// =============================================================================

/**
 * @typedef {('active'|'completed'|'canceled')} NN_TaskStatus
 */

/**
 * @typedef {Object} NN_CheckItem
 * @property {string} id
 * @property {string} text
 * @property {boolean} isDone
 * @property {string=} notes           任意。サブタスクの説明（Google Tasks の notes に同期）。
 * @property {?string=} dueDate         任意。YYYY-MM-DD（サブタスクの期限）。
 * @property {string=} googleTaskId     任意。サーバが Google Tasks 子タスク id を書き戻す。
 */

/**
 * @typedef {Object} NN_Task
 * @property {string} id                    UUIDv4。クライアント発行。
 * @property {string} area
 * @property {string} project
 * @property {string} heading
 * @property {string} title
 * @property {string} notes
 * @property {NN_TaskStatus} status
 * @property {?string} startDate            YYYY-MM-DD or null (= いつでも)
 * @property {?string} dueDate              YYYY-MM-DD or null
 * @property {NN_CheckItem[]} checkItems
 * @property {string} updatedAt             ISO 8601。サーバが最終的にスタンプ。
 * @property {number} sortOrder             heading 内の並び順 (新規 = max + 1024)
 * @property {?string} completedAt          ISO 8601 or null
 * @property {string} googleTaskId          N 列。サーバ管理。Google Tasks の task id。
 * @property {string} calendarEventId       O 列。サーバ管理。Calendar イベント id。
 * @property {string} repeatRule            P 列。繰り返し: 空/none/daily/weekly/monthly/yearly。
 */

/**
 * @typedef {{ type: 'upsert', task: NN_Task } | { type: 'delete', id: string }} NN_Op
 */

const NN_SHEET_NAME = 'tasks';
const NN_PROP_SHEET_ID = 'NICENOTES_SHEET_ID';
const NN_PROP_TASKLIST_ID = 'NICENOTES_TASKLIST_ID';
const NN_PROP_TASKLIST_MAP = 'NICENOTES_TASKLIST_MAP';
const NN_PROP_CALENDAR_ID = 'NICENOTES_CALENDAR_ID';
const NN_TASKLIST_TITLE = 'NiceNotes3';
const NN_TASKLIST_TITLE_PREFIX = 'NiceNotes3 · ';
const NN_TASKLIST_FALLBACK_NAME = 'NiceNotes3 · 未分類';
const NN_TASKLIST_TITLE_MAX = 200;
const NN_CALENDAR_NAME = 'NiceNotes3';
const NN_LOCK_TIMEOUT_MS = 10000;

/**
 * GAS エディタで引数なし nn_upsertTask() を実行したときのサンプルペイロード（毎回新しいオブジェクト）。
 * @return {Object}
 */
function nn_editorSampleTask_() {
  return {
    title: 'NiceNotes3 test',
    status: 'active',
    dueDate: '2026-05-20',
    startDate: '',
    area: '',
    project: '',
    heading: '',
    notes: '',
    checkItems: [],
    updatedAt: '',
    sortOrder: 1024,
    completedAt: null,
    googleTaskId: '',
    calendarEventId: '',
    repeatRule: ''
  };
}

/**
 * スプレッドシートの列順。docs/SCHEMA.md と必ず同期させること。
 * インデックス 0 (= A 列) を `id` 主キーとする。
 */
const NN_COLUMNS = [
  'id',              // A
  'area',            // B
  'project',         // C
  'heading',         // D
  'title',           // E
  'notes',           // F
  'status',          // G
  'startDate',       // H
  'dueDate',         // I
  'checkItems',      // J  JSON 文字列。空は "[]"
  'updatedAt',       // K  ISO 8601
  'sortOrder',       // L  数値
  'completedAt',     // M  ISO 8601 or 空
  'googleTaskId',    // N  Google Tasks API task id（サーバ管理）
  'calendarEventId', // O  Calendar イベント id（サーバ管理）
  'repeatRule'       // P  繰り返し（Google Tasks recurrence / アプリ内ルーティーン）
];

const NN_STATUS_VALUES = ['active', 'completed', 'canceled'];

/**
 * NiceNotes のマスタースプレッドシートを初期化・マイグレーションする。
 * 再実行可能: ヘッダーが NN_COLUMNS より短い場合は N/O など欠落列のみ追加する。
 *
 * @return {{ ok: true, spreadsheetId: string, url: string, sheetName: string }}
 */
function nn_initSpreadsheet() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty(NN_PROP_SHEET_ID);
  let ss;

  if (id) {
    ss = SpreadsheetApp.openById(id);
  } else {
    ss = SpreadsheetApp.create('NiceNotes3 Master');
    id = ss.getId();
    props.setProperty(NN_PROP_SHEET_ID, id);
  }

  let sheet = ss.getSheetByName(NN_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(NN_SHEET_NAME);
  }

  ss.getSheets().forEach(function (s) {
    if (s.getName() !== NN_SHEET_NAME) {
      try {
        ss.deleteSheet(s);
      } catch (e) {
        /* 最後の 1 枚は消せない */
      }
    }
  });

  const width = Math.max(sheet.getLastColumn(), NN_COLUMNS.length);
  const headerRow = sheet.getRange(1, 1, 1, width).getValues()[0];
  let i;
  for (i = 0; i < NN_COLUMNS.length; i++) {
    if (headerRow[i] !== NN_COLUMNS[i]) {
      sheet.getRange(1, i + 1).setValue(NN_COLUMNS[i]);
    }
  }

  const lastCol = NN_COLUMNS.length;
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, lastCol)
    .setFontWeight('bold')
    .setBackground('#f0f0f0');

  const lastDataRow = sheet.getLastRow();
  const dataRowCount = Math.max(lastDataRow > 1 ? lastDataRow - 1 : 0, 1);

  const statusCol = NN_COLUMNS.indexOf('status') + 1;
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(NN_STATUS_VALUES, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, statusCol, dataRowCount, 1).setDataValidation(statusRule);

  ['startDate', 'dueDate'].forEach(function (name) {
    const c = NN_COLUMNS.indexOf(name) + 1;
    sheet.getRange(2, c, dataRowCount, 1).setNumberFormat('yyyy-mm-dd');
  });

  ['updatedAt', 'completedAt', 'googleTaskId', 'calendarEventId', 'repeatRule'].forEach(function (name) {
    const c = NN_COLUMNS.indexOf(name) + 1;
    sheet.getRange(2, c, dataRowCount, 1).setNumberFormat('@');
  });

  const sortCol = NN_COLUMNS.indexOf('sortOrder') + 1;
  sheet.getRange(2, sortCol, dataRowCount, 1).setNumberFormat('0');

  sheet.autoResizeColumns(1, lastCol);

  const url = ss.getUrl();
  Logger.log('NiceNotes3 spreadsheet ready: ' + url);
  return { ok: true, spreadsheetId: id, url: url, sheetName: NN_SHEET_NAME };
}

// --- Lock & time -------------------------------------------------------------

function nn_lock_(fn) {
  const lock = LockService.getScriptLock();
  const ok = lock.tryLock(NN_LOCK_TIMEOUT_MS);
  if (!ok) {
    throw new Error('NN_E_LOCK_TIMEOUT: could not acquire script lock');
  }
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function nn_now_() {
  return new Date().toISOString();
}

// --- Spreadsheet I/O ---------------------------------------------------------

function nn_openSheet_() {
  const id = PropertiesService.getScriptProperties().getProperty(NN_PROP_SHEET_ID);
  if (!id) {
    throw new Error('NN_E_NO_SHEET: run nn_initSpreadsheet first');
  }
  const ss = SpreadsheetApp.openById(id);
  const sheet = ss.getSheetByName(NN_SHEET_NAME);
  if (!sheet) {
    throw new Error('NN_E_NO_TASK_SHEET: sheet "tasks" missing');
  }
  return sheet;
}

function nn_findRow_(sheet, taskId) {
  const finder = sheet.getRange('A:A').createTextFinder(String(taskId))
    .matchEntireCell(true);
  const cell = finder.findNext();
  if (!cell) {
    return -1;
  }
  return cell.getRow();
}

function nn_cellStr_(v) {
  if (v === null || v === undefined) {
    return '';
  }
  return String(v).trim();
}

function nn_cellDateOrNull_(v) {
  if (v === null || v === '') {
    return null;
  }
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime())) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  const s = nn_cellStr_(v);
  return s === '' ? null : s;
}

function nn_rowToTask_(row) {
  let checkItems = [];
  const raw = nn_cellStr_(row[NN_COLUMNS.indexOf('checkItems')]);
  if (raw) {
    try {
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed)) {
        checkItems = parsed;
      }
    } catch (e) {
      Logger.log('nn_rowToTask_: invalid checkItems JSON');
    }
  }
  const so = row[NN_COLUMNS.indexOf('sortOrder')];
  let sortOrder = typeof so === 'number' && !isNaN(so) ? so : parseFloat(so);
  if (isNaN(sortOrder)) {
    sortOrder = 1024;
  }
  const completedRaw = nn_cellStr_(row[NN_COLUMNS.indexOf('completedAt')]);
  return {
    id: nn_cellStr_(row[0]),
    area: nn_cellStr_(row[NN_COLUMNS.indexOf('area')]),
    project: nn_cellStr_(row[NN_COLUMNS.indexOf('project')]),
    heading: nn_cellStr_(row[NN_COLUMNS.indexOf('heading')]),
    title: nn_cellStr_(row[NN_COLUMNS.indexOf('title')]),
    notes: nn_cellStr_(row[NN_COLUMNS.indexOf('notes')]),
    status: /** @type {NN_TaskStatus} */ (nn_cellStr_(row[NN_COLUMNS.indexOf('status')]) || 'active'),
    startDate: nn_cellDateOrNull_(row[NN_COLUMNS.indexOf('startDate')]),
    dueDate: nn_cellDateOrNull_(row[NN_COLUMNS.indexOf('dueDate')]),
    checkItems: checkItems,
    updatedAt: nn_cellStr_(row[NN_COLUMNS.indexOf('updatedAt')]),
    sortOrder: sortOrder,
    completedAt: completedRaw === '' ? null : completedRaw,
    googleTaskId: nn_cellStr_(row[NN_COLUMNS.indexOf('googleTaskId')]),
    calendarEventId: nn_cellStr_(row[NN_COLUMNS.indexOf('calendarEventId')]),
    repeatRule: nn_cellStr_(row[NN_COLUMNS.indexOf('repeatRule')])
  };
}

function nn_taskToRow_(task) {
  let checkJson = '[]';
  if (task.checkItems && task.checkItems.length) {
    try {
      checkJson = JSON.stringify(task.checkItems);
    } catch (e) {
      checkJson = '[]';
    }
  }
  return [
    task.id || '',
    task.area || '',
    task.project || '',
    task.heading || '',
    task.title || '',
    task.notes || '',
    task.status || 'active',
    task.startDate || '',
    task.dueDate || '',
    checkJson,
    task.updatedAt || '',
    task.sortOrder != null ? task.sortOrder : 1024,
    task.completedAt || '',
    task.googleTaskId || '',
    task.calendarEventId || '',
    task.repeatRule || ''
  ];
}

function nn_nextSortOrder_(sheet, heading) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return 1024;
  }
  const h = heading || '';
  const idxHeading = NN_COLUMNS.indexOf('heading');
  const idxSort = NN_COLUMNS.indexOf('sortOrder');
  const data = sheet.getRange(2, 1, lastRow - 1, NN_COLUMNS.length).getValues();
  let max = 0;
  let r;
  for (r = 0; r < data.length; r++) {
    if (nn_cellStr_(data[r][idxHeading]) === h) {
      const so = data[r][idxSort];
      const n = typeof so === 'number' ? so : parseFloat(so);
      if (!isNaN(n) && n > max) {
        max = n;
      }
    }
  }
  return max + 1024;
}

// --- Google Tasks / Calendar -------------------------------------------------

function nn_normalizeProjectKey_(project) {
  if (project == null || project === '') {
    return '';
  }
  return String(project).replace(/\s+/g, ' ').trim();
}

function nn_taskListTitleForProject_(project) {
  const key = nn_normalizeProjectKey_(project);
  const base = key === '' ? NN_TASKLIST_FALLBACK_NAME : NN_TASKLIST_TITLE_PREFIX + key;
  if (base.length <= NN_TASKLIST_TITLE_MAX) {
    return base;
  }
  return base.substring(0, NN_TASKLIST_TITLE_MAX - 1) + '…';
}

function nn_loadTaskListMap_(props) {
  const raw = props.getProperty(NN_PROP_TASKLIST_MAP);
  if (!raw) {
    return {};
  }
  try {
    const o = JSON.parse(raw);
    return typeof o === 'object' && o !== null ? o : {};
  } catch (e) {
    return {};
  }
}

function nn_saveTaskListMap_(props, map) {
  props.setProperty(NN_PROP_TASKLIST_MAP, JSON.stringify(map));
}

/**
 * プロジェクト（空は未分類）ごとの Google Task リスト id を返す。
 * @param {*} project
 * @return {string}
 */
function nn_getOrCreateTaskListForProject_(project) {
  const props = PropertiesService.getScriptProperties();
  const key = nn_normalizeProjectKey_(project);
  let map = nn_loadTaskListMap_(props);
  let id = map[key];
  if (id) {
    try {
      Tasks.Tasklists.get(id);
      return id;
    } catch (e) {
      Logger.log('nn_getOrCreateTaskListForProject_: stale map id for "' + key + '"');
      delete map[key];
      nn_saveTaskListMap_(props, map);
    }
  }
  if (key === '') {
    const legacy = props.getProperty(NN_PROP_TASKLIST_ID);
    if (legacy) {
      try {
        Tasks.Tasklists.get(legacy);
        map[''] = legacy;
        nn_saveTaskListMap_(props, map);
        return legacy;
      } catch (e) {
        Logger.log('nn_getOrCreateTaskListForProject_: legacy NICENOTES_TASKLIST_ID invalid');
      }
    }
  }
  const title = nn_taskListTitleForProject_(project);
  const created = Tasks.Tasklists.insert({ title: title });
  map[key] = created.id;
  nn_saveTaskListMap_(props, map);
  if (key === '') {
    props.setProperty(NN_PROP_TASKLIST_ID, created.id);
  }
  return created.id;
}

function nn_allKnownTaskListIds_() {
  const props = PropertiesService.getScriptProperties();
  const map = nn_loadTaskListMap_(props);
  const ids = {};
  Object.keys(map).forEach(function (k) {
    if (map[k]) {
      ids[map[k]] = true;
    }
  });
  const legacy = props.getProperty(NN_PROP_TASKLIST_ID);
  if (legacy) {
    ids[legacy] = true;
  }
  return Object.keys(ids);
}

function nn_removeGoogleTaskCascadeFromList_(listId, taskId) {
  if (!listId || !taskId) {
    return;
  }
  try {
    const resp = Tasks.Tasks.list(listId, { parent: taskId });
    const items = resp.items || [];
    let i;
    for (i = 0; i < items.length; i++) {
      try {
        Tasks.Tasks.remove(listId, items[i].id);
      } catch (e) {
        Logger.log('nn_removeGoogleTaskCascadeFromList_ child: ' + e);
      }
    }
  } catch (e) {
    Logger.log('nn_removeGoogleTaskCascadeFromList_ list children: ' + e);
  }
  try {
    Tasks.Tasks.remove(listId, taskId);
  } catch (e) {
    Logger.log('nn_removeGoogleTaskCascadeFromList_ parent: ' + e);
  }
}

function nn_tryRemoveGoogleTaskEverywhereCascade_(taskId) {
  if (!taskId) {
    return;
  }
  const lists = nn_allKnownTaskListIds_();
  let i;
  for (i = 0; i < lists.length; i++) {
    try {
      nn_removeGoogleTaskCascadeFromList_(lists[i], taskId);
    } catch (e) {
      /* wrong list or already gone */
    }
  }
}

function nn_tryRemoveSubtaskEverywhere_(subTaskId) {
  if (!subTaskId) {
    return;
  }
  const lists = nn_allKnownTaskListIds_();
  let i;
  for (i = 0; i < lists.length; i++) {
    try {
      Tasks.Tasks.remove(lists[i], subTaskId);
    } catch (e) {
      /* not in this list */
    }
  }
}

function nn_buildGoogleSubtaskBody_(item, task) {
  const body = {
    title: String((item && item.text) || '').trim()
  };
  const noteStr = item && item.notes != null ? nn_cellStr_(String(item.notes)) : '';
  body.notes = noteStr;
  const subDue = item && item.dueDate != null ? nn_cellStr_(String(item.dueDate)) : '';
  if (/^\d{4}-\d{2}-\d{2}$/.test(subDue)) {
    body.due = subDue + 'T00:00:00.000Z';
  } else {
    body.due = null;
  }
  if (task.status === 'completed' || (item && item.isDone)) {
    body.status = 'completed';
    body.completed = task.completedAt || nn_now_();
  } else {
    body.status = 'needsAction';
    body.completed = null;
  }
  return body;
}

/**
 * checkItems を親タスク配下のサブタスクとして同期し、各要素に googleTaskId を付与する。
 * @param {Object} task
 * @param {string} listId
 * @param {string} parentId
 */
function nn_syncCheckItemsToGoogleSubtasks_(task, listId, parentId) {
  if (!parentId || !listId) {
    return;
  }
  const rawItems = task.checkItems || [];
  const items = rawItems.filter(function (ci) {
    return ci && String(ci.text || '').trim();
  });
  const desired = {};
  let i;
  for (i = 0; i < items.length; i++) {
    if (items[i].googleTaskId) {
      desired[items[i].googleTaskId] = true;
    }
  }
  try {
    const resp = Tasks.Tasks.list(listId, { parent: parentId });
    const existing = resp.items || [];
    for (i = 0; i < existing.length; i++) {
      if (!desired[existing[i].id]) {
        try {
          Tasks.Tasks.remove(listId, existing[i].id);
        } catch (e) {
          Logger.log('nn_syncCheckItemsToGoogleSubtasks_ remove orphan: ' + e);
        }
      }
    }
  } catch (e) {
    Logger.log('nn_syncCheckItemsToGoogleSubtasks_ list: ' + e);
  }

  let previous = null;
  for (i = 0; i < items.length; i++) {
    const ci = items[i];
    const resource = nn_buildGoogleSubtaskBody_(ci, task);
    if (ci.googleTaskId) {
      try {
        const patched = Tasks.Tasks.patch(resource, listId, ci.googleTaskId);
        ci.googleTaskId = patched.id || ci.googleTaskId;
        previous = ci.googleTaskId;
      } catch (e) {
        Logger.log('nn_syncCheckItemsToGoogleSubtasks_ patch sub, re-insert: ' + e);
        nn_tryRemoveSubtaskEverywhere_(ci.googleTaskId);
        ci.googleTaskId = '';
        const optIns = { parent: parentId };
        if (previous) {
          optIns.previous = previous;
        }
        const ins = Tasks.Tasks.insert(resource, listId, optIns);
        ci.googleTaskId = ins.id;
        previous = ins.id;
      }
    } else {
      const opt = { parent: parentId };
      if (previous) {
        opt.previous = previous;
      }
      const ins = Tasks.Tasks.insert(resource, listId, opt);
      ci.googleTaskId = ins.id;
      previous = ins.id;
    }
  }
}

function nn_getOrCreateCalendar_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty(NN_PROP_CALENDAR_ID);
  if (id) {
    const existing = CalendarApp.getCalendarById(id);
    if (existing) {
      return existing;
    }
  }
  const cal = CalendarApp.createCalendar(NN_CALENDAR_NAME);
  props.setProperty(NN_PROP_CALENDAR_ID, cal.getId());
  return cal;
}

/**
 * @param {Object} task
 * @return {string[]|null}
 */
/**
 * @param {string} s
 * @return {Object|null}
 */
function nn_tryParseRepeatJson_(s) {
  const t = nn_cellStr_(s);
  if (!t || t.charAt(0) !== '{') {
    return null;
  }
  try {
    const o = JSON.parse(t);
    if (o && o.v === 1 && o.f) {
      return o;
    }
  } catch (e) {
    /* ignore */
  }
  return null;
}

/**
 * @param {number} jsWd 0=Sun..6=Sat
 * @return {string}
 */
function nn_weekdayJsToByday_(jsWd) {
  const map = ['SU', 'MO', 'TU', 'WE', 'TH', 'FR', 'SA'];
  const i = parseInt(jsWd, 10);
  return map[i] || 'MO';
}

/**
 * @param {Object} o
 * @return {string[]|null}
 */
function nn_jsonRepeatToRruleLines_(o) {
  const f = String(o.f || '').toLowerCase();
  if (f === 'daily') {
    return ['RRULE:FREQ=DAILY'];
  }
  if (f === 'weekly') {
    const wds = Array.isArray(o.wd) ? o.wd : [];
    const parts = [];
    let i;
    for (i = 0; i < wds.length; i++) {
      const x = parseInt(wds[i], 10);
      if (x >= 0 && x <= 6) {
        parts.push(nn_weekdayJsToByday_(x));
      }
    }
    if (parts.length) {
      return ['RRULE:FREQ=WEEKLY;BYDAY=' + parts.join(',')];
    }
    return ['RRULE:FREQ=WEEKLY'];
  }
  if (f === 'monthly') {
    const mm = String(o.mm || 'dom');
    if (mm === 'nth' && o.nth != null && o.wd != null) {
      const n = parseInt(o.nth, 10);
      const wd = parseInt(o.wd, 10);
      if (n >= 1 && n <= 5 && wd >= 0 && wd <= 6) {
        return ['RRULE:FREQ=MONTHLY;BYDAY=' + n + nn_weekdayJsToByday_(wd)];
      }
    }
    const dom = parseInt(o.dom, 10);
    if (dom >= 1 && dom <= 31) {
      return ['RRULE:FREQ=MONTHLY;BYMONTHDAY=' + dom];
    }
    return ['RRULE:FREQ=MONTHLY'];
  }
  if (f === 'yearly') {
    return ['RRULE:FREQ=YEARLY'];
  }
  return null;
}

function nn_repeatRuleToRecurrence_(task) {
  const r = nn_cellStr_(task.repeatRule || '');
  if (!r || r === 'none') {
    return null;
  }
  const jo = nn_tryParseRepeatJson_(r);
  if (jo) {
    const lines = nn_jsonRepeatToRruleLines_(jo);
    return lines;
  }
  let rr;
  switch (r) {
    case 'daily':
      rr = 'RRULE:FREQ=DAILY';
      break;
    case 'weekly':
      rr = 'RRULE:FREQ=WEEKLY';
      break;
    case 'monthly':
      rr = 'RRULE:FREQ=MONTHLY';
      break;
    case 'yearly':
      rr = 'RRULE:FREQ=YEARLY';
      break;
    default:
      return null;
  }
  return [rr];
}

function nn_prefixNotes_(task) {
  const parts = [];
  if (task.area) {
    parts.push(task.area);
  }
  if (task.project) {
    parts.push(task.project);
  }
  if (task.heading) {
    parts.push(task.heading);
  }
  const pre = parts.length ? '[' + parts.join(' / ') + ']\n' : '';
  return pre + (task.notes || '');
}

/**
 * Google Tasks insert/patch 用リソース (active / completed 用)。canceled は remove。
 */
function nn_buildGoogleTasksBody_(task) {
  const body = {
    title: task.title || '',
    notes: nn_prefixNotes_(task)
  };
  if (task.status === 'completed') {
    body.status = 'completed';
    body.completed = task.completedAt || nn_now_();
    if (task.dueDate) {
      body.due = task.dueDate + 'T00:00:00.000Z';
    }
  } else {
    body.status = 'needsAction';
    body.completed = null;
    if (task.dueDate) {
      body.due = task.dueDate + 'T00:00:00.000Z';
    } else {
      body.due = null;
    }
  }
  const rec = nn_repeatRuleToRecurrence_(task);
  if (rec) {
    body.recurrence = rec;
  } else if (task.googleTaskId) {
    body.recurrence = [];
  }
  return body;
}

/**
 * @return {string} googleTaskId or ""
 */
function nn_syncToGoogleTasks_(task) {
  const listId = nn_getOrCreateTaskListForProject_(task.project);

  if (task.status === 'canceled') {
    if (task.googleTaskId) {
      nn_tryRemoveGoogleTaskEverywhereCascade_(task.googleTaskId);
    }
    const cis = task.checkItems || [];
    let j;
    for (j = 0; j < cis.length; j++) {
      if (cis[j] && cis[j].googleTaskId) {
        nn_tryRemoveSubtaskEverywhere_(cis[j].googleTaskId);
      }
      if (cis[j]) {
        delete cis[j].googleTaskId;
      }
    }
    return '';
  }

  const resource = nn_buildGoogleTasksBody_(task);
  let parentId = '';

  if (task.googleTaskId) {
    try {
      const patched = Tasks.Tasks.patch(resource, listId, task.googleTaskId);
      parentId = patched.id || task.googleTaskId;
    } catch (e) {
      Logger.log('nn_syncToGoogleTasks_ patch parent failed, re-insert: ' + e);
      nn_tryRemoveGoogleTaskEverywhereCascade_(task.googleTaskId);
      const ins = Tasks.Tasks.insert(resource, listId);
      parentId = ins.id;
    }
  } else {
    const inserted = Tasks.Tasks.insert(resource, listId);
    parentId = inserted.id;
  }

  nn_syncCheckItemsToGoogleSubtasks_(task, listId, parentId);
  return parentId;
}

/**
 * Parses ymd as strict calendar date YYYY-MM-DD only. Invalid dates return null (no fallback to today).
 * @param {*} ymd
 * @return {Date|null}
 */
function nn_parseYmdAsLocalDate_(ymd) {
  if (ymd == null || ymd === '') {
    return null;
  }
  const s = String(ymd).trim();
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s);
  if (!m) {
    Logger.log('nn_parseYmdAsLocalDate_: invalid format (expected YYYY-MM-DD): ' + s);
    return null;
  }
  const y = parseInt(m[1], 10);
  const mo = parseInt(m[2], 10);
  const d = parseInt(m[3], 10);
  if (isNaN(y) || isNaN(mo) || isNaN(d) || mo < 1 || mo > 12 || d < 1 || d > 31) {
    Logger.log('nn_parseYmdAsLocalDate_: out of range: ' + s);
    return null;
  }
  const dt = new Date(y, mo - 1, d);
  if (dt.getFullYear() !== y || dt.getMonth() !== mo - 1 || dt.getDate() !== d) {
    Logger.log('nn_parseYmdAsLocalDate_: not a valid calendar date: ' + s);
    return null;
  }
  return dt;
}

function nn_calendarDescription_(task) {
  return nn_prefixNotes_(task);
}

/**
 * @return {string} calendarEventId or ""
 */
function nn_syncToCalendar_(task) {
  const cal = nn_getOrCreateCalendar_();

  if (task.status !== 'active' || !task.dueDate) {
    if (task.calendarEventId) {
      try {
        const evOld = cal.getEventById(task.calendarEventId);
        if (evOld) {
          evOld.deleteEvent();
        }
      } catch (e) {
        Logger.log('nn_syncToCalendar_ delete: ' + e);
      }
    }
    return '';
  }

  const start = nn_parseYmdAsLocalDate_(task.dueDate);
  if (!start) {
    Logger.log('nn_syncToCalendar_: invalid dueDate, skipping calendar upsert: ' + String(task.dueDate));
    if (task.calendarEventId) {
      try {
        const evBad = cal.getEventById(task.calendarEventId);
        if (evBad) {
          evBad.deleteEvent();
        }
      } catch (e) {
        Logger.log('nn_syncToCalendar_ delete (invalid due): ' + e);
      }
    }
    return '';
  }

  const endExclusive = new Date(start.getTime());
  endExclusive.setDate(endExclusive.getDate() + 1);
  const desc = nn_calendarDescription_(task);

  if (task.calendarEventId) {
    try {
      const ev = cal.getEventById(task.calendarEventId);
      if (ev) {
        ev.setTitle(task.title || '');
        ev.setDescription(desc);
        ev.setAllDayDates(start, endExclusive);
        return task.calendarEventId;
      }
    } catch (e) {
      Logger.log('nn_syncToCalendar_ update failed, recreate: ' + e);
    }
  }

  const created = cal.createAllDayEvent(task.title || '', start, endExclusive, {
    description: desc
  });
  return created.getId();
}

/**
 * シート行に対応する Google Tasks（親＋孤立した子）をベストエフォートで削除する。
 * @param {NN_Task} task
 */
function nn_deleteGoogleSideForTask_(task) {
  if (task.googleTaskId) {
    nn_tryRemoveGoogleTaskEverywhereCascade_(task.googleTaskId);
  }
  const cis = task.checkItems || [];
  let i;
  for (i = 0; i < cis.length; i++) {
    if (cis[i] && cis[i].googleTaskId) {
      nn_tryRemoveSubtaskEverywhere_(cis[i].googleTaskId);
    }
  }
}

function nn_deleteCalendarEvent_(calendarEventId) {
  if (!calendarEventId) {
    return;
  }
  try {
    const cal = nn_getOrCreateCalendar_();
    const ev = cal.getEventById(calendarEventId);
    if (ev) {
      ev.deleteEvent();
    }
  } catch (e) {
    Logger.log('nn_deleteCalendarEvent_: ' + e);
  }
}

// --- Public API --------------------------------------------------------------

/**
 * upsert 用に task を検証・正規化する。
 * - id が空・未定義ならサーバで UUID を採番（クライアント未送信 / エディタ検証向け）。
 * - id が数値などでも文字列に正規化する。
 * @param {*} task
 * @return {NN_Task}
 */
function nn_validateAndNormalizeTaskForUpsert_(task) {
  if (!task || typeof task !== 'object') {
    throw new Error('NN_E_BAD_TASK: task object is required（引数なしで試す場合は nn_upsertTask() を実行するか、nn_debugUpsertSample() を使ってください）');
  }
  let id = task.id;
  if (id != null && id !== '') {
    id = String(id).trim();
  }
  if (!id) {
    id = Utilities.getUuid();
  }
  task.id = id;
  if (!task.title || !String(task.title).trim()) {
    throw new Error('NN_E_TITLE_REQUIRED: task.title must not be empty');
  }
  return /** @type {NN_Task} */ (task);
}

/**
 * エディタから 1 クリックで upsert を試す（nn_upsertTask() 引数なしと同じ）。
 * @return {NN_Task}
 */
function nn_debugUpsertSample() {
  return nn_upsertTask();
}

function nn_getAllTasks() {
  const sheet = nn_openSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }
  const values = sheet.getRange(2, 1, lastRow - 1, NN_COLUMNS.length).getValues();
  return values
    .filter(function (r) {
      return nn_cellStr_(r[0]);
    })
    .map(nn_rowToTask_);
}

function nn_upsertTaskUnlocked_(task) {
  const sheet = nn_openSheet_();
  const row = nn_findRow_(sheet, task.id);
  const now = nn_now_();

  if (row !== -1) {
    const existingVals = sheet.getRange(row, 1, 1, NN_COLUMNS.length).getValues()[0];
    const existing = nn_rowToTask_(existingVals);
    task.googleTaskId = existing.googleTaskId || '';
    task.calendarEventId = existing.calendarEventId || '';
    const exMap = {};
    let ix;
    const exItems = existing.checkItems || [];
    for (ix = 0; ix < exItems.length; ix++) {
      const ex = exItems[ix];
      if (ex && ex.id) {
        exMap[String(ex.id)] = ex;
      }
    }
    const tItems = task.checkItems || [];
    for (ix = 0; ix < tItems.length; ix++) {
      const ci = tItems[ix];
      if (!ci || !ci.id) {
        continue;
      }
      const ex = exMap[String(ci.id)];
      if (!ex) {
        continue;
      }
      if (!ci.googleTaskId && ex.googleTaskId) {
        ci.googleTaskId = ex.googleTaskId;
      }
      const cn = ci.notes != null ? nn_cellStr_(String(ci.notes)) : '';
      const en = ex.notes != null ? nn_cellStr_(String(ex.notes)) : '';
      if (!cn && en) {
        ci.notes = en;
      }
      const cd = ci.dueDate != null ? nn_cellStr_(String(ci.dueDate)) : '';
      const ed = ex.dueDate != null ? nn_cellStr_(String(ex.dueDate)) : '';
      if (!cd && ed) {
        ci.dueDate = ed;
      }
    }
  } else {
    task.googleTaskId = '';
    task.calendarEventId = '';
  }

  task.updatedAt = now;
  if (task.status !== 'active') {
    if (!task.completedAt) {
      task.completedAt = now;
    }
  } else {
    task.completedAt = null;
  }

  if (row === -1) {
    const so = task.sortOrder;
    if (so == null || so === '' || so === 0 || (typeof so === 'number' && isNaN(so))) {
      task.sortOrder = nn_nextSortOrder_(sheet, task.heading);
    }
  }

  const rowVals = nn_taskToRow_(task);

  if (row === -1) {
    sheet.appendRow(rowVals);
  } else {
    sheet.getRange(row, 1, 1, NN_COLUMNS.length).setValues([rowVals]);
  }

  const finalRow = row === -1 ? sheet.getLastRow() : row;

  let gid = '';
  let cid = '';
  try {
    gid = nn_syncToGoogleTasks_(task) || '';
  } catch (e) {
    Logger.log('nn_upsertTask Tasks sync: ' + e);
  }
  try {
    cid = nn_syncToCalendar_(task) || '';
  } catch (e) {
    Logger.log('nn_upsertTask Calendar sync: ' + e);
  }

  task.googleTaskId = gid;
  task.calendarEventId = cid;

  const idxN = NN_COLUMNS.indexOf('googleTaskId') + 1;
  const idxJ = NN_COLUMNS.indexOf('checkItems') + 1;
  const rowSynced = nn_taskToRow_(task);
  sheet.getRange(finalRow, idxN, 1, 2).setValues([[gid, cid]]);
  sheet.getRange(finalRow, idxJ).setValue(rowSynced[NN_COLUMNS.indexOf('checkItems')]);

  return task;
}

function nn_deleteTaskUnlocked_(id) {
  const sheet = nn_openSheet_();
  const row = nn_findRow_(sheet, id);
  if (row === -1) {
    return { id: id, deleted: true };
  }
  const vals = sheet.getRange(row, 1, 1, NN_COLUMNS.length).getValues()[0];
  const task = nn_rowToTask_(vals);

  nn_deleteGoogleSideForTask_(task);
  nn_deleteCalendarEvent_(task.calendarEventId);

  sheet.deleteRow(row);
  return { id: id, deleted: true };
}

function nn_upsertTask(task) {
  if (arguments.length === 0) {
    task = nn_editorSampleTask_();
    Logger.log('nn_upsertTask: 引数なしのためエディタ用サンプルを使用しました（本番で引数省略は非推奨）');
  }
  nn_validateAndNormalizeTaskForUpsert_(task);
  return nn_lock_(function () {
    return nn_upsertTaskUnlocked_(task);
  });
}

function nn_deleteTask(id) {
  if (id == null || id === '') {
    throw new Error('NN_E_BAD_ID: id is required');
  }
  const sid = String(id).trim();
  if (!sid) {
    throw new Error('NN_E_BAD_ID: id is required');
  }
  return nn_lock_(function () {
    return nn_deleteTaskUnlocked_(sid);
  });
}

function nn_batchSync(ops) {
  if (!ops || !ops.length) {
    return { results: [] };
  }
  return nn_lock_(function () {
    const results = [];
    let i;
    for (i = 0; i < ops.length; i++) {
      const op = ops[i];
      try {
        if (op.type === 'delete') {
          nn_deleteTaskUnlocked_(op.id);
          results.push({ id: op.id, status: 'ok' });
        } else if (op.type === 'upsert' && op.task) {
          nn_validateAndNormalizeTaskForUpsert_(op.task);
          const t = nn_upsertTaskUnlocked_(op.task);
          results.push({ id: t.id, status: 'ok', task: t });
        } else {
          results.push({ id: '', status: 'conflict', error: 'NN_E_BAD_OP: unknown op' });
        }
      } catch (e) {
        const msg = e && e.message ? e.message : String(e);
        const rid = op && op.type === 'delete' ? op.id : (op.task && op.task.id);
        results.push({ id: rid || '', status: 'conflict', error: msg });
      }
    }
    return { results: results };
  });
}
