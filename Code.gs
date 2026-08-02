// =============================================
// まもるさん — Google Apps Script API（アクセスキー制・ユーザー分離版）
// =============================================
// 【この方式の要点】
//  ・1人1つの「アクセスキー」で認証する。キーが無い/違うリクエストは全拒否。
//  ・データはユーザーごとの別タブ（例：タスク__u001）に保存し、ユーザー同士は見えない。
//  ・スプレッドシートの所有者（＝あなた）は全員分を見られる（受け入れる妥協点）。
//
// 【最初にやること】（Apps Scriptエディタ上で1回だけ）
//  1) 関数 addUser を選び、実行 → 実行ログに発行キーが出る（下の addUser 内で名前を書き換えて実行）
//  2) そのキーを控える（自分用）
//  3) 既存データがある場合は migrateLegacyData を実行（下部の説明参照）
//  4) 「デプロイ」→「デプロイを管理」→ 新バージョンとして再デプロイ
// =============================================

const SS = SpreadsheetApp.getActiveSpreadsheet();

function getSheet(name, headers) {
  let sheet = SS.getSheetByName(name);
  if (!sheet) {
    sheet = SS.insertSheet(name);
    if (headers) sheet.appendRow(headers);
  }
  return sheet;
}

// ユーザー別タブを取得（例：base='タスク', userId='u001' → 'タスク__u001'）
function userSheet(base, headers, userId) {
  return getSheet(base + '__' + userId, headers);
}

function fmtDate(v) {
  if (!v) return '';
  const d = new Date(v);
  if (isNaN(d.getTime())) return String(v);
  return `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}`;
}
function pad(n) { return String(n).padStart(2,'0'); }

function response(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// =============================================
// 認証：アクセスキー → ユーザーID
// ユーザー タブ：キー | ユーザーID | 名前 | 登録日
// =============================================
const USER_HEADERS = ['キー', 'ユーザーID', '名前', '登録日'];

function resolveUser(key) {
  if (!key) return null;
  const sheet = getSheet('ユーザー', USER_HEADERS);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(key) && rows[i][0] !== '') {
      return { userId: String(rows[i][1]), name: String(rows[i][2]) };
    }
  }
  return null;
}

// 新規ユーザーを発行（エディタから手動実行）。名前を書き換えて実行するとキーが発行される。
function addUser(name) {
  const displayName = name || '（名前未設定）';
  const sheet = getSheet('ユーザー', USER_HEADERS);
  // ユーザーID採番：u001, u002, ...
  const rows = sheet.getDataRange().getValues();
  const n = rows.length; // ヘッダ含む行数 → 次の番号
  const userId = 'u' + String(n).padStart(3, '0');
  const key = 'k_' + Utilities.getUuid().replace(/-/g, '');
  sheet.appendRow([key, userId, displayName, fmtDate(new Date())]);
  Logger.log('発行しました → 名前:%s / ユーザーID:%s / キー:%s', displayName, userId, key);
  return { userId, key, name: displayName };
}

// =============================================
// GET / POST（入口で必ず認証）
// =============================================
function doGet(e) {
  try {
    const action = e.parameter.action;
    // キー疎通確認（セットアップ画面用。データには触れない）
    const u = resolveUser(e.parameter.key);
    if (action === 'checkKey') {
      return response(u ? { ok: true, name: u.name, userId: u.userId } : { error: 'unauthorized' });
    }
    if (!u) return response({ error: 'unauthorized' });
    const uid = u.userId;
    let data;
    if      (action === 'getTasks')    data = getTasks(uid);
    else if (action === 'getDaily')    data = getDaily(uid);
    else if (action === 'getSettings') data = getSettings(uid);
    else data = { error: 'unknown action' };
    return response(data);
  } catch (err) {
    return response({ error: err.message });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const u = resolveUser(body.key);
    if (!u) return response({ error: 'unauthorized' });
    const uid = u.userId;
    const action = body.action;
    let result;
    if      (action === 'saveTask')         result = saveTask(uid, body.data);
    else if (action === 'updateTask')       result = updateTask(uid, body.data);
    else if (action === 'deleteTask')       result = deleteTask(uid, body.data);
    else if (action === 'saveDaily')        result = saveDaily(uid, body.data);
    else if (action === 'resetAll')         result = resetAll(uid);
    else if (action === 'deduplicate')      result = deduplicateTasks(uid);
    else if (action === 'saveTaskOrder')    result = saveTaskOrder(uid, body.data);
    else if (action === 'saveSettings')     result = saveSettings(uid, body.data);
    else if (action === 'deleteSampleData') result = deleteSampleData(uid);
    else if (action === 'insertSampleData') result = insertSampleData(uid);
    else if (action === 'replaceAllTasks')  result = replaceAllTasks(uid, body.data);
    else if (action === 'replaceAllDaily')  result = replaceAllDaily(uid, body.data);
    else if (action === 'archiveTasks')     result = archiveTasks(uid);
    else if (action === 'archiveDaily')     result = archiveDaily(uid);
    else if (action === 'saveVoice')        result = saveVoice(uid, body.data);
    else if (action === 'saveUserName')     result = { ok: true }; // 名前はユーザータブ管理。互換のためok返す
    else result = { error: 'unknown action' };
    return response(result);
  } catch (err) {
    return response({ error: err.message });
  }
}

// =============================================
// タスク
// 列：タスクID | カテゴリ | タスク名 | 締切日 | ピン固定 | 完了 | 完了日 | かかった時間 | いつ実行したか | 作成日 | 突発対応 | 並び順 | インデント
// =============================================
const TASK_HEADERS = [
  'タスクID', 'カテゴリ', 'タスク名（詳細・やること）',
  '締切日', 'ピン固定', '完了', '完了日', 'かかった時間', 'いつ実行したか', '作成日', '突発対応', '並び順', 'インデント'
];

function getTasks(uid) {
  const sheet = userSheet('タスク', TASK_HEADERS, uid);
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return { tasks: [] };
  const tasks = rows.slice(1).map((r, i) => ({
    id:       r[0],
    name:     r[1],
    color:    '',
    subs:     r[2] ? r[2].split('\n').filter(Boolean) : [],
    deadline: fmtDate(r[3]),
    pinned:   r[4] === true || r[4] === 'TRUE',
    done:     r[5] === true || r[5] === 'TRUE',
    doneDate: fmtDate(r[6]),
    time:     r[7] || '',
    when:     r[8] || '',
    created:  r[9] || '',
    tokk:     r[10] === true || r[10] === 'TRUE' || false,
    order:    (r[11] !== undefined && r[11] !== '') ? Number(r[11]) : i,
    indent:   r[12] === true || r[12] === 'TRUE' || false,
  }));
  tasks.sort((a, b) => a.order - b.order);
  return { tasks };
}

function saveTask(uid, data) {
  const sheet = userSheet('タスク', TASK_HEADERS, uid);
  const id = 'T' + Date.now();
  const lastRow = sheet.getLastRow();
  const newOrder = lastRow;
  sheet.appendRow([
    id,
    data.name     || '',
    (data.subs    || []).join('\n'),
    data.deadline || '',
    false,
    data.done     || false,
    data.doneDate || '',
    data.time     || '',
    data.when     || '',
    fmtDate(new Date()),
    data.tokk     || false,
    newOrder,
    data.indent   || false,
  ]);
  return { ok: true, id };
}

function deduplicateTasks(uid) {
  const sheet = userSheet('タスク', TASK_HEADERS, uid);
  const rows = sheet.getDataRange().getValues();
  const seen = {};
  const toDelete = [];
  for (let i = 1; i < rows.length; i++) {
    const id = String(rows[i][0]);
    if (!id) continue;
    if (seen[id]) toDelete.push(i + 1);
    else seen[id] = true;
  }
  for (let i = toDelete.length - 1; i >= 0; i--) sheet.deleteRow(toDelete[i]);
  return { ok: true, deleted: toDelete.length };
}

function updateTask(uid, data) {
  const sheet = userSheet('タスク', TASK_HEADERS, uid);
  const rows = sheet.getDataRange().getValues();
  let firstFound = -1;
  const toDelete = [];
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.id)) {
      if (firstFound === -1) firstFound = i + 1;
      else toDelete.push(i + 1);
    }
  }
  if (firstFound === -1) return { error: 'task not found' };
  for (let i = toDelete.length - 1; i >= 0; i--) sheet.deleteRow(toDelete[i]);
  const r = firstFound;
  const cur = sheet.getRange(r, 1, 1, 13).getValues()[0];
  sheet.getRange(r, 1, 1, 13).setValues([[
    data.id,
    data.name     !== undefined ? data.name     : cur[1],
    data.subs     !== undefined ? (data.subs||[]).join('\n') : cur[2],
    data.deadline !== undefined ? data.deadline : cur[3],
    data.pinned   !== undefined ? data.pinned   : cur[4],
    data.done     !== undefined ? data.done     : cur[5],
    data.doneDate !== undefined ? data.doneDate : cur[6],
    data.time     !== undefined ? data.time     : cur[7],
    data.when     !== undefined ? data.when     : cur[8],
    cur[9],
    data.tokk     !== undefined ? data.tokk     : cur[10],
    data.order    !== undefined ? data.order    : (cur[11] !== '' ? cur[11] : 0),
    data.indent   !== undefined ? data.indent   : (cur[12] || false),
  ]]);
  return { ok: true };
}

function deleteTask(uid, data) {
  const sheet = userSheet('タスク', TASK_HEADERS, uid);
  const rows = sheet.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][0]) === String(data.id)) sheet.deleteRow(i + 1);
  }
  return { ok: true };
}

function saveTaskOrder(uid, data) {
  const sheet = userSheet('タスク', TASK_HEADERS, uid);
  const rows = sheet.getDataRange().getValues();
  const orderMap = {};
  (data.order || []).forEach(o => { orderMap[String(o.id)] = o.order; });
  for (let i = 1; i < rows.length; i++) {
    const id = String(rows[i][0]);
    if (orderMap[id] !== undefined) sheet.getRange(i + 1, 12).setValue(orderMap[id]);
  }
  return { ok: true };
}

// =============================================
// 日次記録
// =============================================
const DAILY_HEADERS = [
  '日付', 'しんどさ（体）', 'しんどさ（心）', '昼休み',
  '時間外在校等時間（分）', '持ち帰り業務（分）', '時間外合計（分）', 'メモ'
];

function getDaily(uid) {
  const sheet = userSheet('日次記録', DAILY_HEADERS, uid);
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return { daily: [] };
  const daily = rows.slice(1).map(r => ({
    date:     fmtDate(r[0]),
    moodBody: r[1] || '',
    moodMind: r[2] || '',
    mood:     r[1] || r[2] || '',
    lunch:    r[3] || '',
    schoolOt: r[4] || 0,
    takehome: r[5] || 0,
    ot:       r[6] || 0,
    memo:     r[7] || '',
  }));
  return { daily };
}

function saveDaily(uid, data) {
  const sheet = userSheet('日次記録', DAILY_HEADERS, uid);
  const rows = sheet.getDataRange().getValues();
  const targetDate = data.date || fmtDate(new Date());
  const schoolOt = data.schoolOt || 0;
  const takehome = data.takehome || 0;
  const totalOt  = data.ot || (schoolOt + takehome);
  for (let i = 1; i < rows.length; i++) {
    if (fmtDate(rows[i][0]) === targetDate) {
      const r = i + 1;
      sheet.getRange(r, 1, 1, 8).setValues([[
        targetDate, data.moodBody || data.mood || '', data.moodMind || '',
        data.lunch || '', schoolOt, takehome, totalOt, data.memo || '',
      ]]);
      return { ok: true, updated: true };
    }
  }
  sheet.appendRow([
    targetDate, data.moodBody || data.mood || '', data.moodMind || '',
    data.lunch || '', schoolOt, takehome, totalOt, data.memo || '',
  ]);
  return { ok: true, updated: false };
}

// =============================================
// リセット（そのユーザーの領域だけ）
// =============================================
function resetAll(uid) {
  ['タスク__' + uid, '日次記録__' + uid, '設定__' + uid].forEach(name => {
    const s = SS.getSheetByName(name);
    if (s) SS.deleteSheet(s);
  });
  return { ok: true };
}

// =============================================
// ユーザー設定（そのユーザーの端末間で共通）
// =============================================
const SETTINGS_HEADERS = ['キー', '値'];

function getSettings(uid) {
  const sheet = userSheet('設定', SETTINGS_HEADERS, uid);
  const rows = sheet.getDataRange().getValues();
  const settings = {};
  for (let i = 1; i < rows.length; i++) settings[rows[i][0]] = rows[i][1];
  return { settings };
}

function saveSettings(uid, data) {
  const sheet = userSheet('設定', SETTINGS_HEADERS, uid);
  const rows = sheet.getDataRange().getValues();
  Object.keys(data || {}).forEach(key => {
    let found = false;
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === key) { sheet.getRange(i + 1, 2).setValue(String(data[key])); found = true; break; }
    }
    if (!found) sheet.appendRow([key, String(data[key])]);
  });
  return { ok: true };
}

// =============================================
// サンプルデータ
// =============================================
function deleteSampleData(uid) {
  const taskSheet = userSheet('タスク', TASK_HEADERS, uid);
  const taskRows = taskSheet.getDataRange().getValues();
  let taskDeleted = 0;
  for (let i = taskRows.length - 1; i >= 1; i--) {
    if (String(taskRows[i][0]).startsWith('S')) { taskSheet.deleteRow(i + 1); taskDeleted++; }
  }
  const sampleDates = new Set(SAMPLE_ROWS.map(s => s[0]));
  const dailySheet = userSheet('日次記録', DAILY_HEADERS, uid);
  const dailyRows = dailySheet.getDataRange().getValues();
  let dailyDeleted = 0;
  for (let i = dailyRows.length - 1; i >= 1; i--) {
    if (sampleDates.has(fmtDate(dailyRows[i][0]))) { dailySheet.deleteRow(i + 1); dailyDeleted++; }
  }
  return { ok: true, taskDeleted, dailyDeleted };
}

function insertSampleData(uid) {
  const sheet = userSheet('日次記録', DAILY_HEADERS, uid);
  const rows = sheet.getDataRange().getValues();
  const existingDates = new Set(rows.slice(1).map(r => fmtDate(r[0])));
  let count = 0;
  SAMPLE_ROWS.forEach(s => {
    if (!existingDates.has(s[0])) { sheet.appendRow(s); count++; }
  });
  return { ok: true, count };
}

const SAMPLE_ROWS = [
  ['2026-02-02','ok','ok','ok',30,0,30,''],['2026-02-03','hard','hard','ng',90,0,135,''],
  ['2026-02-04','ok','ok','ok',0,60,60,''],['2026-02-05','hard','hard','ng',60,0,105,''],
  ['2026-02-06','ok','great','ok',0,0,0,''],['2026-02-09','hard','hard','20',60,0,90,''],
  ['2026-02-10','ok','ok','ok',30,0,30,''],['2026-02-12','hard','limit','ng',120,0,165,''],
  ['2026-02-13','ok','hard','10',30,60,105,''],['2026-02-16','great','great','ok',0,0,0,''],
  ['2026-02-17','hard','hard','ng',90,0,135,''],['2026-02-18','ok','ok','ok',30,0,30,''],
  ['2026-02-19','hard','hard','ng',60,120,225,''],['2026-02-20','ok','ok','10',0,0,15,''],
  ['2026-03-02','hard','hard','ng',90,0,135,''],['2026-03-03','ok','ok','ok',30,0,30,''],
  ['2026-03-05','hard','hard','10',30,0,45,''],['2026-03-06','ok','ok','ok',30,0,30,''],
  ['2026-03-07','hard','limit','ng',90,0,135,''],['2026-03-10','ok','ok','10',0,60,75,''],
  ['2026-03-11','hard','hard','ng',60,0,105,''],['2026-03-12','limit','limit','ng',120,0,165,''],
  ['2026-03-13','ok','hard','10',30,0,45,''],['2026-03-14','hard','hard','ng',120,120,285,''],
  ['2026-03-17','great','ok','ok',0,0,0,''],['2026-03-18','hard','hard','ng',120,0,165,''],
  ['2026-03-19','ok','ok','10',30,0,45,''],['2026-03-24','great','great','ok',0,0,0,''],
  ['2026-03-25','ok','ok','ok',30,0,30,''],['2026-03-26','hard','hard','ng',90,60,195,''],
  ['2026-04-03','ok','ok','10',60,0,75,''],['2026-04-06','ok','great','ok',30,0,30,''],
  ['2026-04-07','hard','hard','ng',120,0,165,''],['2026-04-08','hard','limit','ng',90,60,195,''],
  ['2026-04-09','ok','hard','10',30,0,45,''],['2026-04-10','ok','ok','ok',0,0,0,''],
  ['2026-04-13','hard','hard','ng',60,0,105,''],['2026-04-14','great','great','ok',0,0,0,''],
  ['2026-04-15','hard','limit','ng',120,0,165,''],['2026-04-16','ok','ok','10',30,0,45,''],
  ['2026-04-17','hard','hard','ng',90,0,135,''],['2026-04-20','ok','great','ok',0,60,60,''],
  ['2026-04-22','hard','hard','ng',60,0,105,''],['2026-04-23','ok','ok','10',0,0,15,''],
  ['2026-04-25','limit','limit','ng',120,60,225,''],['2026-04-28','hard','hard','ng',90,0,135,''],
  ['2026-04-30','ok','ok','20',30,0,60,''],['2026-05-07','ok','ok','ok',30,0,30,''],
  ['2026-05-08','hard','hard','ng',60,0,105,''],['2026-05-11','ok','ok','10',0,0,15,''],
  ['2026-05-13','great','ok','ok',30,0,30,''],['2026-05-14','ok','ok','10',30,0,45,''],
  ['2026-05-15','hard','hard','ng',90,60,195,''],['2026-05-18','ok','great','ok',0,0,0,''],
  ['2026-05-19','ok','ok','10',30,0,45,''],['2026-05-20','ok','ok','ok',0,60,60,''],
  ['2026-05-21','ok','ok','10',30,0,45,''],['2026-05-22','ok','hard','20',60,0,90,''],
  ['2026-05-25','hard','hard','ng',90,0,135,''],['2026-05-26','ok','ok','20',60,0,90,''],
  ['2026-05-27','hard','hard','ng',90,60,195,''],['2026-05-28','ok','ok','ok',60,0,60,''],
  ['2026-05-29','hard','hard','ng',120,0,165,''],['2026-06-01','ok','ok','ok',30,0,30,''],
  ['2026-06-02','ok','ok','ok',30,0,30,''],['2026-06-03','hard','hard','ng',60,0,105,''],
  ['2026-06-04','ok','ok','10',0,0,15,''],['2026-06-05','great','great','ok',0,0,0,''],
];

// =============================================
// Undo：一括置換（そのユーザーの領域だけ）
// =============================================
function replaceAllTasks(uid, data) {
  const sheet = userSheet('タスク', TASK_HEADERS, uid);
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);
  const list = data.tasks || [];
  if (list.length === 0) return { ok: true, count: 0 };
  const rows = list.map((t, i) => [
    t.id, t.name || '', (t.subs || []).join('\n'),
    t.deadline || '', t.pinned || false, t.done || false,
    t.doneDate || '', t.time || '', t.when || '',
    t.created || fmtDate(new Date()), t.tokk || false,
    t.order !== undefined ? t.order : i, t.indent || false,
  ]);
  sheet.getRange(2, 1, rows.length, TASK_HEADERS.length).setValues(rows);
  return { ok: true, count: rows.length };
}

function replaceAllDaily(uid, data) {
  const sheet = userSheet('日次記録', DAILY_HEADERS, uid);
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);
  const list = data.daily || [];
  if (list.length === 0) return { ok: true, count: 0 };
  const rows = list.map(d => [
    d.date, d.moodBody || d.mood || '', d.moodMind || '',
    d.lunch || '', d.schoolOt || 0, d.takehome || 0,
    d.ot || ((d.schoolOt||0) + (d.takehome||0)), d.memo || '',
  ]);
  sheet.getRange(2, 1, rows.length, DAILY_HEADERS.length).setValues(rows);
  return { ok: true, count: rows.length };
}

// =============================================
// アーカイブ：完了済みタスクを「タスク_アーカイブ__uid」へ移動
// =============================================
function archiveTasks(uid) {
  const src = userSheet('タスク', TASK_HEADERS, uid);
  const rows = src.getDataRange().getValues();
  const archive = userSheet('タスク_アーカイブ', TASK_HEADERS, uid);
  let count = 0;
  // 下から走査：完了（6列目=index5がtrue/TRUE）の行を退避して削除
  for (let i = rows.length - 1; i >= 1; i--) {
    const done = rows[i][5] === true || rows[i][5] === 'TRUE';
    if (done) {
      archive.appendRow(rows[i]);
      src.deleteRow(i + 1);
      count++;
    }
  }
  return { ok: true, count };
}

// アーカイブ：日次記録をすべて「日次記録_アーカイブ__uid」へ移動
function archiveDaily(uid) {
  const src = userSheet('日次記録', DAILY_HEADERS, uid);
  const rows = src.getDataRange().getValues();
  const archive = userSheet('日次記録_アーカイブ', DAILY_HEADERS, uid);
  let count = 0;
  for (let i = rows.length - 1; i >= 1; i--) {
    archive.appendRow(rows[i]);
    src.deleteRow(i + 1);
    count++;
  }
  return { ok: true, count };
}

// =============================================
// 声を届ける：本人の「声__uid」タブに追記
// =============================================
const VOICE_HEADERS = ['日付', '内容', '受信日時'];

function saveVoice(uid, data) {
  const sheet = userSheet('声', VOICE_HEADERS, uid);
  sheet.appendRow([
    (data && data.date) || fmtDate(new Date()),
    (data && data.content) || '',
    fmtDate(new Date()),
  ]);
  return { ok: true };
}

// =============================================
// 既存データ移行（エディタから手動実行）
// 旧：無印タブ（タスク / 日次記録 / 設定）を、あなたのユーザーID付きタブへリネーム。
// 使い方：下の 'u001' を、addUser で発行された「あなたのユーザーID」に書き換えて実行。
// =============================================
function migrateLegacyData() {
  const uid = 'u001'; // ← あなたのユーザーIDに変更してから実行
  [['タスク', TASK_HEADERS], ['日次記録', DAILY_HEADERS], ['設定', SETTINGS_HEADERS]].forEach(([base, headers]) => {
    const legacy = SS.getSheetByName(base);
    const target = base + '__' + uid;
    if (legacy && !SS.getSheetByName(target)) {
      legacy.setName(target);
      Logger.log('移行: %s → %s', base, target);
    } else if (legacy && SS.getSheetByName(target)) {
      Logger.log('スキップ（移行先が既に存在）: %s', target);
    } else {
      Logger.log('対象なし: %s', base);
    }
  });
  return { ok: true };
}
