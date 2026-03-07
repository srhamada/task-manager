// ============================================================
// Google Apps Script — 社労士業務管理ツール
// スプレッドシート: TODO / 記録 / 給与計算記録
// ============================================================

// --------------- 定数 ---------------
var SHEET_TODO    = 'TODO';
var SHEET_RECORD  = '記録';
var SHEET_PAYROLL = '給与計算記録';

// 給与計算系の業務種別（ここに追加すれば分岐が増やせる）
var PAYROLL_CATEGORIES = ['給与計算', '賞与計算', '給与修正・再計算'];

// --------------- doGet: シートデータを返す ---------------
// パラメータ ?sheet=記録 で取得先を切り替え可能（デフォルト: TODO）
function doGet(e) {
  var sheetName = (e && e.parameter && e.parameter.sheet) ? e.parameter.sheet : SHEET_TODO;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    return jsonResponse_({ error: 'シート「' + sheetName + '」が見つかりません' });
  }

  Logger.log('[doGet] シート: ' + sheetName);

  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var result = [];

  for (var i = 1; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      obj[headers[j]] = data[i][j];
    }
    result.push(obj);
  }

  Logger.log('[doGet] 件数: ' + result.length);

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// --------------- doPost: アクション振り分け ---------------
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    Logger.log('[doPost] action=' + (data.action || '(なし)'));

    // ── TODO完了アクション（完了処理を一括実行） ──
    if (data.action === 'completeTodo') {
      return handleCompleteTodo_(ss, data);
    }

    // ── 行更新アクション ──
    if (data.action === 'updateRow') {
      return handleUpdateRow_(ss, data);
    }

    // ── 従来の新規追加処理 ──
    var sheetName = data.sheet || SHEET_TODO;
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      return jsonResponse_({ error: 'シート「' + sheetName + '」が見つかりません' });
    }

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    // IDの自動採番（TODO新規追加時）
    if (sheetName === SHEET_TODO && !data['ID']) {
      data['ID'] = getNextId_(sheet, headers);
    }

    var row = headers.map(function(h) {
      return data[h] !== undefined ? data[h] : '';
    });

    sheet.appendRow(row);
    Logger.log('[doPost] 新規追加: シート=' + sheetName);

    return jsonResponse_({ success: true });

  } catch (err) {
    Logger.log('[doPost] エラー: ' + err.message);
    return jsonResponse_({ error: err.message });
  }
}

// ===============================================================
// TODO完了処理（一括）
// ブラウザ → completeTodo → TODO更新＋記録保存＋記録済セット
// ===============================================================
function handleCompleteTodo_(ss, data) {
  var todoSheet = ss.getSheetByName(SHEET_TODO);
  if (!todoSheet) return jsonResponse_({ error: 'TODOシートが見つかりません' });

  var allData = todoSheet.getDataRange().getValues();
  var headers = allData[0];
  var idCol = headers.indexOf('ID');
  if (idCol === -1) return jsonResponse_({ error: 'ID列が見つかりません' });

  var targetId = String(data.id);
  Logger.log('[completeTodo] 対象ID: ' + targetId);

  // --- 対象行を探す ---
  var rowNum = -1;
  for (var i = 1; i < allData.length; i++) {
    if (String(allData[i][idCol]) === targetId) {
      rowNum = i + 1;
      break;
    }
  }
  if (rowNum === -1) {
    Logger.log('[completeTodo] ID=' + targetId + ' が見つかりません');
    return jsonResponse_({ error: '該当ID(' + targetId + ')がTODOシートに見つかりません' });
  }
  Logger.log('[completeTodo] ID=' + targetId + ' → 行' + rowNum);

  // --- 二重保存チェック ---
  var recordedCol = headers.indexOf('記録済');
  if (recordedCol !== -1) {
    var recordedVal = todoSheet.getRange(rowNum, recordedCol + 1).getValue();
    if (recordedVal !== '' && recordedVal !== null && recordedVal !== undefined) {
      Logger.log('[completeTodo] 既に記録済 → スキップ');
      return jsonResponse_({ success: true, skipped: true, message: '既に記録済です' });
    }
  }

  // --- TODOシートの状態を「完了」に更新 ---
  var statusCol = headers.indexOf('状態');
  if (statusCol !== -1) {
    todoSheet.getRange(rowNum, statusCol + 1).setValue('完了');
  }

  // 更新日もセット
  var updatedCol = headers.indexOf('更新日');
  if (updatedCol !== -1) {
    todoSheet.getRange(rowNum, updatedCol + 1).setValue(new Date());
  }

  // --- 最新の行データを取得 ---
  var colCount = headers.length;
  var todoValues = todoSheet.getRange(rowNum, 1, 1, colCount).getValues()[0];

  // ヘルパー：列名から値を取得
  function tv(colName) {
    var idx = headers.indexOf(colName);
    return idx !== -1 ? todoValues[idx] : '';
  }

  // --- 業務種別を判定 ---
  var category = String(tv('業務種別'));
  var isPayroll = PAYROLL_CATEGORIES.indexOf(category) !== -1;
  var targetSheetName = isPayroll ? SHEET_PAYROLL : SHEET_RECORD;
  Logger.log('[completeTodo] 業務種別: "' + category + '" → 保存先: ' + targetSheetName);

  // --- 記録シートへ保存 ---
  var targetSheet = ss.getSheetByName(targetSheetName);
  if (!targetSheet) {
    Logger.log('[completeTodo] 保存先シート「' + targetSheetName + '」が見つかりません！');
    return jsonResponse_({ error: '保存先シート「' + targetSheetName + '」が見つかりません' });
  }

  if (isPayroll) {
    // --- 給与計算記録シートへ保存 ---
    var payrollMonth = '';
    var dueVal = tv('期限');
    if (dueVal) {
      var dueDate = new Date(dueVal);
      if (!isNaN(dueDate.getTime())) {
        payrollMonth = dueDate.getFullYear() + '-' + ('0' + (dueDate.getMonth() + 1)).slice(-2);
      } else {
        payrollMonth = String(dueVal);
      }
    }
    var workMinutes = convertToMinutes_(tv('作業時間'));
    var newId = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMddHHmmss');

    var payrollRow = [
      newId,                 // A: ID
      tv('ID'),              // B: 元タスクID
      tv('クライアント'),    // C
      payrollMonth,          // D: 給与対象月
      tv('タスク内容'),      // E
      tv('担当者'),          // F
      new Date(),            // G: 完了日
      workMinutes,           // H: 作業時間（分）
      '',                    // I: 人数（空欄）
      '',                    // J: メモ（空欄）
      '',                    // K: （空欄）
      tv('開始時刻'),        // L: 開始時刻
      tv('終了時刻'),        // M: 終了時刻
      tv('作業時間')         // N: 作業時間
    ];
    targetSheet.appendRow(payrollRow);
    Logger.log('[completeTodo] 給与計算記録に追加完了: 元ID=' + tv('ID') + ', クライアント=' + tv('クライアント'));

  } else {
    // --- 記録シートへ保存 ---
    var recordRow = [
      tv('ID'),             // A: ID
      tv('クライアント'),    // B
      tv('業務種別'),        // C
      tv('タスク内容'),      // D
      tv('期限'),            // E
      tv('担当者'),          // F
      tv('状態'),            // G（完了）
      tv('優先度'),          // H
      tv('作成日'),          // I
      tv('更新日'),          // J
      new Date(),            // K: 完了日時
      tv('開始時刻'),        // L
      tv('終了時刻'),        // M
      tv('作業時間'),        // N
      'TODO',                // O: 元シート
      rowNum                 // P: 元行番号
    ];
    targetSheet.appendRow(recordRow);
    Logger.log('[completeTodo] 記録シートに追加完了: ID=' + tv('ID') + ', 業務種別=' + category + ', タスク=' + tv('タスク内容'));
  }

  // --- 記録済に日時を入れる ---
  if (recordedCol !== -1) {
    var now = new Date();
    todoSheet.getRange(rowNum, recordedCol + 1).setValue(now);
    Logger.log('[completeTodo] TODO行' + rowNum + ' 記録済に日時セット: ' + now);
  }

  Logger.log('[completeTodo] 完了処理 すべて成功');
  return jsonResponse_({
    success: true,
    savedTo: targetSheetName,
    todoId: targetId,
    category: category
  });
}

// --------------- 行更新処理（完了以外のステータス変更用） ---------------
function handleUpdateRow_(ss, data) {
  var sheet = ss.getSheetByName(SHEET_TODO);
  if (!sheet) return jsonResponse_({ error: 'TODOシートが見つかりません' });

  var allData = sheet.getDataRange().getValues();
  var headers = allData[0];
  var idCol = headers.indexOf('ID');
  if (idCol === -1) return jsonResponse_({ error: 'ID列が見つかりません' });

  var targetId = String(data.id);
  Logger.log('[updateRow] 対象ID: ' + targetId + ', 状態: ' + data['状態']);

  for (var i = 1; i < allData.length; i++) {
    if (String(allData[i][idCol]) === targetId) {
      var rowNum = i + 1;
      Logger.log('[updateRow] ID=' + targetId + ' → 行' + rowNum);

      // 送信されたフィールドをヘッダーに基づいて更新
      for (var j = 0; j < headers.length; j++) {
        var colName = headers[j];
        if (colName === 'ID' || colName === '作成日' || colName === '記録済') continue;
        if (data[colName] !== undefined) {
          sheet.getRange(rowNum, j + 1).setValue(data[colName]);
        }
      }

      return jsonResponse_({ success: true });
    }
  }

  return jsonResponse_({ error: '該当ID(' + targetId + ')が見つかりません' });
}

// --------------- ユーティリティ ---------------

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getNextId_(sheet, headers) {
  var idCol = headers.indexOf('ID');
  if (idCol === -1) return 1;
  var data = sheet.getDataRange().getValues();
  var maxId = 0;
  for (var i = 1; i < data.length; i++) {
    var v = parseInt(data[i][idCol], 10);
    if (v > maxId) maxId = v;
  }
  return maxId + 1;
}

function convertToMinutes_(val) {
  if (val === '' || val === null || val === undefined) return '';

  if (typeof val === 'number') {
    if (val < 1 && val > 0) {
      return Math.round(val * 24 * 60);
    }
    return val;
  }

  if (val instanceof Date) {
    return val.getHours() * 60 + val.getMinutes();
  }

  var str = String(val);
  var match = str.match(/^(\d+):(\d+)$/);
  if (match) {
    return parseInt(match[1], 10) * 60 + parseInt(match[2], 10);
  }

  var num = parseInt(str, 10);
  if (!isNaN(num)) return num;

  return '';
}
