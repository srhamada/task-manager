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

// --------------- doPost: 新規追加 / 行更新 ---------------
function doPost(e) {
  var data = JSON.parse(e.postData.contents);
  var ss = SpreadsheetApp.getActiveSpreadsheet();

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

  return jsonResponse_({ success: true });
}

// --------------- 行更新処理 ---------------
function handleUpdateRow_(ss, data) {
  var sheet = ss.getSheetByName(SHEET_TODO);
  if (!sheet) return jsonResponse_({ error: 'TODOシートが見つかりません' });

  var allData = sheet.getDataRange().getValues();
  var headers = allData[0];
  var idCol = headers.indexOf('ID');
  if (idCol === -1) return jsonResponse_({ error: 'ID列が見つかりません' });

  var targetId = String(data.id);
  Logger.log('[handleUpdateRow_] 対象ID: ' + targetId + ', 状態: ' + data['状態']);

  for (var i = 1; i < allData.length; i++) {
    if (String(allData[i][idCol]) === targetId) {
      var rowNum = i + 1;
      Logger.log('[handleUpdateRow_] ID=' + targetId + ' を行' + rowNum + 'で発見');

      // 送信されたフィールドをヘッダーに基づいて更新
      for (var j = 0; j < headers.length; j++) {
        var colName = headers[j];
        // ID・作成日・記録済は上書きしない
        if (colName === 'ID' || colName === '作成日' || colName === '記録済') continue;
        if (data[colName] !== undefined) {
          sheet.getRange(rowNum, j + 1).setValue(data[colName]);
        }
      }

      // 状態が「完了」なら記録シートへ保存
      if (data['状態'] === '完了') {
        Logger.log('[handleUpdateRow_] 完了検知 → saveCompletedTask_ 呼び出し');
        saveCompletedTask_(ss, sheet, allData, headers, i, rowNum);
      }

      return jsonResponse_({ success: true, saved: data['状態'] === '完了' });
    }
  }

  Logger.log('[handleUpdateRow_] ID=' + targetId + ' が見つかりません');
  return jsonResponse_({ error: '該当ID(' + targetId + ')が見つかりません' });
}

// --------------- 完了時の記録保存（メイン分岐） ---------------
function saveCompletedTask_(ss, todoSheet, allData, headers, dataIndex, rowNum) {
  // K列（記録済）チェック — 二重保存防止
  var recordedCol = headers.indexOf('記録済');
  if (recordedCol !== -1) {
    var recordedVal = todoSheet.getRange(rowNum, recordedCol + 1).getValue();
    if (recordedVal !== '' && recordedVal !== null) {
      Logger.log('[saveCompletedTask_] 記録済のためスキップ（行' + rowNum + '）');
      return;
    }
  }

  // 最新の行データを取得（直前のsetValueを反映）
  var todoValues = todoSheet.getRange(rowNum, 1, 1, headers.length).getValues()[0];

  // 業務種別で分岐
  var categoryCol = headers.indexOf('業務種別');
  var category = categoryCol !== -1 ? String(todoValues[categoryCol]) : '';
  Logger.log('[saveCompletedTask_] 業務種別: "' + category + '", 行: ' + rowNum);

  if (PAYROLL_CATEGORIES.indexOf(category) !== -1) {
    Logger.log('[saveCompletedTask_] → 給与計算記録シートへ保存');
    saveToPayrollRecordSheet_(ss, headers, todoValues, rowNum);
  } else {
    Logger.log('[saveCompletedTask_] → 記録シートへ保存');
    saveToRecordSheet_(ss, headers, todoValues, rowNum);
  }

  // 記録済に日時を入れる
  if (recordedCol !== -1) {
    var now = new Date();
    todoSheet.getRange(rowNum, recordedCol + 1).setValue(now);
    Logger.log('[saveCompletedTask_] 記録済に日時セット: ' + now);
  }
}

// --------------- 記録シートへ保存 ---------------
function saveToRecordSheet_(ss, todoHeaders, todoValues, todoRowNum) {
  var sheet = ss.getSheetByName(SHEET_RECORD);
  if (!sheet) {
    Logger.log('[saveToRecordSheet_] 記録シートが見つかりません！');
    return;
  }

  // ヘルパー：TODO列の値を取得
  function tv(colName) {
    var idx = todoHeaders.indexOf(colName);
    return idx !== -1 ? todoValues[idx] : '';
  }

  // 記録シートの列構成に合わせて行を組み立て
  // A:ID  B:クライアント  C:業務種別  D:タスク内容  E:期限  F:担当者
  // G:状態  H:優先度  I:作成日  J:更新日  K:完了日時
  // L:開始時刻  M:終了時刻  N:作業時間  O:元シート  P:元行番号
  var row = [
    tv('ID'),             // A: ID
    tv('クライアント'),    // B
    tv('業務種別'),        // C
    tv('タスク内容'),      // D
    tv('期限'),            // E
    tv('担当者'),          // F
    tv('状態'),            // G
    tv('優先度'),          // H
    tv('作成日'),          // I
    tv('更新日'),          // J
    new Date(),            // K: 完了日時
    tv('開始時刻'),        // L
    tv('終了時刻'),        // M
    tv('作業時間'),        // N
    'TODO',                // O: 元シート
    todoRowNum             // P: 元行番号
  ];

  sheet.appendRow(row);
  Logger.log('[saveToRecordSheet_] 記録シートに追加完了: ID=' + tv('ID') + ', 業務種別=' + tv('業務種別') + ', タスク=' + tv('タスク内容'));
}

// --------------- 給与計算記録シートへ保存 ---------------
function saveToPayrollRecordSheet_(ss, todoHeaders, todoValues, todoRowNum) {
  var sheet = ss.getSheetByName(SHEET_PAYROLL);
  if (!sheet) return;

  function tv(colName) {
    var idx = todoHeaders.indexOf(colName);
    return idx !== -1 ? todoValues[idx] : '';
  }

  // 給与対象月：期限(E列) から yyyy-MM 形式に変換
  var payrollMonth = '';
  var dueVal = tv('期限');
  if (dueVal) {
    var dueDate = new Date(dueVal);
    if (!isNaN(dueDate.getTime())) {
      var y = dueDate.getFullYear();
      var m = ('0' + (dueDate.getMonth() + 1)).slice(-2);
      payrollMonth = y + '-' + m;
    } else {
      // Date変換できない場合はそのまま入れる
      payrollMonth = String(dueVal);
    }
  }

  // 作業時間（分）への変換
  var workMinutes = convertToMinutes_(tv('作業時間'));

  // 新規ID（日時ベース）
  var newId = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMddHHmmss');

  // 給与計算記録シートの列構成に合わせて行を組み立て
  // A:ID  B:元タスクID  C:クライアント  D:給与対象月  E:タスク内容
  // F:担当者  G:完了日  H:作業時間（分）  I:人数  J:メモ  K:(空)
  // L:開始時刻  M:終了時刻  N:作業時間
  var row = [
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

  sheet.appendRow(row);
}

// --------------- ユーティリティ ---------------

// JSON レスポンスを返す
function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// 次のIDを自動採番
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

// 作業時間を分に変換
// 対応形式:
//   数値（そのまま分として扱う）
//   "1:30" や "01:30" → 90分
//   Date型（時刻として 1899/12/30 ベース）→ 分に換算
//   空・不正値 → 空文字
function convertToMinutes_(val) {
  if (val === '' || val === null || val === undefined) return '';

  // 数値の場合はそのまま分として返す
  if (typeof val === 'number') {
    // Sheets の時刻型は小数（例: 0.0625 = 1:30）の場合がある
    if (val < 1 && val > 0) {
      // 小数 → 24時間制の分に変換
      return Math.round(val * 24 * 60);
    }
    return val;
  }

  // Date型（Sheetsの時刻セルはDateになる）
  if (val instanceof Date) {
    return val.getHours() * 60 + val.getMinutes();
  }

  // 文字列 "H:MM" or "HH:MM"
  var str = String(val);
  var match = str.match(/^(\d+):(\d+)$/);
  if (match) {
    return parseInt(match[1], 10) * 60 + parseInt(match[2], 10);
  }

  // 数値文字列
  var num = parseInt(str, 10);
  if (!isNaN(num)) return num;

  return '';
}
