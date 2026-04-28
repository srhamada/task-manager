// ============================================================
// Google Apps Script — 社労士業務管理ツール
// スプレッドシート: TODO / 記録 / 給与計算記録 / クライアント / 行政問い合わせ記録
// ============================================================

// --------------- 定数 ---------------
var SHEET_TODO    = 'TODO';
var SHEET_RECORD  = '記録';
var SHEET_PAYROLL = '給与計算記録';
var SHEET_CLIENT  = 'クライアント';
var SHEET_INQUIRY = '行政問い合わせ記録';
var SHEET_CONSULT = '相談記録';
var SHEET_BUSY    = '算定年更管理';
var SHEET_STRESS  = 'ストレスチェック管理';
var SHEET_ACTIVITY = 'アクティビティログ';
var SHEET_HOLIDAY  = '休日マスタ';
var SHEET_MESSAGE  = '一言メッセージ';

// 給与計算系の業務種別（ここに追加すれば分岐が増やせる）
var PAYROLL_CATEGORIES = ['給与計算', '賞与計算', '給与修正・再計算', '会計入力'];

// --------------- doGet: シートデータを返す ---------------
// パラメータ ?sheet=記録 で取得先を切り替え可能（デフォルト: TODO）
// パラメータ ?action=getConsultTodos で相談記録の要対応データのみ返す
function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : '';

  // ── 相談記録TODO取得（O列=TRUE かつ H列≠完了） ──
  if (action === 'getConsultTodos') {
    return handleGetConsultTodos_();
  }

  // ── 算定年更管理データ取得 ──
  if (action === 'getBusySeasonRecords') {
    return handleGetBusySeasonRecords_(e);
  }

  // ── ストレスチェック管理データ取得 ──
  if (action === 'getStressCheckRecords') {
    return handleGetStressCheckRecords_();
  }

  // ── アクティビティログ取得（直近20件） ──
  if (action === 'getActivityLog') {
    return handleGetActivityLog_();
  }

  // ── 休日マスタ取得 ──
  if (action === 'getHolidayMaster') {
    return handleGetHolidayMaster_();
  }

  // ── 担当者状態取得 ──
  if (action === 'getMemberStatuses') {
    return handleGetMemberStatuses_();
  }

  // ── 一言メッセージ取得 ──
  if (action === 'getMessages') {
    return handleGetMessages_();
  }

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

// --------------- 相談記録TODO取得 ---------------
function handleGetConsultTodos_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_CONSULT);
  if (!sheet) {
    return jsonResponse_({ error: 'シート「' + SHEET_CONSULT + '」が見つかりません' });
  }

  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  Logger.log('[getConsultTodos] ヘッダー: ' + JSON.stringify(headers));

  // ヘッダーから列インデックスを取得
  var flagCol   = headers.indexOf('要対応フラグ');
  var statusCol = headers.indexOf('ステータス');
  Logger.log('[getConsultTodos] 要対応フラグ列=' + flagCol + ', ステータス列=' + statusCol);

  var result = [];
  for (var i = 1; i < data.length; i++) {
    // 要対応フラグ: TRUE / "TRUE" / true を全て吸収
    var flagVal = (flagCol >= 0) ? String(data[i][flagCol]).toUpperCase().trim() : '';
    var statusVal = (statusCol >= 0) ? String(data[i][statusCol]).trim() : '';

    if (flagVal !== 'TRUE') continue;
    if (statusVal === '完了') continue;

    var obj = { _rowIndex: i + 1 }; // スプレッドシートの行番号
    for (var j = 0; j < headers.length; j++) {
      obj[headers[j]] = data[i][j];
    }
    result.push(obj);
  }

  Logger.log('[getConsultTodos] 対象件数: ' + result.length);
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// --------------- doPost: アクション振り分け ---------------
function doPost(e) {
  try {
    // URLSearchParams(data=JSON) または application/json の両方に対応
    var rawContents = (e.parameter && e.parameter.data) ? e.parameter.data : e.postData.contents;
    Logger.log('[doPost] ★受信rawデータ: ' + rawContents);
    var data = JSON.parse(rawContents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    Logger.log('[doPost] action=' + (data.action || '(なし)'));
    Logger.log('[doPost] ★data全体: ' + JSON.stringify(data));

    // ── アクティビティログ追加 ──
    if (data.action === 'logActivity') {
      return handleLogActivity_(ss, data);
    }

    // ── TODO完了アクション（完了処理を一括実行） ──
    if (data.action === 'completeTodo') {
      Logger.log('[doPost] ★completeTodo呼出前: 開始時刻=' + data['開始時刻'] + ', 終了時刻=' + data['終了時刻'] + ', 作業時間=' + data['作業時間']);
      return handleCompleteTodo_(ss, data);
    }

    // ── 相談記録TODO完了（O列をFALSEに） ──
    if (data.action === 'completeConsultTodo') {
      return handleCompleteConsultTodo_(ss, data);
    }

    // ── 担当者状態更新 ──
    if (data.action === 'setMemberStatus') {
      return handleSetMemberStatus_(data);
    }

    // ── 一言メッセージ投稿 ──
    if (data.action === 'postMessage') {
      return handlePostMessage_(data);
    }

    // ── 一言メッセージ削除 ──
    if (data.action === 'deleteMessage') {
      return handleDeleteMessage_(data);
    }

    // ── 算定年更管理 一括保存 ──
    if (data.action === 'saveBusySeasonRecords') {
      return handleSaveBusySeasonRecords_(ss, data);
    }

    // ── ストレスチェック管理 新規追加 ──
    if (data.action === 'saveStressCheckRecord') {
      return handleSaveStressCheckRecord_(ss, data);
    }

    // ── ストレスチェック管理 更新 ──
    if (data.action === 'updateStressCheckRecord') {
      return handleUpdateStressCheckRecord_(ss, data);
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
    var nowJST = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');

    // IDの自動採番（TODO新規追加時）
    if (sheetName === SHEET_TODO && !data['ID']) {
      data['ID'] = getNextId_(sheet, headers);
    }

    // クライアントシート: client_id・作成日時・更新日時・有効フラグを自動設定
    if (sheetName === SHEET_CLIENT) {
      if (!data['client_id']) {
        data['client_id'] = getNextIdByCol_(sheet, headers, 'client_id');
      }
      if (!data['作成日時']) data['作成日時'] = nowJST;
      if (!data['更新日時']) data['更新日時'] = nowJST;
      if (data['有効フラグ'] === undefined || data['有効フラグ'] === '') data['有効フラグ'] = 'TRUE';
      Logger.log('[doPost] クライアント新規追加: client_id=' + data['client_id']);
    }

    // 行政問い合わせ記録シート: inquiry_id・作成日時・更新日時・有効フラグを自動設定
    if (sheetName === SHEET_INQUIRY) {
      if (!data['inquiry_id']) {
        data['inquiry_id'] = getNextIdByCol_(sheet, headers, 'inquiry_id');
      }
      if (!data['作成日時']) data['作成日時'] = nowJST;
      if (!data['更新日時']) data['更新日時'] = nowJST;
      if (data['有効フラグ'] === undefined || data['有効フラグ'] === '') data['有効フラグ'] = 'TRUE';
      Logger.log('[doPost] 行政問い合わせ新規追加: inquiry_id=' + data['inquiry_id']);
    }

    // 相談記録シート: ヘッダー名の揺れ吸収
    if (sheetName === SHEET_CONSULT) {
      Logger.log('[doPost] ★相談記録 ヘッダー: ' + JSON.stringify(headers));
      // 「作業時間（分）」→「作業時間」のマッピング
      if (data['作業時間（分）'] !== undefined && data['作業時間'] === undefined) {
        data['作業時間'] = data['作業時間（分）'];
      }
      // 「要対応フラグ（TRUE / FALSE）」→「要対応フラグ」のマッピング
      if (data['要対応フラグ（TRUE / FALSE）'] !== undefined && data['要対応フラグ'] === undefined) {
        data['要対応フラグ'] = data['要対応フラグ（TRUE / FALSE）'];
      }
    }

    var row = headers.map(function(h) {
      return data[h] !== undefined ? data[h] : '';
    });

    // appendRow 実行前ログ
    var lastRowBefore = sheet.getLastRow();
    Logger.log('[doPost] ★appendRow実行前: シート=' + sheetName + ', 最終行=' + lastRowBefore);
    Logger.log('[doPost] ★書き込み配列: ' + JSON.stringify(row));

    sheet.appendRow(row);

    // appendRow 実行後ログ
    var lastRowAfter = sheet.getLastRow();
    Logger.log('[doPost] ★appendRow実行後: 最終行=' + lastRowAfter);

    if (lastRowAfter <= lastRowBefore) {
      Logger.log('[doPost] ❌ appendRow後に行が増えていない');
      return jsonResponse_({ success: false, error: 'appendRowが反映されませんでした', sheetName: sheetName });
    }

    Logger.log('[doPost] ✅ 新規追加成功: シート=' + sheetName + ', 行=' + lastRowAfter);
    return jsonResponse_({
      success: true,
      sheetName: sheetName,
      writtenRow: lastRowAfter,
      savedData: row,
      headers: headers
    });

  } catch (err) {
    Logger.log('[doPost] ❌ エラー: ' + err.message);
    return jsonResponse_({ success: false, error: err.message });
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

  // 更新日もセット（JST）
  var updatedCol = headers.indexOf('更新日');
  var todoNowJST = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
  if (updatedCol !== -1) {
    todoSheet.getRange(rowNum, updatedCol + 1).setValue(todoNowJST);
  }

  // --- 最新の行データを取得 ---
  var colCount = headers.length;
  var todoValues = todoSheet.getRange(rowNum, 1, 1, colCount).getValues()[0];

  // ヘルパー：列名から値を取得
  function tv(colName) {
    var idx = headers.indexOf(colName);
    return idx !== -1 ? todoValues[idx] : '';
  }

  // --- 業務種別を判定（C列を参照、trimして比較） ---
  var categoryRaw = tv('業務種別');
  var category = String(categoryRaw || '').trim();
  var isPayroll = PAYROLL_CATEGORIES.indexOf(category) !== -1;
  var targetSheetName = isPayroll ? SHEET_PAYROLL : SHEET_RECORD;
  Logger.log('[completeTodo] 業務種別(元値): "' + categoryRaw + '"');
  Logger.log('[completeTodo] 業務種別(trim後): "' + category + '"');
  Logger.log('[completeTodo] 分岐結果: ' + (isPayroll ? '給与計算記録' : '記録') + ' → 保存先: ' + targetSheetName);

  // --- 記録シートへ保存 ---
  var targetSheet = ss.getSheetByName(targetSheetName);
  if (!targetSheet) {
    Logger.log('[completeTodo] 保存先シート「' + targetSheetName + '」が見つかりません！');
    return jsonResponse_({ error: '保存先シート「' + targetSheetName + '」が見つかりません' });
  }

  // --- POSTデータから作業時間を取得（フロントから直接渡される） ---
  var postWorkStart   = data['開始時刻'] || '';
  var postWorkEnd     = data['終了時刻'] || '';
  var postWorkMinutes = data['作業時間'] || '';
  Logger.log('[completeTodo] POST受信: 開始時刻=' + postWorkStart + ', 終了時刻=' + postWorkEnd + ', 作業時間=' + postWorkMinutes);

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
    var workMinutes = convertToMinutes_(postWorkMinutes);
    var newId = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMddHHmmss');

    var nowJST = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
    var payrollRow = [
      newId,                 // A: ID
      tv('ID'),              // B: 元タスクID
      tv('クライアント'),    // C
      payrollMonth,          // D: 給与対象月
      tv('タスク内容'),      // E
      tv('担当者'),          // F
      nowJST,                // G: 完了日（JST）
      workMinutes,           // H: 作業時間（分）
      '',                    // I: 人数（空欄）
      '',                    // J: メモ（空欄）
      '',                    // K: （空欄）
      postWorkStart,         // L: 開始時刻（POSTから）
      postWorkEnd,           // M: 終了時刻（POSTから）
      postWorkMinutes        // N: 作業時間（POSTから）
    ];
    Logger.log('[completeTodo] ★給与計算記録 書込直前: payrollRow[11]=' + payrollRow[11] + ', [12]=' + payrollRow[12] + ', [13]=' + payrollRow[13]);
    Logger.log('[completeTodo] ★payrollRow全体: ' + JSON.stringify(payrollRow));
    targetSheet.appendRow(payrollRow);
    Logger.log('[completeTodo] 給与計算記録に追加完了: 元ID=' + tv('ID') + ', クライアント=' + tv('クライアント'));

  } else {
    // --- 記録シートへ保存 ---
    var nowJST2 = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
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
      nowJST2,               // J: 更新日（JST）
      nowJST2,               // K: 完了日時（JST）
      postWorkStart,         // L: 開始時刻（POSTから）
      postWorkEnd,           // M: 終了時刻（POSTから）
      postWorkMinutes,       // N: 作業時間（POSTから）
      'TODO',                // O: 元シート
      rowNum                 // P: 元行番号
    ];
    Logger.log('[completeTodo] ★記録シート 書込直前: recordRow[11]=' + recordRow[11] + ', [12]=' + recordRow[12] + ', [13]=' + recordRow[13]);
    Logger.log('[completeTodo] ★recordRow全体: ' + JSON.stringify(recordRow));
    targetSheet.appendRow(recordRow);
    Logger.log('[completeTodo] 記録シートに追加完了: ID=' + tv('ID') + ', 業務種別=' + category + ', タスク=' + tv('タスク内容'));
  }

  // --- 記録済に日時を入れる ---
  if (recordedCol !== -1) {
    var recordedNowJST = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
    todoSheet.getRange(rowNum, recordedCol + 1).setValue(recordedNowJST);
    Logger.log('[completeTodo] TODO行' + rowNum + ' 記録済に日時セット: ' + recordedNowJST);
  }

  Logger.log('[completeTodo] 完了処理 すべて成功');
  return jsonResponse_({
    success: true,
    savedTo: targetSheetName,
    todoId: targetId,
    category: category
  });
}

// --------------- 行更新処理（既存行のステータス等を更新） ---------------
// appendRow は使わず、必ず該当IDの行を探して上書きする
function handleUpdateRow_(ss, data) {
  // シート名の指定があればそちらを使う（なければTODOシート）
  var sheetName = data.sheet || SHEET_TODO;
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return jsonResponse_({ error: 'シート「' + sheetName + '」が見つかりません' });

  var allData = sheet.getDataRange().getValues();
  var headers = allData[0];

  // 検索キー列の指定（デフォルトは 'ID'）
  var keyCol = data.updateKey || 'ID';
  var idCol = headers.indexOf(keyCol);
  if (idCol === -1) return jsonResponse_({ error: keyCol + '列が見つかりません（シート: ' + sheetName + '）' });

  var targetId = String(data[keyCol] || data.id || '');
  Logger.log('[updateRow] シート=' + sheetName + ', キー=' + keyCol + ', 対象ID=' + targetId);

  for (var i = 1; i < allData.length; i++) {
    if (String(allData[i][idCol]) === targetId) {
      var rowNum = i + 1;
      Logger.log('[updateRow] ID=' + targetId + ' → 行' + rowNum);

      // 送信されたフィールドをヘッダーに基づいて更新
      for (var j = 0; j < headers.length; j++) {
        var colName = headers[j];
        // ID・作成日・記録済は変更しない
        if (colName === keyCol || colName === '作成日' || colName === '記録済') continue;
        // 更新日・更新日時はGAS側の現在日時（JST）を強制使用
        if (colName === '更新日' || colName === '更新日時') {
          var updateNowJST = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
          sheet.getRange(rowNum, j + 1).setValue(updateNowJST);
          continue;
        }
        if (data[colName] !== undefined) {
          sheet.getRange(rowNum, j + 1).setValue(data[colName]);
        }
      }

      Logger.log('[updateRow] 行' + rowNum + ' の更新完了');
      return jsonResponse_({ success: true, updatedRow: rowNum
      });
    }
  }

  return jsonResponse_({ error: '該当ID(' + targetId + ')がシート「' + sheetName + '」に見つかりません' });
}

// --------------- 相談記録TODO完了 ---------------
function handleCompleteConsultTodo_(ss, data) {
  var sheet = ss.getSheetByName(SHEET_CONSULT);
  if (!sheet) return jsonResponse_({ success: false, error: 'シート「' + SHEET_CONSULT + '」が見つかりません' });

  var allData = sheet.getDataRange().getValues();
  var headers = allData[0];
  var idCol     = headers.indexOf('ID');
  var flagCol   = headers.indexOf('要対応フラグ');
  var statusCol = headers.indexOf('ステータス');
  var updateCol = headers.indexOf('更新日時');

  Logger.log('[completeConsultTodo] ID列=' + idCol + ', 要対応フラグ列=' + flagCol + ', ステータス列=' + statusCol);

  if (idCol === -1) return jsonResponse_({ success: false, error: 'ID列が見つかりません' });
  if (flagCol === -1) return jsonResponse_({ success: false, error: '要対応フラグ列が見つかりません' });

  var targetId = String(data.consultId || data.id || '');
  Logger.log('[completeConsultTodo] 対象ID=' + targetId);

  for (var i = 1; i < allData.length; i++) {
    if (String(allData[i][idCol]) === targetId) {
      var rowNum = i + 1;
      // O列を FALSE に
      sheet.getRange(rowNum, flagCol + 1).setValue('FALSE');
      Logger.log('[completeConsultTodo] 行' + rowNum + ' 要対応フラグ → FALSE');

      // H列(ステータス)を「完了」に
      if (statusCol >= 0) {
        sheet.getRange(rowNum, statusCol + 1).setValue('完了');
        Logger.log('[completeConsultTodo] 行' + rowNum + ' ステータス → 完了');
      }

      // K列(更新日時)を現在時刻に
      if (updateCol >= 0) {
        var nowJST = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
        sheet.getRange(rowNum, updateCol + 1).setValue(nowJST);
      }

      return jsonResponse_({ success: true, updatedRow: rowNum, sheetName: SHEET_CONSULT });
    }
  }

  return jsonResponse_({ success: false, error: '該当ID(' + targetId + ')が見つかりません' });
}

// --------------- 休日マスタ ---------------
function handleGetHolidayMaster_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_HOLIDAY);
  if (!sheet) {
    return ContentService.createTextOutput(JSON.stringify([]))
      .setMimeType(ContentService.MimeType.JSON);
  }
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var result = [];
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      var val = data[i][j];
      // 日付列はyyyy-MM-dd文字列に変換
      if (headers[j] === '日付' && val instanceof Date) {
        val = Utilities.formatDate(val, 'Asia/Tokyo', 'yyyy-MM-dd');
      }
      obj[headers[j]] = val;
    }
    result.push(obj);
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// --------------- アクティビティログ ---------------

// ログ追加
function handleLogActivity_(ss, data) {
  var sheet = ss.getSheetByName(SHEET_ACTIVITY);
  if (!sheet) {
    // シートがなければ作成
    var headers = ['日時','操作種別','対象種別','クライアント名','件名','詳細','実行者','元データID','元シート','備考'];
    sheet = ss.insertSheet(SHEET_ACTIVITY);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  var nowJST = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
  var row = [
    nowJST,
    data['操作種別'] || '',
    data['対象種別'] || '',
    data['クライアント名'] || '',
    data['件名'] || '',
    data['詳細'] || '',
    data['実行者'] || '',
    data['元データID'] || '',
    data['元シート'] || '',
    data['備考'] || ''
  ];
  sheet.appendRow(row);
  return jsonResponse_({ success: true });
}

// ログ取得（直近20件、新しい順）
function handleGetActivityLog_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_ACTIVITY);
  if (!sheet) {
    return ContentService.createTextOutput(JSON.stringify([]))
      .setMimeType(ContentService.MimeType.JSON);
  }
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var result = [];
  // 末尾（最新）から20件取得
  for (var i = data.length - 1; i >= 1 && result.length < 20; i--) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      obj[headers[j]] = data[i][j];
    }
    result.push(obj);
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// --------------- 算定年更管理 ---------------

// データ取得（年度フィルタ対応）
function handleGetBusySeasonRecords_(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_BUSY);
  if (!sheet) {
    // シートが無ければ空配列を返す（フロント側で初回読込時に対応）
    Logger.log('[getBusySeason] シート未作成');
    return ContentService.createTextOutput(JSON.stringify([]))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var year = (e && e.parameter && e.parameter.year) ? String(e.parameter.year) : '';
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var result = [];

  for (var i = 1; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      var val = data[i][j];
      // Date型はJST文字列に変換（JSON.stringifyのUTC変換を防止）
      if (val instanceof Date) {
        val = Utilities.formatDate(val, 'Asia/Tokyo', 'yyyy/MM/dd');
      }
      obj[headers[j]] = val;
    }
    // 年度フィルタ
    if (year && String(obj['年度']) !== year) continue;
    result.push(obj);
  }

  Logger.log('[getBusySeason] 年度=' + (year || '全件') + ', 件数=' + result.length);
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// 一括保存（全行を差し替え）
function handleSaveBusySeasonRecords_(ss, data) {
  var rows = data.rows;
  if (!Array.isArray(rows)) {
    return jsonResponse_({ success: false, error: 'rows が配列ではありません' });
  }

  var sheet = ss.getSheetByName(SHEET_BUSY);
  var headers = ['年度','顧問先区分','顧問先名',
    '年度更新_資料回収','年度更新_データ作成','年度更新_申告書作成','年度更新_申告',
    '年度更新_納付書作成','年度更新_納付額通知','年度更新_公文書通知',
    '算定基礎_データ作成','算定基礎_申告書作成','算定基礎_申請',
    '算定基礎_結果取込','算定基礎_保険料通知','算定基礎_公文書通知',
    'コメント','更新日'];

  // シートが無ければ作成
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_BUSY);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    Logger.log('[saveBusySeason] シート新規作成');
  }

  var year = data.year ? String(data.year) : '';
  Logger.log('[saveBusySeason] 年度=' + year + ', 行数=' + rows.length);

  // 既存データ読み込み
  var existingData = sheet.getDataRange().getValues();
  var existingHeaders = existingData[0];
  var yearCol = existingHeaders.indexOf('年度');

  // 指定年度以外の行を保持
  var keepRows = [];
  for (var i = 1; i < existingData.length; i++) {
    if (year && String(existingData[i][yearCol]) !== year) {
      keepRows.push(existingData[i]);
    }
  }

  // 新しい行を構築
  var nowJST = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
  var newRows = rows.map(function(row) {
    return headers.map(function(h) {
      if (h === '更新日') return nowJST;
      return row[h] !== undefined ? row[h] : '';
    });
  });

  // 全データ = ヘッダー + 他年度 + 今回年度
  var allRows = [headers].concat(keepRows).concat(newRows);

  // シートをクリアして書き直し
  sheet.clearContents();
  if (allRows.length > 0) {
    var range = sheet.getRange(1, 1, allRows.length, headers.length);
    // 日付列（D〜P: 4列目〜16列目）をテキスト形式にしてDate自動変換を防止
    for (var ci = 3; ci <= 15; ci++) {
      sheet.getRange(2, ci + 1, Math.max(allRows.length - 1, 1), 1).setNumberFormat('@');
    }
    range.setValues(allRows);
  }

  Logger.log('[saveBusySeason] 保存完了: 全' + (allRows.length - 1) + '行（今回' + newRows.length + '行）');
  return jsonResponse_({ success: true, savedCount: newRows.length, totalCount: allRows.length - 1 });
}

// --------------- ストレスチェック管理 ---------------

// データ取得（全件）
function handleGetStressCheckRecords_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_STRESS);
  if (!sheet) {
    Logger.log('[getStressCheck] シート未作成');
    return ContentService.createTextOutput(JSON.stringify([]))
      .setMimeType(ContentService.MimeType.JSON);
  }
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    return ContentService.createTextOutput(JSON.stringify([]))
      .setMimeType(ContentService.MimeType.JSON);
  }
  var headers = data[0];
  var result = [];
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      var val = data[i][j];
      if (val instanceof Date) {
        val = Utilities.formatDate(val, 'Asia/Tokyo', 'yyyy-MM-dd');
      }
      obj[headers[j]] = val;
    }
    result.push(obj);
  }
  Logger.log('[getStressCheck] 件数: ' + result.length);
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// 新規追加（ID・作成日・更新日を自動セット）
function handleSaveStressCheckRecord_(ss, data) {
  var sheet = ss.getSheetByName(SHEET_STRESS);
  if (!sheet) {
    return jsonResponse_({ success: false, error: 'シート「' + SHEET_STRESS + '」が見つかりません' });
  }
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var nowJST = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');

  data['ID'] = getNextId_(sheet, headers);
  data['作成日'] = nowJST;
  data['更新日'] = nowJST;

  var row = headers.map(function(h) {
    return data[h] !== undefined ? data[h] : '';
  });
  sheet.appendRow(row);
  Logger.log('[saveStressCheck] 新規追加: ID=' + data['ID'] + ', 事業所=' + data['事業所名']);
  return jsonResponse_({ success: true, id: data['ID'] });
}

// 更新（ID指定で1行更新、更新日はGAS側で自動セット）
function handleUpdateStressCheckRecord_(ss, data) {
  var sheet = ss.getSheetByName(SHEET_STRESS);
  if (!sheet) {
    return jsonResponse_({ success: false, error: 'シート「' + SHEET_STRESS + '」が見つかりません' });
  }
  var allData = sheet.getDataRange().getValues();
  var headers = allData[0];
  var idCol = headers.indexOf('ID');
  if (idCol === -1) {
    return jsonResponse_({ success: false, error: 'ID列が見つかりません' });
  }
  // 「重要フラグ」を含む更新の場合、列が存在するか先に確認
  if (data['重要フラグ'] !== undefined && headers.indexOf('重要フラグ') === -1) {
    Logger.log('[updateStressCheck] ❌ 「重要フラグ」列が見つかりません。ヘッダー: ' + JSON.stringify(headers));
    return jsonResponse_({ success: false, error: 'スプレッドシートの「ストレスチェック管理」シートに「重要フラグ」列がありません。シートのヘッダー行に「重要フラグ」列を追加してください。' });
  }

  var targetId = String(data['ID'] || '');
  for (var i = 1; i < allData.length; i++) {
    if (String(allData[i][idCol]) === targetId) {
      var rowNum = i + 1;
      var nowJST = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
      for (var j = 0; j < headers.length; j++) {
        var colName = headers[j];
        if (colName === 'ID' || colName === '作成日') continue;
        if (colName === '更新日') {
          sheet.getRange(rowNum, j + 1).setValue(nowJST);
          continue;
        }
        if (data[colName] !== undefined) {
          sheet.getRange(rowNum, j + 1).setValue(data[colName]);
        }
      }
      Logger.log('[updateStressCheck] 行' + rowNum + ' 更新完了: ID=' + targetId);
      return jsonResponse_({ success: true, updatedRow: rowNum });
    }
  }
  return jsonResponse_({ success: false, error: '該当ID(' + targetId + ')が見つかりません' });
}

// --------------- 担当者状態（一時保存） ---------------

var VALID_MEMBER_STATUSES = ['normal', 'phone', 'toilet', 'smoke', 'clean', 'meal'];

function handleGetMemberStatuses_() {
  var cache = CacheService.getScriptCache();
  var raw = cache.get('memberStatuses');
  var statuses = raw ? JSON.parse(raw) : {};
  return jsonResponse_(statuses);
}

function handleSetMemberStatus_(data) {
  var name = data.name || '';
  var status = data.status || 'normal';
  if (VALID_MEMBER_STATUSES.indexOf(status) === -1) status = 'normal';
  if (!name) return jsonResponse_({ success: false, error: 'name is required' });

  var cache = CacheService.getScriptCache();
  var raw = cache.get('memberStatuses');
  var statuses = raw ? JSON.parse(raw) : {};
  var nowJST = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
  statuses[name] = { status: status, updatedAt: nowJST };
  // CacheService最大6時間（21600秒）
  cache.put('memberStatuses', JSON.stringify(statuses), 21600);
  return jsonResponse_({ success: true, name: name, status: status, updatedAt: nowJST });
}

// --------------- 一言メッセージ（スプレッドシート保存） ---------------

function handleGetMessages_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_MESSAGE);
  if (!sheet) return jsonResponse_([]);

  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return jsonResponse_([]); // ヘッダーのみ

  var messages = [];
  for (var i = 1; i < data.length; i++) {
    var ts = data[i][1];
    if (ts instanceof Date) {
      ts = Utilities.formatDate(ts, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm');
    }
    messages.push({
      id: String(data[i][0]),
      ts: String(ts || ''),
      from: String(data[i][2] || ''),
      to: String(data[i][3] || ''),
      text: String(data[i][4] || '')
    });
  }
  // 新しい順にソートして最新3件
  messages.sort(function(a, b) { return a.ts > b.ts ? -1 : a.ts < b.ts ? 1 : 0; });
  return jsonResponse_(messages.slice(0, 3));
}

function handlePostMessage_(data) {
  var text = (data.text || '').trim();
  if (!text) return jsonResponse_({ success: false, error: 'text is required' });
  var from = data.from || '';
  var to = data.to || '';
  if (!from || !to) return jsonResponse_({ success: false, error: 'from and to are required' });

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_MESSAGE);
  if (!sheet) return jsonResponse_({ success: false, error: 'シートが見つかりません' });

  var id = String(new Date().getTime());
  var nowJST = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm');
  sheet.appendRow([id, nowJST, from, to, text]);

  return jsonResponse_({ success: true, message: { id: id, ts: nowJST, from: from, to: to, text: text } });
}

function handleDeleteMessage_(data) {
  var id = String(data.id || '');
  if (!id) return jsonResponse_({ success: false, error: 'id is required' });

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_MESSAGE);
  if (!sheet) return jsonResponse_({ success: false, error: 'シートが見つかりません' });

  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]) === id) {
      sheet.deleteRow(i + 1);
      return jsonResponse_({ success: true });
    }
  }
  return jsonResponse_({ success: false, error: 'not found' });
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

// 指定カラム名でID採番（クライアント・行政問い合わせ用）
function getNextIdByCol_(sheet, headers, colName) {
  var idCol = headers.indexOf(colName);
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
