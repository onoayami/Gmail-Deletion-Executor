// ==========================================
// 0. スプレッドシートを開いた時の処理 (カスタムメニューの追加)
// ==========================================
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // メニューバーの「ヘルプ」の横に独自メニューを追加
  ui.createMenu('【設定手動アップデート 🗑️💨】')
    .addItem('設定を手動で反映させる', 'executeAllDeletionsManually')
    .addSeparator()
    .addItem('一時全停止（全ての設定をOFFにする）', 'emergencyStop')
    .addToUi();
}

// --- メニューからの手動実行用 ---
function executeAllDeletionsManually() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('手動実行', '今すぐメールの自動削除処理を開始しますか？\n※結果は「手動発火ログ」タブに記録されます。', ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES) {
    executeAllDeletions(true); // true = 手動実行のフラグ
    ui.alert('完了', '処理が完了しました。「手動発火ログ」をご確認ください。', ui.ButtonSet.OK);
  }
}

// --- 処理停止用（チェックボックスをすべて外す） ---
function emergencyStop() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('処理停止', '自動削除を停止するため、すべての設定のチェックボックスをOFF(外す)にしますか？', ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES) {
    var labelSheet = ss.getSheetByName("削除ラベル設定");
    if (labelSheet && labelSheet.getLastRow() > 1) {
      labelSheet.getRange(2, 2, labelSheet.getLastRow() - 1, 1).uncheck();
    }
    var senderSheet = ss.getSheetByName("削除送信元設定");
    if (senderSheet && senderSheet.getLastRow() > 1) {
      senderSheet.getRange(2, 2, senderSheet.getLastRow() - 1, 1).uncheck();
    }
    ui.alert('停止完了', 'すべての設定チェックを外し、今後の削除処理を停止状態にしました。', ui.ButtonSet.OK);
  }
}

// ==========================================
// 1. ラベルを指定して削除する機能
// ==========================================
function deleteUnstarredMailsInLabel(isManual) {
  // --- 設定箇所 ---
  // シート名を指定してください
  var sheetName = "削除ラベル設定"; // ラベル名が設定されているシート名
  
  // スプレッドシートから全データを取得
  // A列に「ラベル名」、B列に「チェックボックス (TRUE/FALSE)」がある想定
  // getActiveSpreadsheet() で紐づいているスプレッドシートを直接取得します
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var data = sheet.getDataRange().getValues();
  
  if (!data || data.length === 0) {
    console.log("スプレッドシートにデータがありません。");
    return;
  }
  
  // 期間指定の検索条件を取得（削除設定シートから）
  var dateCondition = getDateCondition();
  // もし【削除設定】シートが見つからずエラー(null)が返ってきた場合は処理を終了
  if (dateCondition === null) {
    return;
  }
  
  // 1行目（見出し）をスキップするため、2行目（row = 1）から順番にチェックしていく
  for (var row = 1; row < data.length; row++) {
    var labelName = data[row][0]; // A列: ラベル名
    var isTarget = data[row][1];  // B列: 実行対象かどうかの判定 (チェックボックスならTRUE/FALSE)
    var rowNumber = row + 1;
    
    // ラベル名が空の場合、またはチェックが付いていない(TRUEではない)場合はスキップ
    if (!labelName || isTarget !== true) {
      continue;
    }
    
    console.log("【" + labelName + "】の処理を開始します。(行: " + rowNumber + ")");
    
    // 検索条件: 指定ラベルがあり、スターなし、かつ期間条件を追加
    var searchQuery = "label:" + labelName + " -is:starred" + dateCondition;
    
    // 検索条件に一致するスレッドを取得 (一度の実行で最大100件処理する設定)
    var threads = GmailApp.search(searchQuery, 0, 100);
    
    // スレッドが存在する場合、ゴミ箱へ移動
    if (threads.length > 0) {
      for (var i = 0; i < threads.length; i++) {
        threads[i].moveToTrash(); // ゴミ箱へ移動
      }
      console.log("  -> " + threads.length + "件のスレッドをゴミ箱に移動しました。");
      if (isManual) {
        writeLog("ラベル", labelName, threads.length + "件 削除しました");
        sheet.getRange(rowNumber, 3).setValue(""); // 成功時は備考(C列)をクリア
      }
    } else {
      console.log("  -> 削除対象のメールはありませんでした。");
      if (isManual) {
        writeLog("ラベル", labelName, "削除対象のメールなし");
        sheet.getRange(rowNumber, 3).setValue("該当するラベルが見つかりませんでした");
      }
    }
  }
}

// ==========================================
// 2. 送信元（From）を指定して削除する機能
// ==========================================
function deleteUnstarredMailsBySender(isManual) {
  // --- 設定箇所 ---
  var sheetName = "削除送信元設定"; // 送信元が設定されているシート名
  
  // スプレッドシートから全データを取得
  // A列に「送信元アドレス」、B列に「チェックボックス (TRUE/FALSE)」がある想定
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  
  // シートが存在しない場合は終了
  if (!sheet) {
    console.log("「" + sheetName + "」シートが見つかりません。");
    return;
  }
  
  var data = sheet.getDataRange().getValues();
  if (!data || data.length === 0) {
    console.log("スプレッドシートにデータがありません。");
    return;
  }
  
  // 期間指定の検索条件を取得（削除設定シートから）
  var dateCondition = getDateCondition();
  // もし【削除設定】シートが見つからずエラー(null)が返ってきた場合は処理を終了
  if (dateCondition === null) {
    return;
  }
  
  // 1行目（見出し）をスキップ
  for (var row = 1; row < data.length; row++) {
    var senderAddress = data[row][0]; // A列: 送信元アドレス（例: example@test.com）
    var isTarget = data[row][1];      // B列: 実行フラグ
    var rowNumber = row + 1;
    
    // 空欄、またはチェックが付いていない場合はスキップ
    if (!senderAddress || isTarget !== true) {
      continue;
    }
    
    console.log("【送信元: " + senderAddress + "】の処理を開始します。(行: " + rowNumber + ")");
    
    // 検索条件: 指定の送信元であり、スターなし、かつ期間条件を追加
    var searchQuery = "from:" + senderAddress + " -is:starred" + dateCondition;
    
    // 検索条件に一致するスレッドを取得 (最大100件)
    var threads = GmailApp.search(searchQuery, 0, 100);
    
    if (threads.length > 0) {
      for (var i = 0; i < threads.length; i++) {
        threads[i].moveToTrash(); // ゴミ箱へ移動
      }
      console.log("  -> " + threads.length + "件のスレッドをゴミ箱に移動しました。");
      if (isManual) {
        writeLog("送信元", senderAddress, threads.length + "件 削除しました");
        sheet.getRange(rowNumber, 3).setValue(""); // 成功時は備考(C列)をクリア
      }
    } else {
      console.log("  -> 削除対象のメールはありませんでした。");
      if (isManual) {
        writeLog("送信元", senderAddress, "削除対象のメールなし");
        sheet.getRange(rowNumber, 3).setValue("該当する送信元が見つかりませんでした");
      }
    }
  }
}

// ==========================================
// 3. 上記の2つの機能（ラベル削除＆送信元削除）をまとめて実行する関数
// トリガー（定期実行）には、この関数を設定してください。
// ==========================================
function executeAllDeletions(isManual) {
  // 手動実行ボタンから呼ばれた場合は isManual === true。
  // 定期トリガーから呼ばれた場合はイベントオブジェクトが入るため手動ではないと判定。
  var manualFlag = (isManual === true);

  console.log("--- 【開始】ラベルに基づく削除処理 ---");
  deleteUnstarredMailsInLabel(manualFlag);
  
  console.log("--- 【開始】送信元に基づく削除処理 ---");
  deleteUnstarredMailsBySender(manualFlag);
  
  console.log("--- すべての自動削除処理が完了しました ---");
}

// ==========================================
// 4. 手動発火ログを書き込む機能
// ==========================================
function writeLog(targetType, targetName, resultText) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = "手動発火ログ";
  var sheet = ss.getSheetByName(sheetName);
  
  // シートが無ければ自動作成する
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(["実行日時", "種類", "対象名", "処理結果"]);
    sheet.getRange("A1:D1").setBackground("#eeeeee").setFontWeight("bold");
    sheet.setColumnWidth(1, 150);
    sheet.setColumnWidth(3, 150);
  }
  
  // 現在の日時を取得して追記
  var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss");
  sheet.appendRow([timestamp, targetType, targetName, resultText]);
}

// ==========================================
// 共通処理: 「【削除設定】」シートから期間条件を取得する
// ==========================================
function getDateCondition() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("【削除設定】");
  
  // 1. シートが存在しない場合は、「【削除設定】」シートを新しく自動作成してアラートを出す
  if (!sheet) {
    sheet = ss.insertSheet("【削除設定】");
    
    // 見出しなどを自動セット
    sheet.getRange("A3").setValue("何日前のものを削除するか");
    sheet.getRange("A4").setValue("指定日以前を削除するか");
    
    var statusRange = sheet.getRange("A1");
    statusRange.setValue("【エラー】\nB3 (〇日前) または B4 (指定日) の設定がありません。\n安全のため処理を停止しました。");
    statusRange.setBackground("#ffcccc"); // ピンク色に設定
    statusRange.setVerticalAlignment("middle");
    statusRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW); // セル内で折り返さず表示する
    sheet.setColumnWidth(1, 200); // A列の幅を少し広げる
    
    console.log("「【削除設定】」シートが存在しなかったため作成し、処理を停止しました。");
    return null; 
  }
  
  // B3に「〇日前のもの」、B4に「yyyy/mm/dd以前」が入力されていると想定
  var daysOld = sheet.getRange("B3").getValue();
  var specificDate = sheet.getRange("B4").getValue();
  
  var dateQuery = "";
  var settingText = "";
  
  // 1. yyyy/mm/dd (B4) が指定されている場合（こちらを優先）
  if (specificDate) {
    if (specificDate instanceof Date) {
      var y = specificDate.getFullYear();
      var m = ("0" + (specificDate.getMonth() + 1)).slice(-2);
      var d = ("0" + specificDate.getDate()).slice(-2);
      dateQuery = " before:" + y + "/" + m + "/" + d;
      settingText = y + "年" + m + "月" + d + "日 以前のメールを削除";
    } else {
      dateQuery = " before:" + String(specificDate).replace(/-/g, '/');
      settingText = String(specificDate) + " 以前のメールを削除";
    }
  } 
  // 2. 自動計算で〇日前 (B3) が数値で指定されている場合
  else if (daysOld && !isNaN(daysOld) && String(daysOld).trim() !== "") {
    dateQuery = " older_than:" + parseInt(daysOld, 10) + "d";
    settingText = parseInt(daysOld, 10) + " 日より前のメールを削除";
  }
  
  var statusRange = sheet.getRange("A1");
  statusRange.setVerticalAlignment("middle");
  statusRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW); // セル内で折り返さず表示する
  
  // どちらも未入力の場合はエラー（全件削除を防ぐ安全対策）にするか判定
  if (dateQuery === "") {
    statusRange.setValue("【エラー】\nB3 (〇日前) または B4 (指定日) の設定がありません。\n安全のため処理を停止しました。");
    statusRange.setBackground("#ffcccc"); // ピンク色
    return null;
  } else {
    // 正常に設定されている場合
    statusRange.setValue("✅ 以下の設定で運用中（" + settingText + "）");
    statusRange.setBackground("#ccffcc"); // 薄い緑色
  }
  
  return dateQuery;
}