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
    try {
      // 処理開始のお知らせを画面の右下に小さくトースト表示
      SpreadsheetApp.getActiveSpreadsheet().toast('メールの削除処理を開始しました。完了までしばらくお待ちください...', '処理中 ⏳', -1);
      
      executeAllDeletions(true); // true = 手動実行のフラグ
      
      SpreadsheetApp.getActiveSpreadsheet().toast('処理が完了しました！', '完了 🎉', 5);
      ui.alert('完了', '処理が完了しました。「手動発火ログ」をご確認ください。', ui.ButtonSet.OK);
    } catch (e) {
      // 万が一エラーで止まった場合
      SpreadsheetApp.getActiveSpreadsheet().toast('処理中にエラーが発生しました。', 'エラー ❌', 10);
      ui.alert('エラー', '処理中にエラーが発生しました:\n' + e.message, ui.ButtonSet.OK);
    }
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
function deleteUnstarredMailsInLabel(isManual, dateCondition) {
  // --- 設定箇所 ---
  // シート名を指定してください
  var sheetName = "削除ラベル設定"; // ラベル名が設定されているシート名
  
  // スプレッドシートから対象シートを取得
  // A列に「ラベル名」、B列に「チェックボックス (TRUE/FALSE)」がある想定
  // getActiveSpreadsheet() で紐づいているスプレッドシートを直接取得します
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  
  // シートが存在しない場合は終了（null のまま getDataRange を呼ぶとエラーになるため）
  if (!sheet) {
    console.log("「" + sheetName + "」シートが見つかりません。");
    return;
  }
  
  var data = sheet.getDataRange().getValues();
  if (!data || data.length === 0) {
    console.log("スプレッドシートにデータがありません。");
    return;
  }
  
  // 期間条件が渡されていない場合（GASエディタでこの関数を単体実行した場合など）は、ここで取得する
  if (!dateCondition) {
    dateCondition = getDateCondition();
    if (dateCondition === null) {
      return; // 期間設定が無効な場合は安全のため中止
    }
  }
  
  // 1行目（見出し）をスキップするため、2行目（row = 1）から順番にチェックしていく
  for (var row = 1; row < data.length; row++) {
    var labelName = data[row][0]; // A列: ラベル名
    var isTarget = data[row][1];  // B列: 実行対象かどうかの判定 (チェックボックスならTRUE/FALSE)
    var individualDays = data[row][2]; // C列: 個別指定日数
    var rowNumber = row + 1;
    
    // ラベル名が空の場合、またはチェックが付いていない(TRUEではない)場合はスキップ
    if (!labelName || isTarget !== true) {
      continue;
    }
    
    // 日付条件を決定（C列の個別設定に数字があればそちらを優先）
    var individualQuery = buildOlderThanQuery(individualDays);
    var localDateCondition = (individualQuery !== "") ? individualQuery : dateCondition;
    
    console.log("【" + labelName + "】の処理を開始します。(行: " + rowNumber + ")");
    
    // 検索条件: 指定ラベルがあり、スターなし、かつ期間条件を追加（ラベル名にスペースが含まれても壊れないよう引用符で囲む）
    var searchQuery = 'label:"' + labelName + '" -is:starred' + localDateCondition;
    
    // 検索条件に一致するスレッドを取得 (一度の実行で最大100件処理する設定)
    var threads = GmailApp.search(searchQuery, 0, 100);
    
    // スレッドが存在する場合、ゴミ箱へ移動
    if (threads.length > 0) {
      GmailApp.moveThreadsToTrash(threads); // 取得したスレッドをまとめてゴミ箱へ移動
      console.log("  -> " + threads.length + "件のスレッドをゴミ箱に移動しました。");
      if (isManual) {
        writeLog("ラベル", labelName, threads.length + "件 削除しました");
        sheet.getRange(rowNumber, 4).setValue(""); // 成功時は備考(D列)をクリア
      }
    } else {
      console.log("  -> 削除対象のメールはありませんでした。");
      if (isManual) {
        writeLog("ラベル", labelName, "削除対象のメールなし");
        sheet.getRange(rowNumber, 4).setValue("ラベル、または削除の対象となるメールが見つかりませんでした");
      }
    }
  }
}

// ==========================================
// 2. 送信元（From）を指定して削除する機能
// ==========================================
function deleteUnstarredMailsBySender(isManual, dateCondition) {
  // --- 設定箇所 ---
  var sheetName = "削除送信元設定"; // 送信元が設定されているシート名
  
  // スプレッドシートから対象シートを取得
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
  
  // 期間条件が渡されていない場合（GASエディタでこの関数を単体実行した場合など）は、ここで取得する
  if (!dateCondition) {
    dateCondition = getDateCondition();
    if (dateCondition === null) {
      return; // 期間設定が無効な場合は安全のため中止
    }
  }
  
  // 1行目（見出し）をスキップ
  for (var row = 1; row < data.length; row++) {
    var senderAddress = data[row][0]; // A列: 送信元アドレス（例: example@test.com）
    var isTarget = data[row][1];      // B列: 実行フラグ
    var individualDays = data[row][2]; // C列: 個別指定日数
    var rowNumber = row + 1;
    
    // 空欄、またはチェックが付いていない場合はスキップ
    if (!senderAddress || isTarget !== true) {
      continue;
    }
    
    // 日付条件を決定（C列の個別設定に数字があればそちらを優先）
    var individualQuery = buildOlderThanQuery(individualDays);
    var localDateCondition = (individualQuery !== "") ? individualQuery : dateCondition;
    
    console.log("【送信元: " + senderAddress + "】の処理を開始します。(行: " + rowNumber + ")");
    
    // 検索条件: 指定の送信元であり、スターなし、かつ期間条件を追加
    var searchQuery = "from:" + senderAddress + " -is:starred" + localDateCondition;
    
    // 検索条件に一致するスレッドを取得 (最大100件)
    var threads = GmailApp.search(searchQuery, 0, 100);
    
    if (threads.length > 0) {
      GmailApp.moveThreadsToTrash(threads); // 取得したスレッドをまとめてゴミ箱へ移動
      console.log("  -> " + threads.length + "件のスレッドをゴミ箱に移動しました。");
      if (isManual) {
        writeLog("送信元", senderAddress, threads.length + "件 削除しました");
        sheet.getRange(rowNumber, 4).setValue(""); // 成功時は備考(D列)をクリア
      }
    } else {
      console.log("  -> 削除対象のメールはありませんでした。");
      if (isManual) {
        writeLog("送信元", senderAddress, "削除対象のメールなし");
        sheet.getRange(rowNumber, 4).setValue("送信元メールアドレス、または削除の対象となるメールが見つかりませんでした");
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

  // 期間条件は最初に1回だけ取得し、各削除処理で使い回す
  var dateCondition = getDateCondition();
  // シートが無い・未設定などで null が返ってきた場合は、安全のため全処理を中止
  if (dateCondition === null) {
    console.log("--- 期間条件が取得できなかったため、処理を中止しました ---");
    return;
  }

  console.log("--- 【開始】ラベルに基づく削除処理 ---");
  deleteUnstarredMailsInLabel(manualFlag, dateCondition);
  
  console.log("--- 【開始】送信元に基づく削除処理 ---");
  deleteUnstarredMailsBySender(manualFlag, dateCondition);
  
  console.log("--- すべての自動削除処理が完了しました ---");
}

// ==========================================
// 共通処理: 個別日数(数値)から older_than 検索クエリを組み立てる
// 数値でない・空・0以下の場合は空文字を返す（未設定扱い）
// ==========================================
function buildOlderThanQuery(days) {
  if (days === "" || days === null || days === undefined) {
    return "";
  }
  if (isNaN(days) || String(days).trim() === "") {
    return "";
  }
  var numDays = parseInt(days, 10);
  // 0や負の数は「未設定」と同じ扱いにし、全件削除などの暴走を防ぐ
  if (isNaN(numDays) || numDays <= 0) {
    return "";
  }
  return " older_than:" + numDays + "d";
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
  
  // 現在の日時を取得して追記（2行目に挿入して最新を上に）
  var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss");
  sheet.insertRowAfter(1); // 1行目（見出し）の直後に新しい行を挿入
  sheet.getRange(2, 1, 1, 4).setValues([[timestamp, targetType, targetName, resultText]]); // 2行目にデータを書き込む
}

// ==========================================
// 共通処理: 「その他削除設定」シートから期間条件を取得する
// ==========================================
function getDateCondition() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("その他削除設定");
  
  // 1. シートが存在しない場合は、「その他削除設定」シートを新しく自動作成してアラートを出す
  if (!sheet) {
    sheet = ss.insertSheet("その他削除設定");
    
    // 見出しなどを自動セット
    sheet.getRange("A2").setValue("何日前のものを削除するか");
    
    var statusRange = sheet.getRange("A4:C4");
    statusRange.merge(); // A4からC4を結合
    statusRange.setValue("【エラー】\nB2 (〇日前) の設定がありません。\n安全のため処理を停止しました。");
    statusRange.setBackground("#ffcccc"); // ピンク色に設定
    statusRange.setVerticalAlignment("middle");
    statusRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW); // セル内で折り返さず表示する
    sheet.setColumnWidth(1, 200); // A列の幅を少し広げる
    
    console.log("「その他削除設定」シートが存在しなかったため作成し、処理を停止しました。");
    return null; 
  }
  
  // B2に「〇日前のもの」が入力されていると想定
  var daysOld = sheet.getRange("B2").getValue();
  
  var dateQuery = "";
  var settingText = "";
  
  // 自動計算で〇日前 (B2) が数値で指定されている場合
  dateQuery = buildOlderThanQuery(daysOld);
  if (dateQuery !== "") {
    settingText = parseInt(daysOld, 10) + " 日より前のメールを削除";
  }
  
  var statusRange = sheet.getRange("A4:C4");
  statusRange.merge(); // 実行時にA4:C4の結合状態を確保
  statusRange.setVerticalAlignment("middle");
  statusRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW); // セル内で折り返さず表示する
  
  // 未入力の場合はエラー（全件削除を防ぐ安全対策）にするか判定
  if (dateQuery === "") {
    statusRange.setValue("🚨エラー🚨\nB2 (〇日前) の設定がありません。\n安全のため処理を停止しました。");
    statusRange.setBackground("#ffcccc"); // ピンク色
    return null;
  } else {
    // 正常に設定されている場合
    statusRange.setValue("♻️「" + settingText + "」の設定で運用中");
    statusRange.setBackground("#ccffcc"); // 薄い緑色
  }
  
  return dateQuery;
}