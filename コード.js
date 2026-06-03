// ==========================================
// 1. ラベルを指定して削除する機能
// ==========================================
function deleteUnstarredMailsInLabel() {
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
    
    // 検索条件: 指定ラベルがあり、かつスターがついていない
    var searchQuery = "label:" + labelName + " -is:starred";
    
    // 検索条件に一致するスレッドを取得 (一度の実行で最大100件処理する設定)
    var threads = GmailApp.search(searchQuery, 0, 100);
    
    // スレッドが存在する場合、ゴミ箱へ移動
    if (threads.length > 0) {
      for (var i = 0; i < threads.length; i++) {
        threads[i].moveToTrash(); // ゴミ箱へ移動
      }
      console.log("  -> " + threads.length + "件のスレッドをゴミ箱に移動しました。");
    } else {
      console.log("  -> 削除対象のメールはありませんでした。");
    }
  }
}

// ==========================================
// 2. 送信元（From）を指定して削除する機能
// ==========================================
function deleteUnstarredMailsBySender() {
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
    
    // 検索条件: 指定の送信元であり、かつスターがついていない
    var searchQuery = "from:" + senderAddress + " -is:starred";
    
    // 検索条件に一致するスレッドを取得 (最大100件)
    var threads = GmailApp.search(searchQuery, 0, 100);
    
    if (threads.length > 0) {
      for (var i = 0; i < threads.length; i++) {
        threads[i].moveToTrash(); // ゴミ箱へ移動
      }
      console.log("  -> " + threads.length + "件のスレッドをゴミ箱に移動しました。");
    } else {
      console.log("  -> 削除対象のメールはありませんでした。");
    }
  }
}

// ==========================================
// 3. 上記の2つの機能（ラベル削除＆送信元削除）をまとめて実行する関数
// トリガー（定期実行）には、この関数を設定してください。
// ==========================================
function executeAllDeletions() {
  console.log("--- 【開始】ラベルに基づく削除処理 ---");
  deleteUnstarredMailsInLabel();
  
  console.log("--- 【開始】送信元に基づく削除処理 ---");
  deleteUnstarredMailsBySender();
  
  console.log("--- すべての自動削除処理が完了しました ---");
}