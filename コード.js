function deleteUnstarredMailsInLabel() {
  // --- 設定箇所 ---
  // スプレッドシートのIDと、シート名を指定してください
  var spreadsheetId = "スプレッドシートのIDをここに入力"; // 例: "1abcdefghijklmnopqrstuvwxyz..."
  var sheetName = "シート1"; // ラベル名が設定されているシート名
  
  // スプレッドシートから全データを取得
  // A列に「ラベル名」、B列に「チェックボックス (TRUE/FALSE)」がある想定
  var sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);
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