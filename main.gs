function myFunction() {
  // ログにステップを記録しながら実行
  try {
    Logger.log("myFunction 開始");
    // 進捗インジケーターを表示してから計算を実行
    showProgressIndicator();
    Logger.log("myFunction 終了");
  } catch (error) {
    Logger.log("myFunction でエラー発生: " + error.toString());
    SpreadsheetApp.getUi().alert("エラーが発生しました: " + error.toString());
  }
}

// 進捗インジケーターを表示する関数
function showProgressIndicator() {
  Logger.log("showProgressIndicator 開始");
  try {
    // HTMLダイアログを作成
    var html = HtmlService.createHtmlOutput(`
      <html>
        <head>
          <style>
            body {
              font-family: Arial, sans-serif;
              text-align: center;
              padding: 20px;
            }
            .spinner {
              margin: 20px auto;
              border: 6px solid #f3f3f3;
              border-top: 6px solid #3498db;
              border-radius: 50%;
              width: 40px;
              height: 40px;
              animation: spin 2s linear infinite;
            }
            @keyframes spin {
              0% { transform: rotate(0deg); }
              100% { transform: rotate(360deg); }
            }
            .progress-text {
              margin-top: 15px;
              font-size: 14px;
            }
          </style>
        </head>
        <body>
          <h3>経路最適化を実行中...</h3>
          <div class="spinner"></div>
          <div class="progress-text" id="status">データをチェック中...</div>
          
          <script>
            // 進捗状況のテキストを更新する関数
            function updateStatus(text) {
              document.getElementById('status').textContent = text;
            }
            
            // 計算を開始
            google.script.run
              .withSuccessHandler(function(result) {
                // 成功時は自動的にダイアログを閉じる
                updateStatus("処理が完了しました: " + result);
                // 閉じる前に少し待機して成功メッセージを表示
                setTimeout(function() {
                  google.script.host.close();
                }, 1000);
              })
              .withFailureHandler(function(error) {
                // エラー時はメッセージを表示
                updateStatus('エラーが発生しました: ' + error);
                setTimeout(function() {
                  google.script.host.close();
                }, 3000);
              })
              .calculateOptimizedRouteWithProgress();
          </script>
        </body>
      </html>
    `)
    .setWidth(300)
    .setHeight(200);
    
    // モーダルダイアログとして表示
    SpreadsheetApp.getUi().showModalDialog(html, "ルート計算中");
    Logger.log("ダイアログを表示しました");
  } catch (error) {
    Logger.log("showProgressIndicator でエラー発生: " + error.toString());
    SpreadsheetApp.getUi().alert("ダイアログ表示中にエラーが発生しました: " + error.toString());
  }
}

// 進捗状況を更新しながら計算を実行する関数
function calculateOptimizedRouteWithProgress() {
  Logger.log("calculateOptimizedRouteWithProgress 開始");
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = sheet.getDataRange().getValues();
    var ui = SpreadsheetApp.getUi();
    
    // エラー検出用の変数
    var errorMessages = [];
    var firstCount = 0;
    var lastCount = 0;
    
    // 進捗状況1: データの検証
    Logger.log("データ検証開始");
    
    // 2行目以降（ヘッダーは1行目）をチェック
    // ※列番号: D=3, E=4, F=5, G=6 (0-indexed)
    for (var i = 1; i < data.length; i++) {
      var rowIndex = i + 1; // シート上の行番号
      var address = data[i][3];    // 列D：住所
      var sendTarget = data[i][4]; // 列E：送迎対象
      var firstPickup = data[i][5]; // 列F：最初に送迎
      var lastPickup = data[i][6];  // 列G：最後に送迎
      
      Logger.log("行 " + rowIndex + " のチェック: 送迎対象=" + sendTarget + ", 最初=" + firstPickup + ", 最後=" + lastPickup);
      
      // 送迎対象(E列)が未チェックなのに、F列またはG列がチェックされている場合
      if (!sendTarget && (firstPickup || lastPickup)) {
        errorMessages.push("行" + rowIndex + ": 送迎対象(E列)がチェックされていないのに、最初または最後に送迎がチェックされています。");
      }
      // 同じ行で両方チェックされている場合
      if (firstPickup && lastPickup) {
        errorMessages.push("行" + rowIndex + ": 同じ行で最初と最後の送迎の両方がチェックされています。");
      }
      if (firstPickup) {
        firstCount++;
      }
      if (lastPickup) {
        lastCount++;
      }
    }
    
    // 「最初に送迎」と「最後に送迎」は各1件以内にする
    if (firstCount > 1) {
      errorMessages.push("「最初に送迎」のチェックが複数行に設定されています。");
    }
    if (lastCount > 1) {
      errorMessages.push("「最後に送迎」のチェックが複数行に設定されています。");
    }
    
    // エラーがあればポップアップで通知して処理を中断
    if (errorMessages.length > 0) {
      Logger.log("バリデーションエラー: " + errorMessages.join(", "));
      ui.alert("エラー", errorMessages.join("\n"), ui.ButtonSet.OK);
      return "バリデーションエラー";
    }
    
    // 進捗状況2: データのグループ分け
    Logger.log("データのグループ分け開始");
    
    // APIキーの設定
    var apiKey = PropertiesService.getScriptProperties().getProperty('MAPS_API_KEY');
    
    // 出発地（かつ帰着地）はA2セルから取得
    var origin = sheet.getRange("A2").getValue();
    Logger.log("出発地: " + origin);
    if (!origin) {
      Logger.log("出発地が設定されていません");
      return "出発地(A2)が設定されていません";
    }
    
    // 出発時刻はB2セルから取得（未入力の場合は現在時刻）
    var departureTimeCell = sheet.getRange("B2").getValue();
    var departureTime = departureTimeCell ? new Date(departureTimeCell) : new Date();
    Logger.log("出発時刻: " + departureTime);
    
    // 目的地のグループ分け
    var firstGroup = [];
    var middleGroup = [];
    var lastGroup = [];
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][4] === true && data[i][3]) {
        var dest = { address: data[i][3], row: i + 1 };
        if (data[i][5] === true) {
          firstGroup.push(dest);
          Logger.log("最初に送迎: " + dest.address);
        } else if (data[i][6] === true) {
          lastGroup.push(dest);
          Logger.log("最後に送迎: " + dest.address);
        } else {
          middleGroup.push(dest);
          Logger.log("中間送迎: " + dest.address);
        }
      }
    }
    
    var totalStops = firstGroup.length + middleGroup.length + lastGroup.length;
    Logger.log("総訪問先数: " + totalStops);
    if (totalStops === 0) {
      Logger.log("送迎対象の訪問先がありません");
      return "送迎対象の訪問先がありません";
    }
    
    // 最大ウェイポイント数の制限（例：10件）
    var maxWaypoints = 10;
    if (totalStops > maxWaypoints) {
      Logger.log("訪問先数が制限(" + maxWaypoints + ")を超えています: " + totalStops);
      var fixedCount = firstGroup.length + lastGroup.length;
      if (fixedCount > maxWaypoints) {
        firstGroup = firstGroup.slice(0, maxWaypoints);
        middleGroup = [];
        lastGroup = [];
      } else {
        var allowedMiddle = maxWaypoints - fixedCount;
        middleGroup = middleGroup.slice(0, allowedMiddle);
      }
    }
    
    // 送迎順設定: 固定グループがある場合は最適化なし、ない場合は中間グループのみで最適化を利用
    var usingOptimization = false;
    var routeStops = [];
    if (firstGroup.length > 0 || lastGroup.length > 0) {
      routeStops = firstGroup.concat(middleGroup, lastGroup);
      Logger.log("固定順序使用: 最初=" + firstGroup.length + ", 中間=" + middleGroup.length + ", 最後=" + lastGroup.length);
    } else {
      if (middleGroup.length > 0) {
        usingOptimization = true;
        routeStops = middleGroup;
        Logger.log("最適化使用: 中間地点数=" + middleGroup.length);
      }
    }
    
    // 進捗状況3: Google Maps APIリクエスト準備
    Logger.log("Google Maps APIリクエスト準備");
    
    // ウェイポイント文字列作成
    var addresses = routeStops.map(function(dest) { return dest.address; });
    var waypoints;
    if (usingOptimization) {
      waypoints = "optimize:true|" + addresses.join("|");
    } else {
      waypoints = addresses.join("|");
    }
    
    // 各パラメータを個別にエンコードしてリクエストURLを作成
    var originParam = "origin=" + encodeURIComponent(origin);
    var destinationParam = "destination=" + encodeURIComponent(origin);
    var waypointsParam = "waypoints=" + encodeURIComponent(waypoints);
    var departureTimeParam = "departure_time=" + Math.floor(departureTime.getTime() / 1000);
    var keyParam = "key=" + apiKey;
    
    var directionsUrl = "https://maps.googleapis.com/maps/api/directions/json?" +
                        originParam + "&" + destinationParam + "&" +
                        waypointsParam + "&" + departureTimeParam + "&" +
                        keyParam;
    
    // URLのログ出力（APIキーは一部マスク）
    var logUrl = directionsUrl.replace(apiKey, apiKey.substring(0, 5) + "...");
    Logger.log("Google Maps API URL: " + logUrl);
    
    // 進捗状況4: Google Maps APIリクエスト実行
    Logger.log("Google Maps APIリクエスト実行");
    
    try {
      var response = UrlFetchApp.fetch(directionsUrl);
      var responseCode = response.getResponseCode();
      Logger.log("API レスポンスコード: " + responseCode);
      
      if (responseCode !== 200) {
        Logger.log("API エラー: HTTPステータス " + responseCode);
        return "Google Maps API エラー: HTTPステータス " + responseCode;
      }
      
      var json = JSON.parse(response.getContentText());
      Logger.log("API ステータス: " + json.status);
      
      if (json.status !== "OK") {
        Logger.log("API ステータスエラー: " + json.status + (json.error_message ? " - " + json.error_message : ""));
        return "Google Maps API エラー: " + json.status + (json.error_message ? " - " + json.error_message : "");
      }
      
      if (!json.routes || json.routes.length === 0) {
        Logger.log("APIレスポンスにルートがありません");
        return "ルートが見つかりませんでした";
      }
      
      if (!json.routes[0].legs || json.routes[0].legs.length === 0) {
        Logger.log("APIレスポンスに経路の詳細がありません");
        return "経路の詳細が取得できませんでした";
      }
    } catch (error) {
      Logger.log("API呼び出し中にエラー: " + error.toString());
      return "Google Maps API呼び出し中にエラー: " + error.toString();
    }
    
    // 進捗状況5: 結果処理
    Logger.log("APIレスポンス処理開始");
    
    // 最適化を利用している場合、APIからの最適化結果に沿って並び替え
    if (usingOptimization) {
      var optimizedOrder = json.routes[0].waypoint_order || [];
      Logger.log("最適化順序: " + JSON.stringify(optimizedOrder));
      var orderedStops = [];
      optimizedOrder.forEach(function(idx) {
        orderedStops.push(routeStops[idx]);
      });
      routeStops = orderedStops;
    }
    
    // 進捗状況6: スプレッドシートに結果を書き込み
    Logger.log("シートへの書き込み開始");
    
    // 出発時刻から各legの所要時間を累積して到着予定時刻を計算、シートへ出力
    var currentTime = new Date(departureTime);
    for (var i = 0; i < routeStops.length; i++) {
      var leg = json.routes[0].legs[i];
      if (leg) {
        var durationMs = leg.duration.value * 1000;
        currentTime = new Date(currentTime.getTime() + durationMs);
        var arrivalTimeStr = Utilities.formatDate(currentTime, Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm");
        
        Logger.log("訪問先 " + (i + 1) + ": 行=" + routeStops[i].row + ", 到着=" + arrivalTimeStr + ", 所要時間=" + leg.duration.text);
        
        sheet.getRange(routeStops[i].row, 8).setValue(i + 1);       // 訪問順 (H列)
        sheet.getRange(routeStops[i].row, 9).setValue(arrivalTimeStr); // 到着予定時刻 (I列)
        sheet.getRange(routeStops[i].row, 10).setValue(leg.duration.text); // 所要時間 (J列)
      } else {
        Logger.log("訪問先 " + (i + 1) + " の経路情報がありません");
        sheet.getRange(routeStops[i].row, 8).setValue(i + 1);
        sheet.getRange(routeStops[i].row, 9).setValue("エラー");
        sheet.getRange(routeStops[i].row, 10).setValue("エラー");
      }
    }
    
    // 進捗状況7: 完了
    Logger.log("処理完了、Google Maps URL作成");
    
    // Google Mapsでルート確認用のURL作成（出発地 → 各目的地 → 出発地）
    var mapsUrl = "https://www.google.com/maps/dir/" +
                  encodeURIComponent(origin) + "/" +
                  routeStops.map(function(dest) { return encodeURIComponent(dest.address); }).join("/") +
                  "/" + encodeURIComponent(origin);
                  
    // ユーザーがボタンをクリックしたタイミングで新規タブをオープン
    openGoogleMaps(mapsUrl);
    
    return "処理完了";
  } catch (error) {
    Logger.log("calculateOptimizedRouteWithProgress でエラー発生: " + error.toString() + "\n" + error.stack);
    return "エラーが発生しました: " + error.toString();
  }
}

// 進捗状況をクライアント側に通知する関数
function updateProgressStatus(statusText) {
  Logger.log("進捗状況更新: " + statusText);
  // この関数は直接呼び出されても何も行いません
  // 本来はクライアント側のHTMLから呼び出されるコールバック用の関数です
  return statusText;
}

function openGoogleMaps(url) {
  Logger.log("openGoogleMaps 開始");
  try {
    // ダイアログ内に「Google Mapsを開く」ボタンを配置し、ユーザー操作で新規タブをオープン
    var html = HtmlService.createHtmlOutput(
      '<html><head><base target="_blank"></head><body>' +
      '<p>下のボタンをクリックしてGoogle Mapsのルートを開いてください。</p>' +
      '<button onclick="openMap()">Google Mapsを開く</button>' +
      '<script>' +
      'function openMap() {' +
      '  window.open("' + url + '", "_blank");' +
      '  google.script.host.close();' +
      '}' +
      '</script>' +
      '</body></html>'
    ).setWidth(300).setHeight(150);
    
    SpreadsheetApp.getUi().showModalDialog(html, "Google Mapsを開く");
    Logger.log("Google Maps ダイアログを表示しました");
  } catch (error) {
    Logger.log("openGoogleMaps でエラー発生: " + error.toString());
    SpreadsheetApp.getUi().alert("Google Maps URLを開く際にエラーが発生しました: " + error.toString());
  }
}