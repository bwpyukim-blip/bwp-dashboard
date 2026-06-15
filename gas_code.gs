// ════════════════════════════════════════════════════════════
// 活動報告ダッシュボード — Google Apps Script
// ════════════════════════════════════════════════════════════
//
// 【設定手順】
//  1. Googleスプレッドシートを開く
//  2. メニュー「拡張機能」→「Apps Script」をクリック
//  3. このコードを貼り付けて保存（Ctrl+S）
//  4. 「デプロイ」→「新しいデプロイ」→「ウェブアプリ」
//     ・実行ユーザー: 自分
//     ・アクセスできるユーザー: 全員（匿名を含む）
//  5. 表示されたURLを index.html の GAS_URL に貼る
//
// ════════════════════════════════════════════════════════════

// ★ シート名を実際のシート名に合わせて変更してください
var SHEET_NAME = 'フォームの回答 1';

// ─────────────────────────────────────────────────────────
// スプレッドシートの列定義（ヘッダー行と対応）
// ─────────────────────────────────────────────────────────
// 列番号（1始まり）とJSONキーのマッピング
var COLUMN_MAP = {
  3:  'date',          // Q1. 活動日
  4:  'university',    // Q2. 所属大学
  5:  'team',          // Q3. 所属チーム名
  6:  'country',       // 支援国
  7:  'project',       // Q4. 所属プロジェクト
  8:  'name',          // Q5. 報告担当者氏名
  9:  'taskTotal',     // 今回のタスク総数
  10: 'taskDone',      // 今回のタスク達成数
  11: 'plan',          // P：Plan（今日の目標）
  12: 'do1',           // D：Do その１（活動内容）
  13: 'do2',           // D：Do その２（一言）
  14: 'check1',        // C：Check その１（よかった点）
  15: 'check2',        // C：Check その２（課題）
  16: 'action',        // A：Action（次回への改善）
  17: 'memo',          // 運営への共有・相談
};

// ─────────────────────────────────────────────────────────
// メイン: GETリクエストでJSONを返す
// ─────────────────────────────────────────────────────────
function doGet(e) {
  try {
    var data = getReportData();

    // クエリパラメータによるフィルタリング
    if (e && e.parameter) {
      var p = e.parameter;

      // ?name=田中 → 氏名で絞り込み
      if (p.name) {
        data = data.filter(function(r) {
          return r.name === p.name;
        });
      }

      // ?team=〇〇チーム → チーム名で絞り込み
      if (p.team) {
        data = data.filter(function(r) {
          return r.team === p.team;
        });
      }

      // ?project=〇〇 → プロジェクト名で絞り込み
      if (p.project) {
        data = data.filter(function(r) {
          return r.project === p.project;
        });
      }

      // ?days=7 → 直近N日分
      if (p.days) {
        var limitMs = parseInt(p.days) * 86400000;
        var now     = Date.now();
        data = data.filter(function(r) {
          var d = new Date(r.date);
          return !isNaN(d) && (now - d.getTime()) <= limitMs;
        });
      }

      // ?limit=20 → 件数上限
      if (p.limit) {
        data = data.slice(0, parseInt(p.limit));
      }
    }

    var result = {
      status:    'ok',
      count:     data.length,
      updatedAt: new Date().toISOString(),
      data:      data,
    };

    return buildResponse(result);

  } catch (err) {
    return buildResponse({ status: 'error', message: err.message });
  }
}

// ─────────────────────────────────────────────────────────
// スプレッドシートからデータを取得して整形
// ─────────────────────────────────────────────────────────
function getReportData() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    throw new Error('シート「' + SHEET_NAME + '」が見つかりません。SHEET_NAMEを確認してください。');
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return []; // データ行なし

  var lastCol = sheet.getLastColumn();
  // 2行目以降を全取得
  var values  = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  var records = [];

  values.forEach(function(row) {
    // 空行スキップ（名前と活動日が両方空の行）
    var nameVal = row[7]  || ''; // 列8 (0始まりで7)
    var dateVal = row[2]  || ''; // 列3 (0始まりで2)
    if (!nameVal && !dateVal) return;

    var record = {};

    // COLUMN_MAPに従って列を取得
    Object.keys(COLUMN_MAP).forEach(function(colNum) {
      var key = COLUMN_MAP[colNum];
      var val = row[parseInt(colNum) - 1]; // 0始まりに変換

      // 日付型は文字列に変換
      if (val instanceof Date) {
        val = formatDate(val);
      }
      // 数値はそのまま
      record[key] = (val === null || val === undefined) ? '' : val;
    });

    // タスク達成率を計算（数値の場合のみ）
    var total = parseInt(record.taskTotal) || 0;
    var done  = parseInt(record.taskDone)  || 0;
    record.taskRate = total > 0 ? Math.round((done / total) * 100) : null;

    records.push(record);
  });

  // 活動日の新しい順に並び替え
  records.sort(function(a, b) {
    return new Date(b.date) - new Date(a.date);
  });

  return records;
}

// ─────────────────────────────────────────────────────────
// ユーティリティ
// ─────────────────────────────────────────────────────────
function formatDate(d) {
  var y  = d.getFullYear();
  var m  = String(d.getMonth() + 1).padStart('0', 2);
  var dd = String(d.getDate()).padStart('0', 2);
  return y + '-' + m + '-' + dd;
}

function buildResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─────────────────────────────────────────────────────────
// デバッグ用: スクリプトエディタで直接実行して確認できる
// ─────────────────────────────────────────────────────────
function debugRun() {
  var data = getReportData();
  Logger.log('取得件数: ' + data.length);
  if (data.length > 0) {
    Logger.log('最新1件: ' + JSON.stringify(data[0], null, 2));
  }
}
