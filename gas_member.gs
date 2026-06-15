// ════════════════════════════════════════════════════════════
// 活動報告ダッシュボード — Google Apps Script（メンバーフォーム用）
// ════════════════════════════════════════════════════════════
//
// 【設定手順】
//  1. メンバーフォームと連携したGoogleスプレッドシートを開く
//  2. メニュー「拡張機能」→「Apps Script」をクリック
//  3. このコードを貼り付けて保存（Ctrl+S）
//  4. 「デプロイ」→「新しいデプロイ」→「ウェブアプリ」
//     ・実行ユーザー: 自分
//     ・アクセスできるユーザー: 全員（匿名を含む）
//  5. 表示されたURLを index.html / university.html の
//     GAS_MEMBER に貼る
//
// ════════════════════════════════════════════════════════════

// ★ シート名を実際の名前に変更してください
var SHEET_NAME = 'フォームの回答 1';

// ─────────────────────────────────────────────────────────
// 列定義（ヘッダー行と対応）
// A=1始まりで左から順番に番号を振っています
// ─────────────────────────────────────────────────────────
// 列番号: JSONキー: 説明
var COLUMN_MAP = {
  1:  'timestamp',    // タイムスタンプ
  2:  'date',         // Q1. 活動日
  3:  'university',   // Q2. 所属大学
  4:  'grade',        // Q3. 学年
  5:  'country',      // Q4. 支援国（所属チーム国）
  6:  'project',      // Q3. 所属プロジェクトチーム名
  7:  'meeting',      // Q4. 今週のミーティング参加状況
  8:  'tasks',        // Q5. 今週参加した活動（タスク）
  9:  'taskDone',     // 今週完了したタスク数
  10: 'satisfaction', // 今週の活動満足度
  11: 'motivation',   // 現在のモチベーション
  12: 'workload',     // 今週の活動負荷（忙しさ）
  13: 'plan',         // P：Plan（今日の目標）
  14: 'do1',          // D：Do その１（活動内容）
  15: 'do2',          // D：Do その２（一言）
  16: 'check1',       // C：Check その１（よかった点）
  17: 'check2',       // C：Check その２（課題）
  18: 'daySatisfaction', // 今日の満足度
  19: 'action',       // A：Action（次回への改善）
  20: 'memo',         // 運営への共有・相談
  21: 'taskDetail',   // 今回参加した活動やタスクの内容
  22: 'fullname',     // フルネーム（カタカナ）
};

// ─────────────────────────────────────────────────────────
// メイン: GETリクエストでJSONを返す
// ─────────────────────────────────────────────────────────
function doGet(e) {
  try {
    var data = getReportData();

    if (e && e.parameter) {
      var p = e.parameter;

      // ?university=〇〇大学 → 大学名で絞り込み
      if (p.university) {
        data = data.filter(function(r) {
          return r.university === p.university;
        });
      }

      // ?project=〇〇 → プロジェクト名で絞り込み
      if (p.project) {
        data = data.filter(function(r) {
          return r.project === p.project;
        });
      }

      // ?name=〇〇 → フルネーム（カタカナ）で絞り込み
      if (p.name) {
        data = data.filter(function(r) {
          return r.fullname === p.name || r.fullname.includes(p.name);
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
  if (lastRow < 2) return [];

  var lastCol = Math.max(sheet.getLastColumn(), 22);
  var values  = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var records = [];

  values.forEach(function(row) {
    // 空行スキップ（活動日とフルネームが両方空の行）
    var dateVal     = row[1]  || ''; // 列2 (0始まりで1)
    var fullnameVal = row[21] || ''; // 列22 (0始まりで21)
    if (!dateVal && !fullnameVal) return;

    var record = { type: 'member' };

    Object.keys(COLUMN_MAP).forEach(function(colNum) {
      var key = COLUMN_MAP[colNum];
      var val = row[parseInt(colNum) - 1];

      if (val instanceof Date) {
        val = formatDate(val);
      }

      // 数値項目は数値型に変換
      var numKeys = ['taskDone', 'satisfaction', 'motivation', 'workload', 'daySatisfaction'];
      if (numKeys.indexOf(key) !== -1) {
        val = parseInt(val) || 0;
      }

      record[key] = (val === null || val === undefined) ? '' : val;
    });

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
