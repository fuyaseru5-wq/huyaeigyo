/**
 * フヤセル営業集計ツール (GAS版)
 * * 使い方:
 * 1. スプレッドシートのメニュー「拡張機能」>「Apps Script」を開く
 * 2. このコードをすべて貼り付けて保存
 * 3. スプレッドシートをリロードするとメニューに「フヤセル集計」が出現します
 */

// --- 設定エリア (列がずれたらここを修正してください) ---
const CONFIG = {
  // 列番号 (A=1, B=2 ... P=16, Q=17, R=18, X=24)
  COL_DATE: 16,   // P列: 入金日/決済日
  COL_STATUS: 17, // Q列: 状態
  COL_NOTE: 18,   // R列: 備考
  COL_NAME: 24,   // X列: 担当者
  
  // 判定条件
  STATUS_KEYWORD: '決済完了',
  SPLIT_LIMIT: 24, // この回数以上の分割をピックアップ
  
  // 出力先シート名
  OUTPUT_SHEET: '📊営業集計結果'
};

/**
 * スプレッドシートを開いた時にメニューを追加
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('フヤセル集計')
    .addItem('集計を実行する', 'main')
    .addToUi();
}

/**
 * メイン処理
 */
function main() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet(); // アクティブなシートを読み込む
  
  // データの取得 (2行目から最終行まで)
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('データがありません。');
    return;
  }
  
  // 高速化のためデータを一括取得
  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  
  // 集計用オブジェクト
  const monthlyStats = {};
  
  // 1行ずつ解析
  data.forEach(row => {
    // 配列のインデックスは0始まりなので、列番号-1 します
    const status = String(row[CONFIG.COL_STATUS - 1]);
    const name = String(row[CONFIG.COL_NAME - 1]).trim();
    
    // 「決済完了」以外はスキップ
    if (status.indexOf(CONFIG.STATUS_KEYWORD) === -1) return;
    if (!name) return; // 名前がない場合スキップ

    // 日付から年月を取得
    const dateVal = row[CONFIG.COL_DATE - 1];
    let monthKey = "不明な期間";
    
    if (dateVal instanceof Date) {
      monthKey = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "yyyy年MM月");
    } else if (String(dateVal).match(/\d{1,2}月/)) {
       const today = new Date();
       monthKey = today.getFullYear() + "年" + String(dateVal).split("月")[0] + "月";
    }

    // 集計初期化
    if (!monthlyStats[monthKey]) {
      monthlyStats[monthKey] = { total: 0, agents: {} };
    }
    
    if (!monthlyStats[monthKey].agents[name]) {
      monthlyStats[monthKey].agents[name] = { count: 0, highSplits: [] };
    }

    // カウントアップ
    monthlyStats[monthKey].total++;
    monthlyStats[monthKey].agents[name].count++;

    // 備考から分割数を抽出
    const note = String(row[CONFIG.COL_NOTE - 1]);
    const match = note.match(/(\d+)分割/);
    if (match) {
      const splitNum = parseInt(match[1], 10);
      if (splitNum >= CONFIG.SPLIT_LIMIT) {
        monthlyStats[monthKey].agents[name].highSplits.push(splitNum + "分割");
      }
    }
  });

  // 結果を出力する
  outputResults(ss, monthlyStats);
}

/**
 * 集計結果をシートに書き出す
 */
function outputResults(ss, monthlyStats) {
  let outSheet = ss.getSheetByName(CONFIG.OUTPUT_SHEET);
  if (outSheet) {
    outSheet.clear();
  } else {
    outSheet = ss.insertSheet(CONFIG.OUTPUT_SHEET);
  }
  
  // タイトル行
  outSheet.getRange("A1").setValue("フヤセル営業集計レポート")
    .setFontSize(16).setFontWeight("bold");
  outSheet.getRange("A2").setValue("実行日時: " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss"));

  let currentRow = 4;

  // 月ごとにソート
  const sortedMonths = Object.keys(monthlyStats).sort((a, b) => a < b ? 1 : -1);

  sortedMonths.forEach(month => {
    const stats = monthlyStats[month];
    
    // 月ヘッダー
    outSheet.getRange(currentRow, 1).setValue(`■ ${month} (全体: ${stats.total}本)`)
      .setFontWeight("bold").setBackground("#e6f2ff").setFontSize(12);
    outSheet.getRange(currentRow, 1, 1, 6).merge();
    currentRow++;

    // テーブルヘッダー
    const headers = ["順位", "担当者", "獲得本数", "シェア(%)", "グラフ", "特記事項 (インセンティブ対象)"];
    outSheet.getRange(currentRow, 1, 1, 6).setValues([headers])
      .setBackground("#f3f4f6").setFontWeight("bold").setBorder(true, true, true, true, true, true);
    currentRow++;

    // 担当者ソート
    const sortedAgents = Object.keys(stats.agents).map(name => {
      return { name: name, ...stats.agents[name] };
    }).sort((a, b) => b.count - a.count);

    // データ書き込み
    sortedAgents.forEach((agent, index) => {
      const rank = index + 1;
      const share = stats.total > 0 ? (agent.count / stats.total) : 0;
      
      let remarks = "";
      if (agent.highSplits.length > 0) {
        const summary = {};
        agent.highSplits.forEach(s => { summary[s] = (summary[s] || 0) + 1; });
        const parts = [];
        for (let key in summary) parts.push(`${key}(${summary[key]})`);
        remarks = "内: " + parts.join(", ");
      }

      outSheet.getRange(currentRow, 1).setValue(rank);
      outSheet.getRange(currentRow, 2).setValue(agent.name);
      outSheet.getRange(currentRow, 3).setValue(agent.count);
      outSheet.getRange(currentRow, 4).setValue(share).setNumberFormat("0.0%");
      
      const color = rank === 1 ? "#F59E0B" : "#3B82F6"; 
      const formula = `=SPARKLINE(${agent.count}, {"charttype","bar";"max",${stats.total};"color1","${color}"})`;
      outSheet.getRange(currentRow, 5).setFormula(formula);
      
      outSheet.getRange(currentRow, 6).setValue(remarks).setFontColor("#DC2626").setFontWeight("bold");

      currentRow++;
    });

    currentRow += 2;
  });

  outSheet.setColumnWidth(1, 50);
  outSheet.setColumnWidth(2, 120);
  outSheet.setColumnWidth(3, 80);
  outSheet.setColumnWidth(4, 80);
  outSheet.setColumnWidth(5, 150);
  outSheet.setColumnWidth(6, 300);
}
