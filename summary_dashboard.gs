/**
 * 売上サマリーダッシュボード
 *
 * 「売上データ」シートを月ごとに集計し、「月次サマリー」シートへ書き出して
 * 月次推移の棒グラフを作成する。setDailyTrigger() で毎朝9時の自動実行を登録できる。
 *
 * 注意: 時間主導型トリガーから実行するため、スプレッドシートに紐づいた
 * コンテナバインドスクリプトとして設置すること。
 */

// 入力元・出力先のシート名
const SOURCE_SHEET_NAME = '売上データ';
const SUMMARY_SHEET_NAME = '月次サマリー';

// トリガー設定（実行する関数名と実行時刻）
const TRIGGER_HANDLER = 'runDashboard';
const TRIGGER_HOUR = 9;

/**
 * メイン処理。売上データを月ごとに集計し、サマリーシートとグラフを更新する。
 */
function runDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sourceSheet = ss.getSheetByName(SOURCE_SHEET_NAME);
  if (!sourceSheet) {
    throw new Error('シート「' + SOURCE_SHEET_NAME + '」が見つかりません。');
  }

  // 出力先シートがなければ新規作成する
  const summarySheet =
    ss.getSheetByName(SUMMARY_SHEET_NAME) || ss.insertSheet(SUMMARY_SHEET_NAME);

  const rows = aggregateByMonth_(sourceSheet, ss.getSpreadsheetTimeZone());
  writeSummary_(summarySheet, rows);
  createChart_(summarySheet, rows.length);

  Logger.log('集計完了: ' + rows.length + ' か月分');
}

/**
 * 毎朝9時に runDashboard を実行するトリガーを登録する。
 * 重複登録を避けるため、既存の同名トリガーは削除してから作り直す。
 * 初回に一度だけ手動で実行すればよい。
 */
function setDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    if (trigger.getHandlerFunction() === TRIGGER_HANDLER) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // atHour は指定時刻から1時間の間のいずれかに実行される。nearMinute で9:00 付近に寄せる
  ScriptApp.newTrigger(TRIGGER_HANDLER)
    .timeBased()
    .everyDays(1)
    .atHour(TRIGGER_HOUR)
    .nearMinute(0)
    .create();

  Logger.log('毎日 ' + TRIGGER_HOUR + ' 時のトリガーを登録しました。');
}

/**
 * 売上データの全行を読み込み、月ごとの合計売上と件数を集計する。
 * 見出し行など日付として解釈できない行は自動的に読み飛ばす。
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 売上データシート
 * @param {string} timeZone 日付を月に変換する際のタイムゾーン
 * @return {Array<Array>} [月ラベル, 合計売上, 件数] の配列（月の昇順）
 */
function aggregateByMonth_(sheet, timeZone) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) {
    return [];
  }

  // A列:日付 / B列:担当者名 / C列:商品名 / D列:金額 の4列を全行読み込む
  const values = sheet.getRange(1, 1, lastRow, 4).getValues();

  // キー: 'yyyy-MM'（文字列のまま昇順に並べれば時系列順になる）
  const totals = {};
  let skipped = 0;

  for (const row of values) {
    const monthKey = toMonthKey_(row[0], timeZone);
    if (!monthKey) {
      continue; // 見出し行・空行など
    }

    const amount = row[3];
    if (typeof amount !== 'number' || !isFinite(amount)) {
      skipped++;
      continue;
    }

    if (!totals[monthKey]) {
      totals[monthKey] = { sum: 0, count: 0 };
    }
    totals[monthKey].sum += amount;
    totals[monthKey].count++;
  }

  if (skipped > 0) {
    Logger.log('金額が数値でない行を ' + skipped + ' 件スキップしました。');
  }

  return Object.keys(totals)
    .sort()
    .map(function (key) {
      const year = key.slice(0, 4);
      const month = Number(key.slice(5)); // 「01」ではなく「1」と表示する
      return [year + '年' + month + '月', totals[key].sum, totals[key].count];
    });
}

/**
 * 日付セルの値を 'yyyy-MM' 形式のキーに変換する。変換できない場合は null を返す。
 * Date 型と、'2026/01/05' のような文字列の両方に対応する。
 *
 * @param {Date|string|*} value 日付セルの値
 * @param {string} timeZone タイムゾーン
 * @return {string|null}
 */
function toMonthKey_(value, timeZone) {
  if (value instanceof Date) {
    if (isNaN(value.getTime())) {
      return null;
    }
    return Utilities.formatDate(value, timeZone, 'yyyy-MM');
  }

  if (typeof value === 'string') {
    const match = value.trim().match(/^(\d{4})[\/\-年](\d{1,2})/);
    if (match) {
      const month = Number(match[2]);
      if (month >= 1 && month <= 12) {
        return match[1] + '-' + ('0' + month).slice(-2);
      }
    }
  }

  return null;
}

/**
 * サマリーシートをクリアし、見出しと集計結果を書き込む。
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 月次サマリーシート
 * @param {Array<Array>} rows aggregateByMonth_ の戻り値
 */
function writeSummary_(sheet, rows) {
  // clear() ではグラフが消えないため、既存のグラフも削除して重複を防ぐ
  sheet.getCharts().forEach(function (chart) {
    sheet.removeChart(chart);
  });
  sheet.clear();

  sheet.getRange(1, 1, 1, 3).setValues([['月', '合計売上', '件数']]).setFontWeight('bold');

  if (rows.length === 0) {
    return;
  }

  // 「2026年1月」が日付として自動変換されないよう、A列を書式なしテキストにしてから書き込む
  sheet.getRange(2, 1, rows.length, 1).setNumberFormat('@');
  sheet.getRange(2, 1, rows.length, 3).setValues(rows);
  sheet.getRange(2, 2, rows.length, 1).setNumberFormat('#,##0');
}

/**
 * 月ごとの合計売上を縦棒グラフとしてサマリーシートに配置する。
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 月次サマリーシート
 * @param {number} rowCount データ行数（見出し行は含まない）
 */
function createChart_(sheet, rowCount) {
  if (rowCount === 0) {
    return; // データがなければグラフは作らない
  }

  // A列(月)とB列(合計売上)を見出し行込みで指定する
  const chart = sheet
    .newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(sheet.getRange(1, 1, rowCount + 1, 2))
    .setNumHeaders(1)
    .setPosition(1, 5, 0, 0) // E1 セルを起点に配置
    .setOption('title', '月次売上推移')
    .setOption('legend.position', 'none')
    .setOption('hAxis.title', '月')
    .setOption('vAxis.title', '合計売上')
    .build();

  sheet.insertChart(chart);
}
