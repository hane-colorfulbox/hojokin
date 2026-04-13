/**
 * 案件管理表に集計ダッシュボードタブを作成・更新
 *
 * 手動実行 or トリガーで定期更新
 */

// ============================================================
// 設定
// ============================================================
const DASHBOARD_CONFIG = {
  SPREADSHEET_ID: '1YKHps9kq7gQ9kZIXXyiukfq_qMHN56NxP1J_f-9hQpU',
  SOURCE_SHEET: 'シート1',
  DASHBOARD_SHEET: 'ダッシュボード',
  HEADER_ROW: 2,

  // 列番号（1始まり）
  COL_CB_STAFF: 1,    // A列
  COL_COMPANY: 2,     // B列
  COL_VENDOR: 3,      // C列
  COL_TEMPLATE: 4,    // D列
  COL_ROUND: 5,       // E列
  COL_STATUS: 6,      // F列
};


// ============================================================
// メイン処理
// ============================================================

/**
 * ダッシュボードを更新（手動実行 or トリガー）
 */
function updateDashboard() {
  const ss = SpreadsheetApp.openById(DASHBOARD_CONFIG.SPREADSHEET_ID);
  const source = ss.getSheetByName(DASHBOARD_CONFIG.SOURCE_SHEET);

  if (!source) {
    Logger.log('シート1が見つかりません');
    return;
  }

  // ダッシュボードシートを取得 or 作成
  let dashboard = ss.getSheetByName(DASHBOARD_CONFIG.DASHBOARD_SHEET);
  if (!dashboard) {
    dashboard = ss.insertSheet(DASHBOARD_CONFIG.DASHBOARD_SHEET);
  } else {
    dashboard.clear();
  }

  // データ取得（ヘッダー行の次から）
  const lastRow = source.getLastRow();
  if (lastRow <= DASHBOARD_CONFIG.HEADER_ROW) {
    dashboard.getRange(1, 1).setValue('データがありません');
    return;
  }

  const dataRange = source.getRange(
    DASHBOARD_CONFIG.HEADER_ROW + 1, 1,
    lastRow - DASHBOARD_CONFIG.HEADER_ROW,
    6 // A〜F列
  );
  const data = dataRange.getDisplayValues();

  // B列（企業名）が空の行を除外
  const rows = data.filter(row => row[1] !== '');

  // 集計
  const statusCounts = countBy_(rows, 5);   // F列: ステータス
  const vendorCounts = countBy_(rows, 2);   // C列: 支援事業者
  const templateCounts = countBy_(rows, 3); // D列: 申請枠
  const staffCounts = countBy_(rows, 0);    // A列: CB担当
  const roundCounts = countBy_(rows, 4);    // E列: 公募回

  // ダッシュボードに書き込み
  let currentRow = 1;

  // タイトル
  dashboard.getRange(currentRow, 1).setValue('案件管理ダッシュボード');
  dashboard.getRange(currentRow, 1).setFontSize(14).setFontWeight('bold');
  currentRow++;

  dashboard.getRange(currentRow, 1).setValue('最終更新: ' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm'));
  currentRow++;

  dashboard.getRange(currentRow, 1).setValue('総案件数: ' + rows.length + '件');
  dashboard.getRange(currentRow, 1).setFontWeight('bold');
  currentRow += 2;

  // ステータス別
  currentRow = writeSection_(dashboard, currentRow, 'ステータス別', statusCounts);
  currentRow++;

  // 支援事業者別
  currentRow = writeSection_(dashboard, currentRow, '支援事業者別', vendorCounts);
  currentRow++;

  // 申請枠別
  currentRow = writeSection_(dashboard, currentRow, '申請枠別', templateCounts);
  currentRow++;

  // CB担当別
  currentRow = writeSection_(dashboard, currentRow, 'CB担当別', staffCounts);
  currentRow++;

  // 公募回別
  currentRow = writeSection_(dashboard, currentRow, '公募回別', roundCounts);

  // 列幅調整
  dashboard.setColumnWidth(1, 300);
  dashboard.setColumnWidth(2, 100);

  // ヘッダー行の色設定
  dashboard.getRange(1, 1, 1, 2).setBackground('#1a73e8').setFontColor('#ffffff');

  Logger.log('ダッシュボード更新完了');
}


// ============================================================
// ヘルパー
// ============================================================

/**
 * 指定列でグループ化してカウント
 */
function countBy_(rows, colIndex) {
  const counts = {};
  for (const row of rows) {
    const key = row[colIndex] || '（未設定）';
    counts[key] = (counts[key] || 0) + 1;
  }

  // 件数の降順でソート
  return Object.entries(counts).sort((a, b) => b[1] - a[1]);
}


/**
 * ダッシュボードにセクションを書き込み
 * @returns {number} 次の行番号
 */
function writeSection_(sheet, startRow, title, data) {
  // セクションタイトル
  sheet.getRange(startRow, 1).setValue(title);
  sheet.getRange(startRow, 1).setFontWeight('bold').setFontSize(11);
  sheet.getRange(startRow, 1, 1, 2).setBackground('#e8f0fe');
  startRow++;

  // ヘッダー
  sheet.getRange(startRow, 1).setValue('項目');
  sheet.getRange(startRow, 2).setValue('件数');
  sheet.getRange(startRow, 1, 1, 2).setFontWeight('bold');
  startRow++;

  // データ
  for (const [key, count] of data) {
    sheet.getRange(startRow, 1).setValue(key);
    sheet.getRange(startRow, 2).setValue(count);
    startRow++;
  }

  return startRow;
}
