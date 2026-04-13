/**
 * 案件管理表に集計ダッシュボードタブを作成・更新
 *
 * 横軸: 公募回（締め切り）
 * 縦軸: 支援事業者名
 *
 * データソース:
 *   1. 案件管理表（シート1）
 *   2. クラフトバンク管理表（別スプレッドシート）
 *
 * 手動実行 or トリガーで定期更新
 */

// ============================================================
// 設定
// ============================================================
const DASHBOARD_CONFIG = {
  // 案件管理表
  SPREADSHEET_ID: '1YKHps9kq7gQ9kZIXXyiukfq_qMHN56NxP1J_f-9hQpU',
  SOURCE_SHEET: '2026案件一覧',
  DASHBOARD_SHEET: 'ダッシュボード',
  HEADER_ROW: 2,

  // 案件管理表の列番号（1始まり）
  COL_COMPANY: 2,     // B列: お客様企業名
  COL_VENDOR: 3,      // C列: 支援事業者名
  COL_ROUND: 5,       // E列: 公募回（締め切り）

  // クラフトバンク管理表
  CB_SPREADSHEET_ID: '1dn6HMJMdFJNQljGRXPPX6RLfVkltLoguKjcb4uDFTpQ',
  CB_HEADER_ROW: 2,
  CB_COL_COMPANY: 2,   // B列: 社名
  CB_COL_TEMPLATE: 10, // J列: 申請枠（回次情報を含む）
  CB_VENDOR_NAME: 'クラフトバンク',
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

  // --- 案件管理表からデータ取得 ---
  const allRows = getMainSheetRows_(source);

  // --- クラフトバンクからデータ取得 ---
  const cbRows = getCraftbankRows_();

  // 全データを結合: [支援事業者名, 公募回]
  const combinedRows = [];

  // 案件管理表の公募回一覧を先に収集（マッチング用）
  const mainRounds = [];
  for (const row of allRows) {
    const round = row[4] || '（未設定）';
    if (!mainRounds.includes(round)) {
      mainRounds.push(round);
    }
  }

  for (const row of allRows) {
    combinedRows.push({
      vendor: row[2] || '（未設定）',  // C列: 支援事業者名
      round: row[4] || '（未設定）',   // E列: 公募回
    });
  }

  for (const row of cbRows) {
    // クラフトバンクの「1次」を案件管理表の「1次（5/12締切）」にマッチさせる
    const matched = mainRounds.find(r => r.startsWith(row.round));
    combinedRows.push({
      vendor: DASHBOARD_CONFIG.CB_VENDOR_NAME,
      round: matched || row.round,
    });
  }

  if (combinedRows.length === 0) {
    dashboard.getRange(1, 1).setValue('データがありません');
    return;
  }

  // 公募回（締め切り）の一覧を取得（出現順）
  const roundSet = [];
  for (const row of combinedRows) {
    if (!roundSet.includes(row.round)) {
      roundSet.push(row.round);
    }
  }

  // 支援事業者名の一覧を取得（出現順）
  const vendorSet = [];
  for (const row of combinedRows) {
    if (!vendorSet.includes(row.vendor)) {
      vendorSet.push(row.vendor);
    }
  }

  // クロス集計: 支援事業者名 x 公募回 → 件数
  const cross = {};
  for (const row of combinedRows) {
    const key = row.vendor + '|||' + row.round;
    cross[key] = (cross[key] || 0) + 1;
  }

  // 公募回ごとの合計
  const roundTotals = {};
  for (const round of roundSet) {
    roundTotals[round] = 0;
  }
  for (const row of combinedRows) {
    roundTotals[row.round]++;
  }

  // --- ダッシュボードに書き込み ---
  let currentRow = 1;

  // タイトル
  dashboard.getRange(currentRow, 1).setValue('案件管理ダッシュボード');
  dashboard.getRange(currentRow, 1).setFontSize(14).setFontWeight('bold');
  currentRow++;

  dashboard.getRange(currentRow, 1).setValue('最終更新: ' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm'));
  currentRow++;

  dashboard.getRange(currentRow, 1).setValue('総案件数: ' + combinedRows.length + '件');
  dashboard.getRange(currentRow, 1).setFontWeight('bold');
  currentRow += 2;

  // クロス集計表
  const tableStartRow = currentRow;
  const tableStartCol = 1;

  // ヘッダー行: 支援事業者名 | 公募回1 | 公募回2 | ... | 合計
  dashboard.getRange(currentRow, tableStartCol).setValue('支援事業者名');
  for (let i = 0; i < roundSet.length; i++) {
    dashboard.getRange(currentRow, tableStartCol + 1 + i).setValue(roundSet[i]);
  }
  dashboard.getRange(currentRow, tableStartCol + 1 + roundSet.length).setValue('合計');

  // ヘッダー行のスタイル
  const headerRange = dashboard.getRange(currentRow, tableStartCol, 1, roundSet.length + 2);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#1a73e8');
  headerRange.setFontColor('#ffffff');
  headerRange.setHorizontalAlignment('center');
  currentRow++;

  // データ行: 支援事業者ごと
  for (const vendor of vendorSet) {
    dashboard.getRange(currentRow, tableStartCol).setValue(vendor);
    let vendorTotal = 0;

    for (let i = 0; i < roundSet.length; i++) {
      const key = vendor + '|||' + roundSet[i];
      const count = cross[key] || 0;
      if (count > 0) {
        dashboard.getRange(currentRow, tableStartCol + 1 + i).setValue(count);
      }
      vendorTotal += count;
    }

    dashboard.getRange(currentRow, tableStartCol + 1 + roundSet.length).setValue(vendorTotal);
    currentRow++;
  }

  // 合計行
  dashboard.getRange(currentRow, tableStartCol).setValue('合計');
  dashboard.getRange(currentRow, tableStartCol).setFontWeight('bold');
  let grandTotal = 0;

  for (let i = 0; i < roundSet.length; i++) {
    const total = roundTotals[roundSet[i]];
    dashboard.getRange(currentRow, tableStartCol + 1 + i).setValue(total);
    grandTotal += total;
  }
  dashboard.getRange(currentRow, tableStartCol + 1 + roundSet.length).setValue(grandTotal);

  // 合計行のスタイル
  const totalRange = dashboard.getRange(currentRow, tableStartCol, 1, roundSet.length + 2);
  totalRange.setFontWeight('bold');
  totalRange.setBackground('#e8f0fe');

  // データ部分の中央揃え（数値列）
  const dataRows = vendorSet.length;
  if (dataRows > 0 && roundSet.length > 0) {
    dashboard.getRange(tableStartRow + 1, tableStartCol + 1, dataRows, roundSet.length + 1)
      .setHorizontalAlignment('center');
    dashboard.getRange(currentRow, tableStartCol + 1, 1, roundSet.length + 1)
      .setHorizontalAlignment('center');
  }

  // 列幅調整
  dashboard.setColumnWidth(1, 250);
  for (let i = 0; i < roundSet.length + 1; i++) {
    dashboard.setColumnWidth(2 + i, 150);
  }

  // 罫線
  const tableRange = dashboard.getRange(tableStartRow, tableStartCol, dataRows + 2, roundSet.length + 2);
  tableRange.setBorder(true, true, true, true, true, true);

  Logger.log('ダッシュボード更新完了');
}


// ============================================================
// データ取得
// ============================================================

/**
 * 案件管理表（シート1）からデータ行を取得
 * @returns {string[][]} B列が空でないデータ行
 */
function getMainSheetRows_(source) {
  const lastRow = source.getLastRow();
  if (lastRow <= DASHBOARD_CONFIG.HEADER_ROW) return [];

  const dataRange = source.getRange(
    DASHBOARD_CONFIG.HEADER_ROW + 1, 1,
    lastRow - DASHBOARD_CONFIG.HEADER_ROW,
    6
  );
  const data = dataRange.getDisplayValues();
  return data.filter(row => row[1] !== '');
}


/**
 * クラフトバンク管理表からデータを取得
 * @returns {Object[]} { round } の配列
 */
function getCraftbankRows_() {
  try {
    const cbSs = SpreadsheetApp.openById(DASHBOARD_CONFIG.CB_SPREADSHEET_ID);
    const cbSheet = cbSs.getSheets()[0]; // 最初のシート

    const lastRow = cbSheet.getLastRow();
    if (lastRow <= DASHBOARD_CONFIG.CB_HEADER_ROW) return [];

    const dataRange = cbSheet.getRange(
      DASHBOARD_CONFIG.CB_HEADER_ROW + 1, 1,
      lastRow - DASHBOARD_CONFIG.CB_HEADER_ROW,
      DASHBOARD_CONFIG.CB_COL_TEMPLATE
    );
    const data = dataRange.getDisplayValues();

    const rows = [];
    for (const row of data) {
      const company = row[DASHBOARD_CONFIG.CB_COL_COMPANY - 1]; // B列: 社名
      if (!company) continue;

      // J列（申請枠）から公募回を抽出
      // 例: "インボ_1" → "1次", "通常_2" → "2次"
      const template = row[DASHBOARD_CONFIG.CB_COL_TEMPLATE - 1] || '';
      const round = parseRoundFromTemplate_(template);

      rows.push({ round: round });
    }

    Logger.log('クラフトバンク: ' + rows.length + '件取得');
    return rows;

  } catch (e) {
    Logger.log('クラフトバンク取得エラー: ' + e.message);
    return [];
  }
}


/**
 * 申請枠の文字列から公募回を抽出
 * 例: "インボ_1" → "1次", "通常_2" → "2次"
 * @param {string} template - 申請枠の文字列
 * @returns {string} 公募回
 */
function parseRoundFromTemplate_(template) {
  // "_数字" のパターンを探す
  const match = template.match(/[_](\d+)/);
  if (match) {
    return match[1] + '次';
  }
  return '（未設定）';
}
