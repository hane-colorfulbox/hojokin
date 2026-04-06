/**
 * 送客フォーム回答 → 管理シート自動転記 + Driveフォルダ作成
 *
 * 機能:
 *   1. Googleフォームの回答を案件管理シートに自動転記
 *   2. 顧客用のGoogle Driveフォルダを自動作成
 *   3. サブフォルダ（資料/出力）も自動作成
 *
 * トリガー設定:
 *   1. Google Apps Script エディタ → 時計アイコン（トリガー）
 *   2. 「トリガーを追加」→ 関数: onFormSubmit
 *   3. イベントソース: スプレッドシートから → フォーム送信時
 *
 *   ※ フォームの回答先スプレッドシートにこのスクリプトを紐付ける
 */

// ============================================================
// 設定（★ 実際の値に変更してください）
// ============================================================
const FORM_CONFIG = {
  // 案件管理スプレッドシートID
  MANAGEMENT_SHEET_ID: 'ここにスプレッドシートIDを入力',

  // 案件管理シート名
  MANAGEMENT_SHEET_NAME: '案件管理',

  // 顧客フォルダの親フォルダID
  PARENT_FOLDER_ID: 'ここにフォルダIDを入力',

  // 自動作成するサブフォルダ名
  SUB_FOLDERS: ['01_提出資料', '02_作成書類', '03_その他'],

  // フォーム回答の列マッピング（0始まり）
  // ★ フォームの質問順に合わせて変更してください
  FORM_COLUMNS: {
    TIMESTAMP: 0,       // タイムスタンプ（自動）
    COMPANY_NAME: 1,    // 会社名
    CONTACT_NAME: 2,    // 担当者名
    EMAIL: 3,           // メールアドレス
    PHONE: 4,           // 電話番号
    TEMPLATE_TYPE: 5,   // 申請枠（通常枠/インボイス枠）
  },

  // 案件管理シートへの転記マッピング
  // { 管理シートの列番号(1始まり): フォーム回答の列番号(0始まり) }
  TRANSFER_MAP: {
    1: 1,   // A列 ← 会社名
    2: 2,   // B列 ← 担当者名
    3: 3,   // C列 ← メールアドレス
    4: 4,   // D列 ← 電話番号
  },
};


// ============================================================
// メイン処理
// ============================================================

/**
 * フォーム送信時に呼ばれるメイン関数
 * @param {Object} e - フォーム送信イベント
 */
function onFormSubmit(e) {
  try {
    const row = e.values;
    const companyName = row[FORM_CONFIG.FORM_COLUMNS.COMPANY_NAME];

    if (!companyName) {
      Logger.log('会社名が空です');
      return;
    }

    Logger.log(`新規送客: ${companyName}`);

    // 1. 管理シートに転記
    const managementRow = addToManagementSheet(row);
    Logger.log(`管理シート: 行${managementRow}に追加`);

    // 2. Driveフォルダ作成
    const folderUrl = createClientFolder(companyName);
    Logger.log(`フォルダ作成: ${folderUrl}`);

    // 3. 管理シートにフォルダURLを記録
    if (managementRow > 0 && folderUrl) {
      recordFolderUrl(managementRow, folderUrl);
    }

  } catch (error) {
    Logger.log(`エラー: ${error.message}`);
    // エラー通知メール（任意）
    // MailApp.sendEmail('your-email@example.com', 'GASエラー通知', error.message);
  }
}


// ============================================================
// 管理シート操作
// ============================================================

/**
 * フォーム回答を案件管理シートに転記
 * @param {Array} formRow - フォーム回答の行データ
 * @returns {number} 追加した行番号
 */
function addToManagementSheet(formRow) {
  const ss = SpreadsheetApp.openById(FORM_CONFIG.MANAGEMENT_SHEET_ID);
  const sheet = ss.getSheetByName(FORM_CONFIG.MANAGEMENT_SHEET_NAME);

  if (!sheet) {
    throw new Error(`シート「${FORM_CONFIG.MANAGEMENT_SHEET_NAME}」が見つかりません`);
  }

  // 重複チェック（同じ会社名がすでにあるか）
  const companyName = formRow[FORM_CONFIG.FORM_COLUMNS.COMPANY_NAME];
  const existingData = sheet.getDataRange().getValues();
  for (let i = 1; i < existingData.length; i++) {
    if (String(existingData[i][0]).trim() === companyName.trim()) {
      Logger.log(`重複検知: ${companyName}（行${i + 1}）→ スキップ`);
      return i + 1;
    }
  }

  // 新しい行を追加
  const newRow = sheet.getLastRow() + 1;
  const transferMap = FORM_CONFIG.TRANSFER_MAP;

  for (const [sheetCol, formCol] of Object.entries(transferMap)) {
    const value = formRow[formCol];
    if (value !== undefined && value !== '') {
      sheet.getRange(newRow, Number(sheetCol)).setValue(value);
    }
  }

  // 登録日を記録
  sheet.getRange(newRow, Object.keys(transferMap).length + 1).setValue(
    new Date().toLocaleDateString('ja-JP')
  );

  return newRow;
}


/**
 * 管理シートにDriveフォルダURLを記録
 */
function recordFolderUrl(row, folderUrl) {
  try {
    const ss = SpreadsheetApp.openById(FORM_CONFIG.MANAGEMENT_SHEET_ID);
    const sheet = ss.getSheetByName(FORM_CONFIG.MANAGEMENT_SHEET_NAME);

    // フォルダURL列（★ 実際の列番号に変更）
    const FOLDER_URL_COL = 10;
    sheet.getRange(row, FOLDER_URL_COL).setValue(folderUrl);
  } catch (e) {
    Logger.log(`フォルダURL記録エラー: ${e.message}`);
  }
}


// ============================================================
// Drive操作
// ============================================================

/**
 * 顧客用Driveフォルダを作成（サブフォルダ付き）
 * @param {string} companyName - 会社名
 * @returns {string} フォルダのURL
 */
function createClientFolder(companyName) {
  const parentFolder = DriveApp.getFolderById(FORM_CONFIG.PARENT_FOLDER_ID);

  // 既存チェック
  const existing = parentFolder.getFoldersByName(companyName);
  if (existing.hasNext()) {
    const folder = existing.next();
    Logger.log(`フォルダ既存: ${companyName}`);
    return folder.getUrl();
  }

  // メインフォルダ作成
  const clientFolder = parentFolder.createFolder(companyName);

  // サブフォルダ作成
  for (const subName of FORM_CONFIG.SUB_FOLDERS) {
    clientFolder.createFolder(subName);
  }

  Logger.log(`フォルダ作成完了: ${companyName} (サブフォルダ${FORM_CONFIG.SUB_FOLDERS.length}個)`);
  return clientFolder.getUrl();
}


// ============================================================
// テスト・ユーティリティ
// ============================================================

/**
 * テスト用: ダミーデータでフォーム送信をシミュレート
 */
function testFormSubmit() {
  const dummyEvent = {
    values: [
      new Date().toISOString(),   // タイムスタンプ
      'テスト株式会社',            // 会社名
      '山田 太郎',                // 担当者名
      'test@example.com',         // メールアドレス
      '03-1234-5678',            // 電話番号
      'インボイス枠',             // 申請枠
    ]
  };

  onFormSubmit(dummyEvent);
}


/**
 * テスト用: フォルダ作成のみテスト
 */
function testCreateFolder() {
  const url = createClientFolder('テスト株式会社_削除OK');
  Logger.log(`作成されたフォルダ: ${url}`);
}
