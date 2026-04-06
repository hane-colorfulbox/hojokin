/**
 * フォーム送信時 → Driveフォルダ自動作成 + URL記録
 *
 * 機能:
 *   1. フォーム送信時に顧客用Driveフォルダを自動作成
 *   2. サブフォルダ（提出資料/作成書類/その他）も作成
 *   3. 案件管理シートのR列にフォルダURLを記録
 *
 * ※ フォーム回答の転記はGoogleフォームが自動で行うため、
 *   このスクリプトではフォルダ作成とURL記録のみ行う
 *
 * トリガー設定:
 *   1. 案件管理スプレッドシートを開く
 *   2. メニュー「拡張機能」→「Apps Script」
 *   3. このコードを貼り付け
 *   4. 時計アイコン（トリガー）→「トリガーを追加」
 *   5. 関数: onFormSubmit
 *   6. イベントソース: スプレッドシートから → フォーム送信時
 */

// ============================================================
// 設定
// ============================================================
const FORM_CONFIG = {
  // 顧客フォルダの親フォルダID
  // https://drive.google.com/drive/folders/1L_nt38A9IFbZLaDZl1_nKxcbSrUlwKjX
  PARENT_FOLDER_ID: '1L_nt38A9IFbZLaDZl1_nKxcbSrUlwKjX',

  // 自動作成するサブフォルダ名
  SUB_FOLDERS: ['01_提出資料', '02_作成書類', '03_その他'],

  // フォーム回答の列インデックス（0始まり、e.valuesの添字）
  // 実際の管理表 Form_Responses シートの列構成に対応
  COL_COMPANY_NAME: 11,   // L列: お客様企業名
  COL_TEMPLATE_TYPE: 10,  // K列: IT導入補助金の申請枠

  // フォルダURL記録先の列番号（1始まり）
  FOLDER_URL_COL: 18,     // R列
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
    const companyName = row[FORM_CONFIG.COL_COMPANY_NAME];

    if (!companyName) {
      Logger.log('会社名が空です');
      return;
    }

    Logger.log(`新規送客: ${companyName}`);

    // 1. Driveフォルダ作成（連番付き）
    const folderUrl = createClientFolder(companyName);
    Logger.log(`フォルダ作成: ${folderUrl}`);

    // 2. 管理シートのR列にフォルダURLを記録
    if (folderUrl) {
      const sheet = e.source.getActiveSheet();
      const lastRow = sheet.getLastRow();
      sheet.getRange(lastRow, FORM_CONFIG.FOLDER_URL_COL).setValue(folderUrl);
      Logger.log(`URL記録: 行${lastRow}`);
    }

  } catch (error) {
    Logger.log(`エラー: ${error.message}`);
  }
}


// ============================================================
// Drive操作
// ============================================================

/**
 * 顧客用Driveフォルダを作成（連番+会社名、サブフォルダ付き）
 *
 * 既存フォルダの命名規則に合わせる:
 *   00_フォーム, 01_クラフトバンク, 02_clipLine, ...
 *
 * @param {string} companyName - 会社名
 * @returns {string} フォルダのURL
 */
function createClientFolder(companyName) {
  const parentFolder = DriveApp.getFolderById(FORM_CONFIG.PARENT_FOLDER_ID);

  // 既存フォルダに同じ会社名が含まれていないかチェック
  const allFolders = parentFolder.getFolders();
  let maxNumber = -1;

  while (allFolders.hasNext()) {
    const folder = allFolders.next();
    const name = folder.getName();

    // 既に同じ会社名のフォルダがあればそのURLを返す
    if (name.includes(companyName)) {
      Logger.log(`フォルダ既存: ${name}`);
      return folder.getUrl();
    }

    // 連番の最大値を取得（例: "02_clipLine" → 2）
    const numMatch = name.match(/^(\d+)_/);
    if (numMatch) {
      const num = parseInt(numMatch[1], 10);
      if (num > maxNumber) {
        maxNumber = num;
      }
    }
  }

  // 次の連番でフォルダ作成
  const nextNumber = String(maxNumber + 1).padStart(2, '0');
  const folderName = `${nextNumber}_${companyName}`;
  const clientFolder = parentFolder.createFolder(folderName);

  // サブフォルダ作成
  for (const subName of FORM_CONFIG.SUB_FOLDERS) {
    clientFolder.createFolder(subName);
  }

  Logger.log(`フォルダ作成完了: ${folderName}`);
  return clientFolder.getUrl();
}


// ============================================================
// テスト
// ============================================================

/**
 * テスト用: フォルダ作成のみテスト（実行後に手動削除してください）
 */
function testCreateFolder() {
  const url = createClientFolder('テスト株式会社_削除OK');
  Logger.log(`作成されたフォルダ: ${url}`);
}
