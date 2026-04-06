/**
 * フォーム送信時 → Driveフォルダ自動作成 + リンク記録
 *
 * 機能:
 *   1. フォーム送信時に顧客用Driveフォルダを自動作成
 *   2. サブフォルダ（1.交付申請 / 2.実績報告）を作成
 *   3. シート1の該当行B列（お客様企業名）にフォルダURLをハイパーリンクとして設定
 *
 * ※ フォーム回答の転記は既存GASが行っているため、
 *   このスクリプトではフォルダ作成とリンク記録のみ行う
 *
 * 注意:
 *   既存の onFormSubmit と関数名が衝突するため、
 *   この関数は onFormSubmitFolder という別名にしている。
 *   トリガー設定時は onFormSubmitFolder を選択すること。
 *
 * トリガー設定:
 *   1. 案件管理スプレッドシートのApps Scriptエディタを開く
 *   2. このコードを新しいファイルとして追加
 *   3. 時計アイコン（トリガー）→「トリガーを追加」
 *   4. 関数: onFormSubmitFolder
 *   5. イベントソース: スプレッドシートから → フォーム送信時
 */

// ============================================================
// 設定
// ============================================================
const FOLDER_CONFIG = {
  // 顧客フォルダの親フォルダID
  // https://drive.google.com/drive/folders/1L_nt38A9IFbZLaDZl1_nKxcbSrUlwKjX
  PARENT_FOLDER_ID: '1L_nt38A9IFbZLaDZl1_nKxcbSrUlwKjX',

  // サブフォルダ名（実際の運用ルールに合わせる）
  SUB_FOLDERS: ['1.交付申請', '2.実績報告'],

  // フォーム回答の列インデックス（0始まり、e.valuesの添字）
  // Form_Responses シートの列構成
  COL_COMPANY_NAME: 11,   // L列: お客様企業名
  COL_TEMPLATE_TYPE: 10,  // K列: IT導入補助金の申請枠（インボイス枠/通常枠）

  // シート1の列番号（1始まり）
  SHEET1_COL_COMPANY: 2,  // B列: お客様企業名（ここにリンクを埋め込む）

  // 転記先シート名
  TARGET_SHEET_NAME: 'シート1',
};


// ============================================================
// メイン処理
// ============================================================

/**
 * フォーム送信時にフォルダを作成する関数
 * ※ 既存の onFormSubmit とは別の関数名にしている
 * @param {Object} e - フォーム送信イベント
 */
function onFormSubmitFolder(e) {
  try {
    const row = e.values;
    const companyName = row[FOLDER_CONFIG.COL_COMPANY_NAME];
    const templateType = row[FOLDER_CONFIG.COL_TEMPLATE_TYPE] || '';

    if (!companyName) {
      Logger.log('会社名が空です');
      return;
    }

    Logger.log(`新規送客: ${companyName} (${templateType})`);

    // 1. Driveフォルダ作成
    const result = createClientFolder(companyName, templateType);
    Logger.log(`フォルダ: ${result.url}`);

    // 2. シート1のB列にハイパーリンクを設定
    if (result.url) {
      setFolderLink(companyName, result.url);
    }

  } catch (error) {
    Logger.log(`エラー: ${error.message}`);
  }
}


// ============================================================
// Drive操作
// ============================================================

/**
 * 顧客用Driveフォルダを作成
 *
 * 命名規則（実際の運用に合わせる）:
 *   親フォルダ/
 *     01.㈱〇〇〇〇_インボイス枠/
 *       1.交付申請/
 *       2.実績報告/
 *
 * @param {string} companyName - 会社名
 * @param {string} templateType - 申請枠（インボイス枠/通常枠）
 * @returns {Object} { url, folderName }
 */
function createClientFolder(companyName, templateType) {
  const parentFolder = DriveApp.getFolderById(FOLDER_CONFIG.PARENT_FOLDER_ID);

  // 既存フォルダに同じ会社名が含まれていないかチェック
  const allFolders = parentFolder.getFolders();
  let maxNumber = 0;

  while (allFolders.hasNext()) {
    const folder = allFolders.next();
    const name = folder.getName();

    if (name.includes(companyName)) {
      Logger.log(`フォルダ既存: ${name}`);
      return { url: folder.getUrl(), folderName: name };
    }

    // 連番の最大値を取得（例: "02.clipLine" → 2）
    const numMatch = name.match(/^(\d+)[._]/);
    if (numMatch) {
      const num = parseInt(numMatch[1], 10);
      if (num > maxNumber) {
        maxNumber = num;
      }
    }
  }

  // フォルダ名: 連番.会社名_申請枠
  const nextNumber = String(maxNumber + 1).padStart(2, '0');
  let folderName = `${nextNumber}.${companyName}`;
  if (templateType) {
    folderName += `_${templateType}`;
  }

  const clientFolder = parentFolder.createFolder(folderName);

  // サブフォルダ作成
  for (const subName of FOLDER_CONFIG.SUB_FOLDERS) {
    clientFolder.createFolder(subName);
  }

  Logger.log(`フォルダ作成完了: ${folderName}`);
  return { url: clientFolder.getUrl(), folderName: folderName };
}


/**
 * シート1のB列（お客様企業名）にDriveフォルダのハイパーリンクを設定
 */
function setFolderLink(companyName, folderUrl) {
  try {
    const ss = SpreadsheetApp.openById(
      '1YKHps9kq7gQ9kZIXXyiukfq_qMHN56NxP1J_f-9hQpU'
    );
    const sheet = ss.getSheetByName(FOLDER_CONFIG.TARGET_SHEET_NAME);
    if (!sheet) {
      Logger.log(`シート「${FOLDER_CONFIG.TARGET_SHEET_NAME}」が見つかりません`);
      return;
    }

    // B列から会社名を探す
    const data = sheet.getRange(2, FOLDER_CONFIG.SHEET1_COL_COMPANY, sheet.getLastRow() - 1, 1).getValues();
    for (let i = 0; i < data.length; i++) {
      const cellValue = String(data[i][0]).trim();
      if (cellValue === companyName.trim()) {
        const row = i + 2;
        const cell = sheet.getRange(row, FOLDER_CONFIG.SHEET1_COL_COMPANY);
        // RichTextValueでハイパーリンクを設定（テキストは会社名のまま）
        const richText = SpreadsheetApp.newRichTextValue()
          .setText(companyName)
          .setLinkUrl(folderUrl)
          .build();
        cell.setRichTextValue(richText);
        Logger.log(`リンク設定: ${companyName} (行${row})`);
        return;
      }
    }

    Logger.log(`シート1に「${companyName}」が見つかりません`);
  } catch (e) {
    Logger.log(`リンク設定エラー: ${e.message}`);
  }
}


// ============================================================
// テスト
// ============================================================

/**
 * テスト用: フォルダ作成のみテスト（実行後に手動削除してください）
 */
function testCreateFolder() {
  const result = createClientFolder('テスト株式会社_削除OK', 'インボイス枠');
  Logger.log(`作成されたフォルダ: ${result.url}`);
}
