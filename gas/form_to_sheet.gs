/**
 * フォーム送信時 → Driveフォルダ自動作成 + リンク記録
 *
 * フォルダ構造:
 *   2026/
 *     支援事業者名/
 *       001.顧客企業名_申請枠/
 *         1.交付申請/
 *         2.実績報告/
 *
 * トリガー設定:
 *   関数: onFormSubmitFolder
 *   イベントソース: スプレッドシートから → フォーム送信時
 */

// ============================================================
// 設定
// ============================================================
const FOLDER_CONFIG = {
  // 2026年度の親フォルダID
  // https://drive.google.com/drive/folders/1L_nt38A9IFbZLaDZl1_nKxcbSrUlwKjX
  PARENT_FOLDER_ID: '1L_nt38A9IFbZLaDZl1_nKxcbSrUlwKjX',

  // サブフォルダ名
  SUB_FOLDERS: ['1.交付申請', '2.実績報告'],

  // フォーム回答の列インデックス（0始まり、e.valuesの添字）
  COL_COMPANY_NAME: 11,    // L列: お客様企業名
  COL_TEMPLATE_TYPE: 10,   // K列: IT導入補助金の申請枠
  COL_SUPPORT_COMPANY: 6,  // G列: 支援事業者名（貴社名）

  // シート1の列番号（1始まり）
  SHEET1_COL_COMPANY: 2,   // B列: お客様企業名

  // 転記先シート名
  TARGET_SHEET_NAME: '2026案件一覧',
};


// ============================================================
// メイン処理
// ============================================================

/**
 * フォーム送信時にフォルダを作成する関数
 * @param {Object} e - フォーム送信イベント
 */
function onFormSubmitFolder(e) {
  try {
    const row = e.values;
    const companyName = row[FOLDER_CONFIG.COL_COMPANY_NAME];
    const templateType = row[FOLDER_CONFIG.COL_TEMPLATE_TYPE] || '';
    const supportCompany = row[FOLDER_CONFIG.COL_SUPPORT_COMPANY] || '';

    if (!companyName) {
      Logger.log('会社名が空です');
      return;
    }

    Logger.log(`新規送客: ${companyName} (${templateType}) / 支援事業者: ${supportCompany}`);

    // 1. Driveフォルダ作成（支援事業者フォルダの下に顧客フォルダ）
    const result = createClientFolder(companyName, templateType, supportCompany);
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
 * 構造:
 *   親フォルダ(2026)/
 *     支援事業者名/
 *       001.顧客企業名_申請枠/
 *         1.交付申請/
 *         2.実績報告/
 *
 * @param {string} companyName - 顧客企業名
 * @param {string} templateType - 申請枠（インボイス枠/通常枠）
 * @param {string} supportCompany - 支援事業者名
 * @returns {Object} { url, folderName }
 */
function createClientFolder(companyName, templateType, supportCompany) {
  const parentFolder = DriveApp.getFolderById(FOLDER_CONFIG.PARENT_FOLDER_ID);

  // 1. 支援事業者フォルダを取得 or 作成
  const vendorFolder = getOrCreateVendorFolder_(parentFolder, supportCompany);

  // 2. 支援事業者フォルダ内で既存の顧客フォルダをチェック
  const allFolders = vendorFolder.getFolders();
  let maxNumber = 0;

  while (allFolders.hasNext()) {
    const folder = allFolders.next();
    const name = folder.getName();

    if (name.includes(companyName)) {
      Logger.log(`フォルダ既存: ${name}`);
      return { url: folder.getUrl(), folderName: name };
    }

    // 連番の最大値を取得（例: "002.〇〇株式会社" → 2）
    const numMatch = name.match(/^(\d+)[._]/);
    if (numMatch) {
      const num = parseInt(numMatch[1], 10);
      if (num > maxNumber) {
        maxNumber = num;
      }
    }
  }

  // 3. 顧客フォルダ作成: 連番.企業名_申請枠
  const nextNumber = String(maxNumber + 1).padStart(3, '0');
  let folderName = `${nextNumber}.${companyName}`;
  if (templateType) {
    folderName += `_${templateType}`;
  }

  const clientFolder = vendorFolder.createFolder(folderName);

  // 4. サブフォルダ作成
  for (const subName of FOLDER_CONFIG.SUB_FOLDERS) {
    clientFolder.createFolder(subName);
  }

  Logger.log(`フォルダ作成完了: ${supportCompany}/${folderName}`);
  return { url: clientFolder.getUrl(), folderName: folderName };
}


/**
 * 支援事業者フォルダを取得（なければ作成）
 * @param {Folder} parentFolder - 親フォルダ（2026）
 * @param {string} supportCompany - 支援事業者名
 * @returns {Folder} 支援事業者フォルダ
 */
function getOrCreateVendorFolder_(parentFolder, supportCompany) {
  // 支援事業者名が空の場合は「その他」フォルダに入れる
  const vendorName = supportCompany ? supportCompany.trim() : 'その他';

  // 既存フォルダを検索
  const folders = parentFolder.getFolders();
  while (folders.hasNext()) {
    const folder = folders.next();
    if (folder.getName() === vendorName) {
      Logger.log(`支援事業者フォルダ既存: ${vendorName}`);
      return folder;
    }
  }

  // なければ作成
  const newFolder = parentFolder.createFolder(vendorName);
  Logger.log(`支援事業者フォルダ作成: ${vendorName}`);
  return newFolder;
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
