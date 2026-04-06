/**
 * メール添付ファイル → Google Drive 自動格納
 *
 * 機能:
 *   1. 特定条件のメールから添付ファイルを取得
 *   2. 顧客名でDriveフォルダを探し（なければ作成）、添付を保存
 *   3. 案件管理スプレッドシートにチェックを入れる
 *   4. 処理済みメールにラベルを付けて二重処理を防止
 *
 * トリガー設定:
 *   1. Google Apps Script エディタ → 時計アイコン（トリガー）
 *   2. 「トリガーを追加」→ 関数: processNewEmails
 *   3. イベントソース: 時間主導型 → 5分おき or 15分おき
 *
 * 初回セットアップ:
 *   1. CONFIG の値を実際の環境に合わせて変更
 *   2. setupLabels() を1回手動実行してラベルを作成
 */

// ============================================================
// 設定（★ 実際の値に変更してください）
// ============================================================
const CONFIG = {
  // 資料提出メールの検索条件（Gmail検索クエリ）
  // 例: 特定の件名パターン、送信元、ラベルなど
  SEARCH_QUERY: 'subject:(資料 OR 提出 OR 補助金) has:attachment -label:auto-processed',

  // 資料格納先の親フォルダID（DriveのURL末尾の文字列）
  // 例: https://drive.google.com/drive/folders/XXXXX → XXXXX がID
  PARENT_FOLDER_ID: 'ここにフォルダIDを入力',

  // 案件管理スプレッドシートID
  // https://docs.google.com/spreadsheets/d/1YKHps9kq7gQ9kZIXXyiukfq_qMHN56NxP1J_f-9hQpU/
  MANAGEMENT_SHEET_ID: '1YKHps9kq7gQ9kZIXXyiukfq_qMHN56NxP1J_f-9hQpU',

  // 案件管理シート名（★ 実際のシート名に合わせて変更）
  MANAGEMENT_SHEET_NAME: 'シート1',

  // 案件管理シートの列番号（1始まり）
  COL_COMPANY_NAME: 1,    // 会社名の列
  COL_DOC_RECEIVED: 5,    // 資料受領チェックの列（★要確認）

  // 処理済みラベル名
  PROCESSED_LABEL: 'auto-processed',

  // 1回の実行で処理するメール上限
  MAX_THREADS: 10,
};


// ============================================================
// メイン処理
// ============================================================

/**
 * 新着メールを処理するメイン関数（トリガーから呼ばれる）
 */
function processNewEmails() {
  const threads = GmailApp.search(CONFIG.SEARCH_QUERY, 0, CONFIG.MAX_THREADS);

  if (threads.length === 0) {
    Logger.log('新着メールなし');
    return;
  }

  Logger.log(`${threads.length}件のメールを処理`);
  const parentFolder = DriveApp.getFolderById(CONFIG.PARENT_FOLDER_ID);
  const processedLabel = getOrCreateLabel(CONFIG.PROCESSED_LABEL);

  for (const thread of threads) {
    try {
      processThread(thread, parentFolder);
      thread.addLabel(processedLabel);
      Logger.log(`処理完了: ${thread.getFirstMessageSubject()}`);
    } catch (e) {
      Logger.log(`エラー: ${thread.getFirstMessageSubject()} - ${e.message}`);
    }
  }
}


/**
 * 1つのメールスレッドを処理
 */
function processThread(thread, parentFolder) {
  const messages = thread.getMessages();

  for (const message of messages) {
    const attachments = message.getAttachments();
    if (attachments.length === 0) continue;

    // 送信者・件名から顧客名を推定
    const companyName = extractCompanyName(message);
    if (!companyName) {
      Logger.log(`顧客名を特定できません: ${message.getSubject()}`);
      continue;
    }

    // 顧客フォルダを取得 or 作成
    const companyFolder = getOrCreateFolder(parentFolder, companyName);

    // 添付ファイルを保存
    let savedCount = 0;
    for (const attachment of attachments) {
      const fileName = attachment.getName();
      // 重複チェック（同名ファイルがあればスキップ）
      if (fileExistsInFolder(companyFolder, fileName)) {
        Logger.log(`スキップ（既存）: ${fileName}`);
        continue;
      }
      companyFolder.createFile(attachment);
      savedCount++;
      Logger.log(`保存: ${companyName}/${fileName}`);
    }

    // 案件管理表を更新
    if (savedCount > 0) {
      updateManagementSheet(companyName);
    }
  }
}


// ============================================================
// 顧客名の推定
// ============================================================

/**
 * メールの件名・本文から顧客名（会社名）を推定する
 *
 * ★ 実際の運用に合わせてカスタマイズしてください
 * パターン例:
 *   件名: 「【株式会社○○】資料送付の件」
 *   件名: 「○○株式会社_補助金資料」
 *   本文冒頭: 「株式会社○○の△△です」
 */
function extractCompanyName(message) {
  const subject = message.getSubject();

  // パターン1: 【会社名】形式
  const bracketMatch = subject.match(/【(.+?)】/);
  if (bracketMatch) {
    return cleanCompanyName(bracketMatch[1]);
  }

  // パターン2: 会社名_○○ 形式（アンダースコア区切り）
  const underscoreMatch = subject.match(/^(.+?)[_＿]/);
  if (underscoreMatch) {
    const candidate = underscoreMatch[1].trim();
    if (candidate.includes('株式会社') || candidate.includes('有限会社')) {
      return cleanCompanyName(candidate);
    }
  }

  // パターン3: 「株式会社○○」「○○株式会社」を件名から探す
  const corpMatch = subject.match(/(株式会社[^\s_＿【】]+|[^\s_＿【】]+株式会社)/);
  if (corpMatch) {
    return cleanCompanyName(corpMatch[1]);
  }

  // パターン4: 本文から探す
  const body = message.getPlainBody().substring(0, 500);
  const bodyMatch = body.match(/(株式会社[^\s、。]+|[^\s、。]+株式会社)/);
  if (bodyMatch) {
    return cleanCompanyName(bodyMatch[1]);
  }

  return null;
}


/**
 * 会社名をクリーンアップ（余計な文字を除去）
 */
function cleanCompanyName(name) {
  return name
    .replace(/[\s　]+/g, '')
    .replace(/様$/, '')
    .replace(/御中$/, '')
    .trim();
}


// ============================================================
// Drive操作
// ============================================================

/**
 * 親フォルダ内にサブフォルダを取得 or 作成
 */
function getOrCreateFolder(parentFolder, folderName) {
  const folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  Logger.log(`フォルダ作成: ${folderName}`);
  return parentFolder.createFolder(folderName);
}


/**
 * フォルダ内に同名ファイルが存在するかチェック
 */
function fileExistsInFolder(folder, fileName) {
  const files = folder.getFilesByName(fileName);
  return files.hasNext();
}


// ============================================================
// スプレッドシート操作
// ============================================================

/**
 * 案件管理表の該当行にチェックを入れる
 */
function updateManagementSheet(companyName) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.MANAGEMENT_SHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.MANAGEMENT_SHEET_NAME);
    if (!sheet) {
      Logger.log(`シート「${CONFIG.MANAGEMENT_SHEET_NAME}」が見つかりません`);
      return;
    }

    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      const cellValue = String(data[i][CONFIG.COL_COMPANY_NAME - 1]).trim();
      if (cellValue.includes(companyName) || companyName.includes(cellValue)) {
        const row = i + 1;
        const currentValue = sheet.getRange(row, CONFIG.COL_DOC_RECEIVED).getValue();
        if (!currentValue) {
          sheet.getRange(row, CONFIG.COL_DOC_RECEIVED).setValue('✅');
          sheet.getRange(row, CONFIG.COL_DOC_RECEIVED).setNote(
            `自動処理: ${new Date().toLocaleString('ja-JP')}`
          );
          Logger.log(`管理表更新: ${companyName} (行${row})`);
        }
        return;
      }
    }

    Logger.log(`管理表に「${companyName}」が見つかりません`);
  } catch (e) {
    Logger.log(`管理表更新エラー: ${e.message}`);
  }
}


// ============================================================
// ユーティリティ
// ============================================================

/**
 * Gmailラベルを取得 or 作成
 */
function getOrCreateLabel(labelName) {
  let label = GmailApp.getUserLabelByName(labelName);
  if (!label) {
    label = GmailApp.createLabel(labelName);
    Logger.log(`ラベル作成: ${labelName}`);
  }
  return label;
}


/**
 * 初回セットアップ: ラベルを作成する（1回だけ手動実行）
 */
function setupLabels() {
  getOrCreateLabel(CONFIG.PROCESSED_LABEL);
  Logger.log('セットアップ完了');
}


/**
 * テスト用: 検索クエリにマッチするメール一覧を表示（実際の処理はしない）
 */
function testSearchQuery() {
  const threads = GmailApp.search(CONFIG.SEARCH_QUERY, 0, 5);
  Logger.log(`マッチ: ${threads.length}件`);
  for (const thread of threads) {
    const msg = thread.getMessages()[0];
    Logger.log(`  件名: ${msg.getSubject()}`);
    Logger.log(`  送信者: ${msg.getFrom()}`);
    Logger.log(`  添付: ${msg.getAttachments().length}件`);
    Logger.log(`  推定顧客名: ${extractCompanyName(msg)}`);
    Logger.log('---');
  }
}
