/**
 * ステータス変更時にGoogle Chatへ通知
 *
 * トリガー設定:
 *   関数: onEditNotify
 *   イベントソース: スプレッドシートから → 編集時
 *
 * 【2026-04-17】村上さん依頼で通知を一旦停止中。
 * 再開する場合は NOTIFY_CONFIG.ENABLED を true に戻す。
 */

// ============================================================
// Google Chat Webhook URL（補助金連絡スペース「案件通知」）
// 送客通知（form_to_sheet.gs）からも参照する
// ============================================================
const CHAT_WEBHOOK_URL = 'https://chat.googleapis.com/v1/spaces/AAAAJyNY9qM/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=K6-A-G82-5OYY-cvseY5PJTALjjpknyHicVO-eZarRQ';


// ============================================================
// 設定
// ============================================================
const NOTIFY_CONFIG = {
  // 通知の有効/無効（2026-04-17 村上さん依頼で一旦停止）
  ENABLED: false,

  // 通知するステータス
  NOTIFY_STATUSES: ['申請_準備完了'],

  // 除外する支援事業者（通知しない）
  EXCLUDE_VENDORS: ['クラフトバンク'],

  // シート1の列番号（1始まり）
  COL_CB_STAFF: 1,      // A列: CB担当
  COL_COMPANY: 2,        // B列: お客様企業名
  COL_VENDOR: 3,         // C列: 構成員/支援事業者名
  COL_TEMPLATE: 4,       // D列: 申請枠
  COL_STATUS: 6,         // F列: ステータス

  // 対象シート名
  SHEET_NAME: '2026案件一覧',

  // ヘッダー行
  HEADER_ROW: 2,
};


// ============================================================
// メイン処理
// ============================================================

/**
 * シート編集時にステータス変更を検知して通知
 * @param {Object} e - 編集イベント
 */
function onEditNotify(e) {
  try {
    // 通知が無効化されている場合は即終了
    if (!NOTIFY_CONFIG.ENABLED) return;

    const sheet = e.source.getActiveSheet();

    // シート1以外は無視
    if (sheet.getName() !== NOTIFY_CONFIG.SHEET_NAME) return;

    const range = e.range;
    const row = range.getRow();
    const col = range.getColumn();

    // ヘッダー行以下、F列（ステータス）の変更のみ対象
    if (row <= NOTIFY_CONFIG.HEADER_ROW) return;
    if (col !== NOTIFY_CONFIG.COL_STATUS) return;

    const newValue = e.value;

    // 通知対象のステータスか確認
    if (!NOTIFY_CONFIG.NOTIFY_STATUSES.includes(newValue)) return;

    // 行データを取得
    const companyName = sheet.getRange(row, NOTIFY_CONFIG.COL_COMPANY).getDisplayValue();
    const vendor = sheet.getRange(row, NOTIFY_CONFIG.COL_VENDOR).getDisplayValue();
    const template = sheet.getRange(row, NOTIFY_CONFIG.COL_TEMPLATE).getDisplayValue();
    const cbStaff = sheet.getRange(row, NOTIFY_CONFIG.COL_CB_STAFF).getDisplayValue();

    // 除外する支援事業者か確認
    for (const exclude of NOTIFY_CONFIG.EXCLUDE_VENDORS) {
      if (vendor.includes(exclude)) {
        Logger.log(`通知除外: ${vendor} / ${companyName}`);
        return;
      }
    }

    // Driveフォルダリンクを取得（B列のハイパーリンク）
    const richText = sheet.getRange(row, NOTIFY_CONFIG.COL_COMPANY).getRichTextValue();
    let folderUrl = '';
    if (richText) {
      folderUrl = richText.getLinkUrl() || '';
    }

    // Chat通知を送信
    sendChatNotification_(companyName, vendor, template, cbStaff, newValue, folderUrl);

  } catch (error) {
    Logger.log(`通知エラー: ${error.message}`);
  }
}


/**
 * Google Chatにメッセージを送信
 */
function sendChatNotification_(companyName, vendor, template, cbStaff, status, folderUrl) {
  let message = `📋 *案件ステータス更新*\n\n`;
  message += `*企業名:* ${companyName}\n`;
  message += `*支援事業者:* ${vendor}\n`;
  message += `*申請枠:* ${template}\n`;
  message += `*CB担当:* ${cbStaff || '未割当'}\n`;
  message += `*ステータス:* ${status}\n`;

  if (folderUrl) {
    message += `\n📁 <${folderUrl}|Driveフォルダを開く>`;
  }

  const payload = {
    text: message,
  };

  const options = {
    method: 'post',
    contentType: 'application/json; charset=UTF-8',
    payload: JSON.stringify(payload),
  };

  UrlFetchApp.fetch(CHAT_WEBHOOK_URL, options);
  Logger.log(`Chat通知送信: ${companyName} → ${status}`);
}


/**
 * テスト用: 通知の送信テスト
 */
function testChatNotification() {
  sendChatNotification_(
    'テスト株式会社',
    '株式会社ピスケス',
    '通常枠（5万円～150万円未満）',
    '村上',
    '申請_準備完了',
    'https://drive.google.com/drive/folders/example'
  );
}
