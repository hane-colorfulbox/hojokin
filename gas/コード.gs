function onFormSubmit(e) {
  // --- 設定エリア ---
  const targetSsId = '1YKHps9kq7gQ9kZIXXyiukfq_qMHN56NxP1J_f-9hQpU';
  const targetSheetName = '2026案件一覧'; // 転記先のシート名
  // ----------------

  const responses = e.values;
  if (!responses) return;

  // 1. 各データの抽出（ご提示の列インデックスに合わせて調整）
  // e.values[0]がA列, [7]がH列...となります
  const supportCompany = responses[6]; // G列：支援事業者名（貴社名）
  const supportStaff  = responses[7];  // H列：支援事業者（貴社）担当者氏名
  const supportEmail  = responses[8];  // I列：担当者メールアドレス
  const supportTel    = responses[9];  // J列：担当者電話番号
  const shinseiWaku   = responses[10]; // K列：申請枠を選択してください
  const clientName    = responses[11]; // L列：お客様企業名
  const clientStaff   = responses[12]; // M列：お客様担当者氏名
  const clientEmail   = responses[13]; // N列：お客様担当者メールアドレス
  const clientTel     = responses[14]; // O列：お客様担当者電話番号
  const contactTool   = responses[15]; // P列：お客様連絡希望ツール

  // 2. 転記先「シート1」の列構成に合わせた1行分のデータ作成（A列〜N列）
  // A=0, B=1, C=2...
  let rowData = new Array(14).fill("");

  rowData[1]  = clientName;      // B列：お客様企業名
  rowData[2]  = supportCompany; // C列：構成員/支援事業者名
  rowData[3]  = shinseiWaku;    // D列：申請枠
  rowData[7]  = contactTool;   // H列：希望連絡ツール
  rowData[8]  = clientStaff;   // I列：お客様担当者氏名
  rowData[9]  = clientEmail;   // J列：お客様担当者メールアドレス
  rowData[10] = clientTel;     // K列：お客様担当者電話番号
  rowData[11] = supportStaff;  // L列：支援事業者担当者名
  rowData[12] = supportEmail;  // M列：支援事業者メールアドレス
  rowData[13] = supportTel;    // N列：支援事業者電話番号

  // 3. 転記先シートへの書き込み処理
  const targetSs = SpreadsheetApp.openById(targetSsId);
  const targetSheet = targetSs.getSheetByName(targetSheetName);

  if (targetSheet) {
    // B列（お客様企業名）の実データで最終行を判定
    // getLastRow()はチェックボックスや数式がある行も含めてしまうため、
    // 実際にデータが入っているB列で判定する
    const newRow = getLastDataRow_(targetSheet, 2) + 1;

    // データを書き込み
    targetSheet.getRange(newRow, 1, 1, rowData.length).setValues([rowData]);

    // 上の行からプルダウン（データの入力規則）をコピー
    copyDataValidations_(targetSheet, newRow);

    // チェックボックスと数式も上の行からコピー
    copyCheckboxesAndFormulas_(targetSheet, newRow);

    Logger.log('転記完了: 行' + newRow + ' / ' + clientName);
  } else {
    console.error("転記先シートが見つかりません: " + targetSheetName);
  }
}


/**
 * 指定列の実データがある最終行を取得
 * （空文字・null・undefinedを除く）
 * @param {Sheet} sheet - 対象シート
 * @param {number} col - 列番号（1始まり）
 * @returns {number} 最終データ行（ヘッダー行=2を最低値とする）
 */
function getLastDataRow_(sheet, col) {
  const values = sheet.getRange(1, col, sheet.getLastRow(), 1).getValues();
  for (let i = values.length - 1; i >= 0; i--) {
    if (values[i][0] !== '' && values[i][0] !== null) {
      return i + 1;
    }
  }
  return 2; // ヘッダー行
}


/**
 * 上の行からデータの入力規則（プルダウン）を新しい行にコピー
 * @param {Sheet} sheet - 対象シート
 * @param {number} targetRow - コピー先の行番号
 */
function copyDataValidations_(sheet, targetRow) {
  // 入力規則がある行を探す（3行目を基準とする）
  const sourceRow = 3;
  const maxCol = sheet.getLastColumn();

  const sourceRange = sheet.getRange(sourceRow, 1, 1, maxCol);
  const validations = sourceRange.getDataValidations()[0];

  // 入力規則がある列だけ新しい行に設定
  // フォームからのデータが選択肢にない場合もあるため、警告モードで設定
  for (let col = 0; col < validations.length; col++) {
    if (validations[col] !== null) {
      const rule = validations[col].copy().setAllowInvalid(true).build();
      sheet.getRange(targetRow, col + 1).setDataValidation(rule);
    }
  }
}


/**
 * 上の行からチェックボックスと数式を新しい行にコピー
 * @param {Sheet} sheet - 対象シート
 * @param {number} targetRow - コピー先の行番号
 */
function copyCheckboxesAndFormulas_(sheet, targetRow) {
  const sourceRow = 3;
  const maxCol = sheet.getLastColumn();

  const sourceRange = sheet.getRange(sourceRow, 1, 1, maxCol);
  const formulas = sourceRange.getFormulas()[0];
  const values = sourceRange.getValues()[0];
  const validations = sourceRange.getDataValidations()[0];

  for (let col = 0; col < formulas.length; col++) {
    const cell = sheet.getRange(targetRow, col + 1);

    if (formulas[col] !== '') {
      // 数式がある列 → 数式をコピー（行番号を自動調整するためR1C1ではなくコピーを使う）
      sheet.getRange(sourceRow, col + 1).copyTo(cell, SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
    } else if (values[col] === true || values[col] === false) {
      // チェックボックス（Boolean値）→ チェックボックスを挿入してFalseで初期化
      cell.insertCheckboxes();
      cell.setValue(false);
    }
  }
}
