/**
 * Google Sheets APIラッパー関数群
 */

const SEARCH_SHEET_NAME = "検索条件";
const RESULTS_SHEET_NAME = "結果";

/**
 * アクティブなSpreadsheetを取得
 * @returns {Spreadsheet} - アクティブなスプレッドシート
 */
function getActiveSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * 検索条件シートを取得（存在しなければ作成）
 * @returns {Sheet} - 検索条件シート
 */
function getSearchSheet() {
  const ss = getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SEARCH_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SEARCH_SHEET_NAME, 0);
    initializeSearchSheet(sheet);
  }

  return sheet;
}

/**
 * 結果シートを取得（存在しなければ作成）
 * @returns {Sheet} - 結果シート
 */
function getResultsSheet() {
  const ss = getActiveSpreadsheet();
  let sheet = ss.getSheetByName(RESULTS_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(RESULTS_SHEET_NAME, 1);
    initializeResultsSheet(sheet);
  }

  return sheet;
}

/**
 * 検索条件シートを初期化（ヘッダー行を設定）
 * @param {Sheet} sheet - 検索条件シート
 */
function initializeSearchSheet(sheet) {
  const headers = ["項目", "値"];
  sheet.appendRow(headers);

  // デフォルト値を追加
  const today = new Date();
  const startDate = new Date(today.getFullYear(), today.getMonth(), 1);
  const endDate = today;

  const defaultValues = [
    ["期間（開始日）", formatDateTime(startDate).split(" ")[0]],
    ["期間（終了日）", formatDateTime(endDate).split(" ")[0]],
    ["フィルタ文字列", ""],
    ["保存先フォルダパス", "/"],
  ];

  defaultValues.forEach((value) => {
    sheet.appendRow(value);
  });

  // スタイル調整
  sheet.getRange(1, 1, 1, 2).setFontWeight("bold").setBackground("#e8f0fe");
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 300);
}

/**
 * 結果シートを初期化（ヘッダー行を設定）
 * @param {Sheet} sheet - 結果シート
 */
function initializeResultsSheet(sheet) {
  const headers = [
    "保存対象",
    "メールタイトル",
    "受信日時",
    "添付ファイル名",
    "保存ファイル名",
    "保存先フォルダ名",
    "処理結果",
    "ファイルリンク",
  ];

  sheet.appendRow(headers);

  // ヘッダーをスタイル調整
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight("bold").setBackground("#e8f0fe");

  // 列幅を設定
  sheet.setColumnWidth(1, 80); // 保存対象
  sheet.setColumnWidth(2, 250); // メールタイトル
  sheet.setColumnWidth(3, 180); // 受信日時
  sheet.setColumnWidth(4, 200); // 添付ファイル名
  sheet.setColumnWidth(5, 200); // 保存ファイル名
  sheet.setColumnWidth(6, 200); // 保存先フォルダ名
  sheet.setColumnWidth(7, 150); // 処理結果
  sheet.setColumnWidth(8, 200); // ファイルリンク

  // 1列目にチェックボックスを追加（データ行用）
  sheet.getDataRange().insertCheckboxes();
}

/**
 * 検索条件を取得
 * @returns {Object} - 検索条件オブジェクト
 *   {
 *     startDate: Date,
 *     endDate: Date,
 *     keywords: string[],
 *     folderPath: string
 *   }
 */
function getSearchConditions() {
  const sheet = getSearchSheet();
  const data = sheet.getDataRange().getValues();

  const conditions = {};

  // ヘッダー行をスキップしてデータを取得
  for (let i = 1; i < data.length; i++) {
    const key = data[i][0];
    const value = data[i][1];

    switch (key) {
      case "期間（開始日）":
        conditions.startDate = validateDateInput(value.toString());
        break;
      case "期間（終了日）":
        conditions.endDate = validateDateInput(value.toString());
        break;
      case "フィルタ文字列":
        conditions.keywords = parseKeywords(value.toString());
        break;
      case "保存先フォルダパス":
        conditions.folderPath = value.toString().trim() || "/";
        break;
    }
  }

  return conditions;
}

/**
 * 結果シートをクリア（ヘッダーは保持）
 */
function clearResultsSheet() {
  const sheet = getResultsSheet();
  const lastRow = sheet.getLastRow();

  if (lastRow > 1) {
    sheet.deleteRows(2, lastRow - 1);
  }
}

/**
 * 結果シートに行を追加
 * @param {Object} rowData - 行データ
 *   {
 *     targetCheckbox: boolean,
 *     title: string,
 *     date: Date,
 *     attachmentNames: string,
 *     saveFileName: string,
 *     saveFolder: string,
 *     result: string,
 *     fileLink: string
 *   }
 */
function addResultRow(rowData) {
  const sheet = getResultsSheet();

  const row = [
    rowData.targetCheckbox || true,
    rowData.title || "",
    rowData.date ? formatDateTime(rowData.date) : "",
    rowData.attachmentNames || "",
    rowData.saveFileName || "",
    rowData.saveFolder || "",
    rowData.result || "",
    rowData.fileLink || "",
  ];

  sheet.appendRow(row);
}

/**
 * 結果シートに複数行を一括追加
 * @param {Object[]} rowsData - 行データの配列
 */
function addResultRows(rowsData) {
  const sheet = getResultsSheet();

  rowsData.forEach((rowData) => {
    addResultRow(rowData);
  });
}

/**
 * 結果シートの指定行を更新
 * @param {number} rowIndex - 行番号（1ベース、ヘッダーは1）
 * @param {Object} updateData - 更新データ
 */
function updateResultRow(rowIndex, updateData) {
  const sheet = getResultsSheet();

  if (updateData.result !== undefined) {
    sheet.getRange(rowIndex, 7).setValue(updateData.result);
  }

  if (updateData.fileLink !== undefined) {
    sheet.getRange(rowIndex, 8).setValue(updateData.fileLink);
  }

  if (updateData.saveFileName !== undefined) {
    sheet.getRange(rowIndex, 5).setValue(updateData.saveFileName);
  }

  if (updateData.saveFolder !== undefined) {
    sheet.getRange(rowIndex, 6).setValue(updateData.saveFolder);
  }
}

/**
 * 結果シートの処理対象データを取得（保存対象=trueのみ）
 * @returns {Object[]} - 処理対象データの配列
 */
function getTargetRowsData() {
  const sheet = getResultsSheet();
  const data = sheet.getDataRange().getValues();

  const targetRows = [];

  // ヘッダー行をスキップしてデータを取得
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === true) {
      // 保存対象列がON
      targetRows.push({
        rowIndex: i + 1, // スプレッドシートの行番号（1ベース）
        title: data[i][1],
        attachmentNames: data[i][3],
        saveFileName: data[i][4],
        saveFolder: data[i][5],
      });
    }
  }

  return targetRows;
}

/**
 * 初期構築関数：検索条件シートと結果シートを作成
 * ユーザーが初回に実行するべき関数
 */
function initializeSheets() {
  try {
    const searchSheet = getSearchSheet();
    const resultsSheet = getResultsSheet();

    SpreadsheetApp.getUi().alert(
      "シートの初期化が完了しました！\n\n" +
        "検索条件シートで検索条件を入力してから、\n" +
        "メニューから「検索して一覧表示」を実行してください。",
    );

    logDebug("シート初期化完了");
  } catch (error) {
    logError("シート初期化エラー", error.message);
    SpreadsheetApp.getUi().alert("エラーが発生しました: " + error.message);
  }
}
