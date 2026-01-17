/**
 * Google Sheets APIラッパー関数群 (TypeScript)
 */

const SEARCH_SHEET_NAME = "検索条件";
const RESULTS_SHEET_NAME = "結果";

type SearchConditions = {
  startDate: Date;
  endDate: Date;
  keywords: string[];
  folderPath: string;
};

type ResultRow = {
  targetCheckbox?: boolean;
  title?: string;
  date?: Date;
  attachmentNames?: string;
  saveFileName?: string;
  saveFolder?: string;
  result?: string;
  fileLink?: string;
};

type TargetRow = {
  rowIndex: number;
  title: string;
  attachmentNames: string;
  saveFileName: string;
  saveFolder: string;
};

function getActiveSpreadsheet(): GoogleAppsScript.Spreadsheet.Spreadsheet {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getSearchSheet(): GoogleAppsScript.Spreadsheet.Sheet {
  const ss = getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SEARCH_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SEARCH_SHEET_NAME, 0);
    initializeSearchSheet(sheet);
  }

  return sheet;
}

function getResultsSheet(): GoogleAppsScript.Spreadsheet.Sheet {
  const ss = getActiveSpreadsheet();
  let sheet = ss.getSheetByName(RESULTS_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(RESULTS_SHEET_NAME, 1);
    initializeResultsSheet(sheet);
  }

  return sheet;
}

function initializeSearchSheet(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
): void {
  const headers = ["項目", "値"];
  sheet.appendRow(headers);

  const today = new Date();
  const startDate = new Date(today.getFullYear(), today.getMonth(), 1);
  const endDate = today;

  const defaultValues: [string, string][] = [
    ["期間（開始日）", formatDateTime(startDate).split(" ")[0]],
    ["期間（終了日）", formatDateTime(endDate).split(" ")[0]],
    ["フィルタ文字列", ""],
    ["保存先フォルダパス", "/"],
  ];

  defaultValues.forEach((value) => {
    sheet.appendRow(value);
  });

  sheet.getRange(1, 1, 1, 2).setFontWeight("bold").setBackground("#e8f0fe");
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 300);
}

function initializeResultsSheet(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
): void {
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

  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight("bold").setBackground("#e8f0fe");

  sheet.setColumnWidth(1, 80);
  sheet.setColumnWidth(2, 250);
  sheet.setColumnWidth(3, 180);
  sheet.setColumnWidth(4, 200);
  sheet.setColumnWidth(5, 200);
  sheet.setColumnWidth(6, 200);
  sheet.setColumnWidth(7, 150);
  sheet.setColumnWidth(8, 200);

  sheet.getDataRange().insertCheckboxes();
}

function getSearchConditions(): SearchConditions {
  const sheet = getSearchSheet();
  const data = sheet.getDataRange().getValues();

  const conditions: Partial<SearchConditions> = {};

  for (let i = 1; i < data.length; i++) {
    const key = String(data[i][0]);
    const rawValue = data[i][1];
    const value =
      rawValue === null || rawValue === undefined ? "" : String(rawValue);

    switch (key) {
      case "期間（開始日）":
        conditions.startDate = validateDateInput(value);
        break;
      case "期間（終了日）":
        conditions.endDate = validateDateInput(value);
        break;
      case "フィルタ文字列":
        conditions.keywords = parseKeywords(value);
        break;
      case "保存先フォルダパス":
        conditions.folderPath = value.trim() || "/";
        break;
      default:
        break;
    }
  }

  return conditions as SearchConditions;
}

function clearResultsSheet(): void {
  const sheet = getResultsSheet();
  const lastRow = sheet.getLastRow();

  if (lastRow > 1) {
    sheet.deleteRows(2, lastRow - 1);
  }
}

function addResultRow(rowData: ResultRow): void {
  const sheet = getResultsSheet();

  const row = [
    rowData.targetCheckbox ?? true,
    rowData.title ?? "",
    rowData.date ? formatDateTime(rowData.date) : "",
    rowData.attachmentNames ?? "",
    rowData.saveFileName ?? "",
    rowData.saveFolder ?? "",
    rowData.result ?? "",
    rowData.fileLink ?? "",
  ];

  sheet.appendRow(row);
}

function addResultRows(rowsData: ResultRow[]): void {
  rowsData.forEach((rowData) => addResultRow(rowData));
}

function updateResultRow(
  rowIndex: number,
  updateData: Partial<ResultRow>,
): void {
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

function getTargetRowsData(): TargetRow[] {
  const sheet = getResultsSheet();
  const data = sheet.getDataRange().getValues();

  const targetRows: TargetRow[] = [];

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === true) {
      targetRows.push({
        rowIndex: i + 1,
        title: String(data[i][1] ?? ""),
        attachmentNames: String(data[i][3] ?? ""),
        saveFileName: String(data[i][4] ?? ""),
        saveFolder: String(data[i][5] ?? ""),
      });
    }
  }

  return targetRows;
}

function initializeSheets(): void {
  try {
    getSearchSheet();
    getResultsSheet();

    SpreadsheetApp.getUi().alert(
      "シートの初期化が完了しました！\n\n" +
        "検索条件シートで検索条件を入力してから、\n" +
        "メニューから「検索して一覧表示」を実行してください。",
    );

    logDebug("シート初期化完了");
  } catch (error) {
    const err = error as Error;
    logError("シート初期化エラー", err.message);
    SpreadsheetApp.getUi().alert("エラーが発生しました: " + err.message);
  }
}
