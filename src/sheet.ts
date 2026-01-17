/**
 * Google Sheets APIラッパー関数群 (TypeScript)
 */

const SEARCH_SHEET_NAME = "検索条件";
const RESULTS_SHEET_NAME = "結果";
const FOLDER_LIST_SHEET_NAME = "フォルダ一覧";

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
  folderLink?: string;
  result?: string;
  fileLink?: string;
  body?: string; // メール本文
};

type TargetRow = {
  rowIndex: number;
  title: string;
  attachmentNames: string;
  saveFileName: string;
  saveFolder: string;
  result: string; // 処理結果の列
};

type RowUpdate = {
  rowIndex: number;
  updateData: Partial<ResultRow>;
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
    "メール本文",
    "添付ファイル名",
    "保存ファイル名",
    "保存先フォルダ名",
    "フォルダリンク",
    "処理結果",
    "ファイルリンク",
  ];

  sheet.appendRow(headers);

  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight("bold").setBackground("#e8f0fe");

  sheet.setColumnWidth(1, 80);
  sheet.setColumnWidth(2, 250);
  sheet.setColumnWidth(3, 180);
  sheet.setColumnWidth(4, 400); // 本文用に幅を物初設定
  sheet.setColumnWidth(5, 200);
  sheet.setColumnWidth(6, 200);
  sheet.setColumnWidth(7, 200);
  sheet.setColumnWidth(8, 200);
  sheet.setColumnWidth(9, 150);
  sheet.setColumnWidth(10, 200);
}

function getSearchConditions(): SearchConditions {
  logDebug("getSearchConditions開始");

  const sheet = getSearchSheet();
  logDebug("シート取得完了", sheet.getName());

  const data = sheet.getDataRange().getValues();
  logDebug("=== シートからデータ取得 ===");
  logDebug(`データ範囲: ${data.length}行 x ${data[0]?.length || 0}列`);

  // すべてのデータをログ出力
  for (let i = 0; i < data.length; i++) {
    const col0 = data[i][0];
    const col1 = data[i][1];
    const col0Type = typeof col0;
    const col1Type = typeof col1;
    const col1ToString = Object.prototype.toString.call(col1);

    Logger.log(
      `行${i + 1}: キー=[${col0}](型:${col0Type}), 値=[${col1}](型:${col1Type}, toString:${col1ToString})`,
    );
  }

  const conditions: Partial<SearchConditions> = {};

  for (let i = 1; i < data.length; i++) {
    const key = String(data[i][0]);
    const rawValue = data[i][1];

    logDebug(`処理中: キー=[${key}], 値=[${rawValue}], 型=${typeof rawValue}`);

    switch (key) {
      case "期間（開始日）":
        logDebug("開始日の処理開始", rawValue);
        conditions.startDate = validateDateInput(rawValue);
        logDebug("開始日の処理完了", conditions.startDate);
        break;
      case "期間（終了日）":
        logDebug("終了日の処理開始", rawValue);
        conditions.endDate = validateDateInput(rawValue);
        logDebug("終了日の処理完了", conditions.endDate);
        break;
      case "フィルタ文字列":
        const filterValue =
          rawValue === null || rawValue === undefined ? "" : String(rawValue);
        conditions.keywords = parseKeywords(filterValue);
        break;
      case "保存先フォルダパス":
        const pathValue =
          rawValue === null || rawValue === undefined ? "" : String(rawValue);
        conditions.folderPath = pathValue.trim() || "/";
        break;
      default:
        logDebug(`未知のキー: ${key}`);
        break;
    }
  }

  logDebug("=== 最終的な検索条件 ===", conditions);
  return conditions as SearchConditions;
}

function clearResultsSheet(): void {
  const sheet = getResultsSheet();
  const lastRow = sheet.getLastRow();

  if (lastRow > 1) {
    sheet.deleteRows(2, lastRow - 1);
  }
}

function addResultRows(rowsData: ResultRow[]): void {
  if (rowsData.length === 0) {
    return;
  }

  const sheet = getResultsSheet();
  const startRow = sheet.getLastRow() + 1; // 追加開始行

  // すべての行データを2次元配列に変換
  const batchData = rowsData.map((rowData) => [
    rowData.targetCheckbox ?? true,
    rowData.title ?? "",
    rowData.date ? formatDateTime(rowData.date) : "",
    rowData.body ?? "",
    rowData.attachmentNames ?? "",
    rowData.saveFileName ?? "",
    rowData.saveFolder ?? "",
    rowData.folderLink ?? "",
    rowData.result ?? "",
    rowData.fileLink ?? "",
  ]);

  // バッチ処理ですべての行を一度に追加
  const endRow = startRow + rowsData.length - 1;
  sheet.getRange(startRow, 1, rowsData.length, 10).setValues(batchData);

  // 追加したすべての行の保存対象列にチェックボックスを設定
  sheet.getRange(startRow, 1, rowsData.length, 1).insertCheckboxes();
}

function updateResultRows(updates: RowUpdate[]): void {
  if (updates.length === 0) {
    return;
  }

  const sheet = getResultsSheet();
  const currentData = sheet.getDataRange().getValues();

  // すべての更新データを現在のデータに反映
  updates.forEach((update) => {
    const rowIdx = update.rowIndex - 1; // 0ベースにする

    if (update.updateData.result !== undefined) {
      currentData[rowIdx][8] = update.updateData.result; // 列9
    }

    if (update.updateData.fileLink !== undefined) {
      currentData[rowIdx][9] = update.updateData.fileLink; // 列10
    }

    if (update.updateData.saveFileName !== undefined) {
      currentData[rowIdx][5] = update.updateData.saveFileName; // 列6
    }

    if (update.updateData.saveFolder !== undefined) {
      currentData[rowIdx][6] = update.updateData.saveFolder; // 列7
    }

    if (update.updateData.folderLink !== undefined) {
      currentData[rowIdx][7] = update.updateData.folderLink; // 列8
    }
  });

  // バッチ処理ですべてのデータを一度に書き込み
  sheet.getDataRange().setValues(currentData);
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
        result: String(data[i][7] ?? ""), // 処理結果列（8列目）
      });
    }
  }

  return targetRows;
}

function initializeSheets(): void {
  try {
    const ss = getActiveSpreadsheet();

    // 既存のシートを削除してリセット
    const existingSearchSheet = ss.getSheetByName(SEARCH_SHEET_NAME);
    if (existingSearchSheet) {
      ss.deleteSheet(existingSearchSheet);
    }

    const existingResultsSheet = ss.getSheetByName(RESULTS_SHEET_NAME);
    if (existingResultsSheet) {
      ss.deleteSheet(existingResultsSheet);
    }

    const existingFolderSheet = ss.getSheetByName(FOLDER_LIST_SHEET_NAME);
    if (existingFolderSheet) {
      ss.deleteSheet(existingFolderSheet);
    }

    // 新しいシートを作成
    getSearchSheet();
    getResultsSheet();

    SpreadsheetApp.getUi().alert(
      "シートの初期化が完了しました！\n\n" +
        "検索条件シートで検索条件を入力してから、\n" +
        "メニューから「検索して一覧表示」を実行してください。",
    );

    logDebug("シート初期化完了（既存データをリセット）");
  } catch (error) {
    const err = error as Error;
    logError("シート初期化エラー", err.message);
    SpreadsheetApp.getUi().alert("エラーが発生しました: " + err.message);
  }
}

function displayFolderList(folders: FolderInfo[]): void {
  const ss = getActiveSpreadsheet();

  // 既存のシートがあれば削除
  const existingSheet = ss.getSheetByName(FOLDER_LIST_SHEET_NAME);
  if (existingSheet) {
    ss.deleteSheet(existingSheet);
  }

  // 新しいシートを作成
  const sheet = ss.insertSheet(FOLDER_LIST_SHEET_NAME);

  // ヘッダー行
  const headers = ["選択", "フォルダ名", "フォルダパス", "リンク"];
  sheet.appendRow(headers);

  // ヘッダーのスタイル設定
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight("bold").setBackground("#e8f0fe");

  // フォルダデータをバッチ処理で一度に追加
  if (folders.length > 0) {
    const batchData = folders.map((folder) => [
      false, // チェックボックス（デフォルトは未選択）
      folder.name,
      folder.path,
      folder.link,
    ]);

    sheet.getRange(2, 1, folders.length, 4).setValues(batchData);

    // チェックボックスを設定
    sheet.getRange(2, 1, folders.length, 1).insertCheckboxes();
  }

  // 列幅を調整
  sheet.setColumnWidth(1, 60);
  sheet.setColumnWidth(2, 250);
  sheet.setColumnWidth(3, 300);
  sheet.setColumnWidth(4, 300);

  // シートをアクティブにする
  sheet.activate();
}

function getSelectedFolderFromList(): FolderInfo | null {
  const ss = getActiveSpreadsheet();
  const sheet = ss.getSheetByName(FOLDER_LIST_SHEET_NAME);

  if (!sheet) {
    throw new Error(
      "フォルダ一覧シートが見つかりません。先にフォルダ一覧を表示してください。",
    );
  }

  const data = sheet.getDataRange().getValues();

  // ヘッダー行をスキップして、チェックされた行を探す
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === true) {
      return {
        id: "", // IDは保存していないので空
        name: String(data[i][1]),
        path: String(data[i][2]),
        link: String(data[i][3]),
      };
    }
  }

  return null;
}

function setFolderPathToSearchSheet(folderPath: string): void {
  const sheet = getSearchSheet();
  const data = sheet.getDataRange().getValues();

  // 「保存先フォルダパス」の行を探して更新
  for (let i = 1; i < data.length; i++) {
    const key = String(data[i][0]);
    if (key === "保存先フォルダパス") {
      sheet.getRange(i + 1, 2).setValue(folderPath);
      logDebug("保存先フォルダパス更新", folderPath);
      return;
    }
  }

  throw new Error(
    "検索条件シートに「保存先フォルダパス」の項目が見つかりません。",
  );
}
