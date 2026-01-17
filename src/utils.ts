/**
 * ユーティリティ関数集 (TypeScript)
 */

/**
 * 日付文字列（YYYY-MM-DD）またはDateオブジェクトをDateオブジェクトに変換
 * @throws Error 無効な日付形式の場合
 */
function validateDateInput(dateInput: Date | string): Date {
  // Date オブジェクトの検出（より堅牢な方法）
  // Google Apps Script では instanceof Date が正しく動作しないことがあるため、
  // Object.prototype.toString を使用
  const isDateObject =
    Object.prototype.toString.call(dateInput) === "[object Date]";

  if (isDateObject) {
    const dateObj = dateInput as Date;
    if (!isNaN(dateObj.getTime())) {
      return dateObj;
    } else {
      throw new Error("無効な日付オブジェクトです");
    }
  }

  // 文字列の場合は YYYY-MM-DD フォーマットをチェック
  const dateString = String(dateInput).trim();
  const regex = /^\d{4}-\d{2}-\d{2}$/;
  if (!regex.test(dateString)) {
    throw new Error(
      `無効な日付形式です。YYYY-MM-DD形式で入力してください。入力値: ${dateString}`,
    );
  }

  const date = new Date(dateString);
  if (isNaN(date.getTime())) {
    throw new Error(`無効な日付です。入力値: ${dateString}`);
  }

  return date;
}

/**
 * Dateオブジェクトを YYYY-MM-DD HH:MM:SS 形式の文字列に変換
 */
function formatDateTime(date: Date): string {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  const hours = String(date.getHours()).padStart(2, "0");
  const minutes = String(date.getMinutes()).padStart(2, "0");
  const seconds = String(date.getSeconds()).padStart(2, "0");

  return `${year}-${month}-${day} ${hours}:${minutes}:${seconds}`;
}

/**
 * ファイル名に.pdf拡張子が付いているか確認し、なければ付与
 */
function ensurePdfExtension(fileName: string): string {
  if (!fileName) {
    throw new Error("ファイル名が空です");
  }

  if (!fileName.toLowerCase().endsWith(".pdf")) {
    return `${fileName}.pdf`;
  }

  return fileName;
}

/**
 * キーワード文字列（カンマ区切り）を配列に分割
 */
function parseKeywords(keywordString: string): string[] {
  if (!keywordString || keywordString.trim() === "") {
    return [];
  }

  return keywordString
    .split(",")
    .map((keyword) => keyword.trim())
    .filter((keyword) => keyword.length > 0);
}

/**
 * Google Drive ファイルIDからアクセスリンクを生成
 */
function createDriveLink(fileId: string): string {
  if (!fileId) {
    throw new Error("ファイルIDが空です");
  }

  return `https://drive.google.com/file/d/${fileId}/view`;
}

/**
 * Loggerでデバッグログを出力
 */
function logDebug(message: string, data: unknown = null): void {
  if (data !== null && data !== undefined) {
    Logger.log(`[DEBUG] ${message}: ${JSON.stringify(data)}`);
  } else {
    Logger.log(`[DEBUG] ${message}`);
  }
}

/**
 * Loggerで警告ログを出力
 */
function logWarn(message: string, data: unknown = null): void {
  if (data !== null && data !== undefined) {
    Logger.log(`[WARN] ${message}: ${JSON.stringify(data)}`);
  } else {
    Logger.log(`[WARN] ${message}`);
  }
}

/**
 * Loggerでエラーログを出力
 */
function logError(message: string, data: unknown = null): void {
  if (data !== null && data !== undefined) {
    Logger.log(`[ERROR] ${message}: ${JSON.stringify(data)}`);
  } else {
    Logger.log(`[ERROR] ${message}`);
  }
}
