/**
 * ユーティリティ関数集
 */

/**
 * 日付文字列（YYYY-MM-DD）をDateオブジェクトに変換
 * @param {string} dateString - 日付文字列（YYYY-MM-DD形式）
 * @returns {Date} - Dateオブジェクト
 * @throws {Error} - 無効な日付形式の場合
 */
function validateDateInput(dateString) {
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
 * @param {Date} date - Dateオブジェクト
 * @returns {string} - フォーマットされた日付文字列
 */
function formatDateTime(date) {
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
 * @param {string} fileName - ファイル名
 * @returns {string} - .pdf拡張子が付いたファイル名
 */
function ensurePdfExtension(fileName) {
  if (!fileName) {
    throw new Error("ファイル名が空です");
  }

  if (!fileName.toLowerCase().endsWith(".pdf")) {
    return fileName + ".pdf";
  }

  return fileName;
}

/**
 * キーワード文字列（カンマ区切り）を配列に分割
 * @param {string} keywordString - キーワード文字列（カンマ区切り）
 * @returns {string[]} - キーワード配列
 */
function parseKeywords(keywordString) {
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
 * @param {string} fileId - Google Drive ファイルID
 * @returns {string} - Google Drive ファイルリンク
 */
function createDriveLink(fileId) {
  if (!fileId) {
    throw new Error("ファイルIDが空です");
  }

  return `https://drive.google.com/file/d/${fileId}/view`;
}

/**
 * Loggerでデバッグログを出力
 * @param {string} message - ログメッセージ
 * @param {any} data - ログデータ（オプション）
 */
function logDebug(message, data = null) {
  if (data) {
    Logger.log(`[DEBUG] ${message}: ${JSON.stringify(data)}`);
  } else {
    Logger.log(`[DEBUG] ${message}`);
  }
}

/**
 * Loggerで警告ログを出力
 * @param {string} message - ログメッセージ
 * @param {any} data - ログデータ（オプション）
 */
function logWarn(message, data = null) {
  if (data) {
    Logger.log(`[WARN] ${message}: ${JSON.stringify(data)}`);
  } else {
    Logger.log(`[WARN] ${message}`);
  }
}

/**
 * Loggerでエラーログを出力
 * @param {string} message - ログメッセージ
 * @param {any} data - ログデータ（オプション）
 */
function logError(message, data = null) {
  if (data) {
    Logger.log(`[ERROR] ${message}: ${JSON.stringify(data)}`);
  } else {
    Logger.log(`[ERROR] ${message}`);
  }
}
