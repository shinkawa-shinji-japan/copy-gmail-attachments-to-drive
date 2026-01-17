/**
 * Gmail APIラッパー関数群
 */

/**
 * 検索条件に基づいてメールスレッドを検索
 * @param {Date} startDate - 開始日
 * @param {Date} endDate - 終了日（この日の23:59:59まで包含）
 * @param {string[]} keywords - キーワード配列
 * @returns {GmailThread[]} - メールスレッドの配列
 */
function searchEmails(startDate, endDate, keywords) {
  try {
    // 検索クエリを構築
    let query = buildSearchQuery(startDate, endDate, keywords);
    logDebug("Gmail検索クエリ", query);

    // メールを検索
    const threads = GmailApp.search(query);
    logDebug("検索結果", `${threads.length}件のスレッドが見つかりました`);

    return threads;
  } catch (error) {
    logError("Gmail検索エラー", error.message);
    throw new Error("Gmail検索に失敗しました: " + error.message);
  }
}

/**
 * Gmail検索クエリを構築
 * @param {Date} startDate - 開始日
 * @param {Date} endDate - 終了日
 * @param {string[]} keywords - キーワード配列
 * @returns {string} - Gmail検索クエリ
 */
function buildSearchQuery(startDate, endDate, keywords) {
  // 日付フォーマット: YYYY/MM/DD (Gmailは / 区切り)
  const startStr = Utilities.formatDate(startDate, "Asia/Tokyo", "yyyy/MM/dd");
  const endStr = Utilities.formatDate(endDate, "Asia/Tokyo", "yyyy/MM/dd");

  // 基本クエリ: 添付ファイル付き + 日付範囲
  let query = `is:attachment after:${startStr} before:${endStr}`;

  // キーワードを追加（複数の場合はOR条件）
  if (keywords && keywords.length > 0) {
    const keywordClause = keywords.map((kw) => `subject:(${kw})`).join(" OR ");
    query += ` (${keywordClause})`;
  }

  return query;
}

/**
 * スレッドからメール情報とPDF添付ファイルを抽出
 * @param {GmailThread} thread - Gmailスレッド
 * @returns {Object} - メール情報オブジェクト
 *   {
 *     subject: string,
 *     date: Date,
 *     attachments: [{name: string, blobSource: Blob}]
 *   }
 */
function extractEmailAndAttachments(thread) {
  try {
    const messages = thread.getMessages();
    if (messages.length === 0) {
      return null;
    }

    // 最新のメッセージを取得
    const message = messages[messages.length - 1];

    const subject = message.getSubject();
    const date = message.getDate();
    const attachments = getPdfAttachments(message);

    if (attachments.length === 0) {
      // PDF添付ファイルがない場合はnullを返す
      return null;
    }

    return {
      subject: subject,
      date: date,
      attachments: attachments,
      messageId: message.getId(),
    };
  } catch (error) {
    logError("メール情報抽出エラー", error.message);
    return null;
  }
}

/**
 * メッセージからPDF添付ファイルのみを取得
 * @param {GmailMessage} message - Gmailメッセージ
 * @returns {Object[]} - PDF添付ファイルの配列 [{name: string, blob: Blob}]
 */
function getPdfAttachments(message) {
  try {
    const allAttachments = message.getAttachments();
    const pdfAttachments = [];

    allAttachments.forEach((attachment) => {
      const fileName = attachment.getFileName();

      // PDF形式のみをフィルタ
      if (fileName.toLowerCase().endsWith(".pdf")) {
        pdfAttachments.push({
          name: fileName,
          blob: attachment,
        });
      }
    });

    return pdfAttachments;
  } catch (error) {
    logError("添付ファイル取得エラー", error.message);
    return [];
  }
}

/**
 * メールからファイル名を取得（複数の場合はカンマ区切り）
 * @param {Object} emailData - メール情報オブジェクト
 * @returns {string} - ファイル名（カンマ区切り）
 */
function getAttachmentNamesAsString(emailData) {
  if (!emailData || !emailData.attachments) {
    return "";
  }

  return emailData.attachments.map((att) => att.name).join(", ");
}
