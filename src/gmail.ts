/**
 * Gmail APIラッパー関数群 (TypeScript)
 */

type PdfAttachment = {
  name: string;
  blob: GoogleAppsScript.Base.Blob;
};

type EmailData = {
  subject: string;
  date: Date;
  attachments: PdfAttachment[];
  messageId: string;
  body: string; // メール本文
};

function searchEmails(
  startDate: Date,
  endDate: Date,
  keywords: string[],
): GoogleAppsScript.Gmail.GmailThread[] {
  try {
    const query = buildSearchQuery(startDate, endDate, keywords);
    logDebug("Gmail検索クエリ", query);

    const threads = GmailApp.search(query);
    logDebug("検索結果", `${threads.length}件のスレッドが見つかりました`);

    return threads;
  } catch (error) {
    const err = error as Error;
    logError("Gmail検索エラー", err.message);
    throw new Error("Gmail検索に失敗しました: " + err.message);
  }
}

function buildSearchQuery(
  startDate: Date,
  endDate: Date,
  keywords: string[],
): string {
  // ユーザーのタイムゾーンを取得
  const userTimezone =
    SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  logDebug("buildSearchQuery", { userTimezone, startDate, endDate });

  // 入力された日付をユーザーのタイムゾーンで日付文字列にフォーマット
  const startStr = Utilities.formatDate(startDate, userTimezone, "yyyy/MM/dd");
  const endStr = Utilities.formatDate(endDate, userTimezone, "yyyy/MM/dd");

  logDebug("フォーマット後の日付", { startStr, endStr });

  // has:attachment で添付ファイル有無、filename:pdf でPDF形式に限定
  let query = `has:attachment filename:pdf after:${startStr} before:${endStr}`;

  if (keywords && keywords.length > 0) {
    const keywordClause = keywords.map((kw) => `subject:(${kw})`).join(" OR ");
    query += ` (${keywordClause})`;
  }

  logDebug("最終Gmail検索クエリ", query);
  return query;
}

function extractEmailAndAttachments(
  thread: GoogleAppsScript.Gmail.GmailThread,
): EmailData | null {
  try {
    const messages = thread.getMessages();
    if (messages.length === 0) {
      return null;
    }

    const message = messages[messages.length - 1];

    const subject = message.getSubject();
    const date = new Date(message.getDate().getTime());
    const attachments = getPdfAttachments(message);
    const body = message.getPlainBody();

    if (attachments.length === 0) {
      return null;
    }

    return {
      subject,
      date,
      attachments,
      messageId: message.getId(),
      body,
    };
  } catch (error) {
    const err = error as Error;
    logError("メール情報抽出エラー", err.message);
    return null;
  }
}

function getPdfAttachments(
  message: GoogleAppsScript.Gmail.GmailMessage,
): PdfAttachment[] {
  try {
    const allAttachments = message.getAttachments();
    const pdfAttachments: PdfAttachment[] = [];

    allAttachments.forEach((attachment) => {
      const fileName = attachment.getName();

      if (fileName.toLowerCase().endsWith(".pdf")) {
        pdfAttachments.push({
          name: fileName,
          blob: attachment.copyBlob(),
        });
      }
    });

    return pdfAttachments;
  } catch (error) {
    const err = error as Error;
    logError("添付ファイル取得エラー", err.message);
    return [];
  }
}

function getAttachmentNamesAsString(emailData: EmailData | null): string {
  if (!emailData || !emailData.attachments) {
    return "";
  }

  return emailData.attachments.map((att) => att.name).join(", ");
}
