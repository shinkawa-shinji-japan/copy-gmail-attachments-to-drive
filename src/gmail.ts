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
  const startStr = Utilities.formatDate(startDate, "Asia/Tokyo", "yyyy/MM/dd");
  const endStr = Utilities.formatDate(endDate, "Asia/Tokyo", "yyyy/MM/dd");

  let query = `is:attachment after:${startStr} before:${endStr}`;

  if (keywords && keywords.length > 0) {
    const keywordClause = keywords.map((kw) => `subject:(${kw})`).join(" OR ");
    query += ` (${keywordClause})`;
  }

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

    if (attachments.length === 0) {
      return null;
    }

    return {
      subject,
      date,
      attachments,
      messageId: message.getId(),
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
