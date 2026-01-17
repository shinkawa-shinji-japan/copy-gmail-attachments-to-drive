/**
 * Gmail添付ファイルをGoogleドライブにコピーするスクリプト (TypeScript)
 */

function onOpen(): void {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Gmail添付ファイル")
    .addItem("初期化（シート作成）", "initializeSheets")
    .addSeparator()
    .addItem("検索して一覧表示", "searchAndDisplay")
    .addItem("選択対象をGoogleドライブにコピー", "executeAndCopy")
    .addToUi();
}

function searchAndDisplay(): void {
  try {
    logDebug("=== 検索フロー開始 ===");

    const conditions = getSearchConditions();
    logDebug("検索条件を取得", {
      startDate: formatDateTime(conditions.startDate),
      endDate: formatDateTime(conditions.endDate),
      keywords: conditions.keywords,
      folderPath: conditions.folderPath,
    });

    validateSearchConditions(conditions);

    const threads = searchEmails(
      conditions.startDate,
      conditions.endDate,
      conditions.keywords,
    );

    if (threads.length === 0) {
      SpreadsheetApp.getUi().alert("条件に合うメールが見つかりませんでした。");
      logDebug("検索結果：メール0件");
      return;
    }

    logDebug("メール検索完了", `${threads.length}件のメール`);

    const resultsData: ResultRow[] = [];
    threads.forEach((thread) => {
      const emailData = extractEmailAndAttachments(thread);

      if (emailData) {
        resultsData.push({
          targetCheckbox: true,
          title: emailData.subject,
          date: emailData.date,
          attachmentNames: getAttachmentNamesAsString(emailData),
          saveFileName: "",
          saveFolder: conditions.folderPath,
          result: "",
          fileLink: "",
        });
      }
    });

    logDebug("抽出データ", `${resultsData.length}件のメール`);

    clearResultsSheet();
    addResultRows(resultsData);

    SpreadsheetApp.getUi().alert(
      `検索完了！\n\n${resultsData.length}件のメールが見つかりました。\n` +
        "保存対象のチェック、ファイル名、フォルダを確認してから、\n" +
        "「選択対象をGoogleドライブにコピー」を実行してください。",
    );

    logDebug("=== 検索フロー完了 ===");
  } catch (error) {
    const err = error as Error;
    logError("検索フローエラー", err.message);
    SpreadsheetApp.getUi().alert("エラーが発生しました:\n" + err.message);
  }
}

function executeAndCopy(): void {
  try {
    logDebug("=== 実行フロー開始 ===");

    const targetRows = getTargetRowsData();
    logDebug("処理対象データ取得", `${targetRows.length}件の処理対象`);

    if (targetRows.length === 0) {
      SpreadsheetApp.getUi().alert("処理対象が選択されていません。");
      return;
    }

    let successCount = 0;
    let skipCount = 0;
    let errorCount = 0;

    targetRows.forEach((row, index) => {
      try {
        logDebug(`処理中 (${index + 1}/${targetRows.length}): ${row.title}`);

        const folderId = resolveFolderPath(row.saveFolder);

        const threads = GmailApp.search(`subject:"${row.title}"`);

        if (threads.length === 0) {
          throw new Error("メールが見つかりません");
        }

        const thread = threads[0];
        const emailData = extractEmailAndAttachments(thread);

        if (!emailData || emailData.attachments.length === 0) {
          throw new Error("PDF添付ファイルが見つかりません");
        }

        const copyResults = copyMultipleFiles(
          emailData.attachments,
          folderId,
          row.saveFileName || null,
        );

        copyResults.forEach((copyResult) => {
          if (copyResult.success && copyResult.fileId && copyResult.fileName) {
            const fileLink = generateDriveLink(copyResult.fileId);

            updateResultRow(row.rowIndex, {
              result: "OK",
              fileLink,
              saveFileName: copyResult.fileName,
            });

            successCount++;
            logDebug("ファイルコピー成功", copyResult.fileName);
          } else {
            const errorMessage = copyResult.error ?? "不明なエラー";
            const resultText = errorMessage.includes("既に存在")
              ? "SKIP"
              : "エラー";
            updateResultRow(row.rowIndex, {
              result: `${resultText}: ${errorMessage}`,
            });

            if (resultText === "SKIP") {
              skipCount++;
            } else {
              errorCount++;
            }

            logWarn("ファイルコピー失敗", errorMessage);
          }
        });
      } catch (error) {
        const err = error as Error;
        updateResultRow(row.rowIndex, {
          result: "エラー: " + err.message,
        });

        errorCount++;
        logError(`処理エラー (行${row.rowIndex})`, err.message);
      }
    });

    const message =
      `実行完了！\n\n` +
      `成功: ${successCount}件\n` +
      `スキップ: ${skipCount}件\n` +
      `エラー: ${errorCount}件`;

    SpreadsheetApp.getUi().alert(message);

    logDebug("=== 実行フロー完了 ===", {
      successCount,
      skipCount,
      errorCount,
    });
  } catch (error) {
    const err = error as Error;
    logError("実行フロー全体エラー", err.message);
    SpreadsheetApp.getUi().alert("エラーが発生しました:\n" + err.message);
  }
}

function validateSearchConditions(conditions: SearchConditions): void {
  if (!conditions.startDate) {
    throw new Error("期間（開始日）が未入力です。");
  }

  if (!conditions.endDate) {
    throw new Error("期間（終了日）が未入力です。");
  }

  if (conditions.startDate > conditions.endDate) {
    throw new Error("開始日が終了日より後になっています。");
  }

  if (!conditions.folderPath) {
    throw new Error("保存先フォルダパスが未入力です。");
  }
}
