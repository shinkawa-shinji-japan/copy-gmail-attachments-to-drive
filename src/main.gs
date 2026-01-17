/**
 * Gmail添付ファイルをGoogleドライブにコピーするスクリプト
 * メイン処理とUI制御
 */

/**
 * Google Sheets起動時にカスタムメニューを追加
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Gmail添付ファイル")
    .addItem("初期化（シート作成）", "initializeSheets")
    .addSeparator()
    .addItem("検索して一覧表示", "searchAndDisplay")
    .addItem("選択対象をGoogleドライブにコピー", "executeAndCopy")
    .addToUi();
}

/**
 * 検索フロー実行関数
 * 検索条件シートから条件を取得し、メールを検索して結果シートに表示
 */
function searchAndDisplay() {
  try {
    logDebug("=== 検索フロー開始 ===");

    // 1. 検索条件を取得
    const conditions = getSearchConditions();
    logDebug("検索条件を取得", {
      startDate: formatDateTime(conditions.startDate),
      endDate: formatDateTime(conditions.endDate),
      keywords: conditions.keywords,
      folderPath: conditions.folderPath,
    });

    // 2. 入力値を検証
    validateSearchConditions(conditions);

    // 3. メール検索を実行
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

    // 4. 検索結果を処理
    const resultsData = [];
    threads.forEach((thread, index) => {
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

    // 5. 結果シートをクリアして書き込み
    clearResultsSheet();
    addResultRows(resultsData);

    // 6. 完了通知
    SpreadsheetApp.getUi().alert(
      `検索完了！\n\n${resultsData.length}件のメールが見つかりました。\n` +
        "保存対象のチェック、ファイル名、フォルダを確認してから、\n" +
        "「選択対象をGoogleドライブにコピー」を実行してください。",
    );

    logDebug("=== 検索フロー完了 ===");
  } catch (error) {
    logError("検索フローエラー", error.message);
    SpreadsheetApp.getUi().alert("エラーが発生しました:\n" + error.message);
  }
}

/**
 * 実行フロー実行関数
 * 結果シートから処理対象を取得し、ファイルをGoogleドライブにコピー
 */
function executeAndCopy() {
  try {
    logDebug("=== 実行フロー開始 ===");

    // 1. 処理対象データを取得
    const targetRows = getTargetRowsData();
    logDebug("処理対象データ取得", `${targetRows.length}件の処理対象`);

    if (targetRows.length === 0) {
      SpreadsheetApp.getUi().alert("処理対象が選択されていません。");
      return;
    }

    // 2. 各メールについてファイルをコピー
    let successCount = 0;
    let skipCount = 0;
    let errorCount = 0;

    targetRows.forEach((row, index) => {
      try {
        logDebug(`処理中 (${index + 1}/${targetRows.length}): ${row.title}`);

        // フォルダパスをフォルダIDに解決
        const folderId = resolveFolderPath(row.saveFolder);

        // メール件名からメッセージを取得（簡易実装）
        // 実際にはメールIDを持つ必要があります
        const threads = GmailApp.search(`subject:"${row.title}"`);

        if (threads.length === 0) {
          throw new Error("メールが見つかりません");
        }

        const thread = threads[0];
        const messages = thread.getMessages();
        const message = messages[messages.length - 1];
        const emailData = extractEmailAndAttachments(thread);

        if (!emailData || emailData.attachments.length === 0) {
          throw new Error("PDF添付ファイルが見つかりません");
        }

        // 3. ファイルをコピー
        const copyResults = copyMultipleFiles(
          emailData.attachments,
          folderId,
          row.saveFileName || null,
        );

        // 4. 結果を処理
        copyResults.forEach((copyResult) => {
          if (copyResult.success) {
            // ファイルリンクを生成
            const fileLink = generateDriveLink(copyResult.fileId);

            updateResultRow(row.rowIndex, {
              result: "OK",
              fileLink: fileLink,
              saveFileName: copyResult.fileName,
            });

            successCount++;
            logDebug("ファイルコピー成功", copyResult.fileName);
          } else {
            // スキップまたはエラー
            const resultText = copyResult.error.includes("既に存在")
              ? "SKIP"
              : "エラー";
            updateResultRow(row.rowIndex, {
              result: resultText + ": " + copyResult.error,
            });

            if (resultText === "SKIP") {
              skipCount++;
            } else {
              errorCount++;
            }

            logWarn("ファイルコピー失敗", copyResult.error);
          }
        });
      } catch (error) {
        updateResultRow(row.rowIndex, {
          result: "エラー: " + error.message,
        });

        errorCount++;
        logError(`処理エラー (行${row.rowIndex})`, error.message);
      }
    });

    // 5. 完了通知
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
    logError("実行フロー全体エラー", error.message);
    SpreadsheetApp.getUi().alert("エラーが発生しました:\n" + error.message);
  }
}

/**
 * 検索条件を検証
 * @param {Object} conditions - 検索条件オブジェクト
 * @throws {Error} - 検証失敗時
 */
function validateSearchConditions(conditions) {
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
