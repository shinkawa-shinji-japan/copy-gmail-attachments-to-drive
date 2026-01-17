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
    .addSeparator()
    .addItem("📁 フォルダ一覧を表示", "showFolderList")
    .addItem("✓ 選択フォルダを保存先に設定", "setSelectedFolder")
    .addToUi();
}

function searchAndDisplay(): void {
  try {
    logDebug("=== 検索フロー開始 ===");

    // 既存データをクリア
    clearResultsSheet();
    logDebug("既存データをクリア");

    logDebug("getSearchConditions呼び出し直前");
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

    // 保存先フォルダのIDとリンクを取得
    const folderId = resolveFolderPath(conditions.folderPath);
    const folderLink = generateFolderLink(folderId);

    const resultsData: ResultRow[] = [];
    threads.forEach((thread) => {
      const emailData = extractEmailAndAttachments(thread);

      if (emailData) {
        // 各添付ファイルにつき1行を作成
        emailData.attachments.forEach((attachment) => {
          resultsData.push({
            targetCheckbox: true,
            title: emailData.subject,
            date: emailData.date,
            body: emailData.body,
            attachmentNames: attachment.name,
            saveFileName: "",
            saveFolder: conditions.folderPath,
            folderLink: folderLink,
            result: "",
            fileLink: "",
          });
        });
      }
    });

    logDebug("抽出データ", `${resultsData.length}件の添付ファイル`);

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
    if (err.stack) {
      logError("スタックトレース", err.stack);
    }
    SpreadsheetApp.getUi().alert("エラーが発生しました:\n" + err.message);
  }
}

function executeAndCopy(): void {
  try {
    logDebug("=== 実行フロー開始 ===");

    // 検索条件を取得（デフォルトフォルダパス用）
    const conditions = getSearchConditions();

    const targetRows = getTargetRowsData();
    logDebug("処理対象データ取得", `${targetRows.length}件の処理対象`);

    if (targetRows.length === 0) {
      SpreadsheetApp.getUi().alert("処理対象が選択されていません。");
      return;
    }

    let successCount = 0;
    let skipCount = 0;
    let errorCount = 0;
    let skippedByResultCount = 0;

    // すべての行の処理結果をここで蓄積
    const updates: RowUpdate[] = [];

    targetRows.forEach((row, index) => {
      try {
        logDebug(`処理中 (${index + 1}/${targetRows.length}): ${row.title}`);

        // 処理結果の列が空でない場合はスキップ
        if (row.result && row.result.trim() !== "") {
          logDebug(
            `スキップ（既に処理済み） (行${row.rowIndex})`,
            `現在の結果: ${row.result}`,
          );
          skippedByResultCount++;
          return;
        }

        // 保存先フォルダの決定: 結果シートの値が空欄なら検索条件シートの値を使用
        const targetFolder =
          row.saveFolder && row.saveFolder.trim() !== ""
            ? row.saveFolder
            : conditions.folderPath;

        const folderId = resolveFolderPath(targetFolder);
        const folderLink = generateFolderLink(folderId);

        const threads = GmailApp.search(`subject:"${row.title}"`);

        if (threads.length === 0) {
          throw new Error("メールが見つかりません");
        }

        const thread = threads[0];
        const emailData = extractEmailAndAttachments(thread);

        if (!emailData || emailData.attachments.length === 0) {
          throw new Error("PDF添付ファイルが見つかりません");
        }

        // 添付ファイル名でフィルタ（行の添付ファイル名と一致するものを処理）
        const targetAttachment = emailData.attachments.find(
          (att) => att.name === row.attachmentNames,
        );

        if (!targetAttachment) {
          throw new Error(
            `指定された添付ファイルが見つかりません: ${row.attachmentNames}`,
          );
        }

        // 保存ファイル名の決定
        const saveFileName =
          row.saveFileName && row.saveFileName.trim() !== ""
            ? row.saveFileName
            : targetAttachment.name;

        const copyResult = copyFileToFolder(
          targetAttachment.blob,
          folderId,
          saveFileName,
        );

        if (copyResult.success && copyResult.fileId && copyResult.fileName) {
          const fileLink = generateDriveLink(copyResult.fileId);

          updates.push({
            rowIndex: row.rowIndex,
            updateData: {
              result: "OK",
              fileLink,
              saveFileName: copyResult.fileName,
              folderLink: folderLink,
            },
          });

          successCount++;
          logDebug("ファイルコピー成功", copyResult.fileName);
        } else {
          const errorMessage = copyResult.error ?? "不明なエラー";
          const resultText = errorMessage.includes("既に存在")
            ? "SKIP"
            : "エラー";
          updates.push({
            rowIndex: row.rowIndex,
            updateData: {
              result: `${resultText}: ${errorMessage}`,
            },
          });

          if (resultText === "SKIP") {
            skipCount++;
          } else {
            errorCount++;
          }

          logWarn("ファイルコピー失敗", errorMessage);
        }
      } catch (error) {
        const err = error as Error;
        updates.push({
          rowIndex: row.rowIndex,
          updateData: {
            result: "エラー: " + err.message,
          },
        });

        errorCount++;
        logError(`処理エラー (行${row.rowIndex})`, err.message);
      }
    });

    // バッチ処理ですべての更新を一度に適用
    updateResultRows(updates);

    const message =
      `実行完了！\n\n` +
      `成功: ${successCount}件\n` +
      `スキップ: ${skipCount}件\n` +
      `エラー: ${errorCount}件\n` +
      `既に処理済み: ${skippedByResultCount}件`;

    SpreadsheetApp.getUi().alert(message);

    logDebug("=== 実行フロー完了 ===", {
      successCount,
      skipCount,
      errorCount,
      skippedByResultCount,
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

function showFolderList(): void {
  try {
    const ui = SpreadsheetApp.getUi();

    const result = ui.alert(
      "フォルダ一覧を取得",
      "Google Driveのフォルダ一覧を取得して表示します。\n" +
        "フォルダが多い場合、時間がかかる場合があります。\n\n" +
        "続行しますか？",
      ui.ButtonSet.YES_NO,
    );

    if (result !== ui.Button.YES) {
      return;
    }

    logDebug("=== フォルダ一覧取得開始 ===");

    const folders = listDriveFolders();
    logDebug(`フォルダ取得完了: ${folders.length}件`);

    displayFolderList(folders);

    ui.alert(
      "完了！\n\n" +
        `${folders.length}個のフォルダが見つかりました。\n` +
        "\n" +
        "フォルダ一覧シートで保存先にしたいフォルダを選択（チェック）してから、\n" +
        "メニューから「✓ 選択フォルダを保存先に設定」を実行してください。",
    );

    logDebug("=== フォルダ一覧表示完了 ===");
  } catch (error) {
    const err = error as Error;
    logError("フォルダ一覧取得エラー", err.message);
    SpreadsheetApp.getUi().alert("エラーが発生しました:\n" + err.message);
  }
}

function setSelectedFolder(): void {
  try {
    const selectedFolder = getSelectedFolderFromList();

    if (!selectedFolder) {
      SpreadsheetApp.getUi().alert(
        "フォルダが選択されていません。\n\n" +
          "フォルダ一覧シートで1つのフォルダにチェックを入れてから、\n" +
          "もう一度実行してください。",
      );
      return;
    }

    setFolderPathToSearchSheet(selectedFolder.path);

    SpreadsheetApp.getUi().alert(
      `保存先フォルダを設定しました！\n\n` +
        `フォルダ名: ${selectedFolder.name}\n` +
        `パス: ${selectedFolder.path}\n\n` +
        `検索条件シートの「保存先フォルダパス」が更新されました。`,
    );

    logDebug("フォルダパス設定完了", selectedFolder.path);
  } catch (error) {
    const err = error as Error;
    logError("フォルダ設定エラー", err.message);
    SpreadsheetApp.getUi().alert("エラーが発生しました:\n" + err.message);
  }
}
