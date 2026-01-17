/**
 * Google Drive APIラッパー関数群
 */

/**
 * フォルダパスからフォルダIDを取得（なければ作成）
 * 例："/プロジェクト/PDFフォルダ" → フォルダID
 * @param {string} folderPath - フォルダパス
 * @returns {string} - フォルダID
 * @throws {Error} - フォルダ操作に失敗した場合
 */
function resolveFolderPath(folderPath) {
  try {
    if (!folderPath || folderPath === "/" || folderPath.trim() === "") {
      // ルートフォルダを返す
      return DriveApp.getRootFolder().getId();
    }

    // パスを / で分割し、空白を除去
    const parts = folderPath.split("/").filter((p) => p && p.trim().length > 0);

    let currentFolder = DriveApp.getRootFolder();

    for (const folderName of parts) {
      let nextFolder = null;

      // 同名フォルダを検索
      const subfolders = currentFolder.getFoldersByName(folderName);
      if (subfolders.hasNext()) {
        nextFolder = subfolders.next();
      } else {
        // フォルダが見つからなければ作成
        nextFolder = currentFolder.createFolder(folderName);
        logDebug(`フォルダを作成しました: ${folderName}`);
      }

      currentFolder = nextFolder;
    }

    return currentFolder.getId();
  } catch (error) {
    logError("フォルダパス解析エラー", error.message);
    throw new Error(`フォルダパス解析に失敗しました: ${error.message}`);
  }
}

/**
 * ファイルが指定フォルダに既に存在するかチェック
 * @param {string} folderId - フォルダID
 * @param {string} fileName - ファイル名（拡張子を含む）
 * @returns {boolean} - true: 存在、false: 存在しない
 */
function fileExistsInFolder(folderId, fileName) {
  try {
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFilesByName(fileName);
    return files.hasNext();
  } catch (error) {
    logError("ファイル存在チェックエラー", error.message);
    return false;
  }
}

/**
 * PDFファイルをフォルダにコピー
 * @param {Blob} fileBlob - ファイルのBlob
 * @param {string} folderId - 保存先フォルダID
 * @param {string} newFileName - 保存時のファイル名（拡張子を含む）
 * @returns {Object} - コピー結果
 *   {
 *     success: boolean,
 *     fileId: string (成功時のみ),
 *     fileName: string (成功時のみ),
 *     error: string (失敗時のみ)
 *   }
 */
function copyFileToFolder(fileBlob, folderId, newFileName) {
  try {
    // ファイル名の.pdf拡張子を確認
    const fileName = ensurePdfExtension(newFileName);

    const folder = DriveApp.getFolderById(folderId);

    // ファイルが既に存在するかチェック
    if (fileExistsInFolder(folderId, fileName)) {
      return {
        success: false,
        error: `ファイルが既に存在します: ${fileName}`,
      };
    }

    // ファイルを新規作成
    const createdFile = folder.createFile(fileBlob.setName(fileName));

    logDebug(`ファイルをコピーしました: ${fileName}`, createdFile.getId());

    return {
      success: true,
      fileId: createdFile.getId(),
      fileName: fileName,
    };
  } catch (error) {
    logError("ファイルコピーエラー", error.message);
    return {
      success: false,
      error: `ファイルコピーに失敗しました: ${error.message}`,
    };
  }
}

/**
 * 複数のファイルをフォルダにコピー
 * @param {Object[]} files - ファイルの配列
 *   [{name: string, blob: Blob}, ...]
 * @param {string} folderId - 保存先フォルダID
 * @param {string} [customFileNameBase] - カスタムファイル名の基底（複数ファイル時用）
 * @returns {Object[]} - コピー結果の配列
 */
function copyMultipleFiles(files, folderId, customFileNameBase = null) {
  const results = [];

  files.forEach((file, index) => {
    let fileName = file.name;

    // カスタムファイル名が指定されている場合
    if (customFileNameBase) {
      if (files.length === 1) {
        fileName = customFileNameBase;
      } else {
        // 複数ファイルの場合は番号を付与
        fileName =
          customFileNameBase.replace(/\.pdf$/i, "") + `_${index + 1}.pdf`;
      }
    }

    const result = copyFileToFolder(file.blob, folderId, fileName);
    results.push({
      originalName: file.name,
      ...result,
    });
  });

  return results;
}

/**
 * Google Drive ファイルIDからアクセスリンクを生成（ラッパー）
 * @param {string} fileId - Google Drive ファイルID
 * @returns {string} - Google Drive ファイルリンク
 */
function generateDriveLink(fileId) {
  return createDriveLink(fileId);
}
