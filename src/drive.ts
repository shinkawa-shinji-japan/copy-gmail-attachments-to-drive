/**
 * Google Drive APIラッパー関数群 (TypeScript)
 */

type CopyResult = {
  success: boolean;
  fileId?: string;
  fileName?: string;
  error?: string;
  originalName?: string;
};

type CopyFileInput = {
  name: string;
  blob: GoogleAppsScript.Base.Blob;
};

function resolveFolderPath(folderPath: string): string {
  try {
    if (!folderPath || folderPath === "/" || folderPath.trim() === "") {
      return DriveApp.getRootFolder().getId();
    }

    const parts = folderPath.split("/").filter((p) => p && p.trim().length > 0);

    let currentFolder = DriveApp.getRootFolder();

    for (const folderName of parts) {
      let nextFolder: GoogleAppsScript.Drive.Folder | null = null;

      const subfolders = currentFolder.getFoldersByName(folderName);
      if (subfolders.hasNext()) {
        nextFolder = subfolders.next();
      } else {
        nextFolder = currentFolder.createFolder(folderName);
        logDebug(`フォルダを作成しました: ${folderName}`);
      }

      currentFolder = nextFolder;
    }

    return currentFolder.getId();
  } catch (error) {
    const err = error as Error;
    logError("フォルダパス解析エラー", err.message);
    throw new Error(`フォルダパス解析に失敗しました: ${err.message}`);
  }
}

function fileExistsInFolder(folderId: string, fileName: string): boolean {
  try {
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFilesByName(fileName);
    return files.hasNext();
  } catch (error) {
    const err = error as Error;
    logError("ファイル存在チェックエラー", err.message);
    return false;
  }
}

function copyFileToFolder(
  fileBlob: GoogleAppsScript.Base.Blob,
  folderId: string,
  newFileName: string,
): CopyResult {
  try {
    const fileName = ensurePdfExtension(newFileName);

    const folder = DriveApp.getFolderById(folderId);

    if (fileExistsInFolder(folderId, fileName)) {
      return {
        success: false,
        error: `ファイルが既に存在します: ${fileName}`,
      };
    }

    const createdFile = folder.createFile(fileBlob.setName(fileName));

    logDebug(`ファイルをコピーしました: ${fileName}`, createdFile.getId());

    return {
      success: true,
      fileId: createdFile.getId(),
      fileName,
    };
  } catch (error) {
    const err = error as Error;
    logError("ファイルコピーエラー", err.message);
    return {
      success: false,
      error: `ファイルコピーに失敗しました: ${err.message}`,
    };
  }
}

function copyMultipleFiles(
  files: CopyFileInput[],
  folderId: string,
  customFileNameBase: string | null = null,
): CopyResult[] {
  const results: CopyResult[] = [];

  files.forEach((file, index) => {
    let fileName = file.name;

    if (customFileNameBase) {
      if (files.length === 1) {
        fileName = customFileNameBase;
      } else {
        fileName = `${customFileNameBase.replace(/\.pdf$/i, "")}_${index + 1}.pdf`;
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

function generateDriveLink(fileId: string): string {
  return createDriveLink(fileId);
}
