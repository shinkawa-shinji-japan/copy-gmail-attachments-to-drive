# Copilot Instructions: Gmail Attachments to Drive

## Project Overview

A Google Apps Script tool that extracts PDF attachments from Gmail within a specified date range and copies them to Google Drive folders. The app provides a spreadsheet-based UI using Google Sheets with two main sheets:
- **検索条件 (Search Criteria)**: Input sheet for date ranges, keywords, and destination folder path
- **結果 (Results)**: Output sheet showing search results with checkboxes for selective copying

## Architecture & Data Flow

### Core Module Structure (TypeScript)

**`src/main.ts`** - UI orchestration layer
- `onOpen()`: Creates custom menu "Gmail添付ファイル" with action items
- `searchAndDisplay()`: Workflow for searching emails and populating Results sheet
- `executeAndCopy()`: Workflow for processing checked rows and copying files
- `initializeSheets()`: Creates required sheets with default headers/data

**`src/gmail.ts`** - Gmail API wrapper
- `searchEmails()`: Searches threads using Gmail query syntax (`is:attachment after:... before:...`)
- `buildSearchQuery()`: Constructs Gmail search query with date range and keyword filters
- `extractEmailAndAttachments()`: Extracts subject, date, and PDFs from message threads
- `getPdfAttachments()`: Filters PDF files from message attachments
- Timezone-aware date handling via `SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone()`

**`src/drive.ts`** - Google Drive API wrapper
- `resolveFolderPath()`: Navigates/creates folder hierarchy from path string (e.g., `/year/2026/Jan`)
- `fileExistsInFolder()`: Prevents duplicate file overwrites
- `copyFileToFolder()`: Copies blob to Drive with error handling
- `ensurePdfExtension()`: Ensures `.pdf` extension on filenames

**`src/sheet.ts`** - Google Sheets API wrapper
- `getSearchSheet()` / `getResultsSheet()`: Lazy initialization pattern
- `getSearchConditions()`: Reads rows 2-5 from Search sheet into typed object
- `getTargetRowsData()`: Reads checked rows from Results sheet
- `addResultRows()`: Batch inserts result data with formulas for Drive links
- Handles checkbox state and formula generation for shared Drive links

**`src/utils.ts`** - Common utilities
- `validateDateInput()`: Parses YYYY-MM-DD strings or Date objects
- `formatDateTime()`: Formats dates for display/logging
- `logDebug()` / `logError()`: Logging with contextual objects

## Key Development Patterns

### Build & Deployment

```bash
npm run build          # TypeScript → JavaScript (clasp push requires JS)
npm run push           # Compile + push to GAS project
npm run watch          # Watch mode for development
npm run logs           # View execution logs via clasp
```

The `appsscript.json` declares required OAuth scopes:
- `gmail.readonly` (search, read attachments)
- `spreadsheets` (read/write sheets)
- `drive` (create folders, copy files)

### Error Handling Pattern

All async Google Apps Script API calls are wrapped in try-catch:
```typescript
try {
  // API call
  logDebug("Step context", {detailKey: value});
  return result;
} catch (error) {
  const err = error as Error;
  logError("Operation name", err.message);
  throw new Error(`User-friendly message: ${err.message}`);
}
```

Errors bubble up to `main.ts` functions which display alerts via `SpreadsheetApp.getUi().alert()`.

### Type Definitions

Custom types for data contracts between modules:
- `EmailData`: {subject, date, attachments[], messageId}
- `SearchConditions`: {startDate, endDate, keywords[], folderPath}
- `ResultRow`: Matches Results sheet columns
- `CopyResult`: {success, fileId?, error?}

### Timezone Handling

**Critical**: All date operations respect the spreadsheet's timezone:
- Input dates from sheet are assumed local to spreadsheet timezone
- Gmail query dates formatted with `Utilities.formatDate(date, timezone, "yyyy/MM/dd")`
- No timezone conversion—dates flow directly through system

## Sheets UI Structure

### Search Sheet (検索条件)
```
| 項目              | 値         |
|-------------------|----------|
| 期間（開始日）    | 2026-01-01 |
| 期間（終了日）    | 2026-01-31 |
| フィルタ文字列    | (optional keywords, comma-separated) |
| 保存先フォルダパス | /project/pdfs |
```

### Results Sheet (結果)
Columns: 保存対象 (checkbox) | メールタイトル | 受信日時 | 添付ファイル名 | 保存ファイル名 | 保存先フォルダ名 | 処理結果 | ファイルリンク

## Integration Points & Dependencies

- **GmailApp**: Thread search, message metadata, attachment blobs
- **DriveApp**: Folder navigation, file creation, duplicate detection
- **SpreadsheetApp**: Sheet access, range operations, timezone info
- **Utilities**: Date formatting, type checking

## Critical Considerations

1. **PDF-only filtering**: Non-PDF attachments are silently skipped
2. **Timezone awareness**: Date strings from sheet assume spreadsheet timezone
3. **Duplicate prevention**: Files with identical names in same folder are skipped (not overwritten)
4. **Batch operations**: `addResultRows()` uses batch inserts for performance
5. **Search limitations**: Gmail search syntax supports date ranges but not fine-grained time filtering
