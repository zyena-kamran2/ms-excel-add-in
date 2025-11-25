/* global Office */

Office.onReady(() => {
  // Office.js is ready
});

let debugDialog: Office.Dialog | null = null;
let dialogReady = false;
const debugQueue: string[] = [];

/**
 * Send debug messages to a separate debug.html dialog
 */
function debugLog(message: string) {
  if (!debugDialog) {
    Office.context.ui.displayDialogAsync(
      "https://localhost:3000/debug.html",
      { height: 30, width: 40 },
      (result) => {
        debugDialog = result.value;

        debugDialog.addEventHandler(
          Office.EventType.DialogMessageReceived,
          (arg) => {
            if ("message" in arg && arg.message === "ready") {
              dialogReady = true;
              // Send any queued messages
              debugQueue.forEach((msg) => debugDialog?.messageChild(msg));
              debugQueue.length = 0;
            }
          }
        );
      }
    );
  }

  if (dialogReady) {
    debugDialog?.messageChild(message);
  } else {
    debugQueue.push(message); // Queue messages until dialog is ready
  }

  console.log(message); // Fallback
}
/* global Office */

Office.onReady(() => {
  // Office.js ready
});

export async function cloneWorksheetValues(event: Office.AddinCommands.Event) {
  try {
    await Excel.run(async (context) => {
      const workbook = context.workbook;
      const sheet = workbook.worksheets.getActiveWorksheet();

      debugLog("Loading used range (values + formulas)...");

      const usedRange = sheet.getUsedRange();
      usedRange.load([
        "values",
        "formulas",
        "rowCount",
        "columnCount",
        "address",
        "rowIndex",
        "columnIndex",
      ]);
      await context.sync();

      const values = usedRange.values as any[][];
      const formulas = usedRange.formulas as any[][];
      const rowCount = usedRange.rowCount!;
      const colCount = usedRange.columnCount!;
      const startRow = usedRange.rowIndex!;
      const startCol = usedRange.columnIndex!;

      debugLog(
        `Used range ${usedRange.address}, rows: ${rowCount}, cols: ${colCount}, startRow: ${startRow}, startCol: ${startCol}`
      );

      // Load original sheet name
      sheet.load("name");
      await context.sync();

      const originalName = sheet.name!;
      let newName = `${originalName} - Copy`;

      const sheets = workbook.worksheets;
      let suffix = 1;

      // Generate unique sheet name
      while (true) {
        try {
          sheets.add(newName);
          break;
        } catch (e: any) {
          suffix++;
          newName = `${originalName} - Copy (${suffix})`;
        }
      }

      const newSheet = workbook.worksheets.getItem(newName);
      debugLog(`Created ${newName}`);

      // Target range starts at A1 (0,0) in new sheet
      const targetRange = newSheet.getRangeByIndexes(0, 0, rowCount, colCount);

      // 1️⃣ Copy formats safely
      debugLog("Copying formats...");
      targetRange.copyFrom(usedRange, Excel.RangeCopyType.formats);
      await context.sync();

      // 2️⃣ Copy values
      debugLog("Writing values (values-only copy)...");
      targetRange.values = values;
      await context.sync();

      // 3️⃣ Restore internal formulas selectively
      debugLog("Restoring internal formulas...");

      const externalPatterns: RegExp[] = [
        /\[[^\]]+\]/, // external workbook reference
        /\bWEBSERVICE\s*\(/i,
        /\bFILTERXML\s*\(/i,
        /\bRTD\s*\(/i,
        /https?:\/\//i,
        /::/i,
      ];

      function isFormulaString(cellFormula: any): boolean {
        return typeof cellFormula === "string" && cellFormula.startsWith("=");
      }

      function isExternalFormula(formulaText: string): boolean {
        if (!formulaText || typeof formulaText !== "string") return false;
        return externalPatterns.some((re) => re.test(formulaText));
      }

      // Build boolean map of formulas to keep
      const keepFormula: boolean[][] = [];
      for (let r = 0; r < rowCount; r++) {
        keepFormula[r] = [];
        for (let c = 0; c < colCount; c++) {
          const f = formulas?.[r]?.[c];
          keepFormula[r][c] = isFormulaString(f) && !isExternalFormula(f);
        }
      }

      // Apply formulas in contiguous blocks
      for (let r = 0; r < rowCount; r++) {
        let c = 0;
        while (c < colCount) {
          if (!keepFormula[r][c]) {
            c++;
            continue;
          }
          let runStart = c;
          let runEnd = c + 1;
          while (runEnd < colCount && keepFormula[r][runEnd]) runEnd++;
          const runLen = runEnd - runStart;

          const subRange = newSheet.getRangeByIndexes(r, runStart, 1, runLen);
          const formulasSlice: any[][] = [formulas[r].slice(runStart, runEnd)];

          subRange.formulas = formulasSlice;
          c = runEnd;
        }
      }

      await context.sync();
      debugLog(`Worksheet cloned successfully as "${newName}"`);
    });
  } catch (error: any) {
    debugLog(`Error cloning worksheet with selective formulas: ${error}`);
  } finally {
    event.completed();
  }
}



// Register the function
Office.actions.associate("cloneWorksheetValues", cloneWorksheetValues);
