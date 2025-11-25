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
export async function cloneWorksheetValues(event: Office.AddinCommands.Event) {
  try {
    await Excel.run(async (context) => {
      const workbook = context.workbook;
      const sheet = workbook.worksheets.getActiveWorksheet();

      debugLog("Loading used range (values + formulas)...");

      const usedRange = sheet.getUsedRange();
      // note: load formulas and values
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

      // If usedRange is empty this will throw earlier; assume you already handle empties.
      const values = usedRange.values as any[][];
      const formulas = usedRange.formulas as any[][];
      const rowCount = usedRange.rowCount!;
      const colCount = usedRange.columnCount!;
      const startRow = usedRange.rowIndex!;
      const startCol = usedRange.columnIndex!;

      debugLog(
        `Used range ${usedRange.address}, rows: ${rowCount}, cols: ${colCount}, startRow: ${startRow}, startCol: ${startCol}`
      );

      // Create sheet name (unique)
      sheet.load("name");
      await context.sync();
      const originalName = sheet.name!;
      let newName = `${originalName} - Copy`;
      const sheets = workbook.worksheets;
      let suffix = 1;
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

      // Target range in the new sheet with same offsets
      const targetRange = newSheet.getRangeByIndexes(startRow, startCol, rowCount, colCount);

      // 1) copy formats first (safe)
      debugLog("Copying formats...");
      targetRange.copyFrom(usedRange, Excel.RangeCopyType.formats);
      await context.sync();

      // 2) copy values for the whole block (so external formulas don't re-run)
      debugLog("Writing values (values-only copy)...");
      targetRange.values = values;
      await context.sync();

      // 3) restore formulas selectively
      debugLog("Detecting formula cells to preserve...");

      // Heuristic: consider a formula external if it matches any of these patterns.
      // Customize `externalPatterns` or expose them as a parameter to your add-in.
      const externalPatterns: RegExp[] = [
        /\[[^\]]+\]/, // external workbook reference [Book.xlsx]
        /\bWEBSERVICE\s*\(/i,
        /\bFILTERXML\s*\(/i,
        /\bRTD\s*\(/i,
        /https?:\/\//i,
        /::/i, // odd references
        // add custom function names e.g. /\bMYCUSTOMFUNC\s*\(/i
      ];

      function isFormulaString(cellFormula: any): boolean {
        // Range.formulas returns the formula for formula cells, and the value for non-formula cells.
        // Typically formula strings begin with '='. But formulas array may hold other types;
        // ensure we treat only strings starting with '=' as formulas.
        return typeof cellFormula === "string" && cellFormula.startsWith("=");
      }

      function isExternalFormula(formulaText: string): boolean {
        if (!formulaText || typeof formulaText !== "string") return false;
        // quick lowercased check (already using regex with i flag)
        for (const re of externalPatterns) {
          if (re.test(formulaText)) return true;
        }
        return false;
      }

      // Build boolean map: keepFormula[r][c] = true if we should preserve formula in that cell
      const keepFormula: boolean[][] = [];
      for (let r = 0; r < rowCount; r++) {
        keepFormula[r] = [];
        for (let c = 0; c < colCount; c++) {
          const f = formulas?.[r]?.[c];
          if (isFormulaString(f) && !isExternalFormula(f)) {
            // internal-looking formula and not matched as external => preserve
            keepFormula[r][c] = true;
          } else {
            keepFormula[r][c] = false;
          }
        }
      }

      debugLog("Restoring internal formulas in contiguous blocks...");

      // Helper to take a slice of formulas for a contiguous block
      function sliceFormulasSlice(
        src: any[][],
        rowStart: number,
        colStart: number,
        numRows: number,
        numCols: number
      ) {
        const out: any[][] = [];
        for (let rr = 0; rr < numRows; rr++) {
          const row: any[] = [];
          for (let cc = 0; cc < numCols; cc++) {
            row.push(src[rowStart + rr][colStart + cc]);
          }
          out.push(row);
        }
        return out;
      }

      // We'll iterate row by row and group contiguous formula columns to reduce calls.
      for (let r = 0; r < rowCount; r++) {
        let c = 0;
        while (c < colCount) {
          // skip until we find a formula to keep
          if (!keepFormula[r][c]) {
            c++;
            continue;
          }
          // found start of contiguous run at column c
          let runStart = c;
          let runEnd = c + 1;
          while (runEnd < colCount && keepFormula[r][runEnd]) runEnd++;
          const runLen = runEnd - runStart;

          // create a subrange in target sheet for this run
          const subRange = newSheet.getRangeByIndexes(startRow + r, startCol + runStart, 1, runLen);
          // get the formulas slice for [r, runStart .. runEnd)
          const formulasSlice = sliceFormulasSlice(formulas, r, runStart, 1, runLen);

          // set the formulas (these are strings like "=A1+B1")
          subRange.formulas = formulasSlice;

          // advance
          c = runEnd;
        }
      }

      await context.sync();
      debugLog(`Clone finished: ${newName}`);
    });
  } catch (error: any) {
    debugLog(`Error cloning worksheet with selective formulas: ${error}`);
  } finally {
    event.completed();
  }
}


// Register the function
Office.actions.associate("cloneWorksheetValues", cloneWorksheetValues);
