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

      debugLog("Loading used range...");

      const usedRange = sheet.getUsedRange();
      usedRange.load([
        "values",
        "rowCount",
        "columnCount",
        "address",
        "rowIndex",
        "columnIndex"
      ]);
      await context.sync();

      const values = usedRange.values as (string | number | boolean)[][];
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
      debugLog(`New sheet created: ${newName}`);

      // Step 1: Format copy (safe)
      const targetRange = newSheet.getRangeByIndexes(
        startRow,
        startCol,
        rowCount,
        colCount
      );

      targetRange.copyFrom(usedRange, Excel.RangeCopyType.formats);
      await context.sync();

      debugLog("Copied formatting.");

      // Step 2: Write values
      targetRange.values = values;
      await context.sync();

      debugLog(`Worksheet cloned successfully as "${newName}"`);
    });
  } catch (error: any) {
    debugLog(`Error cloning worksheet: ${error}`);
  } finally {
    event.completed();
  }
}


// Register the function
Office.actions.associate("cloneWorksheetValues", cloneWorksheetValues);
