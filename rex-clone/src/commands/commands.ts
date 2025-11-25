/* global Office, Excel */

Office.onReady(() => {
  debugLog("Add-in ready.");
});

/** Simple debug logger */
let debugDialog: Office.Dialog | null = null;
let dialogReady = false;
const debugQueue: string[] = [];

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


/**
 * Clone the active worksheet, strip external formulas in the copy, and export only that sheet to a new workbook
 */

export async function cloneWorksheetValues(event: Office.AddinCommands.Event) {
  try {
    //
    // 1️⃣ FIRST: Clone the worksheet
    //
    let newName = "";
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

      sheet.load("name");
      await context.sync();

      const originalName = sheet.name!;
      newName = `${originalName} - Copy`;
      const sheets = workbook.worksheets;
      let suffix = 1;

      while (true) {
        try {
          sheets.add(newName);
          break;
        } catch {
          newName = `${originalName} - Copy (${suffix})`;
          suffix++;
        }
      }

      const newSheet = workbook.worksheets.getItem(newName);
      debugLog(`Created ${newName}`);

      const targetRange = newSheet.getRangeByIndexes(0, 0, rowCount, colCount);

      debugLog("Copying formats...");
      targetRange.copyFrom(usedRange, Excel.RangeCopyType.formats);
      await context.sync();

      debugLog("Writing values (values-only copy)...");
      targetRange.values = values;
      await context.sync();

      debugLog("Restoring internal formulas...");

      const externalPatterns: RegExp[] = [
        /\[[^\]]+\]/,
        /\bWEBSERVICE\s*\(/i,
        /\bFILTERXML\s*\(/i,
        /\bRTD\s*\(/i,
        /https?:\/\//i,
        /::/i,
      ];

      function isFormulaString(f: any): boolean {
        return typeof f === "string" && f.startsWith("=");
      }

      function isExternalFormula(text: string): boolean {
        if (!text || typeof text !== "string") return false;
        return externalPatterns.some((re) => re.test(text));
      }

      const keepFormula: boolean[][] = [];
      for (let r = 0; r < rowCount; r++) {
        keepFormula[r] = [];
        for (let c = 0; c < colCount; c++) {
          const f = formulas?.[r]?.[c];
          keepFormula[r][c] = isFormulaString(f) && !isExternalFormula(f);
        }
      }

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

          const subRange = newSheet.getRangeByIndexes(r, runStart, 1, runEnd - runStart);
          const formulasSlice: any[][] = [
            formulas[r].slice(runStart, runEnd),
          ];

          subRange.formulas = formulasSlice;
          c = runEnd;
        }
      }

      await context.sync();
      debugLog(`Worksheet cloned successfully as "${newName}"`);
    });

    //
    // 2️⃣ NOW: Hide all sheets except the new one, export workbook, then restore visibility
    //
    debugLog("Preparing to export only the new sheet...");

    // Store original visibility states
    const sheetVisibility: { name: string; visibility: string }[] = [];

    await Excel.run(async (context) => {
      const workbook = context.workbook;
      const sheets = workbook.worksheets;
      sheets.load("items/name, items/visibility");
      await context.sync();

      // Save current visibility and hide all sheets except the new one
      for (const sheet of sheets.items) {
        sheetVisibility.push({
          name: sheet.name,
          visibility: sheet.visibility
        });

        if (sheet.name !== newName) {
          sheet.visibility = Excel.SheetVisibility.hidden;
        }
      }

      await context.sync();
      debugLog(`Hidden ${sheetVisibility.length - 1} sheets, keeping only "${newName}" visible`);
    });

    //
    // 3️⃣ Export the workbook (now containing only the visible new sheet)
    //
    debugLog("Starting workbook export to Base64...");

    await new Promise<void>((resolve, reject) => {
      Office.context.document.getFileAsync(
        Office.FileType.Compressed,
        { sliceSize: 65536 },
        (result) => {
          if (result.status !== Office.AsyncResultStatus.Succeeded) {
            debugLog(`ERROR: Failed to get file - ${result.error.message}`);
            reject(result.error);
            return;
          }

          const file = result.value;
          const sliceCount = file.sliceCount;
          let slicesReceived = 0;
          const fileContent: number[] = [];

          debugLog(`File has ${sliceCount} slices to read`);

          function readSlice(sliceIndex: number) {
            file.getSliceAsync(sliceIndex, async (sliceResult) => {
              if (sliceResult.status !== Office.AsyncResultStatus.Succeeded) {
                debugLog(`ERROR: Failed to read slice ${sliceIndex}`);
                file.closeAsync();
                reject(sliceResult.error);
                return;
              }

              const byteArray = new Uint8Array(sliceResult.value.data);
              for (let i = 0; i < byteArray.length; i++) {
                fileContent.push(byteArray[i]);
              }

              slicesReceived++;
              debugLog(`Read slice ${slicesReceived}/${sliceCount}`);

              if (slicesReceived === sliceCount) {
                file.closeAsync();

                try {
                  const uint8Array = new Uint8Array(fileContent);
                  const base64Data = arrayBufferToBase64(uint8Array);

                  debugLog("File conversion complete, creating new workbook...");

                  await Excel.run(async (context) => {
                    Excel.createWorkbook(base64Data);
                    await context.sync();
                    debugLog("New workbook created successfully with only the cloned sheet!");
                    resolve();
                  });
                } catch (error: any) {
                  debugLog(`ERROR during workbook creation: ${error.message}`);
                  reject(error);
                }
              } else {
                readSlice(sliceIndex + 1);
              }
            });
          }

          readSlice(0);
        }
      );
    });

    //
    // 4️⃣ Restore original sheet visibility and delete the temp sheet
    //
    await Excel.run(async (context) => {
      const workbook = context.workbook;
      const sheets = workbook.worksheets;

      // Restore original visibility
      for (const sheetInfo of sheetVisibility) {
        try {
          const sheet = sheets.getItem(sheetInfo.name);
          sheet.visibility = sheetInfo.visibility as any;
        } catch (e) {
          debugLog(`Could not restore visibility for ${sheetInfo.name}`);
        }
      }

      await context.sync();
      debugLog("Restored original sheet visibility");

      // Delete the temp sheet
      try {
        const sheetToDelete = workbook.worksheets.getItem(newName);
        sheetToDelete.delete();
        await context.sync();
        debugLog(`Deleted temp sheet: ${newName}`);
      } catch (e) {
        debugLog(`Sheet ${newName} not found for deletion`);
      }
    });

  } catch (error: any) {
    debugLog(`ERROR: ${error.message || error}`);
    if (error.debugInfo) debugLog(`Debug info: ${JSON.stringify(error.debugInfo)}`);
  } finally {
    event.completed();
  }
}

// Helper: convert to Base64
function arrayBufferToBase64(buffer: Uint8Array): string {
  let binary = "";
  const chunkSize = 8192;
  for (let i = 0; i < buffer.byteLength; i += chunkSize) {
    const chunk = buffer.subarray(i, Math.min(i + chunkSize, buffer.byteLength));
    binary += String.fromCharCode.apply(null, Array.from(chunk));
  }
  return btoa(binary);
}

// Register the function
Office.actions.associate("cloneWorksheetValues", cloneWorksheetValues);