Office.onReady((info) => {
  console.log("Taskpane ready");
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;

    document.getElementById("btn-clone")!.addEventListener("click", cloneWorkbookSnapshot);
  }
});

export async function run() {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load("address");
      range.format.fill.color = "yellow";
      await context.sync();
      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}

async function cloneWorkbook() {
  const status = document.getElementById("status")!;
  await Excel.run(async (context) => {
    const wb = context.workbook;

    status.textContent = "Cloning all sheets...";

    // Load all sheets
    const sheets = wb.worksheets;
    sheets.load("items/name");
    await context.sync();

    // For each sheet, duplicate it within the same workbook
    for (const sheet of sheets.items) {
      // Copy the sheet to the end of workbook
      const copySheet = sheet.copy("End"); // valid positionType
      copySheet.load("name");
      await context.sync();

      // Rename the copied sheet
      copySheet.name = sheet.name + "_Copy";

      // Flatten formulas in the copied sheet
      await flattenSheetFormulas(context, copySheet);
    }


    await context.sync();
    status.textContent = "Workbook cloned successfully (all formulas replaced with values)!";
  });
}

// Flatten formulas: replace formulas with their calculated values
async function flattenSheetFormulas(
  context: Excel.RequestContext,
  sheet: Excel.Worksheet
) {
  const used = sheet.getUsedRange();
  used.load(["values", "formulas"]);
  await context.sync();

  const values = used.values;

  // Clear formulas but keep formatting
  used.clear(Excel.ClearApplyTo.contents);

  // Set values
  used.values = values;
}




export async function cloneWorkbookSnapshot() {
  const statusEl = document.getElementById("status");
  const setStatus = (s: string) => { if (statusEl) statusEl.textContent = s; };
  setStatus("Reading source workbook...");

  type SheetSnapshot = {
    name: string;
    rowCount: number;
    columnCount: number;
    values: any[][];
    numberFormats: any[][];
    font: { bold?: boolean; italic?: boolean; size?: number; color?: string; name?: string };
    fillColor: string;
  };

  let snapshot: SheetSnapshot[] = [];

  // STEP 1: Read source workbook
  try {
    snapshot = await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.load("items/name");
      await context.sync();

      const result: SheetSnapshot[] = [];

      for (const sh of sheets.items) {
        const ws = sheets.getItem(sh.name);
        const used = ws.getUsedRange(true);
        used.load(["rowCount", "columnCount", "values", "numberFormat"]);
        used.format.load(["font/bold", "font/italic", "font/size", "font/color", "font/name", "fill/color"]);
        await context.sync();

        result.push({
          name: sh.name,
          rowCount: used.rowCount,
          columnCount: used.columnCount,
          values: used.values,
          numberFormats: used.numberFormat,
          font: {
            bold: used.format.font.bold,
            italic: used.format.font.italic,
            size: used.format.font.size,
            color: used.format.font.color,
            name: used.format.font.name,
          },
          fillColor: used.format.fill.color || "",
        });
      }

      return result;
    });
  } catch (err) {
    console.error("Error reading source workbook:", err);
    setStatus("Failed reading workbook: " + (err as any).message);
    return;
  }

  setStatus("Creating new workbook...");

  // STEP 2: Create new workbook
  try {
    await Excel.run(async (context) => {
      await Excel.createWorkbook(); // creates new workbook in host
      await context.sync();

      
  setStatus("Created new workbook...");

      const newWb = context.workbook;
      const newSheets = newWb.worksheets;

      // Remove default sheet
      try {
        newSheets.getFirst().delete();
        await context.sync();
      } catch { }

      // Add sheets from snapshot
      for (const sh of snapshot) {
        setStatus(`Adding sheet: ${sh.name}`);
        const uniqueName = getUniqueSheetName(newSheets, sh.name);
        const dst = newSheets.add(uniqueName);
        await context.sync(); // important: wait for sheet to be ready

        if (sh.rowCount > 0 && sh.columnCount > 0) {
          const dstRange = dst.getRangeByIndexes(0, 0, sh.rowCount, sh.columnCount);
          dstRange.values = sh.values;
          dstRange.numberFormat = sh.numberFormats;

          dstRange.format.font.bold = sh.font.bold;
          dstRange.format.font.italic = sh.font.italic;
          dstRange.format.font.size = sh.font.size;
          dstRange.format.font.color = sh.font.color;
          dstRange.format.font.name = sh.font.name;
          dstRange.format.fill.color = sh.fillColor;

          // Remove formulas
          dstRange.formulas = dstRange.values;
        }

        setStatus(`Added sheet: ${uniqueName}`);

        await context.sync(); // make sure all writes are applied
      }

      setStatus("Snapshot created — prompting Save As dialog...");
      await newWb.save(Excel.SaveBehavior.prompt);
    });
  } catch (err) {
    console.error("Error creating new workbook:", err);
    setStatus("Failed creating snapshot: " + (err as any).message);
    return;
  }

  setStatus("Saved successfully!");
}

function getUniqueSheetName(existingSheets: Excel.WorksheetCollection, baseName: string): string {
  let name = baseName;
  let counter = 1;
  while (true) {
    try {
      existingSheets.getItem(name);
      name = `${baseName}_${counter++}`;
    } catch {
      break;
    }
  }
  return name;
}

