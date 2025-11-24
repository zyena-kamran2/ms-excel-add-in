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

async function cloneWorkbookSnapshot() {
  const status = document.getElementById("status")!;
  status.textContent = "Creating snapshot...";

  await Excel.run(async (context) => {
    const srcWb = context.workbook;

    // Load source sheet metadata before creating new workbook
    const srcSheets = srcWb.worksheets;
    srcSheets.load("items/name");
    await context.sync();

    // Create brand new workbook (no return object)
    await Excel.createWorkbook();
    await context.sync();

    // context.workbook NOW refers to the new workbook
    const newWb = context.workbook;
    const newSheets = newWb.worksheets;

    // Remove default blank sheet
    newSheets.getFirst().delete();
    await context.sync();

    // Copy each sheet from source
    for (const sheet of srcSheets.items) {
      const srcSheet = srcWb.worksheets.getItem(sheet.name);
      const dstSheet = newSheets.add(sheet.name);

      const used = srcSheet.getUsedRange(true);
      used.load(["values", "formulas", "format/*"]);
      await context.sync();

      if (used.values.length > 0) {
        const dstRange = dstSheet.getRange("A1").getResizedRange(
          used.values.length - 1,
          used.values[0].length - 1
        );

        // Copy formatting
        dstRange.format.fill.color = used.format.fill.color;
        dstRange.format.font.color = used.format.font.color;
        dstRange.format.font.bold = used.format.font.bold;
        dstRange.format.font.size = used.format.font.size;
        dstRange.format.horizontalAlignment = used.format.horizontalAlignment;

        // Copy values
        dstRange.values = used.values;

        // Remove formulas
        dstRange.formulas = dstRange.values;
      }
    }

    await context.sync();

    status.textContent = "Snapshot created — prompting save...";

    // Save the new workbook
    await newWb.save(Excel.SaveBehavior.prompt);
  });

  status.textContent = "Saved successfully!";
}
