/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

Office.onReady(() => {
  // If needed, Office.js is ready to be called.
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
function action(event: Office.AddinCommands.Event) {
  const message: Office.NotificationMessageDetails = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Performed action.",
    icon: "Icon.80x80",
    persistent: true,
  };

  // Show a notification message.
  Office.context.mailbox.item.notificationMessages.replaceAsync(
    "ActionPerformanceNotification",
    message
  );

  // Be sure to indicate when the add-in command function is complete.
  event.completed();
}
/**
 * Clone the active worksheet, copying only the values (not formulas).
 */
export async function cloneWorksheetValues(event: Office.AddinCommands.Event) {
  try {
    await Excel.run(async (context) => {
      const workbook = context.workbook;
      const sheet = workbook.worksheets.getActiveWorksheet();

      // Load the used range (only values)
      const usedRange = sheet.getUsedRange();
      usedRange.load(["values", "rowCount", "columnCount", "address"]);
      await context.sync();

      const values = usedRange.values as (string | number | boolean)[][];
      const rowCount = usedRange.rowCount!;
      const colCount = usedRange.columnCount!;
      console.log(`Used range ${usedRange.address}, rows: ${rowCount}, cols: ${colCount}`);

      // Create a new worksheet
      // Decide on a name (e.g., originalName + " - Copy")
      sheet.load("name");
      await context.sync();
      const originalName = sheet.name!;
      let newName = `${originalName} - Copy`;
      // Make sure name is unique
      const sheets = workbook.worksheets;
      let suffix = 1;
      while (true) {
        try {
          // Try to add with the name
          const newSheet = sheets.add(newName);
          // If no error, break
          break;
        } catch (e: any) {
          // If name exists, change name and try again
          suffix++;
          newName = `${originalName} - Copy (${suffix})`;
        }
      }
      const newSheet = workbook.worksheets.getItem(newName);

      // Write values into new sheet
      // We'll write to A1 with same dimension of usedRange
      const targetRange = newSheet.getRangeByIndexes(0, 0, rowCount, colCount);
      targetRange.values = values;

      // Optionally, copy formats (if you want) – uncomment:
      // targetRange.copyFrom(usedRange, Excel.RangeCopyType.formats);

      await context.sync();
    });
  } catch (error) {
    console.error("Error cloning worksheet:", error);
    // You might want to show a notification to the user
    // Office.ui.displayDialogAsync; // or other UI method

  } finally {
    // Indicate that your command function is complete
    event.completed();
  }
}

// Register the function so Office can call it via the manifest
Office.actions.associate("cloneWorksheetValues", cloneWorksheetValues);
// Register the function with Office.
Office.actions.associate("action", action);
