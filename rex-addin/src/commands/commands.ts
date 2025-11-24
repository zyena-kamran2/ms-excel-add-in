/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

/* commands.ts */
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

// Register the function with Office.
Office.actions.associate("action", action);

export async function cloneSheetValues(event: Office.AddinCommands.Event) {
  try {
    await Excel.run(async (context) => {
      const workbook = context.workbook;
      const srcSheet = workbook.worksheets.getActiveWorksheet();
      srcSheet.load("name");
      const usedRange = srcSheet.getUsedRange();
      usedRange.load(["rowCount", "columnCount", "values"]);
      await context.sync();

      const { rowCount, columnCount, values } = usedRange;
      const srcValues = values as (string | number | boolean | null)[][];

      const newName = `${srcSheet.name}_Clone_${new Date().getTime()}`;
      const newSheet = workbook.worksheets.add(newName);

      const targetRange = newSheet.getRangeByIndexes(0, 0, rowCount, columnCount);
      targetRange.values = srcValues;

      newSheet.activate();

      await context.sync();
    });
  } catch (error) {
    console.error(error);
  } finally {
    event.completed();
  }
}
