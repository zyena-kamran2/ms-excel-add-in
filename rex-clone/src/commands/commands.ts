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

// Fixed approach: Clone worksheet to new workbook using base64 export
// Simplified approach: Clone worksheet by reading data and creating new workbook
// Clone worksheet to new workbook - using window.open approach
export async function cloneWorksheetValues(event: Office.AddinCommands.Event) {
  try {
    debugLog("Starting clone operation...");
    
    // Get the current workbook as a base64 string
    await new Promise<void>((resolve, reject) => {
      Office.context.document.getFileAsync(
        Office.FileType.Compressed,
        { sliceSize: 65536 },
        async (result) => {
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

              // Add slice data to file content
              const sliceData = sliceResult.value.data;
              const byteArray = new Uint8Array(sliceData);
              for (let i = 0; i < byteArray.length; i++) {
                fileContent.push(byteArray[i]);
              }

              slicesReceived++;
              debugLog(`Read slice ${slicesReceived}/${sliceCount}`);

              if (slicesReceived === sliceCount) {
                // All slices received - convert to base64 and create workbook
                file.closeAsync();
                
                try {
                  const uint8Array = new Uint8Array(fileContent);
                  const base64Data = arrayBufferToBase64(uint8Array);
                  
                  debugLog("File conversion complete, creating new workbook...");
                  
                  await Excel.run(async (context) => {
                    // Create new workbook with the base64 data
                    Excel.createWorkbook(base64Data);
                    await context.sync();
                    debugLog("New workbook created successfully!");
                    resolve();
                  });
                } catch (error: any) {
                  debugLog(`ERROR during workbook creation: ${error.message}`);
                  reject(error);
                }
              } else {
                // Read next slice
                readSlice(sliceIndex + 1);
              }
            });
          }

          // Start reading slices
          readSlice(0);
        }
      );
    });

  } catch (error: any) {
    debugLog(`ERROR: ${error.message || error}`);
    if (error.debugInfo) {
      debugLog(`Debug info: ${JSON.stringify(error.debugInfo)}`);
    }
    console.error(error);
  } finally {
    event.completed();
  }
}

// Helper function to convert ArrayBuffer to Base64
function arrayBufferToBase64(buffer: Uint8Array): string {
  let binary = '';
  const len = buffer.byteLength;
  
  // Process in chunks to avoid stack overflow
  const chunkSize = 8192;
  for (let i = 0; i < len; i += chunkSize) {
    const chunk = buffer.subarray(i, Math.min(i + chunkSize, len));
    binary += String.fromCharCode.apply(null, Array.from(chunk));
  }
  
  return btoa(binary);
}
// Register the function
Office.actions.associate("cloneWorksheetValues", cloneWorksheetValues);
