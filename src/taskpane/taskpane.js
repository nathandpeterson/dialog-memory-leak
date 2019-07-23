/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

const DIALOG_URL = 'https://dialog-opener.surge.sh/dialog.html'

Office.onReady(info => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("run").onclick = run;
  }
});



export async function run() {
  Office.context.ui.displayDialogAsync(
    DIALOG_URL, 
    { height: 30, width: 20, displayInIframe: true },
    (asyncResult) => {
      if(asyncResult) {
        console.log('adding event handler')
        const dialog = asyncResult.value

        function processMessage(arg) {
          var messageFromDialog = JSON.parse(arg.message);
          if(messageFromDialog.dialog === false) {
            dialog.close()
          }
        }
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);

      } else {
        console.log('asyncResult failed check', asyncResult)
      }
      
    }
    )
}
