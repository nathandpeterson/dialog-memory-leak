#Overview

This add-in was created to reproduce a memory leak in the Office.js API for Outlook in Edge-webview. This issue is not present in earlier versions of Outlook, which used IE
as the webview renderer.

The original issue was posted [here] (https://github.com/OfficeDev/office-js/issues/584)
with screenshots of gradually increasing memory profile for One Note.

##Main points:
1. When opening a dialog from the taskpane, some portion of memory is allocated in the taskpane script.
2. When closing the dialog from the taskpane, that memory is never freed up.
3. This only occurs when we call `dialog.close()` on the dialog instance, not when the user clicks the close button provided by the UI at the top of the dialog.

All the code in this repro does is open and close a dialog.

Here is a profile from this add-in. [Screen shot with memory leak](screenshot1.png)

And here is a short gif. [Gif of memory leak](clip.gif)

##Environment

* Platform [PC desktop, Mac, iOS, Office Online]: PC desktop
* Host [Excel, Word, PowerPoint, etc.]: Outlook
* Office version number: Version 1906 (Build 11727.20210)
* Operating System: Windows 10 Version 1903 for x64

The key code snippet is reproduced below. Much of this code is adapted from the docs:
```js
Office.context.ui.displayDialogAsync(
    DIALOG_URL, 
    { height: 30, width: 20, displayInIframe: true },
    (asyncResult) => {
      if(asyncResult) {
        const dialog = asyncResult.value

        function processMessage(arg) {
          var messageFromDialog = JSON.parse(arg.message);
          if(messageFromDialog.dialog === false) {
            dialog.close()
          }
        }
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);

      } else {
        console.error('asyncResult failed check', asyncResult)
      }
    }
    )
```

##Instructions

You can replicate this issue by cloning the repo, building, and side-loading the application.
`npm install`
`npm run dev`
Side load the application from `https://localhost:3000/manifest.xml`

I also deployed this add-in via surge. The manifest for the deployed version is at
`https://dialog-opener.surge.sh/manifest.xml`
You can easily test the deployed version of this code by side-loading the manifest.
