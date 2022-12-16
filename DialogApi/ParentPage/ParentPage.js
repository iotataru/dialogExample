"use strict";
// The initialize function must be run each time a new page is loaded
var _dialog;
var _childPageUrl = "https://iotataru.github.io/dialogExample/DialogApi/ChildPage/";

async function writeValues(arg)
{
  console.log('--- Write value called with arg: ', arg);
  //let day = JSON.parse(arg.message);
  let day = arg.message;
  await Excel.run(async (context) => {
    const ws = context.workbook.names.getItem('rngDay').getRange();
    ws.load(['worksheet']);
    await context.sync().catch((error) => {
        console.log("--- Error line 21");
    });
    const workSheetName = ws.worksheet.name;
    const password='password1'
    console.log("--- About to unprotect sheet!");
    await this.toggleSheetProtection(workSheetName, 'unprotect', password);
    console.log(`--- Unprotect complete! Going to write value: ${day}`);
    ws.values = [[day]];
    console.log("--- Value written! Going to sync.");
    await context.sync().catch((error) => {
        console.log("--- Error line 30");
        console.log(error)
    });
    console.log("--- Sync complete. Going to turn protection back on!");
    await this.toggleSheetProtection(workSheetName, 'protect', password);
    console.log("--- Protection is now back on!");
    }).catch((error) => {
        console.log("--- Error line 35");
    });
}

async function toggleSheetProtection(
    sheetName,
    request,
    password) {
    console.log("---parent: toggleSheetProtection")
    await Excel.run(async (context) => {
      //console.log("toggleSheetProtection called: ", context);
      const requiredSheet = context.workbook.worksheets.getItem(sheetName);
      requiredSheet.load('protection');
      await context.sync().catch((error) => {
        console.log("--- Error line 47");
      });
      if (request === 'protect' && !requiredSheet.protection.protected) {
        requiredSheet.protection.protect(
          {
            allowEditObjects: true,
            allowAutoFilter: true,
            allowFormatRows: true,
            allowFormatColumns: true,
          },
          password
        );
      } else if (
        request === 'unprotect' &&
        requiredSheet.protection.protected
      ) {
        requiredSheet.protection.unprotect(password)
      }
  
      await context.sync().catch((error) => {
        console.log("--- Error line 67");
      });
    })
    .catch((error) => {
        console.log("--- Error line 71");
    });
}

function getSettings() {
    var settings = Office.context.document.settings;
    console.log("settings: ", settings);
}

function windowOpen() {
    var urlLaunch = !!(document.getElementById("WindowOpenLaunch").value) ? document.getElementById("WindowOpenLaunch").value : _childPageUrl;
    window.open(urlLaunch);
}

function getCurentSource() {
    var source;
    if (!document.querySelector('[title="Office Add-in TwoWayMessageDialogTest"]')) {
        source = window.location.protocol + "//" + window.location.hostname + (window.location.port ? (":" + window.location.port) : "");
    } else {
        source = document.querySelector('[title="Office Add-in TwoWayMessageDialogTest"]').src;
    }
    document.getElementById('currentSource').innerText = "SOURCE: " + source;
    
    var requirementSetDialogOrigin1 = Office.context.requirements.isSetSupported("DialogOrigin", 1.1);
    var requirementSetDialogOrigin2 = Office.context.requirements.isSetSupported("DialogOrigin", 1.2);
    console.log("requirementSetDialogOrigin1: " + requirementSetDialogOrigin1);
    console.log("requirementSetDialogOrigin2: " + requirementSetDialogOrigin2);
}

function showNotification(text) {
    document.getElementById('actionResult').innerText += text;
}

function launchDialogCallback(arg) {
    if (arg.status === "failed") {
        showNotification("launch dialog failed");
    }
    else {
        _dialog = arg.value;
        _dialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, addMessageStatus);
        _dialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogEventReceived, addCloseStatus);
        //setTimeout(messageChildInitial, 5000);
    }
}

function addMessageStatus(arg) {
    console.log("addMessageStatus called with value: ", arg);
    writeValues(arg);
    if (arg.message === "ping!") {
        messageChild("pong!");
    } else if (arg.message === "closeme") {
        closeDialog();
    }
    showNotification(JSON.stringify(arg));
}

function addCloseStatus(arg) {
    showNotification("dialog closed");
}

function launchInlineDialog() {
    var dialogUrl = !!(document.getElementById("InlineLaunch").value) ? document.getElementById("InlineLaunch").value : _childPageUrl;
    Office.context.ui.displayDialogAsync(dialogUrl,
  {height:80, width:50, hideTitle: false, promptBeforeOpen: true, enforceAppDomain: true, displayInIframe:true},
  launchDialogCallback);
}

function launchWindowDialog() {
    var dialogUrl = !!(document.getElementById("WindowLaunch").value) ? document.getElementById("WindowLaunch").value : _childPageUrl;
    Office.context.ui.displayDialogAsync(dialogUrl,
  {height:80, width:50, hideTitle: false, promptBeforeOpen: true, enforceAppDomain: true},
  launchDialogCallback);
}

function launchInlineDialogFromRibbon(args) {
    Office.context.ui.displayDialogAsync(_childPageUrl, { height: 50, width: 30, promptBeforeOpen: true, displayInIframe: true }, launchDialogCallback);

    args.completed();
}

function launchWindowDialogFromRibbon(args) {
    Office.context.ui.displayDialogAsync(_childPageUrl, { height: 50, width: 30, promptBeforeOpen: true, displayInIframe: false }, launchDialogCallback);

    args.completed();
}

function messageChildInitial() {
    messageChild("Initial message for child upon parent's launchDialogCallback");
}

function messageChild() {
    messageChild("");
}

function messageChild(message) {
    var value = document.getElementById("MessageForChild").value;
    if (!value) {
  value = message;
  if (!value) {
      value = "Message For Child";
  }
    }

    _dialog.messageChild(value);
}

function closeDialog() {
    _dialog.close();
}

function redirect() {
    var value = document.getElementById("RedirectWebsite").value;
    if (!value) {
        console.log("Error: need a website in the textbox.");
        return;
    }
    window.location.href = value;
}
