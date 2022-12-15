"use strict";
// The initialize function must be run each time a new page is loaded
var _dialog;
var _childPageUrl = "https://iotataru.github.io/dialogExample/DialogApi/ChildPage/";

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
