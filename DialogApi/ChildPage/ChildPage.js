"use strict";
// The initialize function must be run each time a new page is loaded
window.Office.initialize = function (reason) {
  debugger;
  console.log("window.Office.initialize called");
  RegisterMessageChild();
}

function getCurentSource() {
  var source;
  if (!document.querySelector('[title="Office Add-in TwoWayMessageDialogTest"]')) {
    source = window.location.protocol + "//" + window.location.hostname + (window.location.port ? (":" + window.location.port) : "");
  } else {
    source = document.querySelector('[title="Office Add-in TwoWayMessageDialogTest"]').src;
  }
  document.getElementById('currentSource').innerText = "SOURCE: " + source;
}

function messageParent() {
  var value = document.getElementById("MessageForParent").value;
  if (!value) {
      value = "Message For Parent";
  }

  Office.context.ui.messageParent(value);
}

function showNotification(text) {
  document.getElementById('actionResult').innerText += text;
}

function addMessageStatus(arg) {
  showNotification(arg.message);
}

function RegisterMessageChild() {
  try {
    Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, addMessageStatus, onRegisterMessageComplete);
  } catch (e) {
    console.log("MESSAGE CHILD NOT SUPPORTED!!!");
  }
}

function onRegisterMessageComplete(asyncResult) {
  document.getElementById('actionResult').innerText += asyncResult.status;
  if(asyncResult.status != Office.AsyncResultStatus.Succeeded) {
    document.getElementById('actionResult').innerText += asyncResult.error.message;
  }
}

function redirect() {
  var value = document.getElementById("RedirectWebsite").value;
  if (!value) {
      console.log("Error: need a website in the textbox.");
      return;
  }
  window.location.href = value;
}
