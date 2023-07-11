/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("subject-button").onclick = setSubjectText;
    document.getElementById("body-button").onclick = setBodyText;
  }
});

export async function setSubjectText() {
  // Get a reference to the current message
  const subject = Office.context.mailbox.item.subject;
  document.getElementById("subject-text").innerHTML = "<b>Subject:</b> <br/>" + subject;
}

export async function setBodyText() {
  // Get a reference to the current message
  const item = Office.context.mailbox.item;

  item.body.getAsync("text", function (result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      document.getElementById("body-text").innerHTML = "<b>Body:</b> <br/>" + result.value;
    }
  });
}
