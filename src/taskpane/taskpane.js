/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
/* global document, fetch, Office */

import env from "../../config";
import { allComponents, provideFluentDesignSystem } from "@fluentui/web-components";
provideFluentDesignSystem().register(allComponents);

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("compose-copy-paste-button").onclick = setBodyInCompose;
    document.getElementById("open-reply-button").onclick = openReplyToCurrentEmail;
    document.getElementById("full-functionality-button").onclick = replyAIMain;
  }
});

export async function replyAIMain() {
  document.getElementById("response-text").innerHTML = "<fluent-progress-ring></fluent-progress-ring>";
  const userInput = document.getElementById("user-input").value;
  const apiPrompt = await buildPrompt(userInput);
  // document.getElementById("test-text").innerHTML = apiPrompt;
  setOpenAiText(apiPrompt);
}

export async function getSubjectText() {
  if (Office.context.mailbox.item.displayReplyForm != undefined) {
    // read mode
    return Office.context.mailbox.item.subject;
  } else {
    // compose mode
    return new Promise((resolve, reject) => {
      Office.context.mailbox.item.subject.getAsync(function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed) {
          reject(new Error("Failed to retrieve subject text."));
        } else {
          resolve(asyncResult.value);
        }
      });
    });
  }
}

export async function getBodyText() {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.body.getAsync("text", function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(new Error("Failed to retrieve body text."));
      }
    });
  });
}

export async function userInput() {
  const txt1 = document.getElementById("tbinput");
  const out1 = document.getElementById("output1");
  out1.innerHTML = txt1.value;
  return txt1.value;
}

export async function buildPrompt(userInput) {
  const stringBuilder = [];
  stringBuilder.push("Email Subject: " + (await getSubjectText()) + " ");
  stringBuilder.push("Email Body: " + (await getBodyText()) + " ");
  stringBuilder.push("Response Instructions: " + userInput + " ");
  stringBuilder.push(
    `Request: I need assistance in crafting a response to this email. Please help me by providing a coherent 
    and formal reply based on the given subject line and email body. Use the information provided in 
    Response Instructions to compose the email; do not add any new information or hallucinate any 
    details. Ensure that the response addresses the sender's concerns, but do not address the concern 
    if there is no answer provided in Response Instructions.`
  );
  stringBuilder.push(
    `Please do not generate any new information or fabricate any details beyond what is provided 
    in the subject line, email body, and any additional context. The response should be based 
    solely on the given information and should not add any speculative or false information.`
  );
  stringBuilder.push(
    `once you finished the email, end with EMAILISFINISHED 
    -------------- start writing ------------------------
    `
  );

  const result = stringBuilder.join("");
  // document.getElementById("prompt-text").innerHTML = "<b>Prompt:</b> <br/>" + result;
  return result;
}

export async function setOpenAiText(apiPrompt) {
  const body = {
    prompt: apiPrompt,
    temperature: 0.69,
    top_p: 0.5,
    frequency_penalty: 0,
    presence_penalty: 0,
    max_tokens: 100,
    stop: ["EMAILISFINISHED"],
  };
  const response = await fetch(env.OPENAI_ENDPOINT, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "api-key": env.OPENAI_KEY,
    },
    body: JSON.stringify(body),
  });
  // document.getElementById("openai-status-text").innerHTML = "<b>Status:</b> <br/>" + response.status;
  const json = await response.json();
  // const jsonString = JSON.stringify(json, null, 2);
  // document.getElementById("openai-text").innerHTML = "<b>Response:</b> <br/>" + jsonString;

  document.getElementById("response-text").innerHTML = "<b>Response:</b> <br/>" + json.choices[0].text;
  return json.choices[0].text;
}

// redundant
export async function setBodyInCompose() {
  const testString = "ahhhhhhhhhhhhhhhhhhhhhhh it actually workeddd??!!!";
  Office.context.mailbox.item.body.prependAsync(testString);
  document.getElementById("compose-copy-paste-text").innerHTML = "<b>Executed</b>";
}

export async function openReplyToCurrentEmail() {
  const testString = "ahhhhhhhhhhhhhhhhhhhhhhh it actually workeddd??!!!";
  Office.context.mailbox.item.displayReplyForm(testString);
}
