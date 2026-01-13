console.log("PhishCheck loaded Successfully");

/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */
/* <backup code> <remove this line if needed>
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  /**
   * Insert your Outlook code here
   */
/* <backup code> <remove this line if needed>
  const item = Office.context.mailbox.item;
  let insertAt = document.getElementById("item-subject");
  let label = document.createElement("b").appendChild(document.createTextNode("Subject: "));
  insertAt.appendChild(label);
  insertAt.appendChild(document.createElement("br"));
  insertAt.appendChild(document.createTextNode(item.subject));
  insertAt.appendChild(document.createElement("br"));
}
<remove this line if needed> */ 

/* global Office */

/* <backup code 2> <delete this line if needed>
Office.onReady(() => {
  const btn = document.getElementById("analyseBtn");
  if (btn) btn.onclick = analyseEmail;
});

function analyseEmail() {
  const item = Office.context.mailbox.item;

  const emailData = {
    subject: item.subject || "",
    sender: item.from ? item.from.emailAddress : "Unknown",
    body: ""
  };

  item.body.getAsync(Office.CoercionType.Text, (result) => {
    emailData.body =
      result.status === Office.AsyncResultStatus.Succeeded ? result.value : "";

    const output = document.getElementById("output");
    if (output) {
      output.textContent = JSON.stringify(emailData, null, 2);
    }

    console.log("Extracted Email Data:", emailData);
  });
}
<backup code 2> <delete this line if needed> */

/* global Office */

/* <backup code 3> <remove this line if needed>
Office.onReady(() => {
  const btn = document.getElementById("analyseBtn");
  if (btn) btn.onclick = analyseEmail;

  setOutput("Ready. Open an email and click “Analyse Email”.");
});

function setOutput(text) {
  const output = document.getElementById("output");
  if (output) output.textContent = text;
}

function analyseEmail() {
  const item = Office.context.mailbox.item;
  if (!item) {
    setOutput("Error: No email item found. Please open an email first.");
    return;
  }

  setOutput("Analyzing email…");

  const emailData = {
    subject: item.subject || "",
    sender: item.from ? item.from.emailAddress : "Unknown",
    body: ""
  };

  item.body.getAsync(Office.CoercionType.Text, (result) => {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      setOutput("Error: Unable to read email body.");
      console.error("Body extraction failed:", result.error);
      return;
    }

    // Trim body to avoid UI freezing for long emails
    const rawBody = result.value || "";
    emailData.body = rawBody.length > 3000 ? rawBody.slice(0, 3000) + "\n\n[Body trimmed]" : rawBody;

    console.log("Extracted Email Data:", emailData);

    setOutput(JSON.stringify(emailData, null, 2));
  });
}
<backup code 3> <remove this line if needed> */

import { extractEmailData } from "./core/emailExtractor";

/* global Office */

Office.onReady(() => {
  const btn = document.getElementById("analyseBtn");
  if (btn) btn.onclick = analyseEmail;

  setOutput("Ready. Open an email and click “Analyse Email”.");
});

function setOutput(text) {
  const output = document.getElementById("output");
  if (output) output.textContent = text;
}

function trimForUI(text, maxLen = 3000) {
  if (!text) return "";
  return text.length > maxLen ? text.slice(0, maxLen) + "\n\n[Body trimmed]" : text;
}

async function analyseEmail() {
  try {
    setOutput("Analyzing email…");

    const emailData = await extractEmailData();

    // Don’t freeze UI with huge bodies
    const uiData = {
      ...emailData,
      body: trimForUI(emailData.body),
    };

    console.log("Extracted Email Data:", emailData);
    setOutput(JSON.stringify(uiData, null, 2));
  } catch (err) {
    console.error(err);
    setOutput(`Error: ${err.message}`);
  }
}
