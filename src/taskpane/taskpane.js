console.log("PhishCheck loaded Successfully");

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

// Trim ONLY for display to avoid freezing pane (keep full body internally)
function trimForUI(data) {
  const copy = structuredCloneSafe(data);

  if (copy?.body?.text && copy.body.text.length > 2000) {
    copy.body.text = copy.body.text.slice(0, 2000) + "\n\n[Text trimmed]";
  }
  if (copy?.body?.html && copy.body.html.length > 2000) {
    copy.body.html = copy.body.html.slice(0, 2000) + "\n\n[HTML trimmed]";
  }
  return copy;
}

async function analyseEmail() {
  try {
    setOutput("Analyzing email…");

    const emailData = await extractEmailData();
    console.log("Extracted Email Data:", emailData);

    const uiData = trimForUI(emailData);
    setOutput(JSON.stringify(uiData, null, 2));
  } catch (err) {
    console.error(err);
    setOutput(`Error: ${err.message}`);
  }
}

// Avoid structuredClone issues in some runtimes
function structuredCloneSafe(obj) {
  try {
    return structuredClone(obj);
  } catch {
    return JSON.parse(JSON.stringify(obj));
  }
}
