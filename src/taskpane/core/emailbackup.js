import { extractUrls } from "../utils/urlUtils";

/**
 * Extracts accessible email data from the currently opened Outlook item.
 * Fully client-side (Office.js).
 */
export async function extractEmailData() {
  const item = Office.context.mailbox.item;
  if (!item) {
    throw new Error("No email item found. Please open an email first.");
  }

  const subject = item.subject || "";
  const sender = item.from ? item.from.emailAddress : "Unknown";

  const attachments = (item.attachments || []).map((a) => ({
    name: a.name,
    contentType: a.contentType,
    size: a.size,
  }));

  const body = await getBodyText(item);
  const urls = extractUrls(body);

  // Keep original (full) body for rules; UI can choose to trim.
  return {
    subject,
    sender,
    body,
    urls,
    attachments,
  };
}

function getBodyText(item) {
  return new Promise((resolve, reject) => {
    item.body.getAsync(Office.CoercionType.Text, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value || "");
      } else {
        reject(new Error("Unable to read email body."));
      }
    });
  });
}
