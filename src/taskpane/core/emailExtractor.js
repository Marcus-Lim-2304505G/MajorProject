import { extractUrlsFromText, extractUrlsFromHtml } from "../utils/urlUtils";

/**
 * Hardened extractor:
 * - Body text + HTML
 * - URLs from both
 * - sender + domain
 * - recipients
 * - attachments metadata
 */
export async function extractEmailData() {
  const item = Office.context.mailbox.item;
  if (!item) throw new Error("No email item found. Please open an email first.");

  const subject = item.subject || "";

  const senderEmail = item.from?.emailAddress || "Unknown";
  const senderName = item.from?.displayName || "";
  const senderDomain = getDomain(senderEmail);

  // replyTo can be array-like in some contexts; handle safely
  const replyToEmail = getReplyTo(item);

  const toRecipients = (item.to || []).map(r => r.emailAddress).filter(Boolean);
  const ccRecipients = (item.cc || []).map(r => r.emailAddress).filter(Boolean);

  const attachments = (item.attachments || []).map(a => ({
    name: a.name || "",
    contentType: a.contentType || "",
    size: typeof a.size === "number" ? a.size : null,
    isInline: !!a.isInline,
  }));

  // Body extraction: get both Text and HTML
  const bodyText = await getBody(item, Office.CoercionType.Text);
  const bodyHtml = await getBody(item, Office.CoercionType.Html);

  // URL extraction
  const urlsText = extractUrlsFromText(bodyText);
  const urlsHtml = extractUrlsFromHtml(bodyHtml);
  const urls = dedupe([...urlsText, ...urlsHtml]);

  // Convenience fields for rules/scoring later
  const derived = {
    numUrls: urls.length,
    numAttachments: attachments.length,
  };

  return {
    subject,
    sender: { email: senderEmail, name: senderName, domain: senderDomain },
    replyTo: replyToEmail ? { email: replyToEmail, domain: getDomain(replyToEmail) } : null,
    recipients: { to: toRecipients, cc: ccRecipients },
    body: { text: bodyText, html: bodyHtml },
    urls,
    attachments,
    derived,
  };
}

function getBody(item, coercionType) {
  return new Promise((resolve, reject) => {
    item.body.getAsync(coercionType, result => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value || "");
      } else {
        reject(new Error(`Unable to read email body (${coercionType}).`));
      }
    });
  });
}

function getDomain(email) {
  if (!email || !email.includes("@")) return "";
  return email.split("@").pop().toLowerCase().trim();
}

function getReplyTo(item) {
  try {
    // Some Outlook contexts expose replyTo as an array of recipients
    const rt = item.replyTo;
    if (!rt) return null;
    if (Array.isArray(rt) && rt.length > 0) return rt[0]?.emailAddress || null;
    // Some expose as an object
    if (rt.emailAddress) return rt.emailAddress;
    return null;
  } catch {
    return null;
  }
}

function dedupe(arr) {
  return [...new Set(arr)];
}
