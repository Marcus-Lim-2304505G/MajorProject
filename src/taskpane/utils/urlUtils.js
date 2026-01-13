// Extract http/https URLs from plain text
export function extractUrlsFromText(text) {
  if (!text) return [];
  const urlRegex = /(https?:\/\/[^\s<>"')\]]+)/gi;
  return dedupe(text.match(urlRegex) || []);
}

// Extract URLs from HTML href attributes
export function extractUrlsFromHtml(html) {
  if (!html) return [];

  try {
    const doc = new DOMParser().parseFromString(html, "text/html");
    const anchors = Array.from(doc.querySelectorAll("a[href]"));
    const hrefs = anchors
      .map(a => (a.getAttribute("href") || "").trim())
      .filter(href => href && href !== "#");

    // Keep only http/https and mailto (useful for phishing too)
    const cleaned = hrefs
      .map(normalizeHref)
      .filter(Boolean);

    return dedupe(cleaned);
  } catch {
    return [];
  }
}

function normalizeHref(href) {
  // Some emails use relative links; those are not useful here.
  if (href.startsWith("http://") || href.startsWith("https://")) return href;
  if (href.startsWith("mailto:")) return href;
  return null;
}

function dedupe(arr) {
  return [...new Set(arr)];
}
