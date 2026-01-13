// Extract http/https URLs from text
export function extractUrls(text) {
  if (!text) return [];
  const urlRegex = /(https?:\/\/[^\s<>"')\]]+)/gi;
  const matches = text.match(urlRegex) || [];
  // Deduplicate
  return [...new Set(matches)];
}
