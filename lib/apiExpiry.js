// Matches the duration options Odoo itself offers when generating an API key.
export const EXPIRY_OPTIONS = ["1 Day", "1 Week", "1 Month", "3 Months", "6 Months", "1 Year", "Persistent Key", "Custom Date"];

// Given a preset option, return the ISO expiry timestamp measured from now.
// "Persistent Key" never expires -> null.
export function computeExpiryFromOption(option, customDate) {
  if (option === "Persistent Key") return null;
  if (option === "Custom Date") return customDate ? new Date(customDate).toISOString() : null;
  const now = new Date();
  const map = {
    "1 Day": () => now.setDate(now.getDate() + 1),
    "1 Week": () => now.setDate(now.getDate() + 7),
    "1 Month": () => now.setMonth(now.getMonth() + 1),
    "3 Months": () => now.setMonth(now.getMonth() + 3),
    "6 Months": () => now.setMonth(now.getMonth() + 6),
    "1 Year": () => now.setFullYear(now.getFullYear() + 1),
  };
  if (map[option]) { map[option](); return now.toISOString(); }
  return null;
}

export function isExpired(apiKeyExpiresAt) {
  if (!apiKeyExpiresAt) return false; // no expiry set = persistent key
  return new Date(apiKeyExpiresAt).getTime() < Date.now();
}

export function expiryLabel(apiKeyExpiresAt) {
  if (!apiKeyExpiresAt) return "Persistent key — no expiry";
  const d = new Date(apiKeyExpiresAt);
  const formatted = d.toLocaleString(undefined, { day: "numeric", month: "short", year: "numeric", hour: "2-digit", minute: "2-digit" });
  return isExpired(apiKeyExpiresAt) ? `Expired ${formatted}` : `API key expires ${formatted}`;
}
