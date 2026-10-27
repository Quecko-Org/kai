// Kai waitlist -> Google Sheet
//
// Receives the JSON that waitlist.html posts and appends one row per lead.
//
// Deploy:
//   1. Open the Sheet, Extensions > Apps Script, paste this file in.
//   2. Deploy > New deployment > Web app.
//      Execute as: Me. Who has access: Anyone.
//   3. Copy the /exec URL into WAITLIST_ENDPOINT in waitlist.html.
//   After editing this file, use Deploy > Manage deployments > Edit >
//   New version, so the /exec URL stays the same.

const SHEET_ID = "1EJ9w7TgzxCa0npU9JNX9zoFGeTcViPufAymO13S1QWs";

// header text -> field in the posted JSON
const COLUMNS = [
  ["Submitted at (UTC)", "submittedAtUtc"],
  ["First name", "firstName"],
  ["Last name", "lastName"],
  ["Email", "email"],
  ["Country of origin", "country"],
  ["Marketing consent", "marketingConsent"],
  ["Consent text shown", "consentTextShown"],
  ["Consent given at (UTC)", "consentGivenAtUtc"],
  ["Source page", "sourcePage"],
];

// rate limits (Apps Script cannot see the caller's IP address)
const SAME_EMAIL_SECONDS = 60;     // one submission per email per minute
const GLOBAL_PER_MINUTE = 30;      // all submissions together per minute

function doPost(e) {
  let data;
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    return reply({ ok: false, error: "bad_request" });
  }

  // honeypot: answer ok so bots learn nothing, but write no row
  if (data.company) return reply({ ok: true });

  const email = String(data.email || "").trim().toLowerCase();
  if (!data.firstName || !data.lastName || !data.country ||
      !/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(email)) {
    return reply({ ok: false, error: "invalid" });
  }

  const cache = CacheService.getScriptCache();
  const emailKey = "email:" + email;
  if (cache.get(emailKey)) return reply({ ok: false, error: "rate_limited" });

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const minuteKey = "minute:" + Math.floor(Date.now() / 60000);
    const count = Number(cache.get(minuteKey) || 0);
    if (count >= GLOBAL_PER_MINUTE) return reply({ ok: false, error: "rate_limited" });
    cache.put(minuteKey, String(count + 1), 120);
    cache.put(emailKey, "1", SAME_EMAIL_SECONDS);

    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
    const headers = ensureHeaders(sheet);
    const row = headers.map((header) => {
      const column = COLUMNS.find(([text]) => text === header);
      return column ? cellValue(column[1], data) : "";
    });
    sheet.appendRow(row);
  } finally {
    lock.releaseLock();
  }

  return reply({ ok: true });
}

// Writes the header row if the Sheet is empty, otherwise uses the one there.
function ensureHeaders(sheet) {
  if (sheet.getLastRow() === 0) {
    const headers = COLUMNS.map(([text]) => text);
    sheet.appendRow(headers);
    return headers;
  }
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    .map((h) => String(h).trim());
}

function cellValue(field, data) {
  const consent = data.marketingConsent === true;
  if (field === "marketingConsent") return consent ? "yes" : "no";
  if (field === "consentTextShown" || field === "consentGivenAtUtc") {
    return consent ? safe(data[field]) : "";
  }
  if (field === "submittedAtUtc") return safe(data.submittedAtUtc || new Date().toISOString());
  return safe(data[field]);
}

// stop a value like "=HYPERLINK(...)" from running as a formula in the Sheet
function safe(value) {
  const text = String(value == null ? "" : value).trim().slice(0, 500);
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

function reply(body) {
  return ContentService.createTextOutput(JSON.stringify(body))
    .setMimeType(ContentService.MimeType.JSON);
}
