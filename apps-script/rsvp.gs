/**
 * RSVP & Ucapan backend for the Adibah & Hazman invitation (v2).
 *
 * Paste this into the Google Sheet's Extensions → Apps Script, then
 * Deploy → New deployment → Web app (Execute as: Me, Who has access: Anyone).
 * Put the web-app URL into SHEET_API in v2/index.html.
 *
 * GET  → { hadir, mungkin, tidak, ucapan: [{nama, ucapan}] }   (newest first)
 * POST → body {nama, kehadiran: "Ya"|"Tidak"|"Mungkin", bilangan, ucapan} → { ok }
 *
 * Only names and wishes are ever returned; delete a row in the sheet to remove a wish.
 */

function sheet_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
}

// Stop text like "=SUM(...)" from being stored as a formula.
function clean_(v, max) {
  var s = String(v == null ? "" : v).trim().slice(0, max);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  var rows = sheet_().getDataRange().getValues().slice(1);
  var out = { hadir: 0, mungkin: 0, tidak: 0, ucapan: [] };
  rows.forEach(function (r) {
    var n = Number(r[3]) || 1;
    if (r[2] === "Ya") out.hadir += n;
    else if (r[2] === "Mungkin") out.mungkin += n;
    else if (r[2] === "Tidak") out.tidak += 1;
    var u = String(r[4]).replace(/^'/, "").trim();
    if (u) out.ucapan.push({ nama: String(r[1]).replace(/^'/, ""), ucapan: u });
  });
  out.ucapan = out.ucapan.reverse().slice(0, 300);
  return json_(out);
}

function doPost(e) {
  var d = JSON.parse(e.postData.contents);
  var nama = clean_(d.nama, 80);
  if (!nama) return json_({ ok: false });
  var k = ["Ya", "Tidak", "Mungkin"].indexOf(d.kehadiran) >= 0 ? d.kehadiran : "Ya";
  var n = Math.max(1, Math.min(20, parseInt(d.bilangan, 10) || 1));
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    sheet_().appendRow([new Date(), nama, k, n, clean_(d.ucapan, 500)]);
  } finally {
    lock.releaseLock();
  }
  return json_({ ok: true });
}
