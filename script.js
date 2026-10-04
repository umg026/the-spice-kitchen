function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    // ---- NEW: visitor tracking goes to the "Track" tab ----
    if (data && data.type === "track") {
      return handleTrack_(data);
    }

    // ---- EXISTING: orders go to the "Orders" tab (unchanged) ----
    const sheet = SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName("Orders");

    if (!sheet) {
      throw new Error('Sheet "Orders" not found.');
    }

    // Convert items array into readable text
    let itemsText = "";

    if (Array.isArray(data.items)) {
      itemsText = data.items
        .map(item => {
          return `${item.name} x ${item.quantity}`;
        })
        .join(" | ");
    } else {
      itemsText = data.items || "";
    }

    sheet.appendRow([
      data.orderId || "",
      data.timestamp || new Date().toISOString(),
      data.customerName || "",
      data.phone || "",
      data.address || "",
      data.postcode || "",
      data.deliveryTime || "",
      itemsText,
      data.total || ""
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({
        success: true
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {

    console.error(error);

    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        error: error.toString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/* =========================================================
   TRACK TAB
   ========================================================= */

const TRACK_TAB_NAME = "Track";

const TRACK_HEADERS = [
  "Server Time", "Client Time (UTC)", "Event", "Label", "Value", "Details",
  "Visitor ID", "Session ID", "Visit #", "New/Returning",
  "IP", "City", "Region", "Country", "ISP",
  "Device", "Device Model", "OS", "OS Version", "Browser", "Browser Version", "In-App Browser",
  "Screen", "Viewport", "Pixel Ratio", "Language", "Timezone", "Connection", "Colour Mode",
  "Referrer", "Source", "Medium", "Campaign", "Term", "Content", "Click ID", "First Source", "Landing URL"
];

function handleTrack_(data) {
  const events = Array.isArray(data.events) ? data.events.slice(0, 50) : [];
  const c = data.ctx || {};
  if (!c.vid || !events.length) return ContentService.createTextOutput("ignored");

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName(TRACK_TAB_NAME);
    if (!sh) {
      sh = ss.getSheets().filter(function (s) {
        return s.getName().toLowerCase() === TRACK_TAB_NAME.toLowerCase();
      })[0];
    }
    if (!sh) sh = ss.insertSheet(TRACK_TAB_NAME);

    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, TRACK_HEADERS.length).setValues([TRACK_HEADERS]).setFontWeight("bold");
      sh.setFrozenRows(1);
    }

    const now = new Date();
    const rows = events.map(function (ev) {
      return [
        now, safe_(ev.t), safe_(ev.e), safe_(ev.l),
        (ev.v === "" || isNaN(Number(ev.v))) ? safe_(ev.v) : Number(ev.v),
        safe_(ev.d),
        safe_(c.vid), safe_(c.sid), safe_(c.visit), safe_(c.returning),
        safe_(c.ip), safe_(c.city), safe_(c.region), safe_(c.country), safe_(c.isp),
        safe_(c.device), safe_(c.model), safe_(c.os), safe_(c.osVer), safe_(c.browser), safe_(c.bver), safe_(c.inApp),
        safe_(c.screen), safe_(c.viewport), safe_(c.dpr), safe_(c.lang), safe_(c.tz), safe_(c.conn), safe_(c.dark),
        safe_(c.ref), safe_(c.source), safe_(c.medium), safe_(c.campaign), safe_(c.term), safe_(c.content), safe_(c.clickId),
        safe_(c.firstSource), safe_(c.landing)
      ];
    });
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, TRACK_HEADERS.length).setValues(rows);
  } finally {
    lock.releaseLock();
  }
  return ContentService.createTextOutput("ok");
}

// The endpoint is public, so stop visitor-supplied text from becoming a spreadsheet formula.
function safe_(v) {
  if (v === null || v === undefined) return "";
  v = String(v).slice(0, 500);
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}