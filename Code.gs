// Google Sheet backend for the result website.
// Sheet 1 columns (row 1 headings): Register No | Name | Year | DOB | Theory | Internal
function doGet(e) {
  var p = e.parameter || {};
  var res = lookup(String(p.reg || '').trim(), String(p.dob || '').trim());
  return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
}

function normDob(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  }
  var s = String(v).trim();
  var m = s.match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4})$/);
  if (m) return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return m ? m[0] : '';
}

function lookup(reg, dob) {
  var bad = { ok: false, msg: 'Register number or date of birth is incorrect. Check and try again.' };
  if (!reg || !dob) return bad;
  var cache = CacheService.getScriptCache();
  var key = 'att_' + reg;
  var tries = Number(cache.get(key) || 0);
  if (tries >= 5) return { ok: false, msg: 'Too many wrong attempts. Try again after 10 minutes.' };

  var rows = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0].getDataRange().getValues();
  var h = rows[0].map(function (x) { return String(x).trim().toLowerCase(); });
  var iR = h.indexOf('register no'), iN = h.indexOf('name'), iY = h.indexOf('year'),
      iD = h.indexOf('dob'), iT = h.indexOf('theory'), iI = h.indexOf('internal');
  if (iR < 0 || iN < 0 || iD < 0 || iT < 0 || iI < 0) return { ok: false, msg: 'Sheet headings are wrong.' };

  for (var r = 1; r < rows.length; r++) {
    if (String(rows[r][iR]).trim() === reg && normDob(rows[r][iD]) === dob) {
      cache.remove(key);
      return {
        ok: true,
        name: String(rows[r][iN]).trim(),
        year: iY >= 0 ? String(rows[r][iY]).trim() : '',
        theory: Number(rows[r][iT]),
        internal: Number(rows[r][iI])
      };
    }
  }
  cache.put(key, String(tries + 1), 600);
  return bad;
}
