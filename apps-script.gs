/**
 * Pinly — odbiornik zapisów na testy beta (beta.html).
 *
 * INSTALACJA (ok. 5 minut, wszystko za darmo):
 *
 *  1. Wejdź na https://sheets.new i utwórz nowy arkusz, np. "Pinly — beta testerzy".
 *  2. W arkuszu: Rozszerzenia → Apps Script.
 *  3. Skasuj cały kod, który tam jest, i wklej ten plik.
 *  4. Zmień SECRET poniżej na własny losowy ciąg znaków.
 *  5. Zapisz (ikona dyskietki).
 *  6. Wdróż → Nowe wdrożenie → koło zębate → Aplikacja internetowa:
 *        - Wykonaj jako:      Ja
 *        - Kto ma dostęp:     Wszyscy   <-- KONIECZNIE, inaczej formularz nie zadziała
 *     → Wdróż → zatwierdź uprawnienia (Google pokaże ostrzeżenie
 *       "aplikacja niezweryfikowana" — to Twój własny skrypt: Zaawansowane → Przejdź do…).
 *  7. Skopiuj "URL aplikacji internetowej" (kończy się na /exec).
 *  8. W pliku beta.html wklej ten URL do ENDPOINT, a swój klucz do SECRET.
 *
 * UWAGA: po każdej zmianie tego kodu zrób Wdróż → Zarządzaj wdrożeniami → edytuj →
 * Wersja: Nowa wersja. Bez tego działa stara wersja skryptu.
 */

// Musi być identyczny jak SECRET w beta.html.
var SECRET = '34NJJ4N33JNER';

var SHEET_NAME = 'Beta testerzy';
var HEADERS = ['Data zapisu', 'E-mail', 'Źródło'];

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
  } catch (err) {
    return json({ ok: false, error: 'busy' });
  }

  try {
    var p = (e && e.parameter) || {};

    // Honeypot: pole niewidoczne dla ludzi. Wypełnione = bot. Udajemy sukces.
    if (p.hp) return json({ ok: true });

    if (SECRET && p.secret !== SECRET) return json({ ok: false, error: 'bad_secret' });

    var email = String(p.email || '').trim().toLowerCase();
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(email) || email.length > 254) {
      return json({ ok: false, error: 'bad_email' });
    }

    var sheet = getSheet();
    if (isDuplicate(sheet, email)) return json({ ok: true, duplicate: true });

    sheet.appendRow([
      new Date(),
      email,
      String(p.source || '').trim().slice(0, 40)
    ]);

    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return json({ ok: true, msg: 'Pinly beta endpoint dziala' });
}

function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function isDuplicate(sheet, email) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;
  var values = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim().toLowerCase() === email) return true;
  }
  return false;
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
