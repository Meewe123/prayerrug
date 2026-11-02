/**
 * Счётчик для страницы PrayerRug (site/index.html).
 * Лежит в Google Таблице: Расширения → Apps Script. Один раз запустить setup(),
 * потом «Развернуть → Новое развёртывание → Веб-приложение»:
 *   Выполнять от имени — «Я», Доступ — «Все».
 * Адрес развёртывания вставить в ENDPOINT в index.html.
 *
 * Листы:
 *   Visits  — каждый переход: время, ролик, первый ли раз с этого телефона, язык, откуда.
 *   Emails  — каждый адрес (повтор того же адреса не пишется), время, ролик.
 *   Summary — по каждому ролику: переходов всего, людей, адресов.
 */

function setup() {
  const ss = SpreadsheetApp.getActive();
  sheet_(ss, 'Visits', ['time', 'video', 'first', 'lang', 'referrer']);
  sheet_(ss, 'Emails', ['time', 'video', 'email']);
  const sum = sheet_(ss, 'Summary', ['video', 'opens', 'people', 'emails']);
  sum.getRange('A2').setFormula(
    "=IFERROR(SORT(UNIQUE(FILTER(Visits!B2:B, Visits!B2:B<>\"\"))), \"\")");
  sum.getRange('B2').setFormula(
    "=ARRAYFORMULA(IF(A2:A=\"\", \"\", COUNTIF(Visits!B2:B, A2:A)))");
  sum.getRange('C2').setFormula(
    "=ARRAYFORMULA(IF(A2:A=\"\", \"\", COUNTIFS(Visits!B2:B, A2:A, Visits!C2:C, TRUE)))");
  sum.getRange('D2').setFormula(
    "=ARRAYFORMULA(IF(A2:A=\"\", \"\", COUNTIF(Emails!B2:B, A2:A)))");
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const ss = SpreadsheetApp.getActive();
    const video = String(d.v || 'direct').slice(0, 24);
    if (d.t === 'visit') {
      ss.getSheetByName('Visits').appendRow([
        new Date(), video, d.first === true, String(d.lang || '').slice(0, 16),
        String(d.ref || '').slice(0, 200)
      ]);
    } else if (d.t === 'email') {
      const email = String(d.email || '').trim().toLowerCase().slice(0, 120);
      if (/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(email)) {
        const sheet = ss.getSheetByName('Emails');
        const last = sheet.getLastRow();
        const known = last > 1
          ? sheet.getRange(2, 3, last - 1, 1).getValues().some(r => r[0] === email)
          : false;
        if (!known) sheet.appendRow([new Date(), video, email]);
      }
    }
  } catch (err) {
    // Ничего не возвращаем наружу: страница ответа всё равно не читает.
  } finally {
    lock.releaseLock();
  }
  return ContentService.createTextOutput('ok');
}

function doGet() {
  return ContentService.createTextOutput('ok');
}

function sheet_(ss, name, header) {
  const sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  sheet.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
  sheet.setFrozenRows(1);
  return sheet;
}
