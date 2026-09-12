function spreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) throw new Error('Run setupProject in the bound spreadsheet first');
  return SpreadsheetApp.openById(id);
}
function validHeaders_(values, name) {
  return values[0].length === SCHEMA[name].length && SCHEMA[name].every((key, i) => values[0][i] === key);
}
function readTable_(ss, name) {
  const sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error('Missing sheet: ' + name);
  const values = sheet.getDataRange().getValues();
  if (!validHeaders_(values, name)) throw new Error('Invalid headers: ' + name);
  return values.slice(1).map((cells, index) => {
    const row = {_row: index + 2};
    SCHEMA[name].forEach((key, i) => { row[key] = typeof cells[i] === 'string' ? cells[i].trim() : cells[i]; });
    return row;
  }).filter(row => SCHEMA[name].some(key => row[key] !== ''));
}
function database_() {
  const ss = spreadsheet_();
  const db = {ss: ss};
  Object.keys(SCHEMA).forEach(name => { db[name] = readTable_(ss, name); });
  return db;
}
function settings_(db) {
  const settings = Object.assign({}, DEFAULT_SETTINGS);
  const seen = new Set();
  db.Settings.forEach(row => {
    if (seen.has(row.Key)) throw new Error('Duplicate setting: ' + row.Key);
    seen.add(row.Key);
    if (Object.prototype.hasOwnProperty.call(DEFAULT_SETTINGS, row.Key)) settings[row.Key] = row.Value;
  });
  ['GREEN_LIMIT', 'ORANGE_LIMIT', 'AUTO_REFRESH_SECONDS'].forEach(key => { settings[key] = Number(settings[key]); });
  if (!(settings.GREEN_LIMIT > 0 && settings.ORANGE_LIMIT > settings.GREEN_LIMIT && settings.AUTO_REFRESH_SECONDS >= 5) ||
      ![settings.GREEN_LIMIT, settings.ORANGE_LIMIT, settings.AUTO_REFRESH_SECONDS].every(Number.isFinite)) throw new Error('Invalid thresholds or refresh interval');
  if (settings.TIMEZONE !== 'Asia/Qatar') throw new Error('TIMEZONE must be Asia/Qatar');
  return settings;
}
function cardContext_(db, cardId, requireActive) {
  const message = 'البطاقة غير موجودة أو غير مفعلة';
  const card = one_(db.Cards, 'CardID', id_(cardId), message);
  const group = one_(db.Groups, 'GroupID', card.GroupID, message);
  if (requireActive && (!active_(card) || !active_(group))) fail_(message);
  if (db.Cards.filter(row => row.GroupID === group.GroupID).length !== 1) fail_('تعذر التحقق من البطاقة، راجع مسؤول النظام');
  if (!Object.prototype.hasOwnProperty.call(STUDENT_SOURCES, group.CardTheme)) throw new Error('Invalid card theme');
  return {card: card, group: group, source: STUDENT_SOURCES[group.CardTheme]};
}
function allowedDestination_(row, theme) { return active_(row) && (row.Availability === theme || row.Availability === 'All'); }
function validateData_(db) {
  const issues = dataIssues_(db);
  if (issues.length) { console.error(JSON.stringify(issues)); fail_('تعذر التحقق من البيانات، راجع مسؤول النظام'); }
}
function appendLog_(db, record) {
  // Preserve snapshots as literal text, including staff input that resembles a formula.
  const values = SCHEMA.Log.map(key => typeof record[key] === 'string' && /^[=+@-]/.test(record[key]) ? "'" + record[key] : record[key]);
  const sheet = db.ss.getSheetByName('Log');
  sheet.getRange(sheet.getLastRow() + 1, 1, 1, values.length).setValues([values]);
}
