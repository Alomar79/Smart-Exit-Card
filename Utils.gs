function fail_(message) { const error = new Error(message); error.isUserError = true; throw error; }
function api_(callback) {
  try { return {ok: true, data: JSON.parse(JSON.stringify(callback()))}; }
  catch (error) {
    console.error(error.stack || error);
    return {ok: false, message: error.isUserError ? error.message : 'تعذر تنفيذ العملية، حاول مرة أخرى'};
  }
}
function locked_(callback) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) fail_('النظام مشغول، حاول مرة أخرى');
  try { const result = callback(); SpreadsheetApp.flush(); return result; }
  finally { lock.releaseLock(); }
}
function id_(value) {
  if (typeof value !== 'string' || !/^[A-Za-z0-9_-]{1,100}$/.test(value)) fail_('تعذر التحقق من الاختيار');
  return value;
}
function active_(row) { return row && String(row.Status).toLowerCase() === 'active'; }
function one_(rows, key, value, message) {
  const matches = rows.filter(row => row[key] === value);
  if (matches.length !== 1) fail_(message || 'تعذر التحقق من البيانات، راجع مسؤول النظام');
  return matches[0];
}
function staffName_(value) {
  if (typeof value !== 'string') fail_('يرجى كتابة اسم المعلمة');
  const name = value.replace(/[\u0000-\u001f\u007f\u200b-\u200f\u202a-\u202e\u2066-\u2069\ufeff]/g, '').replace(/\s+/g, ' ').trim();
  if (!name) fail_('يرجى كتابة اسم المعلمة');
  if (name.length > 100) fail_('يرجى اختصار اسم المعلمة');
  return name;
}
function dateMs_(value) {
  const time = value instanceof Date ? value.getTime() : NaN;
  if (!Number.isFinite(time)) throw new Error('Invalid Log date');
  return time;
}
function logDto_(row, now) {
  const copy = Object.assign({}, row);
  delete copy._row;
  copy.ExitTime = dateMs_(row.ExitTime);
  copy.ReturnTime = row.ReturnTime === '' ? null : dateMs_(row.ReturnTime);
  copy.ElapsedMinutes = Math.max(0, (now.getTime() - copy.ExitTime) / 60000);
  return copy;
}
function personKey_(row) { return row.StudentSource + ':' + row.StudentID; }
