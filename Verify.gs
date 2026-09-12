function dataIssues_(db) {
  const issues = [];
  function add(code, detail) { issues.push({code: code, detail: detail}); }
  Object.keys(SCHEMA).forEach(name => {
    const seen = new Set();
    db[name].forEach(row => {
      const key = row[SCHEMA[name][0]];
      if (!key || seen.has(key)) add('DUPLICATE_ID', name + ' row ' + row._row + ': ' + key);
      seen.add(key);
      if (name !== 'Settings' && name !== 'Log' && !['active','inactive'].includes(String(row.Status).toLowerCase())) add('INVALID_STATUS', name + ': ' + key);
      if (name !== 'Settings' && (typeof key !== 'string' || !/^[A-Za-z0-9_-]{1,100}$/.test(key))) add('INVALID_ID', name + ': ' + key);
    });
  });
  const groups = new Map(db.Groups.map(row => [row.GroupID, row]));
  db.Groups.forEach(group => {
    if (!Object.prototype.hasOwnProperty.call(STUDENT_SOURCES, group.CardTheme)) add('INVALID_THEME', group.GroupID);
    if (!group.GroupName) add('EMPTY_NAME', 'Groups: ' + group.GroupID);
    const count = db.Cards.filter(card => card.GroupID === group.GroupID).length;
    if (count === 0) add('MISSING_CARD', group.GroupID);
    if (count > 1) add('MULTIPLE_CARDS', group.GroupID);
  });
  db.Cards.forEach(card => { if (!groups.has(card.GroupID)) add('CARD_WITHOUT_GROUP', card.CardID); });
  Object.keys(STUDENT_SOURCES).forEach(theme => {
    const source = STUDENT_SOURCES[theme];
    db[source].forEach(student => {
      const group = groups.get(student.GroupID);
      if (!group) add('STUDENT_WITHOUT_GROUP', source + ': ' + student.StudentID);
      else if (group.CardTheme !== theme) add('WRONG_STUDENT_SOURCE', source + ': ' + student.StudentID);
      if (!student.StudentName) add('EMPTY_NAME', source + ': ' + student.StudentID);
    });
  });
  db.Destinations.forEach(row => {
    if (!['Pink','Green','All'].includes(row.Availability)) add('INVALID_AVAILABILITY', row.DestinationID);
    if (!row.DestinationName) add('EMPTY_NAME', 'Destinations: ' + row.DestinationID);
  });
  const cards = new Set(), people = new Set(), activeGroups = new Set();
  db.Log.forEach(row => {
    if (!['ACTIVE','RETURNED'].includes(row.Status)) add('INVALID_LOG_STATUS', row.RecordID);
    if (!['FemaleStudents','MaleStudents'].includes(row.StudentSource)) add('INVALID_STUDENT_SOURCE', row.RecordID);
    if (!(row.ExitTime instanceof Date) || !Number.isFinite(row.ExitTime.getTime())) add('INVALID_EXIT_TIME', row.RecordID);
    if (row.Status === LOG_STATUS.ACTIVE) {
      if (cards.has(row.CardID) || people.has(personKey_(row)) || activeGroups.has(row.GroupID)) add('ACTIVE_CONFLICT', row.RecordID);
      cards.add(row.CardID); people.add(personKey_(row)); activeGroups.add(row.GroupID);
      if (row.ReturnTime !== '' || row.DurationMinutes !== '') add('INVALID_ACTIVE_RETURN', row.RecordID);
      if (!db.Cards.some(card => card.CardID === row.CardID && card.GroupID === row.GroupID) || !groups.has(row.GroupID)) add('ACTIVE_REFERENCE_MISSING', row.RecordID);
    } else if (row.Status === LOG_STATUS.RETURNED &&
      (!(row.ReturnTime instanceof Date) || !Number.isFinite(row.ReturnTime.getTime()) || row.DurationMinutes === '' || !Number.isFinite(Number(row.DurationMinutes)) || Number(row.DurationMinutes) < 0)) add('INVALID_RETURN', row.RecordID);
  });
  try { settings_(db); } catch (error) { add('INVALID_SETTINGS', error.message); }
  return issues;
}
function inspectSetup_(ss) {
  const db = {ss: ss}, issues = [];
  Object.keys(SCHEMA).forEach(name => {
    if (!ss.getSheetByName(name)) issues.push({code:'MISSING_SHEET', detail:name});
    else {
      try { db[name] = readTable_(ss, name); }
      catch (error) { issues.push({code:'INVALID_HEADERS', detail:name}); }
    }
  });
  if (!issues.length) issues.push.apply(issues, dataIssues_(db));
  return {ok: issues.length === 0, summary: issues.length ? 'توجد ملاحظات تحتاج تصحيحًا قبل الاستخدام' : 'الإعداد سليم', issues: issues};
}
function verifyProjectSetup() {
  // Administrative diagnostic only; never included in daily page responses.
  try {
    const report = inspectSetup_(spreadsheet_());
    console.log(JSON.stringify(report, null, 2));
    return report;
  } catch (error) {
    console.error(error.stack || error);
    return {ok:false, summary:'تعذر فحص الإعداد', issues:[{code:'SETUP_ERROR', detail:error.message}]};
  }
}
