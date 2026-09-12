function archiveSheet_(ss, sheet, name, report) {
  const stamp = Utilities.formatDate(new Date(), 'Asia/Qatar', 'yyyyMMdd-HHmmss');
  const base = name + '_Backup_' + stamp;
  let backup = base, suffix = 1;
  while (ss.getSheetByName(backup)) backup = base + '_' + suffix++;
  // Rename the intact original rather than clearing or deleting any historical cells.
  sheet.setName(backup);
  SpreadsheetApp.flush();
  report.archived.push(backup);
  console.log('حُفظت الورقة الأصلية كاملة باسم: ' + backup);
}
function setupProject() {
  return locked_(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error('افتح Apps Script من داخل Google Sheet ثم شغّل setupProject.');
    const report = {created:[], archived:[], seeded:false, addedSettings:[]};
    // Preflight before changing anything: don't strand legacy active exits or overwrite unknown master schemas.
    Object.keys(SCHEMA).forEach(name => {
      const sheet = ss.getSheetByName(name);
      if (!sheet || sheet.getLastRow() === 0) return;
      const values = sheet.getDataRange().getValues();
      if (validHeaders_(values, name)) return;
      if (!['Log','Destinations'].includes(name)) throw new Error('أعمدة غير مطابقة في ' + name + '. لم يتم تغيير البيانات.');
      if (name === 'Log') {
        const statusColumn = values[0].indexOf('Status');
        if (statusColumn >= 0 && values.slice(1).some(row => String(row[statusColumn]).trim().toUpperCase() === 'ACTIVE')) {
          throw new Error('يوجد خروج فعال في Log القديم. سجّل العودة بالنظام القديم أولًا ثم أعد الإعداد. لم يتم ترحيل البيانات.');
        }
      }
    });
    PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());
    Object.keys(SCHEMA).forEach(name => {
      let sheet = ss.getSheetByName(name);
      if (sheet && sheet.getLastRow() > 0 && !validHeaders_(sheet.getDataRange().getValues(), name)) {
        archiveSheet_(ss, sheet, name, report);
        sheet = null;
      }
      if (!sheet) { sheet = ss.insertSheet(name); report.created.push(name); }
      if (sheet.getLastRow() === 0) {
        sheet.getRange(1,1,1,SCHEMA[name].length).setValues([SCHEMA[name]]);
        if (!['Log','Settings'].includes(name)) sheet.getRange(2,1,sheet.getMaxRows()-1,SCHEMA[name].length).setNumberFormat('@');
      }
      sheet.setFrozenRows(1);
      sheet.getRange(1,1,1,SCHEMA[name].length).setFontWeight('bold').setBackground('#e4eee8');
    });
    const props = PropertiesService.getScriptProperties();
    const names = ['FemaleStudents','MaleStudents','Groups','Cards','Destinations'];
    if (names.concat(['Log']).every(name => ss.getSheetByName(name).getLastRow() <= 1) && props.getProperty('GROUPS_SEEDED_SPREADSHEET_ID') !== ss.getId()) {
      const seed = {
        Groups: [['C201-F','طالبات ثاني1','Pink','Active'],['C201-M','طلاب ثاني1','Green','Active'],['C202-F','طالبات ثاني2','Pink','Active'],['C202-M','طلاب ثاني2','Green','Active']],
        Cards: [['CARD-C201-F','C201-F','Active'],['CARD-C201-M','C201-M','Active'],['CARD-C202-F','C202-F','Active'],['CARD-C202-M','C202-M','Active']],
        FemaleStudents: [['FS001','مريم أحمد','C201-F','Active'],['FS002','سارة خالد','C201-F','Active'],['FS003','نور علي','C202-F','Active'],['FS004','ريم سالم','C202-F','Active']],
        MaleStudents: [['MS001','محمد أحمد','C201-M','Active'],['MS002','أحمد سالم','C201-M','Active'],['MS003','خالد علي','C202-M','Active'],['MS004','عمر محمد','C202-M','Active']],
        Destinations: [['D001','حمام الطالبات','Pink','Active'],['D002','حمام الطلاب','Green','Active'],['D003','العيادة','All','Active'],['D004','الإدارة','All','Active']]
      };
      names.forEach(name => ss.getSheetByName(name).getRange(2,1,seed[name].length,SCHEMA[name].length).setValues(seed[name]));
      props.setProperty('GROUPS_SEEDED_SPREADSHEET_ID', ss.getId());
      report.seeded = true;
    }
    const existing = readTable_(ss,'Settings');
    const missing = Object.keys(DEFAULT_SETTINGS).filter(key => !existing.some(row => row.Key === key)).map(key => [key, DEFAULT_SETTINGS[key]]);
    if (missing.length) ss.getSheetByName('Settings').getRange(ss.getSheetByName('Settings').getLastRow()+1,1,missing.length,2).setValues(missing);
    report.addedSettings = missing.map(row => row[0]);
    ss.setSpreadsheetTimeZone('Asia/Qatar');
    const log = ss.getSheetByName('Log');
    log.getRange(2,11,log.getMaxRows()-1,2).setNumberFormat('dd/MM/yyyy hh:mm:ss');
    log.getRange(2,13,log.getMaxRows()-1,1).setNumberFormat('0.00');
    clearDataCache();
    report.verification = inspectSetup_(ss);
    console.log(JSON.stringify(report,null,2));
    return report;
  });
}
