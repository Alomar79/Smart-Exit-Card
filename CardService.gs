function createExit(input) {
  return api_(() => timed_('write.exit', () => locked_(() => {
    input = input || {};
    const ss=spreadsheet_(),db={ss:ss};
    Object.keys(SCHEMA).filter(name=>name!=='Log').forEach(name=>{db[name]=readTable_(ss,name);});
    db.Log=liveRows_(ss);
    validateData_(db); settings_(db);
    const context = cardContext_(db, input.cardId, true);
    const staff = staffName_(input.staffName);
    if (!input.studentId) fail_('يرجى اختيار الاسم');
    if (!input.destinationId) fail_('يرجى اختيار المكان');
    const student = one_(db[context.source], 'StudentID', id_(input.studentId), 'تعذر التحقق من الاسم المختار');
    if (!active_(student) || student.GroupID !== context.group.GroupID) fail_('الاسم غير متاح لهذه البطاقة');
    const destination = one_(db.Destinations, 'DestinationID', id_(input.destinationId), 'المكان غير متاح');
    if (!allowedDestination_(destination, context.group.CardTheme)) fail_('المكان غير متاح لهذه البطاقة');
    const current = db.Log.filter(row => row.Status === LOG_STATUS.ACTIVE);
    if (current.some(row => row.CardID === context.card.CardID || row.GroupID === context.group.GroupID)) fail_('هذه البطاقة مستخدمة حاليًا');
    if (current.some(row => row.StudentSource === context.source && row.StudentID === student.StudentID)) fail_('هذا الاسم مسجل في الخارج حاليًا');
    const record = {
      RecordID: Utilities.getUuid(), CardID: context.card.CardID, GroupID: context.group.GroupID, GroupName: context.group.GroupName,
      StudentSource: context.source, StudentID: student.StudentID, StudentName: student.StudentName, StaffName: staff,
      DestinationID: destination.DestinationID, DestinationName: destination.DestinationName,
      ExitTime: new Date(), ReturnTime: '', DurationMinutes: '', Status: LOG_STATUS.ACTIVE
    };
    appendLog_(db, record);
    return {message:'تم تسجيل الخروج',record:logDto_(record,new Date()),serverNow:Date.now()};
  })));
}
function returnStudent(input) {
  return api_(() => timed_('write.return', () => locked_(() => {
    input = input || {};
    const ss = spreadsheet_();
    const record = one_(logRows_(ss,1,id_(input.recordId)), 'RecordID', input.recordId);
    // The supervisor sends only RecordID; card requests can additionally verify association.
    if (input.cardId !== undefined && record.CardID !== id_(input.cardId)) fail_('الحالة لا تتبع هذه البطاقة');
    if (record.Status !== LOG_STATUS.ACTIVE) fail_('تم تسجيل العودة مسبقًا');
    const now = new Date();
    const elapsed = (now.getTime() - dateMs_(record.ExitTime)) / 60000;
    if (elapsed < 0) throw new Error('ExitTime is in the future');
    ss.getSheetByName('Log').getRange(record._row, SCHEMA.Log.indexOf('ReturnTime') + 1, 1, 3)
      .setValues([[now, Math.round(elapsed * 100) / 100, LOG_STATUS.RETURNED]]);
    record.ReturnTime=now;record.DurationMinutes=Math.round(elapsed*100)/100;record.Status=LOG_STATUS.RETURNED;
    return {message:'تم تسجيل العودة',record:logDto_(record,now),serverNow:now.getTime()};
  })));
}
