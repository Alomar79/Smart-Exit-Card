const SCHEMA = Object.freeze({
  FemaleStudents: ['StudentID', 'StudentName', 'GroupID', 'Status'],
  MaleStudents: ['StudentID', 'StudentName', 'GroupID', 'Status'],
  Groups: ['GroupID', 'GroupName', 'CardTheme', 'Status'],
  Cards: ['CardID', 'GroupID', 'Status'],
  Destinations: ['DestinationID', 'DestinationName', 'Availability', 'Status'],
  Log: ['RecordID', 'CardID', 'GroupID', 'GroupName', 'StudentSource', 'StudentID', 'StudentName', 'StaffName', 'DestinationID', 'DestinationName', 'ExitTime', 'ReturnTime', 'DurationMinutes', 'Status'],
  Settings: ['Key', 'Value']
});
const DEFAULT_SETTINGS = Object.freeze({SCHOOL_NAME: 'مدرسة السلم الخامسة', TIMEZONE: 'Asia/Qatar', GREEN_LIMIT: 10, ORANGE_LIMIT: 15, AUTO_REFRESH_SECONDS: 30});
const STUDENT_SOURCES = Object.freeze({Pink: 'FemaleStudents', Green: 'MaleStudents'});
const LOG_STATUS = Object.freeze({ACTIVE: 'ACTIVE', RETURNED: 'RETURNED'});
