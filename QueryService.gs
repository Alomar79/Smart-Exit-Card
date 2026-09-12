const CATALOG_TTL_SECONDS = 180;
function timed_(operation, callback) {
  const start = Date.now();
  try { return callback(); }
  finally { console.log(JSON.stringify({operation:operation,elapsedMs:Date.now()-start})); }
}
function clearDataCache() {
  // Namespace rotation also invalidates group keys without keeping a key registry.
  PropertiesService.getScriptProperties().setProperty('CATALOG_VERSION', Utilities.getUuid());
  console.log('تم مسح الذاكرة المؤقتة. حدّث الصفحات المفتوحة لجلب القوائم الجديدة.');
  return {ok:true};
}
function indexRows_(rows, key) {
  const map = new Map();
  rows.forEach(row => { if (!row[key] || map.has(row[key])) throw new Error('Duplicate or missing key: '+key); map.set(row[key],row); });
  return map;
}
function requestData_() {
  const ss=spreadsheet_(), tables=new Map(), indexes=new Map();
  const cache=CacheService.getScriptCache();
  const version=PropertiesService.getScriptProperties().getProperty('CATALOG_VERSION') || 'initial';
  function cached(key, read) {
    key='catalog-v2:'+ss.getId()+':'+version+':'+key;
    try { const value=cache.get(key); if(value!==null)return JSON.parse(value); } catch(error) { console.log('Cache read unavailable'); }
    const value=read(), encoded=JSON.stringify(value);
    if(encoded.length*3<95000)try{cache.put(key,encoded,CATALOG_TTL_SECONDS);}catch(error){console.log('Cache write unavailable');}
    return value;
  }
  function table(name) {
    if(!tables.has(name))tables.set(name,cached(name,()=>readTable_(ss,name)));
    return tables.get(name);
  }
  function index(name,key) {
    if(!indexes.has(name))indexes.set(name,indexRows_(table(name),key));
    return indexes.get(name);
  }
  function students(source,groupId) {
    return cached(source+':'+groupId,()=>{
      const rows=readTable_(ss,source);
      indexRows_(rows,'StudentID');
      const groups=new Map();
      rows.forEach(row=>{if(active_(row)){if(!groups.has(row.GroupID))groups.set(row.GroupID,[]);groups.get(row.GroupID).push({StudentID:row.StudentID,StudentName:row.StudentName});}});
      return groups.get(groupId)||[];
    });
  }
  return {ss:ss,table:table,index:index,students:students};
}
function catalog_(request, cardId, allowInactive) {
  const card=request.index('Cards','CardID').get(id_(cardId));
  if(!card)fail_('البطاقة غير موجودة أو غير مفعلة');
  const group=request.index('Groups','GroupID').get(card.GroupID);
  if(!group)fail_('البطاقة غير موجودة أو غير مفعلة');
  const groupCards=new Map();
  request.table('Cards').forEach(row=>groupCards.set(row.GroupID,(groupCards.get(row.GroupID)||0)+1));
  if(groupCards.get(group.GroupID)!==1)throw new Error('Multiple cards per group');
  if(!Object.prototype.hasOwnProperty.call(STUDENT_SOURCES,group.CardTheme))throw new Error('Invalid card theme');
  const canIssue=active_(card)&&active_(group);
  if(!canIssue&&!allowInactive)fail_('البطاقة غير موجودة أو غير مفعلة');
  const destinations=[];
  request.index('Destinations','DestinationID').forEach(row=>{if(allowedDestination_(row,group.CardTheme))destinations.push({DestinationID:row.DestinationID,DestinationName:row.DestinationName});});
  return {group:{GroupName:group.GroupName,CardTheme:group.CardTheme},canIssue:canIssue,
    students:request.students(STUDENT_SOURCES[group.CardTheme],group.GroupID),destinations:destinations};
}
function logRows_(ss, column, value, cardId) {
  const sheet=ss.getSheetByName('Log');
  if(!sheet)throw new Error('Missing Log');
  const last=sheet.getLastRow();
  // TextFinder executes inside Sheets. No historical row data is transferred to Apps Script.
  const matches=last>1?sheet.getRange(2,column,last-1,1).createTextFinder(value).matchEntireCell(true).matchCase(true).useRegularExpression(false).findAll():[];
  let rowNumbers=matches.map(range=>range.getRow());
  if(cardId && rowNumbers.length){
    const cardRows=new Set(sheet.getRange(2,2,last-1,1).createTextFinder(cardId).matchEntireCell(true).matchCase(true).useRegularExpression(false).findAll().map(range=>range.getRow()));
    rowNumbers=rowNumbers.filter(row=>cardRows.has(row));
  }
  const ranges=["'Log'!A1:N1"].concat(rowNumbers.map(row=>"'Log'!A"+row+':N'+row));
  const response=Sheets.Spreadsheets.Values.batchGet(ss.getId(),{ranges:ranges,valueRenderOption:'UNFORMATTED_VALUE',dateTimeRenderOption:'SERIAL_NUMBER'});
  const values=response.valueRanges || [];
  if(!validHeaders_(values[0]&&values[0].values||[[]],'Log'))throw new Error('Invalid Log headers');
  return rowNumbers.map((rowNumber,i)=>{
    const cells=values[i+1]&&values[i+1].values&&values[i+1].values[0]||[];
    const row={_row:rowNumber};
    SCHEMA.Log.forEach((key,j)=>{row[key]=cells[j]===undefined?'':cells[j];});
    ['ExitTime','ReturnTime'].forEach(key=>{
      // Sheets serial dates are local civil time. The spreadsheet timezone is Asia/Qatar (UTC+3).
      if(row[key]!==''){
        if(typeof row[key]!=='number')throw new Error('Invalid date serial');
        row[key]=new Date(Math.round((row[key]-25569)*86400000)-10800000);
      }
    });
    return row;
  }).filter(row=>String(row[SCHEMA.Log[column-1]])===value);
}
function liveRows_(ss,cardId) { return logRows_(ss,SCHEMA.Log.indexOf('Status')+1,LOG_STATUS.ACTIVE,cardId); }
function liveIndex_(rows) {
  const cards=new Map(),people=new Map(),groups=new Map();
  rows.forEach(row=>{
    if(cards.has(row.CardID)||people.has(personKey_(row))||groups.has(row.GroupID))throw new Error('Conflicting active exits');
    cards.set(row.CardID,row);people.set(personKey_(row),row);groups.set(row.GroupID,row);
  });
  return {cards:cards,people:people,groups:groups};
}
