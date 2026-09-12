function include_(name) { return HtmlService.createHtmlOutputFromFile(name).getContent(); }
function doGet(e) {
  try {
    const params = e && e.parameter || {};
    const view = ['card','supervisor','qr'].includes(params.view) ? params.view : (params.view ? 'invalid' : 'home');
    const template = HtmlService.createTemplateFromFile('Index');
    template.qrStyles = view==='qr'?include_('QrStyles'):'';
    template.qrAssets = view==='qr'?include_('QrLibrary')+include_('QrScripts'):'';
    template.boot = JSON.stringify({view: view, cardId: String(params.cardId || '')}).replace(/</g,'\\u003c');
    return template.evaluate().setTitle('بطاقة الخروج الذكية | مدرسة السلم الخامسة').addMetaTag('viewport','width=device-width, initial-scale=1');
  } catch (error) {
    console.error(error.stack || error);
    return HtmlService.createHtmlOutput('<html lang="ar" dir="rtl"><p>تعذر فتح الصفحة، حاول مرة أخرى</p></html>');
  }
}
function getPageData(request) {
  request=request||{};
  return api_(()=>timed_('load.'+(request.view||'home'),()=>{
    const data=requestData_();
    const result={settings:settings_({Settings:data.table('Settings')})};
    if(request.view==='card'){
      const current=liveIndex_(liveRows_(data.ss,id_(request.cardId))).cards.get(id_(request.cardId))||null;
      Object.assign(result,catalog_(data,request.cardId,!!current));
      result.current=current?logDto_(current,new Date()):null;
    }else if(request.view==='supervisor'){
      result.active=liveRows_(data.ss).sort((a,b)=>dateMs_(a.ExitTime)-dateMs_(b.ExitTime)).map(row=>logDto_(row,new Date()));
    }else if(request.view==='qr'){
      const groups=data.index('Groups','GroupID'), cards=data.index('Cards','CardID'), counts=new Map();
      cards.forEach(card=>counts.set(card.GroupID,(counts.get(card.GroupID)||0)+1));
      result.url=ScriptApp.getService().getUrl()||'';result.cards=[];
      cards.forEach(card=>{
        const group=groups.get(card.GroupID);
        if(!group||counts.get(card.GroupID)!==1)throw new Error('Invalid card group');
        if(active_(card)&&active_(group))result.cards.push({CardID:card.CardID,GroupID:group.GroupID,GroupName:group.GroupName,CardTheme:group.CardTheme});
      });
      result.cards.sort((a,b)=>a.GroupID.replace(/-[FM]$/,'').localeCompare(b.GroupID.replace(/-[FM]$/,''),'ar',{numeric:true})||(a.CardTheme===b.CardTheme?0:a.CardTheme==='Pink'?-1:1));
    }else if(request.view!=='home')fail_('الصفحة غير موجودة، استخدم رمز البطاقة الجديد');
    result.serverNow=Date.now();return result;
  }));
}
function getCardState(request) {
  return api_(()=>timed_('poll.card',()=>{
    const current=liveIndex_(liveRows_(spreadsheet_(),id_((request||{}).cardId))).cards.get(id_((request||{}).cardId))||null;
    const now=new Date();return {current:current?logDto_(current,now):null,serverNow:now.getTime()};
  }));
}
function getSupervisorState() {
  return api_(()=>timed_('poll.supervisor',()=>{
    const rows=liveRows_(spreadsheet_()),now=new Date();
    return {active:rows.sort((a,b)=>dateMs_(a.ExitTime)-dateMs_(b.ExitTime)).map(row=>logDto_(row,now)),serverNow:now.getTime()};
  }));
}
