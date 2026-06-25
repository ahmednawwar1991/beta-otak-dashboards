/* Cloud data loader — reads the Excel files LIVE from Google Drive in the browser,
   parses them (SheetJS + extract.js, so hidden sheets/rows stay excluded), builds DATA,
   boots the dashboard, then polls Drive for changes and refreshes in place.
   Configure window.APP_CONFIG = { apiKey, sources:[{id,type}], pollSeconds } in config.js. */
(function(){
  var CFG = window.APP_CONFIG || {sources:[], pollSeconds:10, apiKey:''};
  var sig = null;

  function mediaUrl(id){ return 'https://www.googleapis.com/drive/v3/files/'+id+'?alt=media&supportsAllDrives=true&key='+encodeURIComponent(CFG.apiKey); }
  function metaUrl(id){ return 'https://www.googleapis.com/drive/v3/files/'+id+'?fields=modifiedTime,md5Checksum,size&supportsAllDrives=true&key='+encodeURIComponent(CFG.apiKey); }

  function overlay(html){
    var o=document.getElementById('tut-load');
    if(!o){o=document.createElement('div');o.id='tut-load';
      o.style.cssText='position:fixed;inset:0;z-index:99999;display:flex;align-items:center;justify-content:center;background:#EAEEF6;font-family:system-ui,-apple-system,Segoe UI,sans-serif;color:#3F475E';
      document.body.appendChild(o);}
    o.innerHTML=html; return o;
  }
  function removeOverlay(){var o=document.getElementById('tut-load');if(o)o.remove();}
  function spinner(msg){
    return '<div style="display:flex;flex-direction:column;align-items:center;gap:16px">'
      +'<div style="width:42px;height:42px;border-radius:50%;border:3px solid #DCE2EE;border-top-color:#157A3C;animation:tutspin 1s linear infinite"></div>'
      +'<div style="font-size:13px;font-weight:600">'+msg+'</div>'
      +'<style>@keyframes tutspin{to{transform:rotate(360deg)}}</style></div>';
  }
  function errorCard(msg){
    return '<div style="max-width:560px;text-align:center;padding:30px;background:#fff;border:1px solid #E6EAF3;border-radius:20px;box-shadow:0 20px 44px -14px rgba(20,27,46,.22)">'
      +'<div style="font-size:17px;font-weight:800;color:#151A2D;margin-bottom:8px">Couldn’t load the live data</div>'
      +'<div style="font-size:13px;line-height:1.55;margin-bottom:18px;text-align:left">'+msg+'</div>'
      +'<button id="tut-retry" style="border:none;background:#157A3C;color:#fff;font:inherit;font-weight:700;padding:11px 20px;border-radius:11px;cursor:pointer">Try again</button></div>';
  }
  function hint(code){
    if(code===403) return 'Make sure the file is shared <b>“Anyone with the link → Viewer”</b> and that your API key has the <b>Google Drive API</b> enabled.';
    if(code===404) return 'A file ID in <code>config.js</code> wasn’t found — double-check the IDs.';
    if(code===400) return 'The API key looks wrong — re-check it in <code>config.js</code>.';
    return 'Check your internet connection and try again.';
  }

  function readBuf(buf){return XLSX.read(buf,{type:'array',cellDates:true});}
  async function fetchAll(){
    if(!CFG.apiKey) throw {code:400, msg:'No API key set in <code>config.js</code>.'};
    var items=[];
    for(var i=0;i<CFG.sources.length;i++){
      var s=CFG.sources[i];
      var res=await fetch(mediaUrl(s.id),{cache:'no-store'});
      if(!res.ok) throw {code:res.status, msg:'Drive read failed for <b>'+(s.type||s.id)+'</b> (HTTP '+res.status+').<br>'+hint(res.status)};
      items.push({wb:readBuf(await res.arrayBuffer()), type:s.type, name:s.id});
    }
    return XLSXExtract.fromTyped(items);
  }
  async function signature(){
    var parts=[];
    for(var i=0;i<CFG.sources.length;i++){
      try{ var r=await fetch(metaUrl(CFG.sources[i].id),{cache:'no-store'}); var j=await r.json();
        parts.push((j.md5Checksum||j.modifiedTime||j.size||'')+''); }
      catch(e){ parts.push('?'); }
    }
    return parts.join('~');
  }
  function apply(d){
    if(typeof DATA==='undefined'){window.DATA={};}
    DATA.contracts=d.contracts; DATA.orders=d.orders; DATA.msr=d.msr; DATA.equipment=d.equipment||[];
    DATA.payTotals=d.payTotals||{paid:0,due:0};
    DATA.generated=new Date().toLocaleString();
  }
  function bootOrHook(){
    if(typeof window.onAppData==='function') window.onAppData(DATA);
    else if(typeof boot==='function') boot();
  }
  function rerender(){
    if(typeof window.onAppData==='function') window.onAppData(DATA);
    else if(typeof window.liveRefresh==='function') window.liveRefresh();
    else if(typeof reboot==='function') reboot();
  }
  function toast(msg){
    var t=document.createElement('div');t.textContent=msg;
    t.style.cssText='position:fixed;bottom:22px;left:50%;transform:translateX(-50%);z-index:99999;background:#10152A;color:#fff;font:600 13px system-ui;padding:11px 18px;border-radius:12px;box-shadow:0 18px 40px -12px rgba(20,27,46,.4);opacity:0;transition:opacity .25s';
    document.body.appendChild(t);requestAnimationFrame(function(){t.style.opacity='1';});
    setTimeout(function(){t.style.opacity='0';setTimeout(function(){t.remove();},300);},2600);
  }

  async function init(){
    overlay(spinner('Loading the latest data from Google Drive…'));
    try{
      var d=await fetchAll(); apply(d); removeOverlay(); bootOrHook();
      try{sig=await signature();}catch(e){}
      startPoll();
    }catch(err){
      var o=overlay(errorCard((err&&err.msg)||'Something went wrong.'));
      var btn=o.querySelector('#tut-retry'); if(btn) btn.onclick=init;
    }
  }
  function startPoll(){
    var ms=Math.max(4,(CFG.pollSeconds||10))*1000;
    setInterval(async function(){
      try{
        var s=await signature();
        if(sig!==null && s!==sig){ sig=s; var d=await fetchAll(); apply(d); rerender(); toast('Updated with the latest data'); }
        else { sig=s; }
      }catch(e){}
    }, ms);
  }
  if(document.readyState==='loading') document.addEventListener('DOMContentLoaded',init); else init();
})();
