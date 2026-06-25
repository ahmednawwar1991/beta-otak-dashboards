/* Cloud data loader — reads the Excel files LIVE from Google Drive in the browser,
   parses them (SheetJS + extract.js, so hidden sheets/rows stay excluded), builds DATA,
   boots the dashboard, then polls Drive for changes and refreshes in place.
   Configure window.APP_CONFIG = { apiKey, sources:[{id,type}], pollSeconds } in config.js.

   v2: caches the parsed data in localStorage so navigation between pages is INSTANT
   (boots from cache, then revalidates against Drive in the background). Also adds a
   hide/show toggle for the sidebar (remembered across pages). */
(function(){
  var CFG = window.APP_CONFIG || {sources:[], pollSeconds:10, apiKey:''};
  var sig = null;
  var CACHE_KEY = 'tut_dash_cache_v2';

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

  // ---- localStorage cache (parsed data) — makes navigation instant ----
  function saveCache(d, s){ try{ localStorage.setItem(CACHE_KEY, JSON.stringify({d:d, sig:s, ts:Date.now()})); }catch(e){} }
  function loadCache(){ try{ var r=localStorage.getItem(CACHE_KEY); return r?JSON.parse(r):null; }catch(e){ return null; } }

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
    var c = loadCache();
    if(c && c.d){
      // Instant: boot from cached parsed data — no spinner, no re-download/parse.
      apply(c.d); bootOrHook(); sig = c.sig || null;
      freshenThenPoll();           // quietly revalidate against Drive, then keep polling
    } else {
      // First ever load (or cache cleared): fetch with a spinner.
      overlay(spinner('Loading the latest data from Google Drive…'));
      try{
        var d=await fetchAll(); apply(d); removeOverlay(); bootOrHook();
        try{ sig=await signature(); }catch(e){}
        saveCache(d, sig);
        startPoll();
      }catch(err){
        var o=overlay(errorCard((err&&err.msg)||'Something went wrong.'));
        var btn=o.querySelector('#tut-retry'); if(btn) btn.onclick=init;
      }
    }
  }
  async function freshenThenPoll(){
    try{
      var old=sig, s=await signature();
      if(old===null || s!==old){
        var d=await fetchAll(); apply(d); sig=s; saveCache(d,s); rerender();
        if(old!==null) toast('Updated with the latest data');
      } else { sig=s; }
    }catch(e){}
    startPoll();
  }
  function startPoll(){
    var ms=Math.max(4,(CFG.pollSeconds||10))*1000;
    setInterval(async function(){
      try{
        var s=await signature();
        if(sig!==null && s!==sig){ sig=s; var d=await fetchAll(); apply(d); saveCache(d,s); rerender(); toast('Updated with the latest data'); }
        else { sig=s; }
      }catch(e){}
    }, ms);
  }
  if(document.readyState==='loading') document.addEventListener('DOMContentLoaded',init); else init();

  // ---- Sidebar hide/show toggle (desktop), remembered across pages ----
  (function navToggle(){
    var KEY='tut_nav_collapsed';
    try{ if(localStorage.getItem(KEY)==='1') document.documentElement.classList.add('nav-collapsed'); }catch(e){}
    var st=document.createElement('style');
    st.textContent=
      '@media(min-width:881px){'
      +'.shell{transition:grid-template-columns .26s ease}'
      +'.side{transition:transform .26s ease,opacity .2s ease}'
      +'html.nav-collapsed .shell{grid-template-columns:0 1fr}'
      +'html.nav-collapsed .side{transform:translateX(-110%);opacity:0;pointer-events:none}'
      +'}'
      +'#navToggle{display:inline-grid;place-items:center;width:38px;height:38px;border-radius:11px;'
      +'border:1px solid var(--line-3,#e6eaf3);background:var(--surface,#fff);color:var(--ink-2,#3F475E);'
      +'cursor:pointer;margin-right:10px;flex:0 0 auto;transition:background .15s}'
      +'#navToggle:hover{background:var(--surface-2,#f3f6fc)}'
      +'#navToggle svg{width:18px;height:18px}'
      +'@media(max-width:880px){#navToggle{display:none}}';
    (document.head||document.documentElement).appendChild(st);

    function mount(){
      if(!document.querySelector('.side')) return;     // pages without a sidebar (Executive Overview) get no toggle
      var bar=document.querySelector('.topbar'); if(!bar || document.getElementById('navToggle')) return;
      var btn=document.createElement('button');
      btn.id='navToggle'; btn.type='button'; btn.title='Hide / show menu'; btn.setAttribute('aria-label','Hide or show the menu');
      btn.innerHTML='<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="16" rx="2"/><line x1="9" y1="4" x2="9" y2="20"/></svg>';
      bar.insertBefore(btn, bar.firstChild);
      btn.addEventListener('click',function(){
        var c=document.documentElement.classList.toggle('nav-collapsed');
        try{ localStorage.setItem(KEY, c?'1':'0'); }catch(e){}
      });
    }
    if(document.readyState==='loading') document.addEventListener('DOMContentLoaded',mount); else mount();
  })();
})();
