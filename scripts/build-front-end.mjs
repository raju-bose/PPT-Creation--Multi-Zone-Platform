/**
 * Builds the multi zone front end from the original single file app.
 *
 * It downloads the original East Region app, pulls out the BHF141 slide
 * template and the folder rules, replaces the Google Sheets and Apps Script
 * plumbing with calls to the platform API, adds the zone switcher and writes
 * the result into public/ so Vercel can serve it.
 *
 * Output:
 *   public/platform.html              the app shell
 *   public/assets/app.js              patched application script
 *   public/assets/legacy.css          original stylesheet
 *   public/assets/bhf141.template.json slide template for the server
 *   public/assets/folders.json        sub folder rules per layout
 */
import fs from 'node:fs';
import path from 'node:path';

const SOURCE_REPO = 'raju-bose/PPT-Creation--East-Region-NRO_12345';
const MIRRORS = [
  'https://raw.githubusercontent.com/' + SOURCE_REPO + '/main/index.html',
  'https://cdn.jsdelivr.net/gh/' + SOURCE_REPO + '@main/index.html',
  'https://raw.githack.com/' + SOURCE_REPO + '/main/index.html',
];

const root = process.cwd();
const publicDir = path.join(root, 'public');
const assetsDir = path.join(publicDir, 'assets');
fs.mkdirSync(assetsDir, { recursive: true });

function log() {
  console.log('[front-end]', ...arguments);
}

function isWordChar(c) {
  if (!c) return false;
  const code = c.charCodeAt(0);
  return (
    (code >= 48 && code <= 57) ||
    (code >= 65 && code <= 90) ||
    (code >= 97 && code <= 122) ||
    c === '_' ||
    c === '$'
  );
}

function skipSpace(src, i) {
  while (i < src.length) {
    const ch = src[i];
    if (ch === ' ' || ch === '\n' || ch === '\r' || ch === '\t') i++;
    else break;
  }
  return i;
}

function findBlock(src, from, open, close) {
  let depth = 0;
  let inStr = false;
  let quote = '';
  let esc = false;
  for (let i = from; i < src.length; i++) {
    const ch = src[i];
    if (inStr) {
      if (esc) { esc = false; continue; }
      if (ch === '\\') { esc = true; continue; }
      if (ch === quote) inStr = false;
      continue;
    }
    if (ch === '"' || ch === "'" || ch === '`') { inStr = true; quote = ch; continue; }
    if (ch === open) depth++;
    else if (ch === close) {
      depth--;
      if (depth === 0) return i + 1;
    }
  }
  return -1;
}

function scanValue(src, from) {
  let i = from;
  let depth = 0;
  let inStr = false;
  let quote = '';
  let esc = false;
  while (i < src.length) {
    const ch = src[i];
    if (inStr) {
      if (esc) { esc = false; i++; continue; }
      if (ch === '\\') { esc = true; i++; continue; }
      if (ch === quote) inStr = false;
      i++;
      continue;
    }
    if (ch === '"' || ch === "'" || ch === '`') { inStr = true; quote = ch; i++; continue; }
    if (ch === '(' || ch === '[' || ch === '{') { depth++; i++; continue; }
    if (ch === ')' || ch === ']' || ch === '}') { depth--; i++; if (depth <= 0) return i; continue; }
    if (depth === 0 && (ch === ';' || ch === ',')) return i;
    if (depth === 0 && ch === '\n') {
      const next = skipSpace(src, i + 1);
      const head = src.slice(next, next + 10);
      const keepsGoing =
        head.indexOf('=') === 0 ||
        head.indexOf('+') === 0 ||
        head.indexOf('-') === 0 ||
        head.indexOf('*') === 0 ||
        head.indexOf('&') === 0 ||
        head.indexOf('|') === 0 ||
        head.indexOf('?') === 0 ||
        head.indexOf(':') === 0 ||
        head.indexOf('.') === 0 ||
        head.indexOf('const ') === 0 ||
        head.indexOf('let ') === 0 ||
        head.indexOf('var ') === 0 ||
        head.indexOf('function ') === 0 ||
        head.indexOf('return ') === 0;
      if (!keepsGoing) return i;
    }
    i++;
  }
  return i;
}

function findFunction(src, name) {
  const marker = 'function ' + name + '(';
  let at = src.indexOf(marker);
  while (at >= 0) {
    const prev = at > 0 ? src[at - 1] : ' ';
    const next = src[at + marker.length];
    if (!isWordChar(prev) && !isWordChar(next)) {
      const brace = src.indexOf('{', at);
      const end = findBlock(src, brace, '{', '}');
      if (end > 0) return { start: at, end: end, text: src.slice(at, end) };
    }
    at = src.indexOf(marker, at + 1);
  }
  return null;
}

function findConst(src, name) {
  const marker = name + '=';
  let at = src.indexOf(marker);
  while (at >= 0) {
    const prev = at > 0 ? src[at - 1] : ' ';
    const next = src[at + marker.length];
    if (!isWordChar(prev) && !isWordChar(next)) {
      const eq = at + marker.length;
      const end = scanValue(src, eq);
      const head = src.slice(Math.max(0, at - 12), at);
      let kwAt = -1;
      const keywords = ['const ', 'let ', 'var '];
      for (const kw of keywords) {
        const p = head.lastIndexOf(kw);
        if (p >= 0) {
          const abs = Math.max(0, at - 12) + p;
          if (abs > kwAt) kwAt = abs;
        }
      }
      return { start: kwAt >= 0 ? kwAt : at, eq: eq, end: end, text: src.slice(kwAt >= 0 ? kwAt : at, end) };
    }
    at = src.indexOf(marker, at + 1);
  }
  return null;
}

function replaceFunction(src, name, code) {
  const found = findFunction(src, name);
  if (!found) return null;
  return src.slice(0, found.start) + code + src.slice(found.end);
}

function replaceConst(src, name, code) {
  const found = findConst(src, name);
  if (!found) return null;
  return src.slice(0, found.start) + code + src.slice(found.end);
}

function replaceAllBlocks(src, openTag, closeTag, replacement) {
  let out = src;
  let at = out.indexOf(openTag);
  while (at >= 0) {
    const close = out.indexOf(closeTag, at);
    if (close < 0) break;
    const end = close + closeTag.length;
    out = out.slice(0, at) + replacement + out.slice(end);
    at = out.indexOf(openTag, at + replacement.length);
  }
  return out;
}

async function download() {
  const cacheDir = path.join(root, '.cache');
  fs.mkdirSync(cacheDir, { recursive: true });
  const cached = path.join(cacheDir, 'source-index.html');
  if (process.env.SOURCE_URL) {
    const res = await fetch(process.env.SOURCE_URL);
    if (!res.ok) throw new Error('SOURCE_URL returned ' + res.status);
    const text = await res.text();
    fs.writeFileSync(cached, text);
    return text;
  }
  if (fs.existsSync(cached)) return fs.readFileSync(cached, 'utf8');
  const problems = [];
  for (const url of MIRRORS) {
    for (let attempt = 0; attempt < 3; attempt++) {
      try {
        const res = await fetch(url, { headers: { 'user-agent': 'cai-platform-build' } });
        if (!res.ok) { problems.push(url + ' -> ' + res.status); break; }
        const text = await res.text();
        if (text.indexOf('TPL-JSON') < 0) { problems.push(url + ' -> no template'); break; }
        fs.writeFileSync(cached, text);
        return text;
      } catch (e) {
        problems.push(url + ' -> ' + e.message);
      }
    }
  }
  throw new Error('Could not download the source app: ' + problems.join(' | '));
}

const NEW_CALL = `async function call(action,payload){
  const body=Object.assign({},payload||{},{action:action,code:S.code||'',user:S.myName||''});
  if(action==='createSite'){ body.site=body.site||{}; if(!body.site.region){ body.site.region=(S.region==='all'?'east':S.region); } }
  let res;
  try{
    res=await fetch('/api/app',{method:'POST',credentials:'same-origin',headers:{'content-type':'application/json'},body:JSON.stringify(body)});
  }catch(e){ throw new Error('Cannot reach the platform server.'); }
  let j={};
  try{ j=await res.json(); }catch(e){ throw new Error('Unexpected server response.'); }
  if(j.ok===false){ throw new Error(j.error||('Request failed: '+action)); }
  return j;
}`;

const NEW_PHOTOURL = `function photoUrl(fileId,big){ return '/api/app?f='+encodeURIComponent(fileId); }`;

const NEW_FETCHBLOB = `async function fetchBlob(fileId){
  const r=await call('getPhoto',{fileId:fileId});
  const bin=atob(r.data||'');
  const arr=new Uint8Array(bin.length);
  for(let i=0;i<bin.length;i++){ arr[i]=bin.charCodeAt(i); }
  return new Blob([arr],{type:r.type||'image/jpeg'});
}`;

const NEW_SHOWLOGIN = `function showLogin(msg){
  const last=lsGet('sph-name')||'Admin';
  gate('<h2>Sign in</h2>'+
    '<p class="gate-sub">One platform for North Zone and East Zone site folders.</p>'+
    '<div class="gate-msg" id="g-msg" hidden></div>'+
    '<form id="g-form" novalidate>'+
      '<div class="fld"><label for="g-name">Login ID</label><input id="g-name" autocomplete="username" placeholder="Admin"></div>'+
      '<div class="fld"><label for="g-code">Password</label><input id="g-code" type="password" autocomplete="current-password" value="Admin@2026"></div>'+
      '<button type="submit" class="btn pri">Open platform</button>'+
    '</form>'+
    '<p class="gate-foot">Shared access for every user: ID <b>Admin</b>, password <b>Admin@2026</b>. Your ID is recorded against every photo you upload.</p>');
  const nameInput=document.getElementById('g-name');
  if(nameInput){ nameInput.value=String(last); }
  const box=document.getElementById('g-msg');
  if(box&&msg){ box.textContent=msg; box.hidden=false; }
  const form=document.getElementById('g-form');
  if(!form) return;
  form.onsubmit=async function(e){
    e.preventDefault();
    const btn=form.querySelector('button[type=submit]');
    const raw=String((document.getElementById('g-name')||{}).value||'');
    const id=raw.split(' ').join(' ').trim();
    const pw=String((document.getElementById('g-code')||{}).value||'');
    const msgBox=document.getElementById('g-msg');
    if(!id){ if(msgBox){ msgBox.textContent='Enter your login ID.'; msgBox.hidden=false; } return; }
    if(btn){ btn.disabled=true; btn.textContent='Signing in...'; }
    try{
      const res=await fetch('/api/app',{method:'POST',credentials:'same-origin',headers:{'content-type':'application/json'},body:JSON.stringify({action:'login',payload:{id:id,password:pw}})});
      const j=await res.json();
      if(!j.ok) throw new Error(j.error||'Sign in failed.');
      lsSet('sph-name',id);
      S.myName=id; S.code='';
      enterApp();
    }catch(err){
      if(btn){ btn.disabled=false; btn.textContent='Open platform'; }
      const b2=document.getElementById('g-msg');
      if(b2){ b2.textContent=(err&&err.message)||'Sign in failed.'; b2.hidden=false; }
    }
  };
}`;

const NEW_SHOWSETUP = `function showSetup(msg){
  gate('<h2>Platform not ready</h2><p class="gate-sub">'+(msg||'The platform server is not reachable yet.')+'</p><p class="gate-foot">Ask your administrator to check the deployment.</p>');
}`;

const NEW_MAPSITE = `const mapSite=s=>Object.assign({},s,{region:s.region||'east',subfolders:s.subfolders||[]});`;

const NEW_ISPPTNAME = `function isPptName(n){ return true; }`;

const NEW_BOOT = `(function boot(){
  S.url='/api/app';
  fetch('/api/app',{method:'POST',credentials:'same-origin',headers:{'content-type':'application/json'},body:JSON.stringify({action:'session'})})
    .then(function(r){ return r.json(); })
    .then(function(j){
      if(j&&j.ok&&j.user){ S.myName=j.user.id||(lsGet('sph-name')||'Admin'); enterApp(); }
      else { showLogin(); }
    })
    .catch(function(){ showLogin(); });
})();`;

const ZONE_CODE = `
/* ===== multi zone platform ===== */
(function(){
  var ZONE_LIST=[{id:'all',label:'All zones'},{id:'north',label:'North Zone'},{id:'east',label:'East Zone'}];
  if(!S.region) S.region=lsGet('sph-region')||'all';
  function zoneTick(){
    var bar=document.getElementById('zonebar');
    if(!bar) return;
    var pills=bar.querySelectorAll('[data-zone]');
    for(var i=0;i<pills.length;i++){
      if(pills[i].getAttribute('data-zone')===S.region) pills[i].classList.add('on');
      else pills[i].classList.remove('on');
    }
    var cnt=document.getElementById('zb-count');
    if(cnt){ var n=0; try{ n=sortedSites().length; }catch(e){} cnt.textContent=n+(n===1?' site':' sites'); }
  }
  function mountZoneBar(){
    var bar=document.getElementById('zonebar');
    if(!bar){
      bar=document.createElement('div');
      bar.id='zonebar';
      bar.className='zonebar';
      var html='<div class="zb-inner"><span class="zb-title">Zones</span><span class="zb-pills">';
      for(var i=0;i<ZONE_LIST.length;i++){
        html+='<button type="button" class="zb-pill" data-zone="'+ZONE_LIST[i].id+'">'+ZONE_LIST[i].label+'</button>';
      }
      html+='</span><span class="zb-count" id="zb-count"></span></div>';
      bar.innerHTML=html;
      document.body.insertBefore(bar,document.body.firstChild);
      bar.addEventListener('click',function(e){
        var t=e.target;
        while(t&&t!==bar&&!(t.getAttribute&&t.getAttribute('data-zone'))) t=t.parentNode;
        if(!t||!t.getAttribute) return;
        var z=t.getAttribute('data-zone');
        if(!z) return;
        S.region=z;
        try{ localStorage.setItem('sph-region',z); }catch(e){}
        zoneTick();
        try{ renderAll(); }catch(e){}
      });
    }
    zoneTick();
  }
  try{
    var __ra=renderAll;
    renderAll=function(){ var r=__ra.apply(this,arguments); zoneTick(); return r; };
  }catch(e){}
  try{
    var __ea=enterApp;
    enterApp=function(){ var r=__ea.apply(this,arguments); mountZoneBar(); return r; };
  }catch(e){}
  if(document.readyState==='loading') document.addEventListener('DOMContentLoaded',mountZoneBar);
  else mountZoneBar();
})();
`;

function extractDecl(src, name) {
  const fn = findFunction(src, name);
  if (fn) return fn.text;
  const c = findConst(src, name);
  if (c) return 'const ' + name + '=' + src.slice(c.eq, c.end) + ';';
  return null;
}

function buildRules(appJs, templateJson) {
  const regionStart = appJs.indexOf('const ORT={');
  const regionMarker = appJs.indexOf('const SITE_FOLDERS=');
  if (regionStart < 0 || regionMarker < 0) {
    log('folder rules block not found, folders.json skipped');
    return null;
  }
  const regionEnd = appJs.indexOf('\n', regionMarker);
  const region = appJs.slice(regionStart, regionEnd);
  let tpl = null;
  try {
    tpl = JSON.parse(templateJson);
  } catch (e) {
    log('template json did not parse');
    return null;
  }
  const exports = '\n;return {foldersFor:foldersFor,layoutOf:layoutOf};';
  const extra = [];
  let result = null;
  for (let attempt = 0; attempt < 25; attempt++) {
    try {
      const fn = new Function('TPLX', extra.join('\n') + '\n' + region + exports);
      result = fn(tpl);
      break;
    } catch (e) {
      const message = String(e && e.message ? e.message : '');
      const at = message.indexOf(' is not defined');
      if (at < 0) {
        log('folder rules failed: ' + message);
        return null;
      }
      const name = message.slice(0, at).split(' ').pop();
      const decl = extractDecl(appJs, name);
      if (!decl) {
        log('cannot resolve ' + name);
        return null;
      }
      extra.push(decl);
    }
  }
  if (!result) return null;
  const out = {};
  for (const layout of ['1DU2T', '2DU2T', '3DU3T']) {
    try {
      const lo = result.layoutOf({ layout: layout });
      out[layout] = {
        layout: layout,
        du: lo.du,
        tank: lo.tank,
        prods: lo.prods,
        folders: result.foldersFor(lo),
      };
    } catch (e) {
      log('layout ' + layout + ' failed: ' + e.message);
    }
  }
  return out;
}

async function main() {
  const source = await download();
  let html = source;

  const gsAt = html.indexOf('<script type="text/plain"');
  if (gsAt >= 0) {
    const gsEnd = html.indexOf('</script>', gsAt);
    if (gsEnd > 0) {
      html = html.slice(0, gsAt) + html.slice(gsEnd + 9);
      log('removed the old Apps Script block');
    }
  }

  const tplAt = html.indexOf('TPL-JSON');
  if (tplAt < 0) throw new Error('The source app does not contain the slide template');
  const scriptStart = html.lastIndexOf('<script', tplAt);
  const scriptOpen = html.indexOf('>', scriptStart) + 1;
  const scriptEnd = html.indexOf('</script>', tplAt);
  let appJs = html.slice(scriptOpen, scriptEnd);

  let css = '';
  let styleAt = html.indexOf('<style');
  while (styleAt >= 0) {
    const open = html.indexOf('>', styleAt) + 1;
    const close = html.indexOf('</style>', styleAt);
    if (close < 0) break;
    const block = html.slice(open, close);
    if (block.length > css.length) css = block;
    styleAt = html.indexOf('<style', close);
  }

  const tplMarker = appJs.indexOf('/*TPL-JSON*/');
  const tplStart = appJs.indexOf('{', tplMarker);
  const tplEnd = findBlock(appJs, tplStart, '{', '}');
  const templateJson = appJs.slice(tplStart, tplEnd);
  fs.writeFileSync(path.join(assetsDir, 'bhf141.template.json'), templateJson);
  log('slide template extracted, ' + templateJson.length + ' bytes');

  const rules = buildRules(appJs, templateJson);
  if (rules) {
    fs.writeFileSync(path.join(assetsDir, 'folders.json'), JSON.stringify(rules, null, 1));
    for (const key of Object.keys(rules)) {
      log('layout ' + key + ' has ' + rules[key].folders.length + ' sub folders');
    }
  }

  appJs = replaceFunction(appJs, 'call', NEW_CALL) || appJs;
  appJs = replaceFunction(appJs, 'fetchBlob', NEW_FETCHBLOB) || appJs;
  appJs = replaceFunction(appJs, 'photoUrl', NEW_PHOTOURL) || appJs;
  appJs = replaceFunction(appJs, 'isPptName', NEW_ISPPTNAME) || appJs;
  appJs = replaceFunction(appJs, 'showSetup', NEW_SHOWSETUP) || appJs;
  appJs = replaceConst(appJs, 'mapSite', NEW_MAPSITE) || appJs;

  if (appJs.indexOf('function gate(') >= 0) {
    appJs = replaceFunction(appJs, 'showLogin', NEW_SHOWLOGIN) || appJs;
  } else {
    log('sign in screen helper not found, the original sign in screen is kept');
  }

  const sortedFn = findFunction(appJs, 'sortedSites');
  if (sortedFn) {
    const renamed = sortedFn.text.replace('function sortedSites', 'function __sortedSitesOrig');
    const wrapper =
      `function sortedSites(){` +
      ` const zone=(S.region||'all');` +
      ` const keep=S.sites.filter(function(s){ return zone==='all'||!s.region||s.region===zone; });` +
      ` const all=S.sites; S.sites=keep;` +
      ` try{ return (` + renamed + `)(); } finally { S.sites=all; }` +
      `}`;
    appJs = replaceFunction(appJs, 'sortedSites', wrapper) || appJs;
  } else {
    const sortedConst = findConst(appJs, 'sortedSites');
    if (sortedConst) {
      const value = appJs.slice(sortedConst.eq, sortedConst.end);
      const wrapper =
        `const sortedSites=function(){` +
        ` const zone=(S.region||'all');` +
        ` const keep=S.sites.filter(function(s){ return zone==='all'||!s.region||s.region===zone; });` +
        ` const all=S.sites; S.sites=keep;` +
        ` try{ return (` + value + `)(); } finally { S.sites=all; }` +
        `};`;
      appJs = replaceConst(appJs, 'sortedSites', wrapper) || appJs;
    }
  }

  const bootAt = appJs.indexOf('function boot(');
  if (bootAt >= 0) {
    const brace = appJs.indexOf('{', bootAt);
    const end = findBlock(appJs, brace, '{', '}');
    const parenAt = appJs.lastIndexOf('(', bootAt);
    const between = appJs.slice(parenAt + 1, bootAt);
    if (between.indexOf('function boot') >= 0 && appJs.slice(end, end + 5) === '})();') {
      appJs = appJs.slice(0, parenAt) + NEW_BOOT + appJs.slice(end + 5);
      log('sign in flow replaced');
    } else {
      log('could not rewrite the start up block');
    }
  }

  const lastEnd = appJs.lastIndexOf('})();');
  if (lastEnd > 0) {
    appJs = appJs.slice(0, lastEnd) + ZONE_CODE + '\n' + appJs.slice(lastEnd);
    log('zone switcher added');
  } else {
    log('could not find the app wrapper, the zone switcher was not added');
  }

  appJs = appJs.split('East Region').join('North Zone + East Zone');

  fs.writeFileSync(path.join(assetsDir, 'app.js'), appJs);
  fs.writeFileSync(path.join(assetsDir, 'legacy.css'), css);

  html = html.slice(0, scriptStart) + '<script src="/assets/app.js"></script>' + html.slice(scriptEnd + 9);
  html = replaceAllBlocks(
    html,
    '<style',
    '</style>',
    '<link rel="stylesheet" href="/assets/legacy.css">\n<link rel="stylesheet" href="/assets/platform.css">',
  );

  const titleAt = html.indexOf('<title>');
  const titleEnd = html.indexOf('</title>', titleAt);
  if (titleAt >= 0 && titleEnd > 0) {
    html =
      html.slice(0, titleAt) +
      '<title>PPT Creation Photo Upload - Multi Zone Platform</title>' +
      html.slice(titleEnd + 8);
  }
  html = html.split('East Region').join('North Zone + East Zone');

  fs.writeFileSync(path.join(publicDir, 'platform.html'), html);
  log('front end written to public/platform.html');
}

main().catch((e) => {
  console.error('[front-end] build failed:', e && e.message ? e.message : e);
  process.exit(1);
});
