/* ===== auto fill =====
   1. Table values of the PPT (make, model, serial number, IP) are read from the site's serial-number Excel
      (folder "14.1 Device Serial Numbers": sheets "Device Sr No Sheet" and "CCTV SR NO Sheet") and written into
      every table page of the template. A value that is not in the Excel is shown as "-", never as the
      Bodhgaya template value.
   2. Page 2 (Document Owners / Change History) carries the login name of the person who generated the PPT,
      the first DPR date as the created date and the day of the last generation as the updated date.
   3. Access history and "change my password" for the PPT team.
   This file is appended to public/assets/app.js by scripts/auto-fill.mjs. */
const AF=(function(){
"use strict";
const store={};                                   // siteId → {key, data, loading}
const clean=v=>String(v==null?"":v).replace(/\s+/g," ").trim().replace(/^[,'`]+\s*/,"");
const isBlank=v=>!v || /^(-+|na|n\/a|nil|none|not applicable)$/i.test(v);
const val=v=>{ v=clean(v); return isBlank(v)? "" : v; };
const nk=s=>String(s||"").toLowerCase().replace(/[^a-z0-9]+/g,"");
const IP=/\b(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})\b/;
const clone=o=>JSON.parse(JSON.stringify(o));

/* ---------- read the Excel ---------- */
function parseBook(X,wb){
  const D={devices:[],cams:[],gateway:"",nvrIp:""};
  wb.SheetNames.forEach(sn=>{
    const rows=X.utils.sheet_to_json(wb.Sheets[sn],{header:1,raw:false,defval:""}).map(r=>r.map(clean));
    let hi=-1, kind="";
    for(let i=0;i<Math.min(rows.length,20);i++){ const c=rows[i].map(x=>x.toLowerCase());
      if(c.some(x=>/camera/.test(x)) && c.some(x=>/sr\.?\s*no|serial/.test(x))){ hi=i; kind="cam"; break; }
      if(c.some(x=>/^material|device name|name of device/.test(x)) && c.some(x=>/serial/.test(x))){ hi=i; kind="dev"; break; }
    }
    if(hi<0) return;
    // values written above the table: "Gateway:-10.230.7.65", "NVR IP :-10.249.35.195"
    rows.slice(0,hi).forEach(r=>r.forEach((c,ci)=>{
      const next=()=>{ for(let k=ci+1;k<r.length;k++){ const m=r[k].match(IP); if(m) return m[1]; if(r[k]) return ""; } return ""; };
      if(/^nvr\s*ip/i.test(c)){ const m=c.match(IP); D.nvrIp=(m?m[1]:next())||D.nvrIp; }
      else if(/^gateway/i.test(c) && kind==="dev"){ const m=c.match(IP); D.gateway=(m?m[1]:next())||D.gateway; }
    }));
    const H=rows[hi].map(x=>x.toLowerCase());
    const col=(...res)=>{ for(const re of res){ const i=H.findIndex(x=>re.test(x)); if(i>=0) return i; } return -1; };
    const get=(r,i)=>i<0? "" : val(r[i]);
    if(kind==="dev"){
      const cm=col(/^material/,/device name|name of device/), ck=col(/^make/), cd=col(/^model/), cs=col(/serial/), ci=col(/ip\s*add/), ct=col(/install/);
      rows.slice(hi+1).forEach(r=>{ const name=get(r,cm); if(!name) return;
        D.devices.push({name,key:nk(name),make:get(r,ck),model:get(r,cd),serial:get(r,cs),ip:(get(r,ci).match(IP)||[])[1]||"",date:get(r,ct)}); });
    } else {
      const cn=col(/camera\s*name/,/^location/), ck=col(/^make/), cd=col(/^model/), cs=col(/camera\s*sr/,/serial/), cu=col(/user/), cp=col(/pass/),
        cg=col(/gateway/), ci=col(/camera\s*ip/,/ip\s*add/,/^ip$/), cb=col(/subnet/), cl=col(/^location/), cno=col(/^sr\.?\s*no/,/^s\.?\s*no/);
      const list=[];
      rows.slice(hi+1).forEach(r=>{ const c={name:get(r,cn),make:get(r,ck),model:get(r,cd),serial:get(r,cs),user:get(r,cu),pass:get(r,cp),
          gw:(get(r,cg).match(IP)||[])[1]||"",ip:(get(r,ci).match(IP)||[])[1]||"",subnet:(get(r,cb).match(IP)||[])[1]||"",loc:get(r,cl)};
        if(!c.name && !c.serial && !c.ip) return;
        const no=parseInt(get(r,cno),10); c.no=no>0? no : list.length+1; list.push(c); });
      if(list.length) D.cams=list;
    }
  });
  return D;
}
function merge(a,b){   // later file wins
  if(!a) return b;
  b.devices.forEach(d=>{ const i=a.devices.findIndex(x=>x.key===d.key); if(i>=0) a.devices[i]=d; else a.devices.push(d); });
  if(b.cams.length) a.cams=b.cams; a.gateway=b.gateway||a.gateway; a.nvrIp=b.nvrIp||a.nvrIp;
  return a;
}
function excelFiles(site){
  const ok=new Set([SERIAL_F].concat(TABLE_NAMES));
  const byId={}; (site.subfolders||[]).forEach(f=>byId[f.id]=tplName(f.name));
  return photosOf(site.id).filter(p=>isExcel(p.fileName) && ok.has(byId[p.folderId])).sort((a,b)=>(a.uploadedAt||0)-(b.uploadedAt||0));
}
const fileKey=site=>excelFiles(site).map(p=>p.id+":"+(p.uploadedAt||0)).join("|");
async function load(site){
  const key=fileKey(site); const cur=store[site.id];
  if(cur && cur.key===key && (cur.data!==undefined || cur.loading)) return cur.loading||cur.data;
  const st=store[site.id]={key,data:undefined,files:[]};
  if(!key){ st.data=null; return null; }
  st.loading=(async()=>{
    let D=null; const names=[];
    try{
      await loadScript(LIB.xlsx); const X=window.XLSX;
      for(const p of excelFiles(site)){
        try{ const b=await fetchBlob(p.assetId); const d=parseBook(X,X.read(new Uint8Array(await b.arrayBuffer()),{type:"array"}));
          if(d.devices.length||d.cams.length){ D=merge(D,d); names.push(p.fileName); } }catch(e){ st.err=errText(e); }
      }
    }catch(e){ st.err=errText(e); }
    st.data=D; st.files=names; st.loading=null; return D;
  })();
  return st.loading;
}

/* ---------- look up a device ---------- */
function finder(D){
  const devs=(D&&D.devices)||[];
  const dev=re=>devs.find(x=>re.test(x.key))||null;
  const ONE={router:/^(cisco)?router/, poe:/^(cisco)?poe/, wifi:/wifi|accesspoint/, nvr:/^nvr/, atg:/^atg/, fcc:/^fcc/, ipc:/^ipc/, ams:/^ams/, mms:/^mms/,
    desktop:/^(desktop|bospc|bocpc)/, monitor:/^monitor/, padlock:/^padlock/, padkey:/^padkey/, hooter:/^hooter/, ack:/^ack/,
    epsBase:/^eps.*(base|controller)/, epsP1:/^eps.*p1/, epsD1:/^eps.*d1/, epsP2:/^eps.*p2/, epsD2:/^eps.*d2/};
  const NUM={hdd:"^hdd0?N$", opt:"^opt0?N$", l2:"^(cisco)?l2(switch)?0?N$", fpl:"^fillpipelock0?N$", rpl:"^riserpipelock0?N$", mdl:"^(manual)?dip(pipe)?lock(ingassembly)?0?N$",
    tsen:"^tank(sump)?(leak)?sensor0?N$", dsen:"^(du|dispenser)(pan|sump)*(leak)?sensor0?N$", isen:"^interstitial(leak)?(sensor)?0?N$"};
  const F={
    one:k=>dev(ONE[k]),
    num:(k,n)=>dev(new RegExp(NUM[k].replace("N",String(n)))) || (n===1 && k==="hdd"? dev(/^hdd$/) : null),
    list:k=>[1,2,3,4,5,6].map(n=>F.num(k,n)).filter(Boolean),
    cams:(D&&D.cams)||[],
    cam:n=>((D&&D.cams)||[]).find(c=>c.no===n)||null
  };
  return F;
}

/* ---------- text + grid helpers ---------- */
function estW(t,s){ let w=0; for(const ch of String(t)){
    w+= /[0-9]/.test(ch)?0.507 : /[A-Z]/.test(ch)?(/[IJ]/.test(ch)?0.3:/[MW]/.test(ch)?0.86:0.6) : /[a-z]/.test(ch)?(/[ilj]/.test(ch)?0.25:/[mw]/.test(ch)?0.78:0.49) : /\s/.test(ch)?0.23 : 0.33; }
  return w*s/72; }
const isT=e=>e.k==="t";
const cx=e=>e.al==="c"? e.x+e.w/2 : e.x+(e.tw||e.w)/2;
const cy=e=>e.y+e.h/2;
const grid=pg=>(pg.borders||[]).concat(pg.els.filter(e=>e.k==="r"&&e.f==="000000"));
const uniq=a=>{ a=a.slice().sort((x,y)=>x-y); return a.filter((v,i)=>!i||v-a[i-1]>0.03); };
const vAt=(pg,y)=>uniq(grid(pg).filter(b=>b.w<0.03&&b.h>0.2&&b.y<=y+0.01&&b.y+b.h>=y-0.01).map(b=>b.x));
const hAt=(pg,x)=>uniq(grid(pg).filter(b=>b.h<0.03&&b.w>0.3&&b.x<=x+0.01&&b.x+b.w>=x-0.01).map(b=>b.y));
const colOf=(V,x)=>{ let a=null,b=null; V.forEach(v=>{ if(v<=x) a=v; else if(b===null) b=v; }); return a===null||b===null? null : [a,b]; };
const rowsBelow=(H,y)=>{ const out=[]; for(let i=0;i<H.length-1;i++){ if(H[i]>y) out.push([H[i],H[i+1]]); } return out; };
const find=(pg,re,ok)=>pg.els.find(e=>isT(e)&&re.test(e.t.trim())&&(!ok||ok(e)));
function take(pg,pred){ const out=[]; pg.els=pg.els.filter(e=>{ if(isT(e)&&pred(e)){ out.push(e); return false; } return true; }); return out; }
const inBand=(e,y0,y1)=>cy(e)>y0&&cy(e)<y1;
const textIn=(pg,c,y0,y1,start)=>pg.els.filter(e=>isT(e)&&inBand(e,y0,y1)&&(start? (e.x>=c[0]-0.025&&e.x<c[1]-0.025) : (cx(e)>=c[0]&&cx(e)<c[1])))
  .sort((a,b)=>Math.abs(a.y-b.y)<0.03? a.x-b.x : a.y-b.y).map(e=>e.t.trim()).join(" ");
/* write the lines of one table cell (old text of the cell is removed) */
function setCell(pg,c,y0,y1,arr,o){
  o=o||{};
  const old=take(pg,e=>inBand(e,y0,y1)&&(o.start? (e.x>=c[0]-0.025&&e.x<(o.x1||c[1])-0.025) : (cx(e)>=c[0]&&cx(e)<c[1])));
  const ref=old[0]||o.ref||{h:0.1533,s:11,c:"000000"};
  arr=(arr||[]).map(x=>String(x==null?"":x)).filter((x,i,a)=>x||a.length>1).map(x=>x||"-"); if(!arr.length) arr=["-"];
  let s=o.s||ref.s, pitch=s/59;
  const W=c[1]-c[0]-0.08, Hh=y1-y0-0.03;
  if(arr.length*pitch>Hh){ s=Math.max(5,Math.floor(s*Hh/(arr.length*pitch)*2)/2); pitch=Hh/arr.length; }
  const mid=(y0+y1)/2;
  const widest=Math.max(...arr.map(t=>estW(t,s))); if(widest>W) s=Math.max(5,Math.floor(s*W/widest*2)/2);   // one size for the whole cell
  arr.forEach((t,i)=>{ const si=s;
    const y=mid-ref.h/2+(i-(arr.length-1)/2)*pitch;
    if(o.left) pg.els.push({k:"t",t,x:o.left,y,w:Math.min(W,estW(t,si)),h:ref.h,s:si,c:ref.c,tw:Math.min(W,estW(t,si)),r:ref.r});
    else pg.els.push({k:"t",t,x:c[0]+0.04,y,w:c[1]-c[0]-0.08,h:ref.h,s:si,c:ref.c,tw:Math.min(W,estW(t,si)),al:"c",r:ref.r}); });
  return ref;
}
/* long text → up to n lines that fit the column */
function wrapWords(t,W,s,n){
  if(!t) return ["-"]; if(estW(t,s)<=W || n<2) return [t];
  const words=t.split(/\s+/), out=[]; let cur="";
  words.forEach(w=>{ const nx=(cur+" "+w).trim(); if(cur && estW(nx,s)>W){ out.push(cur); cur=w; } else cur=nx; }); if(cur) out.push(cur);
  return out.length<=n? out : [t];
}
/* one long value (serial / model) → broken after "/" or "-" when the row has room for more lines */
function wrapTok(t,W,s,n){
  if(!t) return ["-"]; if(estW(t,s)<=W || n<2) return [t];
  const tok=t.split(/(?<=[\/\- ])/), out=[]; let cur="";
  tok.forEach(k=>{ if(cur && estW((cur+k).trim(),s)>W){ out.push(cur.trim()); cur=k; } else cur+=k; }); if(cur.trim()) out.push(cur.trim());
  return out.length<=n && out.every(x=>estW(x,s)<=W*1.05)? out : [t];
}
function wrapChars(t,W,s){
  if(!t) return ["-"]; const out=[]; let cur="";
  for(const ch of t){
    if(cur && estW(cur+ch,s)>W){   // break after a "-" or "/" when one is close to the end of the line
      const k=Math.max(cur.lastIndexOf("-"),cur.lastIndexOf("/"));
      if(k>=cur.length-4 && k>0 && k<cur.length-1){ out.push(cur.slice(0,k+1)); cur=cur.slice(k+1)+ch; }
      else { out.push(cur); cur=ch; }
    } else cur+=ch; }
  if(cur) out.push(cur);
  const n=out.length; if(n>1 && out[n-1].length===1){ out[n-1]=out[n-2].slice(-1)+out[n-1]; out[n-2]=out[n-2].slice(0,-1); }
  return out;
}
function wrapIp(t,W,s){
  if(!t) return ["-"]; if(estW(t,s)<=W) return [t];
  const tok=t.split(".").map((x,i)=>i? "."+x : x), out=[]; let cur="";
  tok.forEach(k=>{ if(cur && estW(cur+k,s)>W){ out.push(cur); cur=k; } else cur+=k; }); if(cur) out.push(cur);
  return out;
}
const lastOct=ip=>+String(ip).split(".")[3];
const samePrefix=(a,b)=>a.split(".").slice(0,3).join(".")===b.split(".").slice(0,3).join(".");
function gwOff(ip,gw,sp){ if(!ip||!gw||!samePrefix(ip,gw)) return ""; const d=lastOct(ip)-lastOct(gw); const s=sp?" ":""; return d===0? "GW" : d>0? "GW"+s+"+"+s+d : "GW"+s+"-"+s+(-d); }
function ipRange(gw,mask){ if(!gw||!mask) return null; const a=gw.split(".").map(Number), m=mask.split(".").map(Number);
  if(a.length!==4||m.length!==4) return null; return [a.map((v,i)=>v&m[i]).join("."), a.map((v,i)=>(v&m[i])|(~m[i]&255)).join(".")]; }
const mode=a=>{ const n={}; let best=""; a.filter(Boolean).forEach(v=>{ n[v]=(n[v]||0)+1; if(!best||n[v]>n[best]) best=v; }); return best; };

/* ---------- "Name of Device | Make | Model No. | Serial No." tables ---------- */
function ruleFor(key,F,st){
  const one=d=>d? {make:d.make,model:[d.model],serial:[d.serial]} : {make:"",model:[],serial:[]};
  const many=ds=>{ ds=ds.filter(Boolean); const use=ds.filter(d=>d.serial||d.model); if(!use.length) return {make:(ds.find(d=>d.make)||{}).make||"",model:[],serial:[]};
    const models=use.map(d=>d.model||"-"), u=[...new Set(models)];
    return {make:(ds.find(d=>d.make)||{}).make||"", model:u.length===1? u : models, serial:use.map(d=>d.serial||"-")}; };
  let m;
  if(/^(atg|fcc|ipc|nvr|ams|mms|monitor)$/.test(key)) return one(F.one(key));
  if(/^hooter/.test(key)) return many([F.one("hooter"),F.one("ack")]);
  if(/poe/.test(key)) return one(F.one("poe"));
  if(/router/.test(key)) return one(F.one("router"));
  if((m=key.match(/l2(?:switch)?(\d)$/))) return one(F.num("l2",+m[1]));
  if(key==="hdd"){ st.hdd=(st.hdd||0)+1; return one(F.num("hdd",st.hdd)); }
  if(/^bo[cs]pc$|^desktop/.test(key)) return one(F.one("desktop"));
  if((m=key.match(/^opt(\d)$/))) return one(F.num("opt",+m[1]));
  if(/^fillpipelock/.test(key)) return many(F.list("fpl"));
  if(key==="padkey"||key==="padlock"){ const a=F.one("padkey"), b=F.one("padlock"); const d=(a&&a.serial)? a : (b&&b.serial)? b : (a||b);
    return d? {make:d.make||(a&&a.make)||(b&&b.make)||"",model:[d.model],serial:[d.serial]} : one(null); }
  if(/^riserpipelock/.test(key)) return many(F.list("rpl"));
  if(/^dippipe|^manualdip/.test(key)) return many(F.list("mdl"));
  if(/^tanksump/.test(key)) return many(F.list("tsen"));
  if(/^dispenser/.test(key)) return many(F.list("dsen"));
  if(/^interstitial/.test(key)) return many(F.list("isen"));
  return null;
}
function tableFill(pg,F){
  const hdr=find(pg,/^name of device$/i); if(!hdr) return;
  const V=vAt(pg,hdr.y+0.05);
  const hc=re=>{ const h=find(pg,re,e=>Math.abs(e.y-hdr.y)<0.12); return h? colOf(V,cx(h)) : null; };
  const cN=colOf(V,cx(hdr)), cK=hc(/^make$/i), cM=hc(/^model no\.?$/i), cS=hc(/^serial no\.?$/i);
  if(!cN||!cK||!cM||!cS) return;
  const st={};
  rowsBelow(hAt(pg,(cN[0]+cN[1])/2),hdr.y).forEach(([a,b])=>{
    const key=nk(textIn(pg,cN,a,b)); if(!key) return;
    if(key==="eps"){   // controller in the first part of the row, the four displays in the second
      const base=F.one("epsBase"), disp=["epsP1","epsD1","epsP2","epsD2"].map(k=>F.one(k)).filter(Boolean);
      const all=[base].concat(disp).filter(Boolean);
      const mk=(all.find(d=>d.make)||{}).make||"";
      setCell(pg,cK,a,b,wrapWords(mk,cK[1]-cK[0]-0.1,11,3));
      const inner=hAt(pg,(cS[0]+cS[1])/2).filter(h=>h>a+0.05&&h<b-0.05);
      const parts=inner.length? [[a,inner[0],[base]],[inner[0],b,disp]] : [[a,b,all]];
      parts.forEach(([y0,y1,ds])=>{ ds=ds.filter(d=>d&&(d.model||d.serial));
        setCell(pg,cM,y0,y1,ds.map(d=>d.model||"-")); setCell(pg,cS,y0,y1,ds.map(d=>d.serial||"-")); });
      return;
    }
    const r=ruleFor(key,F,st); if(!r) return;
    const lines=Math.max(1,Math.round((b-a)/0.19));
    setCell(pg,cK,a,b,wrapWords(r.make,cK[1]-cK[0]-0.1,11,Math.min(3,lines)));
    setCell(pg,cM,a,b,r.model.length===1? wrapTok(r.model[0],cM[1]-cM[0]-0.1,11,Math.min(2,lines)) : r.model);
    setCell(pg,cS,a,b,r.serial.length===1? wrapTok(r.serial[0],cS[1]-cS[0]-0.1,11,Math.min(2,lines)) : r.serial);
  });
}

/* ---------- 11. CCTV inventory summary ---------- */
function camGroups(F){
  const g=[]; F.cams.filter(c=>c.serial||c.ip||c.make||c.model).forEach(c=>{ const k=(c.make||"-")+"|"+(c.model||"-"); let x=g.find(y=>y.k===k);
    if(!x) g.push(x={k,make:c.make,model:c.model,list:[]}); x.list.push(c); });
  return g;
}
function fill36(pg,F){
  const hdr=find(pg,/^item description$/i); if(!hdr) return;
  const y0=hdr.y+0.5, V=vAt(pg,y0); if(V.length<7) return;
  const C=i=>[V[i],V[i+1]];
  const groups=camGroups(F), hdds=F.list("hdd"), nvr=F.one("nvr");
  rowsBelow(hAt(pg,(V[1]+V[2])/2),hdr.y).forEach(([a,b])=>{
    const no=parseInt(textIn(pg,C(0),a,b),10); if(!(no>=1&&no<=5)) return;
    let item,make="",model="",qty="";
    if(no===1){ item="NVR"; if(nvr){ make=nvr.make; model=nvr.model; qty=(nvr.make||nvr.model||nvr.serial)? "1" : ""; } }
    else if(no===5){ const h=hdds.find(d=>d.model||d.make)||hdds[0]; const cap=h&&(h.model.match(/^ST(\d{1,2})000/i)||[])[1];
      item="HDD"+(cap? "("+cap+"TB)" : ""); if(h){ make=h.make; model=h.model; qty=String(hdds.filter(d=>d.serial||d.model).length||""); } }
    else { const g=groups[no-2]; item=g? "IP CAMERA" : "-"; if(g){ make=g.make; model=g.model; qty=String(g.list.length); } }
    setCell(pg,C(1),a,b,[item]); setCell(pg,C(2),a,b,[make]);
    setCell(pg,C(3),a,b,wrapChars(model,V[4]-V[3]-0.12,11).slice(0,2).length<=2 && model? wrapChars(model,V[4]-V[3]-0.12,11) : [model]);
    setCell(pg,C(5),a,b,[qty]);
  });
  const cam=find(pg,/^camera\s+1/i);
  if(cam){ const n=F.cams.length; cam.t=n? "Camera 1 – "+n : "Camera"; cam.tw=estW(cam.t,cam.s);
    const c=F.cams.find(x=>x.user||x.pass);
    if(c){ const same=pg.els.filter(e=>isT(e)&&e!==cam&&Math.abs(e.y-cam.y)<0.06&&e.x>cam.x+1).sort((x,y)=>x.x-y.x);
      if(same[0]&&c.user){ same[0].t=c.user; same[0].tw=estW(c.user,same[0].s); } if(same[1]&&c.pass){ same[1].t=c.pass; same[1].tw=estW(c.pass,same[1].s); } } }
}

/* ---------- 12. IP scheme ---------- */
function siteRow(pg,headRe){   // the small "Site Code | IP Range | Gateway | Subnet Mask" table
  const h=find(pg,headRe,e=>e.y<3.2); if(!h) return null;
  const V=vAt(pg,h.y+0.05), rows=rowsBelow(hAt(pg,cx(h)),h.y); if(!rows.length) return null;
  const colBy=re=>{ const e=find(pg,re,x=>Math.abs(x.y-h.y)<0.12); return e? colOf(V,cx(e)) : null; };
  return {row:rows[0],range:colBy(/^ip range$/i),gw:colBy(/^gateway$/i),mask:colBy(/^subnet mask$/i)};
}
function fill50(pg,F,D){
  const gw=(D&&D.gateway)||"";
  const sr=siteRow(pg,/^gateway$/i);
  if(sr&&sr.gw) setCell(pg,sr.gw,sr.row[0],sr.row[1],[gw]);
  const hdr=find(pg,/^item description$/i); if(!hdr) return;
  const V=vAt(pg,hdr.y+0.5); if(V.length<6) return;
  const C=i=>[V[i],V[i+1]]; let sw=0;
  rowsBelow(hAt(pg,(V[1]+V[2])/2),hdr.y).forEach(([a,b])=>{
    if(!/^\d+$/.test(textIn(pg,C(0),a,b))) return;
    const key=nk(textIn(pg,C(1),a,b)); let d=null;
    if(/^(atg|fcc|ams|ipc|mms)$/.test(key)) d=F.one(key);
    else if(key==="router"){ d=F.one("router"); if(!(d&&d.ip)&&gw) d={ip:gw}; }
    else if(/bospc|desktop/.test(key)) d=F.one("desktop");
    else if(/^ciscoswitch/.test(key)){ sw++; d=F.num("l2",sw); }
    else if(key==="poe") d=F.one("poe");
    else if(/accesspoint|wifi/.test(key)) d=F.one("wifi");
    else if(key==="eps") d=F.one("epsBase");
    else return;
    const ip=(d&&d.ip)||"";
    setCell(pg,C(3),a,b,[ip]);
    const off=gwOff(ip,gw,false); if(off) setCell(pg,C(2),a,b,[off]);
  });
}

/* ---------- 13. CCTV IP scheme ---------- */
function fill51(pg,F,D){
  const gw=mode(F.cams.map(c=>c.gw)), mask=mode(F.cams.map(c=>c.subnet));
  const sr=siteRow(pg,/^ip range$/i);
  if(sr){ const rg=ipRange(gw,mask);
    if(sr.range) setCell(pg,sr.range,sr.row[0],sr.row[1],rg? [rg[0]+" -",rg[1]] : ["-"]);
    if(sr.gw) setCell(pg,sr.gw,sr.row[0],sr.row[1],[gw]);
    if(sr.mask) setCell(pg,sr.mask,sr.row[0],sr.row[1],[mask]); }
  const hdr=find(pg,/^item description$/i); if(!hdr) return;
  const V=vAt(pg,hdr.y+0.6); if(V.length<6) return;
  const C=i=>[V[i],V[i+1]];
  const nvr=F.one("nvr"), nvrIp=(D&&D.nvrIp)||(nvr&&nvr.ip)||"";
  rowsBelow(hAt(pg,(V[1]+V[2])/2),hdr.y).forEach(([a,b])=>{
    const no=parseInt(textIn(pg,C(0),a,b),10); if(!(no>0)) return;
    const isNvr=nk(textIn(pg,C(1),a,b))==="nvr";
    const cam=isNvr? null : F.cam(no), ip=isNvr? nvrIp : (cam&&cam.ip)||"";
    setCell(pg,C(3),a,b,[ip]);
    setCell(pg,C(2),a,b,[gwOff(ip,gw,true)]);
    if(!isNvr) setCell(pg,C(4),a,b,[(cam&&(cam.name||cam.loc))||""]);
    else if(F.cams.length) setCell(pg,C(0),a,b,[String(F.cams.length+1)]);
  });
}

/* ---------- 14. consolidated make / model / serial (cameras) – drawn fresh from the Excel ---------- */
function fill52(pg,F){
  const X=[1.155,1.75,2.8267,3.4667,3.8967,5.6867,7.3333], top=1.9, hh=0.3717, ln=0.0133;
  const els=pg.els.filter(e=>isT(e)&&e.y<1.3);
  const borders=(pg.borders||[]).filter(b=>b.f!=="000000");
  const T=(t,i,y,s,c)=>{ const w=X[i+1]-X[i]-0.08; let si=s; const e=estW(t,si); if(e>w) si=Math.max(5,Math.floor(si*w/e*2)/2);
    els.push({k:"t",t,x:X[i]+0.04,y,w,h:0.1533,s:si,c,tw:Math.min(w,estW(t,si)),al:"c"}); };
  els.push({k:"r",x:X[0],y:top,w:X[6]-X[0],h:hh,f:"00A44F"});
  ["Sr. No.","Item Description","Make","Qty.","Model","Serial No"].forEach((t,i)=>T(t,i,top+hh/2-0.077,11,"FFFFFF"));
  let groups=camGroups(F).map(g=>({item:"CCTV IP Camera",make:g.make||"-",model:g.model||"-",qty:String(g.list.length),sn:g.list.map(c=>c.serial||"-")}));
  if(!groups.length) groups=[{item:"CCTV IP Camera",make:"-",model:"-",qty:"-",sn:["-"]}];
  const total=groups.reduce((n,g)=>n+g.sn.length,0);
  const pitch=Math.min(0.183,(10.0-top-hh-0.16*groups.length)/Math.max(1,total)), fs=pitch<0.16? Math.max(6,Math.floor(11*pitch/0.183*2)/2) : 11;
  let y=top+hh; const hl=[top,y];
  groups.forEach((g,i)=>{ const rh=Math.max(0.5,g.sn.length*pitch+0.16), mid=y+rh/2-0.077;
    els.push({k:"r",x:X[0],y,w:X[6]-X[0],h:rh,f:"D9FFEB"});
    T(String(i+1),0,mid,11,"000000"); T(g.item,1,mid,11,"000000"); T(g.make,2,mid,11,"000000"); T(g.qty,3,mid,11,"000000"); T(g.model,4,mid,11,"000000");
    g.sn.forEach((s,j)=>T(s,5,y+rh/2-0.077+(j-(g.sn.length-1)/2)*pitch,fs,"000000"));
    y+=rh; hl.push(y); });
  hl.forEach(v=>borders.push({k:"r",x:X[0],y:v,w:X[6]-X[0]+ln,h:ln,f:"000000"}));
  X.forEach(v=>borders.push({k:"r",x:v,y:top,w:ln,h:y-top,f:"000000"}));
  pg.els=els; pg.borders=borders;
}

/* ---------- 16. Annexure 1 – port details ---------- */
function portFill(pg,F,D){
  const port=find(pg,/^port$/i); if(!port) return;
  const title=nk((find(pg,/switch$/i,e=>e.y<port.y)||find(pg,/^router$/i,e=>e.y<port.y)||{t:""}).t);
  const me=/^router/.test(title)? F.one("router") : /^2nd/.test(title)? F.num("l2",2) : /^poe/.test(title)? F.one("poe") : F.num("l2",1);
  const gw=(D&&D.gateway)||"";
  const none=/^2nd/.test(title)&&!me? "NA" : "-";
  const zone=pg.els.filter(e=>isT(e)&&e.y>1.2&&e.y<port.y-0.03&&!/host name|address|^serial|number:|^make|^router$|switch$|^device connectivity/i.test(e.t.trim())).sort((a,b)=>a.y-b.y);
  const put=(e,t)=>{ if(e){ e.t=t; e.tw=estW(t,e.s); e.w=Math.max(e.w,e.tw); } };
  put(zone[0],(me&&me.ip)||(/^router/.test(title)&&gw)||none); put(zone[1],(me&&me.serial)||none);
  const mk=find(pg,/^make\s*:/i,e=>e.y<port.y);
  if(mk) put(mk,"Make: "+((me&&me.make)||"-")+" ; Model No : "+((me&&me.model)||"-"));
  // rows
  const first=pg.els.filter(e=>isT(e)&&/^\d+$/.test(e.t.trim())&&e.y>port.y+0.3).sort((a,b)=>a.x-b.x||a.y-b.y)[0]; if(!first) return;
  const V=vAt(pg,first.y); if(V.length<8) return;
  const C=i=>[V[i],V[i+1]];
  const isPoe=/^poe/.test(title);
  rowsBelow(hAt(pg,(V[4]+V[5])/2),port.y+0.3).forEach(([a,b])=>{
    const no=parseInt(textIn(pg,C(0),a,b,true),10); if(!(no>0)) return;
    const key=nk(textIn(pg,C(2),a,b,true)); if(!key) return;
    let d=null, m;
    if((m=key.match(/cctv(\d+)$/))){ const c=F.cam(+m[1]); d={ip:c? c.ip:"",serial:c? c.serial:"",loc:c? (c.name||c.loc):"",cam:true}; }
    else if(/^(fcc|atg|ipc|mms|ams|nvr)$/.test(key)) d=F.one(key)||{};
    else if(key==="router"){ d=F.one("router")||{}; if(!d.ip&&gw) d=Object.assign({},d,{ip:gw}); }
    else if(/^bo[cs]pc$|^desktop/.test(key)) d=F.one("desktop")||{};
    else if(key==="eps") d=F.one("epsBase")||{};
    else if((m=key.match(/^opt(\d)$/))) d=F.num("opt",+m[1])||{};
    else if((m=key.match(/^l2switch(\d)$/))) d=F.num("l2",+m[1])||{};
    else if(/^poe/.test(key)) d=F.one("poe")||{};
    else if(/accesspoint|wifi/.test(key)) d=F.one("wifi")||{};
    else return;
    const old=pg.els.find(e=>isT(e)&&inBand(e,a,b)&&e.x>=V[4]-0.025&&e.x<V[6]-0.025)||first;
    take(pg,e=>inBand(e,a,b)&&e.x>=V[4]-0.025&&e.x<V[6]-0.025);
    const ref={h:old.h,s:old.s,c:old.c,r:old.r};
    let s=ref.s, ipL, snL;
    for(;;){ ipL=wrapIp(d.ip||"",V[5]-V[4]-0.05,s); snL=wrapChars(d.serial||"",V[6]-V[5]-0.05,s);
      if(Math.max(ipL.length,snL.length)*s/59<=b-a-0.02 || s<=5) break; s-=0.5; }
    setCell(pg,C(4),a,b,ipL,{ref,s}); setCell(pg,C(5),a,b,snL,{ref,s});
    if(d.cam && isPoe && V.length>11){ const c=C(10);
      // a location word that the template glued to the camera label of the next column ("FX CCTV Electrical")
      pg.els.forEach(e=>{ if(isT(e)&&inBand(e,a,b)&&e.x<c[0]-0.025){ const m=e.t.match(/^((?:VF|FX|DOME) CCTV|DOME)\s+[A-Za-z]/); if(m){ e.t=m[1]; e.tw=estW(e.t,e.s); e.w=e.tw; } } }); setCell(pg,c,a,b,wrapWords(d.loc||"-",c[1]-c[0]-0.05,ref.s,3),{start:true,ref}); }
  });
}

/* ---------- page 2: document owners + change history ---------- */
const MON=["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
const fmtDate=ms=>{ const d=new Date(ms); return String(d.getDate()).padStart(2,"0")+" "+MON[d.getMonth()]+" "+d.getFullYear(); };
const isoDay=ms=>{ const d=new Date(ms); return d.getFullYear()+"-"+String(d.getMonth()+1).padStart(2,"0")+"-"+String(d.getDate()).padStart(2,"0"); };
function firstDpr(site){   // first DPR date: set by the PPT team, else the day of the first upload for the site
  if(site.dprStart && !isNaN(Date.parse(site.dprStart+"T00:00:00"))) return Date.parse(site.dprStart+"T00:00:00");
  const ups=photosOf(site.id).map(p=>+p.uploadedAt).filter(Boolean);
  return ups.length? Math.min(...ups) : (site.pptCreatedAt||site.createdAt||Date.now());
}
function docMeta(site){
  const me=S.canPPT? (S.myName||"") : "";
  const c=firstDpr(site), u=Math.max(site.pptUpdatedAt||Date.now(),c);
  return {author:site.pptCreatedBy||me||"-", crBy:site.pptCreatedBy||me||"-", upBy:site.pptUpdatedBy||me||"-", crDate:fmtDate(c), upDate:fmtDate(u)};
}
function fill2(pg,site){
  const m=docMeta(site);
  const right=(re,t)=>{ const l=find(pg,re); if(!l) return; const v=pg.els.find(e=>isT(e)&&e!==l&&Math.abs(e.y-l.y)<0.06&&e.x>l.x+1); if(v){ v.t=t; v.tw=estW(t,v.s); v.w=Math.max(v.w,v.tw); } };
  right(/^document author$/i,m.author); right(/^release date$/i,m.upDate);
  [["0.0",m.crDate,m.crBy],["1.0",m.upDate,m.upBy]].forEach(([ver,date,by])=>{
    const a=find(pg,new RegExp("^"+ver.replace(".","\\.")+"$"),e=>e.y>4.4&&e.x<2.4); if(!a) return;
    const V=vAt(pg,a.y+0.02), rows=hAt(pg,cx(a)); const row=rows.map((h,i)=>[h,rows[i+1]]).find(r=>r[1]&&cy(a)>r[0]&&cy(a)<r[1]);
    if(V.length<6||!row) return;
    setCell(pg,[V[1],V[2]],row[0],row[1],[date]); setCell(pg,[V[3],V[4]],row[0],row[1],[by]);
  });
}

/* ---------- apply to the pages of one site ---------- */
const TABLES=new Set([12,13,14,15,16,18,19,20,21,23,24,25,28,33]);
function fillPages(pages,site,D){
  const F=finder(D);
  return pages.map(p0=>{
    const n=Math.floor(p0.n), whole=Number.isInteger(p0.n);
    const want=(whole&&(n===2||TABLES.has(n)||n===36||n===50||n===51||n===52||(n>=55&&n<=58)));
    if(!want) return p0;
    const pg=clone(p0);
    try{
      if(n===2) fill2(pg,site);
      else if(TABLES.has(n)) tableFill(pg,F);
      else if(n===36) fill36(pg,F);
      else if(n===50) fill50(pg,F,D);
      else if(n===51) fill51(pg,F,D);
      else if(n===52) fill52(pg,F);
      else portFill(pg,F,D);
    }catch(e){ console.error("auto fill, page "+p0.n,e); return p0; }
    return pg;
  });
}
function variant(V,site,opts){
  const key=fileKey(site); let st=store[site.id];
  if(!st || st.key!==key){ load(site).then(()=>{ if(S.siteId===site.id && !(opts&&opts.docPages)) renderAll(); }).catch(()=>{}); st=store[site.id]; }
  const D=(st && st.key===key && st.data)||null;
  return {pages:fillPages(V.pages,site,D),folders:V.folders};
}
function summary(site){
  const st=store[site.id]; const D=st&&st.data;
  if(!fileKey(site)) return {ok:false,html:'No serial-number Excel in <b>14.1 Device Serial Numbers</b>. Upload the “Device and CCTV Sr no Sheet” there, or every make / model / serial number / IP in the tables is shown as “-”.'};
  if(!st || st.data===undefined) return {ok:true,html:"Reading the serial-number Excel…"};
  if(!D) return {ok:false,html:'The Excel in <b>14.1 Device Serial Numbers</b> could not be read'+(st.err? " ("+esc(st.err)+")" : "")+'. It needs the columns MATERIAL, MAKE, MODEL, SERIAL NO. – the tables are shown with “-”.'};
  const sn=D.devices.filter(d=>d.serial).length, cs=D.cams.filter(c=>c.serial).length;
  return {ok:true,html:'Tables are filled from <b>'+esc(st.files.join(", "))+'</b>: '+sn+' device serial number'+(sn===1?"":"s")+' and '+cs+' camera'+(cs===1?"":"s")+
    '. A device with no entry in the Excel is shown as “-”.'+(D.gateway? "" : '<br><span style="color:var(--ink-2)">The “Gateway:-” cell of the device sheet is empty, so the NRO gateway on page 12 is shown as “-”.</span>')};
}
/* the Generate PPT dialog: first DPR date + what the Excel gave */
function dialog(siteId){
  const site=S.sites.find(s=>s.id===siteId), box=document.getElementById("af-box"); if(!site||!box) return;
  const draw=()=>{ const b=document.getElementById("af-st"); if(!b) return; const s=summary(site); b.className="banner"+(s.ok?"":" bad"); b.innerHTML="<div>"+s.html+"</div>"; };
  box.innerHTML='<div class="banner" id="af-st" style="margin:0 0 12px"></div>'+
    '<div class="fld"><label for="af-dpr">First DPR date <span style="font-weight:400;color:var(--ink-2)">(“created” date on page 2)</span></label><input type="date" id="af-dpr" name="dpr" value="'+isoDay(firstDpr(site))+'" max="'+isoDay(Date.now())+'"></div>'+
    '<p style="font-size:12px">Page 2 shows <b>'+esc(S.myName||"")+'</b> as the person who generated this PPT, and today as the updated date.</p>';
  draw(); load(site).then(draw).catch(draw);
}
/* before the slides are built: read the Excel and record who generates the PPT */
async function prepare(site,opts){
  toast("Reading serial numbers…");
  try{ await load(site); }catch(e){}
  try{ const r=await call("pptGenerated",{id:site.id,dprStart:(opts&&opts.dpr)||""}); if(r&&r.meta){ Object.assign(site,r.meta); decorate(); } }
  catch(e){ toast("Could not record this PPT in the history: "+errText(e)); }
}

/* ---------- access history + passwords (PPT team) ---------- */
const WHEN=ms=>ms? new Date(ms).toLocaleString("en-IN",{timeZone:"Asia/Kolkata",day:"2-digit",month:"short",year:"numeric",hour:"2-digit",minute:"2-digit",hour12:true}) : "—";
const ACT={login:"Signed in",login_failed:"Sign-in refused",open:"Opened the website",logout:"Signed out",ppt:"Generated a PPT",site_created:"Created a site folder",site_deleted:"Deleted a site folder",password:"Password",user:"Logins"};
function wide(){ const f=document.getElementById("mform"); if(f){ f.style.maxWidth="min(96vw,1020px)"; f.style.width="min(96vw,1020px)"; } }
async function historyModal(){
  modal('<h3>Access history</h3><p>Who signed in or opened the website, and when (India time).</p><div id="af-h">Loading…</div>'+
    '<div class="btns" style="justify-content:flex-end;margin-top:12px"><button type="button" class="btn" id="af-hr">Refresh</button><button type="button" class="btn pri" id="m-cancel">Close</button></div>',async()=>{});
  wide(); document.getElementById("m-cancel").onclick=closeModal; document.getElementById("af-hr").onclick=historyModal;
  const box=document.getElementById("af-h");
  try{
    const r=await call("history",{});
    const th='style="text-align:left;padding:7px 10px;position:sticky;top:0;background:var(--surface,#fff);border-bottom:1px solid var(--line,#d9dee7);white-space:nowrap"', td='style="padding:6px 10px;border-bottom:1px solid var(--line,#e6e9ef);vertical-align:top"';
    const users=r.users||[], log=r.log||[];
    box.innerHTML=(S.isOwner? '<div class="af-newuser"><b>Create a PPT team login</b><input id="af-nu" type="text" placeholder="Full name, e.g. Suresh Patil" autocomplete="off" maxlength="60"><button type="button" class="btn pri" id="af-nub">Create login</button><span>The password is made for you and shown once.</span></div>' : '')+
      '<h4 style="margin:6px 0">Last access by user</h4><div style="max-height:26vh;overflow:auto;border:1px solid var(--line,#d9dee7);border-radius:10px"><table style="width:100%;border-collapse:collapse;font-size:13px"><thead><tr><th '+th+'>User</th><th '+th+'>Access</th><th '+th+'>Last seen on the website</th>'+(S.isOwner?'<th '+th+'></th>':'')+'</tr></thead><tbody>'+
      (users.length? users.map(u=>'<tr><td '+td+'><b>'+esc(u.login)+'</b></td><td '+td+'>'+esc(u.role)+(u.owner?" · owner":"")+(u.active===false?" · switched off":"")+'</td><td '+td+'>'+esc(WHEN(u.lastSeen))+'</td>'+(S.isOwner?'<td '+td+'>'+(u.role==="PPT team"?'<button type="button" class="btn" data-af-reset="'+esc(u.login)+'" style="padding:3px 9px;font-size:12px">Reset password</button>'+(u.owner?'':' <button type="button" class="btn" data-af-active="'+esc(u.login)+'" data-on="'+(u.active===false?0:1)+'" style="padding:3px 9px;font-size:12px">'+(u.active===false?'Switch on':'Switch off')+'</button>'):'')+'</td>':'')+'</tr>').join("") : '<tr><td '+td+' colspan="4">No one yet.</td></tr>')+
      '</tbody></table></div>'+
      '<h4 style="margin:14px 0 6px">Activity <span style="font-weight:400;color:var(--ink-2)">(latest '+log.length+')</span></h4><div style="max-height:38vh;overflow:auto;border:1px solid var(--line,#d9dee7);border-radius:10px"><table style="width:100%;border-collapse:collapse;font-size:13px"><thead><tr><th '+th+'>Date &amp; time</th><th '+th+'>User</th><th '+th+'>Access</th><th '+th+'>Activity</th><th '+th+'>Device</th><th '+th+'>IP address</th></tr></thead><tbody>'+
      (log.length? log.map(x=>'<tr><td '+td+' style="white-space:nowrap">'+esc(WHEN(x.at))+'</td><td '+td+'><b>'+esc(x.login)+'</b></td><td '+td+'>'+esc(x.role)+'</td><td '+td+'>'+esc((ACT[x.action]||x.action)+(x.detail&&x.action!=="login"&&x.action!=="open"&&x.action!=="logout"? ": "+x.detail : ""))+'</td><td '+td+'>'+esc(x.device)+'</td><td '+td+'>'+esc(x.ip)+'</td></tr>').join("") : '<tr><td '+td+' colspan="6">Nothing recorded yet.</td></tr>')+
      '</tbody></table></div>';
    const nub=document.getElementById("af-nub");
    if(nub) nub.onclick=async()=>{ const nm=document.getElementById("af-nu").value.trim(); if(!nm){ toast("Enter the person’s name."); return; }
      try{ const x=await call("createUser",{name:nm}); window.prompt("Login created.\nLogin ID: "+x.login+"\nPassword (copy it now – it is shown only once):",x.password); historyModal(); }catch(e){ toast(errText(e)); } };
    box.querySelectorAll("[data-af-active]").forEach(b=>b.onclick=async()=>{ const who=b.getAttribute("data-af-active"), on=b.getAttribute("data-on")==="1";
      if(on && !window.confirm("Switch off the login of "+who+"? They will not be able to sign in.")) return;
      try{ await call("setUserActive",{login:who,active:!on}); historyModal(); }catch(e){ toast(errText(e)); } });
    box.querySelectorAll("[data-af-reset]").forEach(b=>b.onclick=async()=>{ const who=b.getAttribute("data-af-reset");
      if(!window.confirm("Give "+who+" a new password? The old one stops working.")) return;
      try{ const x=await call("resetPassword",{login:who}); window.prompt("New password for "+x.login+" – copy it and give it to them. It is shown only now.",x.password); }catch(e){ toast(errText(e)); } });
  }catch(e){ box.textContent=errText(e); }
}
function passwordModal(){
  modal('<h3>Change my password</h3><p>For <b>'+esc(S.myName)+'</b>. At least 8 characters.</p>'+
    '<div class="fld"><label for="af-p0">Current password</label><input id="af-p0" name="p0" type="password" autocomplete="current-password"></div>'+
    '<div class="fld"><label for="af-p1">New password</label><input id="af-p1" name="p1" type="password" autocomplete="new-password"></div>'+
    '<div class="fld"><label for="af-p2">New password again</label><input id="af-p2" name="p2" type="password" autocomplete="new-password"></div>'+
    '<div class="fld"><span class="err" id="m-err"></span></div>'+
    '<div class="btns" style="justify-content:flex-end"><button type="button" class="btn" id="m-cancel">Cancel</button><button type="submit" class="btn pri">Save password</button></div>',
    async fd=>{ const a=String(fd.get("p0")||""), b=String(fd.get("p1")||""), c=String(fd.get("p2")||""), err=document.getElementById("m-err");
      if(b.length<8){ err.textContent="The new password needs at least 8 characters."; return; }
      if(b!==c){ err.textContent="The two new passwords are not the same."; return; }
      try{ await call("changePassword",{current:a,next:b}); closeModal(); toast("Password changed"); }catch(e){ err.textContent=errText(e); } });
  document.getElementById("m-cancel").onclick=closeModal;
}
function menu(m){
  const h=m.querySelector("#um-hist"), p=m.querySelector("#um-pass");
  if(h) h.onclick=()=>{ m.hidden=true; historyModal(); };
  if(p) p.onclick=()=>{ m.hidden=true; passwordModal(); };
}
/* a "History" tab next to Dashboard / Site Folders / Checklist, for the PPT team */
function role(){
  const nav=document.querySelector("nav.tabs"); if(!nav) return;
  let b=document.getElementById("tab-hist");
  if(!b){ b=document.createElement("button"); b.className="tab"; b.id="tab-hist"; b.type="button";
    b.innerHTML='<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/></svg><span>History</span>';
    b.addEventListener("click",e=>{ e.preventDefault(); e.stopPropagation(); historyModal(); }); nav.appendChild(b); }
  b.style.display=S.canEdit? "" : "none";
  heroButton(); decorate();
}
/* ---------- easy ways to the Generate PPT dialog ---------- */
const PPT_ICON='<svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="3" y="4" width="18" height="12" rx="2"/><path d="M8 20h8M12 16v4M10 8.5v3l3-1.5z"/></svg>';
function zoneSites(){ return S.sites.filter(s=>!S.region||S.region==="all"||(s.region||"east")===S.region); }
/* choose the site first, then the normal Generate PPT dialog opens */
function pickSite(){
  if(!S.canPPT){ toast("Only the PPT team can generate the PPT."); return; }
  const sites=zoneSites();
  if(!sites.length){ toast("Create a site folder first."); return; }
  if(sites.length===1){ pptDialog(sites[0].id); return; }
  const row=s=>{ const st=siteStats(s), left=st.total-st.done;
    return '<button type="button" class="af-pick" data-af-ppt="'+esc(s.id)+'" data-q="'+esc((s.code+" "+s.name+" "+(s.state||"")).toLowerCase())+'"><span class="af-pick-main"><b>'+esc(s.code)+'</b> '+esc(s.name)+'<small>'+(left? left+' of '+st.total+' photo places still empty' : 'All '+st.total+' photo places filled')+'</small></span>'+
      '<span class="af-pick-pct'+(st.pct===100?' full':'')+'">'+st.pct+'%</span><span class="af-pick-go">'+PPT_ICON+'Generate</span></button>'; };
  modal('<h3>Generate PPT</h3><p>Choose the site. Sites with every photo place filled are listed first.</p>'+
    '<div class="fld"><input id="af-pq" type="search" placeholder="Search site code or name" aria-label="Search sites"></div>'+
    '<div class="af-picks" id="af-picks">'+sites.slice().sort((a,b)=>siteStats(b).pct-siteStats(a).pct||String(a.code).localeCompare(String(b.code))).map(row).join("")+'</div>'+
    '<div class="btns" style="justify-content:flex-end;margin-top:12px"><button type="button" class="btn" id="m-cancel">Cancel</button></div>',async()=>{});
  document.getElementById("m-cancel").onclick=closeModal;
  const q=document.getElementById("af-pq");
  q.oninput=()=>{ const v=q.value.trim().toLowerCase(); document.querySelectorAll("#af-picks .af-pick").forEach(b=>{ b.hidden=!!v&&!b.getAttribute("data-q").includes(v); }); };
  q.addEventListener("keydown",e=>{ if(e.key!=="Enter") return; e.preventDefault(); const f=[...document.querySelectorAll("#af-picks .af-pick")].find(x=>!x.hidden); if(f) pptDialog(f.getAttribute("data-af-ppt")); });
}
/* runs after every screen update: a Generate button on each site row, and a floating one inside a site */
function decorate(){
  try{
    document.querySelectorAll("#tbody button.link[data-open-site]").forEach(b=>{
      if(!S.canPPT || b.parentNode.querySelector(".af-rowppt")) return;
      const g=document.createElement("button"); g.type="button"; g.className="af-rowppt"; g.setAttribute("data-af-ppt",b.getAttribute("data-open-site")); g.innerHTML=PPT_ICON+"Generate PPT";
      b.parentNode.style.whiteSpace="nowrap"; b.parentNode.appendChild(g); });
    let fab=document.getElementById("af-fab");
    const site=S.view==="folders"&&S.siteId? S.sites.find(s=>s.id===S.siteId) : null;
    if(site && S.canPPT){
      if(!fab){ fab=document.createElement("button"); fab.type="button"; fab.id="af-fab"; fab.className="af-fab"; document.body.appendChild(fab); }
      const st=siteStats(site), left=st.total-st.done;
      fab.setAttribute("data-af-ppt",site.id);
      fab.innerHTML=PPT_ICON+'<span><b>Generate PPT</b><small>'+esc(site.code)+' · '+(left? left+' photo place'+(left>1?"s":"")+' empty' : 'all photos in')+'</small></span>';
      fab.hidden=false;
    } else if(fab) fab.hidden=true;
    board(); recent(); ready(site); steps(); findButton(); liveChip();
  }catch(e){ console.error(e); }
}
document.addEventListener("click",e=>{ const b=e.target.closest&&e.target.closest("[data-af-ppt]"); if(!b) return; e.preventDefault(); e.stopPropagation(); pptDialog(b.getAttribute("data-af-ppt")); },true);
try{ const ra=renderAll; renderAll=function(){ const r=ra.apply(this,arguments); decorate(); return r; }; }catch(e){}
try{ const sv=setView; setView=function(){ const r=sv.apply(this,arguments); decorate(); return r; }; }catch(e){}
function heroButton(){
  const qa=document.querySelector(".hero .qa"); if(!qa) return;
  let b=document.getElementById("qa-ppt");
  if(!b){ b=document.createElement("button"); b.type="button"; b.className="btn af-heroppt"; b.id="qa-ppt"; b.innerHTML=PPT_ICON+"Generate PPT"; b.addEventListener("click",pickSite); qa.appendChild(b); }
  b.style.display=S.canPPT? "" : "none";
}
/* ---------- what a site still needs before its PPT ---------- */
const RING=(pct,sz)=>{ const r=sz/2-5, c=2*Math.PI*r; return '<svg class="af-ring" width="'+sz+'" height="'+sz+'" viewBox="0 0 '+sz+' '+sz+'" aria-hidden="true"><circle cx="'+sz/2+'" cy="'+sz/2+'" r="'+r+'" fill="none" stroke="currentColor" stroke-opacity=".15" stroke-width="6"/><circle cx="'+sz/2+'" cy="'+sz/2+'" r="'+r+'" fill="none" stroke="currentColor" stroke-width="6" stroke-linecap="round" stroke-dasharray="'+(c*pct/100).toFixed(1)+' '+c.toFixed(1)+'" transform="rotate(-90 '+sz/2+' '+sz/2+')"/><text x="50%" y="50%" dominant-baseline="central" text-anchor="middle" font-size="'+Math.round(sz*0.26)+'" font-weight="700" fill="currentColor">'+pct+'%</text></svg>'; };
const folderOf=(site,t)=>(site.subfolders||[]).find(f=>tplName(f.name)===t)||null;
function emptyFolders(site){
  const st=siteStats(site);
  return st.folders.filter(f=>{ if((st.byF[f.id]||[]).length) return false; const t=tplName(f.name), i=t? FINFO[t] : null;
    return !((i && (i.table || i.opt || !inLayout(site,t))) || OLD_HDD.test(normF(f.name))); });
}
function excelState(site){
  if(!fileKey(site)) return {ok:false,text:"No serial-number Excel yet",sub:"Upload it to 14.1 Device Serial Numbers – the tables are filled from it automatically."};
  const st=store[site.id];
  if(!st || st.key!==fileKey(site) || st.data===undefined){ load(site).then(()=>{ if(S.siteId===site.id) decorate(); }).catch(()=>{}); return {ok:true,text:"Reading the serial-number Excel…",sub:""}; }
  if(!st.data) return {ok:false,text:"The Excel could not be read",sub:"It needs the columns MATERIAL, MAKE, MODEL, SERIAL NO."};
  const d=st.data.devices.filter(x=>x.serial).length, c=st.data.cams.filter(x=>x.serial).length;
  return {ok:true,data:true,text:"Serial numbers read automatically",sub:d+" device serial number"+(d===1?"":"s")+" · "+c+" camera"+(c===1?"":"s")+" from "+st.files.join(", ")};
}
/* the readiness card at the top of a site */
function ready(site){
  const body=document.getElementById("folders-body"); const old=document.getElementById("af-ready");
  if(!site || S.folderId || !body){ if(old) old.remove(); return; }
  const st=siteStats(site), miss=emptyFolders(site), ex=excelState(site);
  const n=t=>{ const f=folderOf(site,t); return f? (st.byF[f.id]||[]).length : 0; };
  const fid=t=>{ const f=folderOf(site,t); return f? f.id : ""; };
  const item=(ok,title,sub,folder,extra)=>'<div class="af-chk '+(ok?"ok":"no")+'"><i>'+(ok?"✓":"!")+'</i><div><b>'+esc(title)+'</b><span>'+esc(sub)+'</span></div>'+(extra||"")+(folder&&!ok? '<button type="button" class="af-mini" data-af-open="'+esc(folder)+'">Open folder</button>' : "")+'</div>';
  const html='<div class="af-ready-top"><div class="af-ready-ring '+(st.pct===100?"full":"")+'">'+RING(st.pct,84)+'</div>'+
    '<div class="af-ready-head"><h3>'+(st.pct===100&&ex.ok? "Ready for the PPT" : "PPT readiness")+'</h3><p>'+st.done+' of '+st.total+' photo places filled'+(miss.length? ' · '+miss.length+' still empty' : '')+(site.pptUpdatedAt? ' · last PPT '+fmtDate(site.pptUpdatedAt)+' by '+esc(site.pptUpdatedBy||"—") : ' · no PPT generated yet')+'</p></div>'+
    '<div class="af-ready-act">'+(ex.data? '<button type="button" class="btn" data-af-data="'+esc(site.id)+'">View Excel data</button>' : '')+(S.canPPT? '<button type="button" class="btn pri" data-af-ppt="'+esc(site.id)+'">'+PPT_ICON+' Generate PPT</button>' : '')+'</div></div>'+
    '<div class="af-chks">'+
      item(!miss.length,"Site photos",miss.length? miss.length+" photo place"+(miss.length>1?"s":"")+" still empty" : "Every photo place has a photo")+
      item(ex.ok,ex.text,ex.sub,fid(SERIAL_F))+
      item(n(CLEAR_F)>0,"Clearance mails",n(CLEAR_F)? n(CLEAR_F)+" file"+(n(CLEAR_F)>1?"s":"")+" uploaded" : "No clearance mail uploaded yet",fid(CLEAR_F))+
      item(n(CHECK_F)>0,"Checklist",n(CHECK_F)? n(CHECK_F)+" file"+(n(CHECK_F)>1?"s":"")+" uploaded" : "No checklist uploaded yet",fid(CHECK_F))+
      item(!!site.address,"Cover page address",site.address? "Set" : "Not set – use “Site details from file”")+
    '</div>'+
    (miss.length? '<div class="af-next"><span>Upload next:</span>'+miss.slice(0,8).map(f=>'<button type="button" class="af-chip" data-af-open="'+esc(f.id)+'">'+esc(f.name)+'</button>').join("")+(miss.length>8? '<em>+'+(miss.length-8)+' more</em>' : '')+'</div>' : '');
  let box=old; if(!box){ box=document.createElement("div"); box.id="af-ready"; box.className="af-ready"; }
  const head=body.querySelector(".sec-h");
  if(head){ if(box.previousElementSibling!==head) head.insertAdjacentElement("afterend",box); } else if(!box.parentNode) body.prepend(box);
  if(box.__h!==html){ box.innerHTML=html; box.__h=html; }
}
/* what was read from the Excel – so it can be checked before the PPT is made */
function dataModal(siteId){
  const site=S.sites.find(x=>x.id===siteId), st=site&&store[site.id], D=st&&st.data; if(!D){ toast("No Excel data for this site yet."); return; }
  const th='style="text-align:left;padding:6px 9px;position:sticky;top:0;background:var(--surface,#fff);border-bottom:1px solid var(--line,#d9dee7);white-space:nowrap"', td='style="padding:5px 9px;border-bottom:1px solid var(--line,#e6e9ef)"';
  const tbl=(heads,rows)=>'<div style="max-height:30vh;overflow:auto;border:1px solid var(--line,#d9dee7);border-radius:10px"><table style="width:100%;border-collapse:collapse;font-size:12.5px"><thead><tr>'+heads.map(h=>'<th '+th+'>'+h+'</th>').join("")+'</tr></thead><tbody>'+rows.map(r=>'<tr>'+r.map(c=>'<td '+td+'>'+esc(c||"—")+'</td>').join("")+'</tr>').join("")+'</tbody></table></div>';
  modal('<h3>Data read from the Excel</h3><p>'+esc(st.files.join(", "))+' · read automatically each time the file changes. These values go into the PPT tables; “—” is shown as “-” in the PPT.</p>'+
    '<h4 style="margin:8px 0 6px">Devices ('+D.devices.length+')</h4>'+tbl(["Material","Make","Model","Serial No.","IP address"],D.devices.map(d=>[d.name,d.make,d.model,d.serial,d.ip]))+
    '<h4 style="margin:14px 0 6px">Cameras ('+D.cams.length+')'+(D.nvrIp? ' · NVR '+esc(D.nvrIp) : '')+'</h4>'+tbl(["No.","Camera name","Make","Model","Serial No.","Camera IP","Gateway"],D.cams.map(c=>[String(c.no),c.name,c.make,c.model,c.serial,c.ip,c.gw]))+
    '<div class="btns" style="justify-content:flex-end;margin-top:12px"><button type="button" class="btn pri" id="m-cancel">Close</button></div>',async()=>{});
  wide(); document.getElementById("m-cancel").onclick=closeModal;
}
/* dashboard: one card per site with what it needs and a Generate button */
function board(){
  const main=document.getElementById("dash-main"); if(!main) return;
  let box=document.getElementById("af-board");
  const sites=zoneSites();
  if(!sites.length){ if(box) box.remove(); return; }
  const rows=sites.map(s=>({s,st:siteStats(s)})).sort((a,b)=>b.st.pct-a.st.pct||b.st.last-a.st.last);
  const show=rows.slice(0,8);
  const html='<div class="af-board-h"><div><h3>PPT status by site</h3><p>Updates by itself as photos and files arrive.</p></div>'+(rows.length>show.length? '<button type="button" class="btn" data-af-find="1">All '+rows.length+' sites</button>' : '')+'</div><div class="af-cards">'+
    show.map(({s,st})=>{ const xl=!!fileKey(s), left=st.total-st.done;
      return '<div class="af-card"><div class="af-card-top"><div class="af-card-ring '+(st.pct===100?"full":"")+'">'+RING(st.pct,56)+'</div><div class="af-card-nm"><b>'+esc(s.code)+'</b><span>'+esc(s.name)+'</span></div></div>'+
        '<div class="af-tags"><span class="af-tag '+(left?"warn":"good")+'">'+(left? left+' photo place'+(left>1?"s":"")+' empty' : 'Photos complete')+'</span><span class="af-tag '+(xl?"good":"bad")+'">'+(xl? 'Excel uploaded' : 'No Excel')+'</span></div>'+
        '<div class="af-card-last">'+(s.pptUpdatedAt? 'Last PPT '+fmtDate(s.pptUpdatedAt)+' · '+esc(s.pptUpdatedBy||"") : 'No PPT generated yet')+'</div>'+
        '<div class="af-card-btns"><button type="button" class="btn" data-af-site="'+esc(s.id)+'">Open</button>'+(S.canPPT? '<button type="button" class="btn pri" data-af-ppt="'+esc(s.id)+'">Generate PPT</button>' : '')+'</div></div>'; }).join("")+'</div>';
  if(!box){ box=document.createElement("div"); box.id="af-board"; box.className="panel af-board"; main.insertBefore(box,main.firstChild); }
  if(box.__h!==html){ box.innerHTML=html; box.__h=html; }
}
/* dashboard (PPT team): who signed in or opened the site last – fetched by itself every minute */
let recentAt=0, recentData=null, recentBusy=false;
function recent(){
  const main=document.getElementById("dash-main"); let box=document.getElementById("af-recent");
  if(!main || !S.canEdit){ if(box) box.remove(); return; }
  if(!box){ box=document.createElement("div"); box.id="af-recent"; box.className="panel af-recent"; const bd=document.getElementById("af-board"); if(bd) bd.insertAdjacentElement("afterend",box); else main.insertBefore(box,main.firstChild); }
  if(!recentBusy && Date.now()-recentAt>60000 && S.view==="dash" && !document.hidden){ recentBusy=true;
    call("history",{limit:40}).then(r=>{ recentData=r; recentAt=Date.now(); }).catch(()=>{ recentAt=Date.now()-30000; }).then(()=>{ recentBusy=false; recent(); }); }
  const r=recentData; let html;
  if(!r) html='<div class="af-board-h"><div><h3>Who was on the site</h3><p>Loading…</p></div></div>';
  else { const seen=(r.log||[]).filter(x=>x.action==="login"||x.action==="open"), last=seen[0];
    const users=(r.users||[]).filter(u=>u.lastSeen).slice(0,6);
    html='<div class="af-board-h"><div><h3>Who was on the site</h3><p>'+(last? 'Last sign-in: <b>'+esc(last.login)+'</b> ('+esc(last.role)+') · '+esc(WHEN(last.at))+' · '+esc(last.device||"") : 'No sign-in recorded yet.')+'</p></div><button type="button" class="btn" id="af-rh">Full history</button></div>'+
      '<div class="af-seen">'+users.map(u=>'<div class="af-seen-u"><span class="af-seen-a">'+esc((u.login[0]||"?").toUpperCase())+'</span><div><b>'+esc(u.login)+'</b><span>'+esc(u.role)+' · '+esc(WHEN(u.lastSeen))+'</span></div></div>').join("")+'</div>'; }
  if(box.__h!==html){ box.innerHTML=html; box.__h=html; const b=document.getElementById("af-rh"); if(b) b.onclick=historyModal; }
}
setInterval(()=>{ if(S.view==="dash" && S.canEdit) recent(); },60000);
/* the three numbered steps on the dashboard become shortcuts */
function steps(){
  document.querySelectorAll(".steps .step").forEach((el,i)=>{ if(el.__af) return; el.__af=1; el.classList.add("af-step"); el.setAttribute("role","button"); el.tabIndex=0;
    const go=()=>{ if(i===0) setView("folders"); else if(i===1){ const b=document.getElementById("qa-upload"); if(b) b.click(); } else if(S.canPPT) pickSite(); else toast("Only the PPT team can generate the PPT."); };
    el.addEventListener("click",go); el.addEventListener("keydown",e=>{ if(e.key==="Enter"||e.key===" "){ e.preventDefault(); go(); } }); });
}
/* quick search: Ctrl+K or the search button → jump to any site, or generate its PPT */
function findSite(){
  const sites=zoneSites(); if(!sites.length){ toast("No site folders yet."); return; }
  const row=s=>{ const st=siteStats(s);
    return '<div class="af-pick" data-q="'+esc((s.code+" "+s.name+" "+(s.state||"")).toLowerCase())+'"><button type="button" class="af-pick-open" data-af-site="'+esc(s.id)+'"><span class="af-pick-main"><b>'+esc(s.code)+'</b> '+esc(s.name)+'<small>'+esc(s.state||"")+(s.state?" · ":"")+st.done+' of '+st.total+' photo places filled</small></span><span class="af-pick-pct'+(st.pct===100?' full':'')+'">'+st.pct+'%</span></button>'+
      (S.canPPT? '<button type="button" class="af-pick-go" data-af-ppt="'+esc(s.id)+'">'+PPT_ICON+'Generate</button>' : '')+'</div>'; };
  modal('<h3>Find a site</h3><div class="fld"><input id="af-pq" type="search" placeholder="Type a site code, name or state" aria-label="Search sites" autocomplete="off"></div>'+
    '<div class="af-picks" id="af-picks">'+sites.slice().sort((a,b)=>String(a.code).localeCompare(String(b.code))).map(row).join("")+'</div>'+
    '<p style="font-size:12px;margin:10px 0 0">Enter opens the first match · Ctrl + K opens this search from anywhere.</p>',async()=>{});
  const q=document.getElementById("af-pq");
  q.addEventListener("keydown",e=>{ if(e.key!=="Enter") return; e.preventDefault(); const f=[...document.querySelectorAll("#af-picks .af-pick")].find(x=>!x.hidden); if(f){ const id=f.querySelector("[data-af-site]").getAttribute("data-af-site"); closeModal(); openSite(id); } });
  q.oninput=()=>{ const v=q.value.trim().toLowerCase(); document.querySelectorAll("#af-picks .af-pick").forEach(b=>{ b.hidden=!!v&&!b.getAttribute("data-q").includes(v); }); };
  q.focus();
}
function findButton(){
  if(document.getElementById("af-find")) return; const th=document.getElementById("b-theme"); if(!th) return;
  const b=document.createElement("button"); b.type="button"; b.id="af-find"; b.className=th.className; b.title="Find a site (Ctrl + K)"; b.setAttribute("aria-label","Find a site");
  b.innerHTML='<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><circle cx="11" cy="11" r="7"/><path d="M20 20l-3.5-3.5"/></svg>';
  b.addEventListener("click",e=>{ e.stopPropagation(); findSite(); }); th.parentNode.insertBefore(b,th);
}
document.addEventListener("keydown",e=>{ if((e.ctrlKey||e.metaKey)&&String(e.key).toLowerCase()==="k"){ const app=document.getElementById("app"); if(app&&!app.hidden){ e.preventDefault(); findSite(); } } });
document.addEventListener("click",e=>{ const t=e.target.closest&&e.target.closest("[data-af-site],[data-af-open],[data-af-data],[data-af-find]"); if(!t) return; e.preventDefault(); e.stopPropagation();
  if(t.hasAttribute("data-af-find")){ findSite(); return; }
  if(t.hasAttribute("data-af-data")){ dataModal(t.getAttribute("data-af-data")); return; }
  if(t.hasAttribute("data-af-open")){ openSite(S.siteId,t.getAttribute("data-af-open")); return; }
  closeModal(); openSite(t.getAttribute("data-af-site")); },true);

/* ---------- data stays fresh by itself: live chip, refresh when the tab comes back ---------- */
let lastSync=0, syncing=false;
try{ const rf=refresh; refresh=async function(){ syncing=true; liveChip(); try{ return await rf.apply(this,arguments); } finally{ syncing=false; if(S.dbState==="ok") lastSync=Date.now(); liveChip(); } }; }catch(e){}
function liveChip(){
  const bar=document.querySelector(".zonebar .zb-inner"); if(!bar) return;
  let c=document.getElementById("af-live");
  if(!c){ c=document.createElement("button"); c.type="button"; c.id="af-live"; c.className="af-live"; c.title="Data refreshes by itself every 20 seconds. Click to refresh now.";
    c.addEventListener("click",()=>{ refresh(true); }); const cnt=bar.querySelector(".zb-count"); if(cnt) bar.insertBefore(c,cnt); else bar.appendChild(c); }
  const app=document.getElementById("app"); c.hidden=!app||app.hidden;
  const off=S.dbState&&S.dbState!=="ok", sec=lastSync? Math.round((Date.now()-lastSync)/1000) : -1;
  c.className="af-live"+(off?" off":"")+(syncing?" busy":"");
  c.innerHTML='<i></i>'+(off? "Offline – retrying" : syncing&&sec<0? "Loading…" : sec<0? "Live" : sec<8? "Live · updated just now" : sec<90? "Live · updated "+sec+" s ago" : "Live · updated "+Math.round(sec/60)+" min ago");
}
setInterval(liveChip,5000);
document.addEventListener("visibilitychange",()=>{ const app=document.getElementById("app"); if(!document.hidden && app && !app.hidden && Date.now()-lastSync>8000) refresh(); });
window.addEventListener("online",()=>{ const app=document.getElementById("app"); if(app&&!app.hidden) refresh(true); });

/* ---------- progress card while the PPT is built ---------- */
let progT=0;
function progress(step,pct,state){
  let p=document.getElementById("af-prog");
  if(!p){ p=document.createElement("div"); p.id="af-prog"; p.className="af-prog"; p.setAttribute("role","status"); document.body.appendChild(p); }
  p.className="af-prog"+(state? " "+state : "");
  p.innerHTML='<div class="af-prog-t"><b>'+(state==="done"? "PPT ready" : state==="bad"? "PPT not built" : "Building the PPT")+'</b><span>'+esc(step)+'</span></div><div class="af-prog-bar'+(pct<0?" run":"")+'"><i style="width:'+(pct<0?40:Math.max(4,pct))+'%"></i></div>';
  p.hidden=false; clearTimeout(progT);
  progT=setTimeout(()=>{ p.hidden=true; },state? (state==="bad"?9000:6000) : 300000);
}
try{ const t0=toast; toast=function(msg){ const m=String(msg); let x;
    if(/^Reading serial numbers/.test(m)){ progress("Reading the serial-number Excel",5); return; }
    if((x=m.match(/^Reading (.+?) …$/))){ progress("Reading "+x[1],12); return; }
    if((x=m.match(/^Building PPT… (\d+) \/ (\d+) photos/))){ const a=+x[1], b=+x[2]; if(!b) progress("Writing the slides",-1); else if(a>=b) progress("All "+b+" photos loaded – writing the slides",-1); else progress("Loading photos · "+a+" of "+b,b? 15+Math.round(75*a/b) : 50); return; }
    const p=document.getElementById("af-prog"), on=p&&!p.hidden&&!/done|bad/.test(p.className);
    if(/^PPT ready/.test(m)){ progress(m.replace(/^PPT ready · /,"")+" · saved to your downloads",100,"done"); return; }
    if(on && /^Could not build the PPT/.test(m)){ progress(m.replace(/^Could not build the PPT: /,""),100,"bad"); return; }
    return t0.apply(this,arguments); }; }catch(e){}

function signedOut(){ try{ fetch("/api/app",{method:"POST",credentials:"same-origin",headers:{"content-type":"application/json"},body:JSON.stringify({action:"logout"})}).catch(()=>{}); }catch(e){} }
return {variant,prepare,dialog,menu,role,signedOut,historyModal,parseBook,fillPages,finder,_store:store};
})();
