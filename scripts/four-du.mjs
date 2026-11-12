/**
 * Adds the 4 DU layouts (4DU2T = 4 DU + 2 Tanks, 4DU3T = 4 DU + 3 Tanks) to the built front end.
 * Runs after build-front-end.mjs. Changes public/assets/app.js and adds the two layouts to folders.json.
 *   - OPT 5–8 photo page after the OPT 1–4 page, OPT table with 8 rows
 *   - DU JB 3 & DU JB 4 termination page
 *   - DU 3 & DU 4 sump sensor page
 */
import fs from 'node:fs';
import path from 'node:path';

const assets = path.join(process.cwd(), 'public', 'assets');
const appPath = path.join(assets, 'app.js');
const log = (...a) => console.log('[four-du]', ...a);

const PATCHES = [
  ['["1DU2T","1 DU + 2 Tanks",1,2]]',
   '["1DU2T","1 DU + 2 Tanks",1,2],["4DU2T","4 DU + 2 Tanks",4,2],["4DU3T","4 DU + 3 Tanks",4,3]]'],
  ['const du=has(/du jb 3|sensor du 3|opt 5/)?3',
   'const du=has(/du jb 4|sensor du 4|opt 7/)?4 : has(/du jb 3|sensor du 3|opt 5/)?3'],
  ['key= (du===3||tank===3)? "3DU3T"',
   'key= du===4? (tank===3? "4DU3T" : "4DU2T") : (du===3||tank===3)? "3DU3T"'],
  ['[[2,2],[3,3],[1,2]].forEach(', '[[2,2],[3,3],[1,2],[4,2],[4,3]].forEach('],
  ['const n=k==="3DU3T"?3:2;', 'const n=/3T$/.test(k)?3:2;', true],
  ['$("#"+p+"-lay").value!=="3DU3T"', '!/3T$/.test($("#"+p+"-lay").value)'],
  ['    // tank products in the captions',
   '    if(n===23 && du===4){ const rowH=0.4483, last=g.els.filter(e=>Math.abs(e.y-3.8127)<0.02);   /* OPT table: 8 rows */\n' +
   '      g.borders.forEach(e=>{ if(e.h>2) e.h+=4*rowH; if(Math.abs(e.y-4.0883)<0.02) e.y+=4*rowH; });\n' +
   '      [5,6,7,8].forEach((k,j)=>{ const y0=4.0883+j*rowH; g.els.push({k:"r",x:1.1017,y:y0,w:6.2983,h:rowH,f:"D9FFEB"}); g.borders.push({k:"r",x:1.0983,y:y0,w:6.305,h:0.0133,f:"000000"});\n' +
   '        last.forEach(e=>{ const c=cloneJ(e); c.y+=(j+1)*rowH; if(/^OPT/.test(c.t)) c.t="OPT "+k; else if(/^SCN/.test(c.t)) c.t="-"; g.els.push(c); }); }); }\n' +
   '    // tank products in the captions'],
  ['    if(n===52) out.push(docPage(',
   '    if(n===22 && du===4){ const o=pg(22); o.n=22.5; o.slots.forEach((sl,i)=>{ sl.f="9.1 OPT "+(5+i); }); [0,1,2,3].forEach(i=>setCap(o,i,"OPT "+(5+i))); out.push(o); }\n' +
   '    if(n===30 && du===4){ const d=pg(30); d.n=30.5; d.slots.forEach((sl,i)=>{ sl.f="9.5.2 DU JB "+(3+(i>>1))+" - Photo "+(1+(i&1)); }); [0,1,2,3].forEach(i=>setCap(d,i,"DU JB "+(3+(i>>1)))); out.push(d); }\n' +
   '    if(n===32 && du===4){ const s2=pg(32); s2.n=32.5; s2.slots[0].f="9.6 Sensor DU 3 Sump"; s2.slots[1].f="9.6 Sensor DU 4 Sump"; setCap(s2,0,"DU 3 Sump"); setCap(s2,1,"DU 4 Sump"); out.push(s2); }\n' +
   '    if(n===52) out.push(docPage('],
];

let app = fs.readFileSync(appPath, 'utf8');
if (app.includes('"4DU2T"')) { log('already applied'); process.exit(0); }
let ok = 0;
for (const [from, to, all] of PATCHES) {
  const count = app.split(from).length - 1;
  if (!count) { log('not applied (text not found): ' + from.slice(0, 60)); continue; }
  if (count > 1 && !all) { log('skipped (found ' + count + ' times): ' + from.slice(0, 60)); continue; }
  app = app.split(from).join(to); ok++;
}
fs.writeFileSync(appPath, app);
log(ok + ' of ' + PATCHES.length + ' changes applied');

// folder lists for the new layouts (used by the Claude / MCP create_site tool)
try {
  const tpl = JSON.parse(fs.readFileSync(path.join(assets, 'bhf141.template.json'), 'utf8'));
  const line = (start) => { const i = app.indexOf(start); return i < 0 ? '' : app.slice(i, app.indexOf('\n', i)) + '\n'; };
  let pre = ['const DEF_PROD', 'const cloneJ', 'const HDD_F', 'const SERIAL_F', 'const ANNEX_F', 'function capEl', 'const GREYLINE', 'const LAYOUTS'].map(line).join('');
  const di = app.indexOf('function docPage'); if (di >= 0) pre += app.slice(di, app.indexOf('\n}', di) + 2) + '\n';
  const region = app.slice(app.indexOf('const ORT='), app.indexOf('const SITE_FOLDERS'));
  const r = new Function('TPLX', pre + region + '\n;return {foldersFor,layoutOf};')(tpl);
  const fp = path.join(assets, 'folders.json');
  const rules = fs.existsSync(fp) ? JSON.parse(fs.readFileSync(fp, 'utf8')) : {};
  for (const layout of ['4DU2T', '4DU3T']) {
    const lo = r.layoutOf({ layout });
    rules[layout] = { layout, du: lo.du, tank: lo.tank, prods: lo.prods, folders: r.foldersFor(lo) };
    log('layout ' + layout + ' has ' + rules[layout].folders.length + ' sub folders');
  }
  fs.writeFileSync(fp, JSON.stringify(rules, null, 1));
} catch (e) {
  log('folders.json not updated: ' + e.message);
}
