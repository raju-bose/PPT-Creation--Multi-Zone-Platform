/**
 * Runs after build-front-end.mjs and four-du.mjs. Changes public/assets/app.js:
 *   - table values of the PPT come from the site's serial-number Excel (scripts/autofill-client.js)
 *   - page 2 of the PPT carries the login name and dates of the person who generated it
 *   - PPT team members have their own login; everyone else uploads photos with the shared Admin login
 *   - an access-history page for the PPT team
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const here = path.dirname(fileURLToPath(import.meta.url));
const appPath = path.join(process.cwd(), 'public', 'assets', 'app.js');
const log = (...a) => console.log('[auto-fill]', ...a);

const PATCHES = [
  // tables + page 2 are filled when the page list of a site is built
  ['const R0=routeFiles(site); const byTpl=R0.byTpl; let extra=R0.extra.slice();\n  const V=variantOf(site);',
   'const R0=routeFiles(site); const byTpl=R0.byTpl; let extra=R0.extra.slice();\n  const V=AF.variant(variantOf(site),site,opts);'],
  // before the slides are built: read the Excel, record who generates the PPT
  ['if(!S.canPPT){ toast("Only the PPT team can generate the PPT."); return; }\n  const docFolders=',
   'if(!S.canPPT){ toast("Only the PPT team can generate the PPT."); return; }\n  await AF.prepare(site,opts);\n  const docFolders='],
  ["modal('<h3>Generate PPT</h3>", "setTimeout(()=>AF.dialog(siteId),0);\n  modal('<h3>Generate PPT</h3>"],
  ['\'<p style="font-size:12px">Tables (serial numbers, IPs, dates, names) are copied from the template as editable text, so you can correct them in PowerPoint.</p>\'+',
   '\'<div id="af-box"></div>\'+'],
  ['try{ await exportPPT(siteId,{pending:!!fd.get("pend")}); }', 'try{ await exportPPT(siteId,{pending:!!fd.get("pend"),dpr:dprDate}); }'],
  ['async fd=>{ closeModal(); try{ await exportPPT(siteId,', 'async fd=>{ const dprDate=String(fd.get("dpr")||""); closeModal(); try{ await exportPPT(siteId,'],
  // roles come from the server: PPT team ("admin") or photo upload
  ['S.canEdit = r.role==="admin" && isPptName(S.myName); S.canPPT = S.canEdit; applyRole();',
   'S.canEdit = r.role==="admin"; S.canPPT = S.canEdit; S.isOwner = !!r.owner; if(r.me) S.myName = r.me; applyRole();'],
  ['function applyRole(){ const q=$("#qa-new"); if(q) q.style.display=S.canEdit?"":"none"; }',
   'function applyRole(){ const q=$("#qa-new"); if(q) q.style.display=S.canEdit?"":"none"; AF.role(); }'],
  // account menu: history + own password for the PPT team; the upload name is the login ID
  ['(S.canEdit?\'\':\'<button class="btn" id="um-name">Change my name</button>\')+\'<button class="btn" id="um-sync">Refresh data</button>',
   '(S.canEdit?\'<button class="btn" id="um-hist">Access history</button><button class="btn" id="um-pass">Change my password</button>\':\'\')+\'<button class="btn" id="um-sync">Refresh data</button>'],
  ['$("#um-out").onclick=()=>{ m.hidden=true; signOut(); };', '$("#um-out").onclick=()=>{ m.hidden=true; signOut(); }; AF.menu(m);'],
  ['function signOut(msg){\n  stopPolling();', 'function signOut(msg){\n  AF.signedOut(); stopPolling();'],
  ['Site engineer · photo upload', 'Photo upload'],
  ['<p class="gate-foot">Shared access for every user: ID <b>Admin</b>, password <b>Admin@2026</b>. Your ID is recorded against every photo you upload.</p>',
   '<p class="gate-foot"><b>PPT team</b> (create folders, generate the PPT): sign in with your own name and your own password.<br><b>Photo upload</b>: ID <b>Admin</b>, password <b>Admin@2026</b>. Every sign-in is recorded in the access history.</p>'],
];

let app = fs.readFileSync(appPath, 'utf8');
if (app.includes('const AF=(function(){')) { log('already applied'); process.exit(0); }
let ok = 0;
for (const [from, to] of PATCHES) {
  const count = app.split(from).length - 1;
  if (count !== 1) { log('NOT applied (found ' + count + ' times): ' + from.slice(0, 70)); continue; }
  app = app.replace(from, () => to); ok++;
}
if (ok !== PATCHES.length) { log('Only ' + ok + ' of ' + PATCHES.length + ' changes fit the front end. Stopping the build.'); process.exit(1); }
// the front end is one closure: the module goes inside it, just before it closes
const end = app.lastIndexOf('})();');
if (end < 0) { log('The end of the front end was not found. Stopping the build.'); process.exit(1); }
app = app.slice(0, end) + fs.readFileSync(path.join(here, 'autofill-client.js'), 'utf8') + '\n' + app.slice(end);
fs.writeFileSync(appPath, app);
log(ok + ' of ' + PATCHES.length + ' changes applied');
