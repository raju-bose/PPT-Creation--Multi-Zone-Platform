import { NextRequest, NextResponse } from 'next/server';
import * as XLSX from 'xlsx';
import {
  APP_COOKIE,
  SESSION_VERSION,
  adminPassword,
  hashPassword,
  loginKey,
  newPassword,
  nowMs,
  sb,
  signSession,
  verifyPassword,
  verifySession,
  zoneLabel,
} from '@/lib/core';

export const runtime = 'nodejs';
export const dynamic = 'force-dynamic';

const MAX_UPLOAD = 12 * 1024 * 1024;
const BUCKET = 'photos';
const SESSION_MS = 1000 * 60 * 60 * 24 * 14;
const VISIT_GAP_MS = 1000 * 60 * 30; // one "opened the website" entry per user per half hour
const MAX_FAILED_LOGINS = 8; // per login ID in 15 minutes

/** Roles: "ppt" = PPT team (own password, can create folders and generate the PPT); "uploader" = shared Admin login (photo upload only). */
type Role = 'ppt' | 'uploader';

function json(data: any, status = 200, cookie?: string) {
  const res = NextResponse.json(data, { status });
  if (cookie) {
    res.cookies.set(APP_COOKIE, cookie, {
      httpOnly: true,
      sameSite: 'lax',
      path: '/',
      maxAge: 60 * 60 * 24 * 14,
    });
  }
  return res;
}

function mapSite(row: any) {
  return {
    id: row.id,
    code: row.code || '',
    name: row.name || '',
    state: row.state || '',
    region: row.region || 'east',
    address: row.address || '',
    latlong: row.latlong || '',
    layout: row.layout || '',
    subfolders: row.subfolders || [],
    createdAt: row.created_at || 0,
    updatedAt: row.updated_at || 0,
    dprStart: row.dpr_start || '',
    pptCreatedBy: row.ppt_created_by || '',
    pptCreatedAt: row.ppt_created_at || 0,
    pptUpdatedBy: row.ppt_updated_by || '',
    pptUpdatedAt: row.ppt_updated_at || 0,
  };
}

function mapPhoto(row: any) {
  return {
    id: row.id,
    siteId: row.site_id,
    folderId: row.folder_id,
    fileId: row.file_id,
    fileName: row.file_name || '',
    sizeBytes: row.size_bytes || 0,
    uploadedAt: row.uploaded_at || 0,
    uploadedBy: row.uploaded_by || '',
    caption: row.caption || '',
  };
}

function version(sites: any[], photos: any[]) {
  let a = 0;
  let b = 0;
  for (const s of sites) a = Math.max(a, Number(s.createdAt) || 0, Number(s.updatedAt) || 0);
  for (const p of photos) b = Math.max(b, Number(p.uploadedAt) || 0);
  return sites.length + '-' + photos.length + '-' + a + '-' + b;
}

/* ---------- access log ---------- */

function clientIp(req: NextRequest): string {
  const fwd = req.headers.get('x-forwarded-for') || '';
  return (fwd.split(',')[0] || req.headers.get('x-real-ip') || '').trim().slice(0, 60);
}

function device(req: NextRequest): string {
  const ua = req.headers.get('user-agent') || '';
  const os = /android/i.test(ua)
    ? 'Android'
    : /iphone|ipad|ios/i.test(ua)
      ? 'iPhone / iPad'
      : /windows/i.test(ua)
        ? 'Windows'
        : /mac os/i.test(ua)
          ? 'Mac'
          : /linux/i.test(ua)
            ? 'Linux'
            : 'Other';
  const browser = /edg\//i.test(ua)
    ? 'Edge'
    : /chrome|crios/i.test(ua)
      ? 'Chrome'
      : /firefox|fxios/i.test(ua)
        ? 'Firefox'
        : /safari/i.test(ua)
          ? 'Safari'
          : 'Browser';
  return os + ' · ' + browser;
}

async function logAccess(req: NextRequest, login: string, role: string, action: string, detail = '') {
  try {
    await sb()
      .from('access_log')
      .insert({
        login_id: String(login || '').slice(0, 80),
        login_key: loginKey(login),
        role,
        action,
        detail: String(detail || '').slice(0, 300),
        ip: clientIp(req),
        device: device(req),
        at: nowMs(),
      });
  } catch {
    // the log must never stop the user's action
  }
}

async function touchUser(login: string, role: Role) {
  try {
    if (role === 'ppt') {
      await sb().from('app_users').update({ last_seen: nowMs() }).eq('login_id', login);
    } else {
      await sb()
        .from('app_users')
        .upsert(
          { login_id: login, display_name: login, role: 'uploader', last_seen: nowMs() },
          { onConflict: 'login_id' },
        );
    }
  } catch {
    // convenience record only
  }
}

/** The PPT team member this login ID belongs to, or null. */
async function findPptUser(id: string): Promise<any | null> {
  const key = loginKey(id);
  if (!key) return null;
  const { data, error } = await sb().from('app_users').select('*').eq('role', 'ppt');
  if (error) throw new Error('The user list could not be read. Please try again.');
  for (const u of data || []) {
    if (loginKey(u.login_id) === key) return u;
    if (Array.isArray(u.aliases) && u.aliases.some((a: string) => loginKey(a) === key)) return u;
  }
  return null;
}

function makeSession(name: string, role: Role, owner: boolean) {
  return { uid: name, name, role, owner, v: SESSION_VERSION, exp: Date.now() + SESSION_MS };
}

function userOut(sess: any) {
  return { id: sess.name, name: sess.name, role: sess.role === 'ppt' ? 'admin' : 'user', owner: !!sess.owner };
}

export async function GET(req: NextRequest) {
  const url = req.nextUrl.searchParams;
  const file = url.get('f');
  if (file) {
    const { data, error } = await sb().storage.from(BUCKET).createSignedUrl(file, 3600);
    if (error || !data) return json({ ok: false, error: 'That file could not be found' }, 404);
    return NextResponse.redirect(data.signedUrl);
  }

  if (url.get('export') === 'excel') {
    const token = url.get('token') || '';
    const mcpOk = !!process.env.MCP_TOKEN && token === process.env.MCP_TOKEN;
    if (!mcpOk && !verifySession(req.cookies.get(APP_COOKIE)?.value)) {
      return json({ ok: false, error: 'Please sign in again.' }, 401);
    }
    return exportExcel(url.get('region') || 'all');
  }

  return json({ ok: false, error: 'Not found' }, 404);
}

async function exportExcel(region: string) {
  const s = sb();
  const siteRes = await s.from('sites').select('*').order('code', { ascending: true });
  const photoRes = await s.from('photos').select('site_id, folder_id');
  const sites = (siteRes.data || []).map(mapSite);
  const photos = photoRes.data || [];

  const rows = sites
    .filter((site: any) => region === 'all' || site.region === region)
    .map((site: any) => {
      const done = new Set(
        photos.filter((p: any) => p.site_id === site.id).map((p: any) => p.folder_id),
      );
      const total = (site.subfolders || []).length;
      const filled = done.size;
      return {
        Zone: zoneLabel(site.region),
        'Site code': site.code,
        'Site name': site.name,
        State: site.state,
        Layout: site.layout,
        Folders: total,
        Uploaded: filled,
        Pending: Math.max(0, total - filled),
        'Progress %': total ? Math.round((filled / total) * 100) : 0,
      };
    });

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(
    rows.length ? rows : [{ Zone: 'No sites yet', 'Site code': '', 'Site name': '' }],
  );
  XLSX.utils.book_append_sheet(wb, ws, 'Site progress');
  const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });

  return new NextResponse(new Uint8Array(buf), {
    status: 200,
    headers: {
      'content-type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'content-disposition': 'attachment; filename="site-progress-' + region + '.xlsx"',
    },
  });
}

async function login(req: NextRequest, payload: any) {
  const id = String(payload.id || '').trim().replace(/\s+/g, ' ').slice(0, 60);
  const password = String(payload.password || '');
  if (!id) return json({ ok: false, error: 'Enter your login ID.' }, 400);

  // slow down password guessing on one login ID
  try {
    const { count } = await sb()
      .from('access_log')
      .select('id', { count: 'exact', head: true })
      .eq('action', 'login_failed')
      .eq('login_key', loginKey(id))
      .gt('at', nowMs() - 15 * 60 * 1000);
    if ((count || 0) >= MAX_FAILED_LOGINS) {
      return json({ ok: false, error: 'Too many wrong passwords. Please wait 15 minutes and try again.' }, 429);
    }
  } catch {
    // if the count cannot be read the normal check still applies
  }

  let member: any = null;
  try {
    member = await findPptUser(id);
  } catch (e: any) {
    return json({ ok: false, error: e?.message || 'Sign in is not available right now.' }, 503);
  }

  if (member) {
    // PPT team: own login ID, own password
    if (member.active === false) {
      await logAccess(req, member.login_id, 'ppt', 'login_failed', 'Login is switched off');
      return json({ ok: false, error: 'This login is switched off. Ask Raju Bose.' }, 403);
    }
    if (!verifyPassword(password, member.pass_hash)) {
      await logAccess(req, member.login_id, 'ppt', 'login_failed', 'Wrong password');
      return json(
        { ok: false, error: 'Wrong password for ' + member.login_id + '. PPT team members sign in with their own password.' },
        401,
      );
    }
    const sess = makeSession(member.login_id, 'ppt', !!member.is_owner);
    await touchUser(member.login_id, 'ppt');
    await logAccess(req, member.login_id, 'ppt', 'login', 'Signed in');
    return json({ ok: true, user: userOut(sess) }, 200, signSession(sess));
  }

  // everyone else (engineers): their own name with the shared password, so the history shows who it was
  if (/^(admin|administrator|user|test|engineer|guest)$/.test(loginKey(id)) || loginKey(id).length < 3) {
    return json({ ok: false, error: 'Enter your own name (for example “Suresh Patil”), not “' + id + '”. The password stays Admin@2026.' }, 400);
  }
  if (password !== adminPassword()) {
    await logAccess(req, id, 'uploader', 'login_failed', 'Wrong password');
    return json({ ok: false, error: 'Wrong access code or session expired. Please sign in again.' }, 401);
  }
  const sess = makeSession(id, 'uploader', false);
  await touchUser(id, 'uploader');
  await logAccess(req, id, 'uploader', 'login', 'Signed in (engineer)');
  return json({ ok: true, user: userOut(sess) }, 200, signSession(sess));
}

export async function POST(req: NextRequest) {
  let body: any = {};
  try {
    const raw = await req.text();
    body = raw ? JSON.parse(raw) : {};
  } catch {
    return json({ ok: false, error: 'Bad request' }, 400);
  }

  const action = String(body.action || '');
  const payload = body.payload && typeof body.payload === 'object' ? body.payload : body;
  const code = String(body.code || '');

  if (action === 'login') return login(req, payload);

  if (action === 'session') {
    const sess = verifySession(req.cookies.get(APP_COOKIE)?.value);
    if (!sess) return json({ ok: false });
    // "opened the website" – at most one entry per user per half hour
    try {
      const { data } = await sb()
        .from('access_log')
        .select('at')
        .eq('login_key', loginKey(sess.name))
        .in('action', ['login', 'open'])
        .order('at', { ascending: false })
        .limit(1);
      const last = data && data[0] ? Number(data[0].at) : 0;
      if (nowMs() - last > VISIT_GAP_MS) await logAccess(req, sess.name, sess.role, 'open', 'Opened the website');
    } catch {
      // ignore
    }
    await touchUser(sess.name, sess.role);
    return json({ ok: true, user: userOut(sess) });
  }

  if (action === 'logout') {
    const sess = verifySession(req.cookies.get(APP_COOKIE)?.value);
    if (sess) await logAccess(req, sess.name, sess.role, 'logout', 'Signed out');
    const res = NextResponse.json({ ok: true });
    res.cookies.set(APP_COOKIE, '', { httpOnly: true, path: '/', maxAge: 0 });
    return res;
  }

  let session = verifySession(req.cookies.get(APP_COOKIE)?.value);
  let freshCookie: string | undefined;
  if (!session && code && code === adminPassword()) {
    // older clients that send the shared code with every call: photo upload only, never a PPT team name
    const id = String(body.user || 'Admin').trim().slice(0, 60) || 'Admin';
    let reserved = true;
    try {
      reserved = !!(await findPptUser(id));
    } catch {
      reserved = true;
    }
    if (!reserved) {
      session = makeSession(id, 'uploader', false);
      freshCookie = signSession(session);
    }
  }
  if (!session) {
    return json(
      { ok: false, error: 'Wrong access code or session expired. Please sign in again.' },
      401,
    );
  }

  const isPpt = session.role === 'ppt';
  const pptOnly = (what: string) =>
    json({ ok: false, error: 'Only the PPT team can ' + what + '. Sign in with your PPT team login.' }, 403);

  try {
    switch (action) {
      case 'list':
        return json(await list(payload, session), 200, freshCookie);
      case 'createSite': {
        // every signed-in user may create a site folder; generating the PPT stays with the PPT team
        const out = await createSite(payload);
        await logAccess(req, session.name, session.role, 'site_created', out.site.code + ' ' + out.site.name);
        return json(out, 200, freshCookie);
      }
      case 'updateSite':
        if (!isPpt) return pptOnly('edit site details');
        return json(await updateSite(payload), 200, freshCookie);
      case 'setCaption':
        return json(await setCaption(payload), 200, freshCookie);
      case 'setSubfolders':
        return json(await setSubfolders(payload, isPpt), 200, freshCookie);
      case 'deleteSite': {
        if (!isPpt) return pptOnly('delete site folders');
        const out = await deleteSite(payload);
        await logAccess(req, session.name, session.role, 'site_deleted', out.label);
        return json({ ok: true }, 200, freshCookie);
      }
      case 'upload': {
        const out = await upload(payload, session);
        await touchUser(session.name, session.role);
        return json(out, 200, freshCookie);
      }
      case 'deletePhoto':
        return json(await deletePhoto(payload), 200, freshCookie);
      case 'getPhoto':
        return json(await getPhoto(payload), 200, freshCookie);
      case 'pptGenerated': {
        if (!isPpt) return pptOnly('generate the PPT');
        const out = await pptGenerated(payload, session);
        await touchUser(session.name, session.role);
        await logAccess(req, session.name, session.role, 'ppt', 'Generated the PPT for ' + out.label);
        return json({ ok: true, meta: out.meta }, 200, freshCookie);
      }
      case 'history':
        if (!isPpt) return pptOnly('see the access history');
        return json(await history(Number(payload.limit) || 500), 200, freshCookie);
      case 'createUser': {
        if (!isPpt || !session.owner) {
          return json({ ok: false, error: 'Only Raju Bose can create a login.' }, 403);
        }
        const out = await createUser(payload);
        await logAccess(req, session.name, session.role, 'user', 'Created the PPT team login ' + out.login);
        return json({ ok: true, login: out.login, password: out.password }, 200, freshCookie);
      }
      case 'setUserActive': {
        if (!isPpt || !session.owner) {
          return json({ ok: false, error: 'Only Raju Bose can switch a login on or off.' }, 403);
        }
        const out = await setUserActive(payload, session);
        await logAccess(req, session.name, session.role, 'user', (out.active ? 'Switched on' : 'Switched off') + ' the login ' + out.login);
        return json({ ok: true }, 200, freshCookie);
      }
      case 'changePassword': {
        if (!isPpt) return pptOnly('change a password');
        await changePassword(payload, session);
        await logAccess(req, session.name, session.role, 'password', 'Changed own password');
        return json({ ok: true }, 200, freshCookie);
      }
      case 'resetPassword': {
        if (!isPpt || !session.owner) {
          return json({ ok: false, error: 'Only Raju Bose can reset a password.' }, 403);
        }
        const out = await resetPassword(payload);
        await logAccess(req, session.name, session.role, 'password', 'Reset the password of ' + out.login);
        return json({ ok: true, login: out.login, password: out.password }, 200, freshCookie);
      }
      default:
        return json({ ok: false, error: 'Unknown action: ' + action }, 400);
    }
  } catch (e: any) {
    return json({ ok: false, error: e?.message || 'Something went wrong' }, 500);
  }
}

async function list(payload: any, session: any) {
  const s = sb();
  const siteRes = await s.from('sites').select('*').order('created_at', { ascending: true });
  const photoRes = await s.from('photos').select('*').order('uploaded_at', { ascending: true });
  if (siteRes.error) throw new Error(siteRes.error.message);
  const sites = (siteRes.data || []).map(mapSite);
  const photos = (photoRes.data || []).map(mapPhoto);
  const ver = version(sites, photos);
  const since = payload && payload.since ? String(payload.since) : '';
  const who = { role: session.role === 'ppt' ? 'admin' : 'user', owner: !!session.owner, me: session.name };
  if (since && since === ver) {
    return { ok: true, ver, same: true, sites: [], photos: [], ...who };
  }
  return { ok: true, ver, same: false, sites, photos, ...who };
}

async function createSite(payload: any) {
  const site = payload.site || payload;
  const code = String(site.code || '').trim();
  if (!code) throw new Error('Site code is required');
  const row = {
    code,
    name: String(site.name || code),
    state: String(site.state || ''),
    region: String(site.region || 'east'),
    address: String(site.address || ''),
    latlong: String(site.latlong || ''),
    layout: String(site.layout || ''),
    subfolders: Array.isArray(site.subfolders) ? site.subfolders : [],
    created_at: nowMs(),
    updated_at: nowMs(),
  };
  const { data, error } = await sb().from('sites').insert(row).select().single();
  if (error) throw new Error(error.message);
  return { ok: true, site: mapSite(data) };
}

function cleanDate(v: unknown): string {
  const s = String(v || '').trim();
  return /^\d{4}-\d{2}-\d{2}$/.test(s) && !isNaN(Date.parse(s + 'T00:00:00Z')) ? s : '';
}

async function updateSite(payload: any) {
  const site = payload.site || {};
  const patch: any = { updated_at: nowMs() };
  const fields = ['name', 'state', 'region', 'address', 'latlong', 'layout'];
  for (const key of fields) {
    if (site[key] !== undefined) patch[key] = String(site[key]);
  }
  if (site.dprStart !== undefined) patch.dpr_start = cleanDate(site.dprStart);
  if (Array.isArray(site.subfolders)) patch.subfolders = site.subfolders;
  const { error } = await sb().from('sites').update(patch).eq('id', String(payload.id));
  if (error) throw new Error(error.message);
  return { ok: true };
}

async function setCaption(payload: any) {
  const { error } = await sb()
    .from('photos')
    .update({ caption: String(payload.caption || '') })
    .eq('id', String(payload.id));
  if (error) throw new Error(error.message);
  return { ok: true };
}

async function setSubfolders(payload: any, isPpt: boolean) {
  const next = Array.isArray(payload.subfolders) ? payload.subfolders : [];
  if (!isPpt) {
    // photo-upload logins may only add the missing PPT subfolders, never remove or rename one
    const { data } = await sb().from('sites').select('subfolders').eq('id', String(payload.id)).maybeSingle();
    const cur: any[] = (data && data.subfolders) || [];
    const kept = cur.every((f: any) => next.some((n: any) => n && n.id === f.id && n.name === f.name));
    if (!kept) throw new Error('Only the PPT team can remove or rename subfolders.');
  }
  const { error } = await sb()
    .from('sites')
    .update({ subfolders: next, updated_at: nowMs() })
    .eq('id', String(payload.id));
  if (error) throw new Error(error.message);
  return { ok: true };
}

async function deleteSite(payload: any) {
  const id = String(payload.id);
  const { data: site } = await sb().from('sites').select('code, name').eq('id', id).maybeSingle();
  const { data: photos } = await sb().from('photos').select('file_id').eq('site_id', id);
  const paths = (photos || []).map((p: any) => p.file_id).filter(Boolean);
  if (paths.length) await sb().storage.from(BUCKET).remove(paths);
  const { error } = await sb().from('sites').delete().eq('id', id);
  if (error) throw new Error(error.message);
  return { ok: true, label: site ? (site.code + ' ' + site.name).trim() : id };
}

async function upload(payload: any, session: any) {
  const siteId = String(payload.siteId || '');
  const folderId = String(payload.folderId || '');
  const fileName = String(payload.fileName || 'photo.jpg');
  const mime = String(payload.type || 'image/jpeg');
  if (!siteId || !folderId) throw new Error('Site and folder are required');
  const buf = Buffer.from(String(payload.data || ''), 'base64');
  if (buf.length > MAX_UPLOAD) throw new Error('That file is too large. The limit is 12 MB.');
  const safe = fileName.replace(/[^A-Za-z0-9._-]/g, '_');
  const path =
    siteId + '/' + folderId + '/' + Date.now() + '-' + Math.random().toString(36).slice(2, 8) + '-' + safe;
  const up = await sb().storage.from(BUCKET).upload(path, buf, { contentType: mime, upsert: true });
  if (up.error) throw new Error(up.error.message);
  const row = {
    site_id: siteId,
    folder_id: folderId,
    file_id: path,
    file_name: fileName,
    mime,
    size_bytes: buf.length,
    uploaded_at: nowMs(),
    uploaded_by: String(session.name || 'Admin'),
    caption: '',
  };
  const { data, error } = await sb().from('photos').insert(row).select().single();
  if (error) throw new Error(error.message);
  return { ok: true, photo: mapPhoto(data) };
}

async function deletePhoto(payload: any) {
  const id = String(payload.id);
  const { data } = await sb().from('photos').select('file_id').eq('id', id).maybeSingle();
  if (data && data.file_id) await sb().storage.from(BUCKET).remove([data.file_id]);
  const { error } = await sb().from('photos').delete().eq('id', id);
  if (error) throw new Error(error.message);
  return { ok: true };
}

async function getPhoto(payload: any) {
  const fileId = String(payload.fileId || '');
  const { data: row } = await sb().from('photos').select('mime').eq('file_id', fileId).maybeSingle();
  const { data, error } = await sb().storage.from(BUCKET).download(fileId);
  if (error || !data) throw new Error('That photo could not be found');
  const buf = Buffer.from(await data.arrayBuffer());
  return {
    ok: true,
    data: buf.toString('base64'),
    type: (row && row.mime) || 'image/jpeg',
  };
}

/**
 * Called each time a PPT team member generates the PPT. Keeps the page-2 history of the site:
 * who created the PPT first, and who generated it last and when.
 */
async function pptGenerated(payload: any, session: any) {
  const id = String(payload.id || '');
  const { data: site, error } = await sb().from('sites').select('*').eq('id', id).maybeSingle();
  if (error || !site) throw new Error('That site could not be found');
  const now = nowMs();
  const patch: any = {
    ppt_updated_by: session.name,
    ppt_updated_at: now,
    updated_at: now,
  };
  if (!site.ppt_created_by) {
    patch.ppt_created_by = session.name;
    patch.ppt_created_at = now;
  }
  const dpr = cleanDate(payload.dprStart);
  if (dpr) patch.dpr_start = dpr;
  const { data: saved, error: e2 } = await sb().from('sites').update(patch).eq('id', id).select().single();
  if (e2) throw new Error(e2.message);
  const m = mapSite(saved);
  return {
    label: (m.code + ' ' + m.name).trim(),
    meta: {
      dprStart: m.dprStart,
      pptCreatedBy: m.pptCreatedBy,
      pptCreatedAt: m.pptCreatedAt,
      pptUpdatedBy: m.pptUpdatedBy,
      pptUpdatedAt: m.pptUpdatedAt,
      updatedAt: m.updatedAt,
    },
  };
}

async function createUser(payload: any) {
  const name = String(payload.name || '').trim().replace(/\s+/g, ' ').slice(0, 60);
  if (name.length < 3) throw new Error('Enter the full name for the login (at least 3 letters).');
  if (loginKey(name) === 'admin') throw new Error('“Admin” is the shared photo-upload login. Choose the person’s own name.');
  if (await findPptUser(name)) throw new Error('A PPT team login with that name already exists.');
  const password = newPassword();
  const { error } = await sb()
    .from('app_users')
    .upsert(
      { login_id: name, display_name: name, role: 'ppt', pass_hash: hashPassword(password), is_owner: false, active: true, last_seen: 0 },
      { onConflict: 'login_id' },
    );
  if (error) throw new Error(error.message);
  return { login: name, password };
}

async function setUserActive(payload: any, session: any) {
  const member = await findPptUser(String(payload.login || ''));
  if (!member) throw new Error('That PPT team login was not found.');
  if (member.login_id === session.name) throw new Error('You cannot switch off your own login.');
  const active = !!payload.active;
  const { error } = await sb().from('app_users').update({ active }).eq('login_id', member.login_id);
  if (error) throw new Error(error.message);
  return { login: member.login_id, active };
}

async function history(limit = 500) {
  const s = sb();
  const users = await s.from('app_users').select('login_id, role, is_owner, active, last_seen').order('last_seen', { ascending: false });
  const log = await s
    .from('access_log')
    .select('login_id, role, action, detail, ip, device, at')
    .order('at', { ascending: false })
    .limit(Math.max(1, Math.min(500, limit)));
  if (log.error) throw new Error(log.error.message);
  return {
    ok: true,
    users: (users.data || []).map((u: any) => ({
      login: u.login_id,
      role: u.role === 'ppt' ? 'PPT team' : 'Engineer',
      owner: !!u.is_owner,
      active: u.active !== false,
      lastSeen: Number(u.last_seen) || 0,
    })),
    log: (log.data || []).map((r: any) => ({
      login: r.login_id,
      role: r.role === 'ppt' ? 'PPT team' : 'Engineer',
      action: r.action,
      detail: r.detail || '',
      ip: r.ip || '',
      device: r.device || '',
      at: Number(r.at) || 0,
    })),
  };
}

async function changePassword(payload: any, session: any) {
  const current = String(payload.current || '');
  const next = String(payload.next || '');
  if (next.length < 8) throw new Error('The new password needs at least 8 characters.');
  if (next === adminPassword()) throw new Error('Choose a password different from the shared upload password.');
  const { data: user } = await sb().from('app_users').select('*').eq('login_id', session.name).maybeSingle();
  if (!user || !verifyPassword(current, user.pass_hash)) throw new Error('Your current password is not correct.');
  const { error } = await sb().from('app_users').update({ pass_hash: hashPassword(next) }).eq('login_id', session.name);
  if (error) throw new Error(error.message);
}

async function resetPassword(payload: any) {
  const member = await findPptUser(String(payload.login || ''));
  if (!member) throw new Error('That PPT team login was not found.');
  const password = newPassword();
  const { error } = await sb().from('app_users').update({ pass_hash: hashPassword(password) }).eq('login_id', member.login_id);
  if (error) throw new Error(error.message);
  return { login: member.login_id, password };
}
