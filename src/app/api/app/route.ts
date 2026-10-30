import { NextRequest, NextResponse } from 'next/server';
import * as XLSX from 'xlsx';
import {
  APP_COOKIE,
  adminPassword,
  baseUrl,
  nowMs,
  sb,
  signSession,
  verifySession,
  zoneLabel,
} from '@/lib/core';

export const runtime = 'nodejs';
export const dynamic = 'force-dynamic';

const MAX_UPLOAD = 12 * 1024 * 1024;
const BUCKET = 'photos';

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
  for (const s of sites) a = Math.max(a, Number(s.createdAt) || 0);
  for (const p of photos) b = Math.max(b, Number(p.uploadedAt) || 0);
  return sites.length + '-' + photos.length + '-' + a + '-' + b;
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

  if (action === 'login') {
    const id = String(payload.id || '').trim().replace(/\s+/g, ' ');
    const password = String(payload.password || '');
    if (!id) return json({ ok: false, error: 'Enter your login ID.' }, 400);
    if (password !== adminPassword()) {
      return json(
        { ok: false, error: 'Wrong access code or session expired. Please sign in again.' },
        401,
      );
    }
    const sess = {
      uid: id,
      name: id,
      role: 'admin',
      exp: Date.now() + 1000 * 60 * 60 * 24 * 14,
    };
    try {
      await sb()
        .from('app_users')
        .upsert(
          { login_id: id, display_name: id, role: 'admin', last_seen: nowMs() },
          { onConflict: 'login_id' },
        );
    } catch {
      // the user record is a convenience, the session still works without it
    }
    return json({ ok: true, user: { id, name: id, role: 'admin' } }, 200, signSession(sess));
  }

  if (action === 'session') {
    const sess = verifySession(req.cookies.get(APP_COOKIE)?.value);
    if (!sess) return json({ ok: false });
    return json({ ok: true, user: { id: sess.name, name: sess.name, role: sess.role || 'admin' } });
  }

  if (action === 'logout') {
    const res = NextResponse.json({ ok: true });
    res.cookies.set(APP_COOKIE, '', { httpOnly: true, path: '/', maxAge: 0 });
    return res;
  }

  let session = verifySession(req.cookies.get(APP_COOKIE)?.value);
  let freshCookie: string | undefined;
  if (!session && code && code === adminPassword()) {
    const id = String(body.user || 'Admin');
    session = {
      uid: id,
      name: id,
      role: 'admin',
      exp: Date.now() + 1000 * 60 * 60 * 24 * 14,
    };
    freshCookie = signSession(session);
  }
  if (!session) {
    return json(
      { ok: false, error: 'Wrong access code or session expired. Please sign in again.' },
      401,
    );
  }

  try {
    switch (action) {
      case 'list':
        return json(await list(payload), 200, freshCookie);
      case 'createSite':
        return json(await createSite(payload), 200, freshCookie);
      case 'updateSite':
        return json(await updateSite(payload), 200, freshCookie);
      case 'setCaption':
        return json(await setCaption(payload), 200, freshCookie);
      case 'setSubfolders':
        return json(await setSubfolders(payload), 200, freshCookie);
      case 'deleteSite':
        return json(await deleteSite(payload), 200, freshCookie);
      case 'upload':
        return json(await upload(payload, session), 200, freshCookie);
      case 'deletePhoto':
        return json(await deletePhoto(payload), 200, freshCookie);
      case 'getPhoto':
        return json(await getPhoto(payload), 200, freshCookie);
      default:
        return json({ ok: false, error: 'Unknown action: ' + action }, 400);
    }
  } catch (e: any) {
    return json({ ok: false, error: e?.message || 'Something went wrong' }, 500);
  }
}

async function list(payload: any) {
  const s = sb();
  const siteRes = await s.from('sites').select('*').order('created_at', { ascending: true });
  const photoRes = await s.from('photos').select('*').order('uploaded_at', { ascending: true });
  if (siteRes.error) throw new Error(siteRes.error.message);
  const sites = (siteRes.data || []).map(mapSite);
  const photos = (photoRes.data || []).map(mapPhoto);
  const ver = version(sites, photos);
  const since = payload && payload.since ? String(payload.since) : '';
  if (since && since === ver) {
    return { ok: true, ver, same: true, sites: [], photos: [], role: 'admin' };
  }
  return { ok: true, ver, same: false, sites, photos, role: 'admin' };
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

async function updateSite(payload: any) {
  const site = payload.site || {};
  const patch: any = { updated_at: nowMs() };
  const fields = ['name', 'state', 'region', 'address', 'latlong', 'layout'];
  for (const key of fields) {
    if (site[key] !== undefined) patch[key] = String(site[key]);
  }
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

async function setSubfolders(payload: any) {
  const { error } = await sb()
    .from('sites')
    .update({ subfolders: payload.subfolders || [], updated_at: nowMs() })
    .eq('id', String(payload.id));
  if (error) throw new Error(error.message);
  return { ok: true };
}

async function deleteSite(payload: any) {
  const id = String(payload.id);
  const { data: photos } = await sb().from('photos').select('file_id').eq('site_id', id);
  const paths = (photos || []).map((p: any) => p.file_id).filter(Boolean);
  if (paths.length) await sb().storage.from(BUCKET).remove(paths);
  const { error } = await sb().from('sites').delete().eq('id', id);
  if (error) throw new Error(error.message);
  return { ok: true };
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
