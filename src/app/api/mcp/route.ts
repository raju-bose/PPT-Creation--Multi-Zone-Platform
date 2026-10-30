import { NextRequest, NextResponse } from 'next/server';
import {
  APP_COOKIE,
  ZONES,
  baseUrl,
  folderList,
  loadRules,
  nowMs,
  sb,
  verifySession,
  zoneLabel,
} from '@/lib/core';

export const runtime = 'nodejs';
export const dynamic = 'force-dynamic';

const PROTOCOL_VERSION = '2024-11-05';
const SERVER_INFO = { name: 'cai-photo-platform', version: '1.0.0' };

function mcpToken(): string {
  return process.env.MCP_TOKEN || '';
}

function authorized(req: NextRequest): boolean {
  const raw = req.headers.get('x-mcp-token') || req.headers.get('authorization') || '';
  const header = raw.replace(/^Bearer\s+/i, '').trim();
  const token = mcpToken();
  if (token && header && header === token) return true;
  return !!verifySession(req.cookies.get(APP_COOKIE)?.value);
}

const TOOLS = [
  {
    name: 'list_zones',
    description:
      'List the zones this platform serves, with the states covered by each zone.',
    inputSchema: { type: 'object', properties: {} },
  },
  {
    name: 'list_sites',
    description:
      'List every site folder with its upload progress. Pass region as north, east or all.',
    inputSchema: {
      type: 'object',
      properties: {
        region: { type: 'string', description: 'north, east or all. Defaults to all.' },
      },
    },
  },
  {
    name: 'get_site',
    description:
      'Full detail for one site: every sub folder, which ones are filled, who uploaded them and what is still open.',
    inputSchema: {
      type: 'object',
      properties: { code: { type: 'string', description: 'Site code, for example NRO-001' } },
      required: ['code'],
    },
  },
  {
    name: 'get_checklist',
    description:
      'Checklist rows for the PPT pack, one row per site folder, with progress and the folders still missing photos.',
    inputSchema: {
      type: 'object',
      properties: {
        region: { type: 'string', description: 'north, east or all.' },
        status: { type: 'string', description: 'pending, complete or all. Defaults to all.' },
      },
    },
  },
  {
    name: 'create_site',
    description:
      'Create a site folder with the standard sub folders for its layout (34 folders for 1DU2T and 2DU2T, 45 for 3DU3T).',
    inputSchema: {
      type: 'object',
      properties: {
        code: { type: 'string', description: 'Site code, for example NRO-001' },
        name: { type: 'string' },
        state: { type: 'string' },
        region: { type: 'string', description: 'north or east' },
        layout: { type: 'string', description: '1DU2T, 2DU2T or 3DU3T. Defaults to 2DU2T.' },
        address: { type: 'string' },
        latlong: { type: 'string' },
      },
      required: ['code', 'region'],
    },
  },
  {
    name: 'upload_photo',
    description:
      'Store a photo into one of a site sub folders. Give either file_url or file_base64, plus the folder name.',
    inputSchema: {
      type: 'object',
      properties: {
        site_code: { type: 'string' },
        folder: { type: 'string', description: 'Sub folder name, for example 05 Product Equipment' },
        file_name: { type: 'string' },
        file_url: { type: 'string', description: 'Public or signed URL of the photo to pull in.' },
        file_base64: { type: 'string', description: 'Base64 encoded image, used when there is no URL.' },
        content_type: { type: 'string' },
      },
      required: ['site_code', 'folder'],
    },
  },
  {
    name: 'list_photos',
    description:
      'Photos stored for a site, with signed links so the images can actually be viewed.',
    inputSchema: {
      type: 'object',
      properties: {
        site_code: { type: 'string' },
        folder: { type: 'string', description: 'Optional sub folder name to narrow the list.' },
      },
      required: ['site_code'],
    },
  },
  {
    name: 'pending_uploads',
    description: 'Site folders still missing photos, least complete first.',
    inputSchema: {
      type: 'object',
      properties: {
        region: { type: 'string' },
        limit: { type: 'number', description: 'How many rows to return. Defaults to 25.' },
      },
    },
  },
  {
    name: 'export_excel',
    description: 'Build the Excel progress report for a zone and return a download link.',
    inputSchema: {
      type: 'object',
      properties: { region: { type: 'string', description: 'north, east or all.' } },
    },
  },
];

async function loadAll() {
  const s = sb();
  const siteRes = await s.from('sites').select('*').order('created_at', { ascending: true });
  const photoRes = await s.from('photos').select('*').order('uploaded_at', { ascending: true });
  if (siteRes.error) throw new Error(siteRes.error.message);
  const sites = (siteRes.data || []).map((r: any) => ({
    id: r.id,
    code: r.code,
    name: r.name,
    state: r.state,
    region: r.region || 'east',
    layout: r.layout,
    address: r.address,
    latlong: r.latlong,
    subfolders: r.subfolders || [],
    createdAt: r.created_at,
  }));
  const photos = (photoRes.data || []).map((r: any) => ({
    id: r.id,
    siteId: r.site_id,
    folderId: r.folder_id,
    fileId: r.file_id,
    fileName: r.file_name,
    uploadedAt: r.uploaded_at,
    uploadedBy: r.uploaded_by,
    caption: r.caption,
  }));
  return { sites, photos };
}

function progress(site: any, photos: any[]) {
  const mine = photos.filter((p) => p.siteId === site.id);
  const done = new Set(mine.map((p) => p.folderId));
  const total = (site.subfolders || []).length;
  const folders = (site.subfolders || []).map((f: any) => ({
    id: f.id,
    name: f.name,
    uploaded: done.has(f.id),
  }));
  return {
    siteCode: site.code,
    siteName: site.name,
    state: site.state,
    region: site.region,
    zone: zoneLabel(site.region),
    layout: site.layout,
    folders: total,
    uploaded: done.size,
    pending: Math.max(0, total - done.size),
    progressPct: total ? Math.round((done.size / total) * 100) : 0,
    lastUpload: mine.length ? Math.max(...mine.map((p) => p.uploadedAt)) : 0,
    status: total && done.size >= total ? 'complete' : 'pending',
    openFolders: folders.filter((f) => !f.uploaded).map((f) => f.name),
  };
}

async function callTool(name: string, args: any, req: NextRequest) {
  const a = args || {};
  const region = String(a.region || 'all');
  const { sites, photos } = await loadAll();
  const scoped = sites.filter((s: any) => region === 'all' || s.region === region);

  if (name === 'list_zones') {
    return {
      zones: Object.keys(ZONES).map((k) => ({
        id: k,
        label: ZONES[k].label,
        states: ZONES[k].states,
      })),
    };
  }

  if (name === 'list_sites') {
    return { sites: scoped.map((s: any) => progress(s, photos)) };
  }

  if (name === 'get_checklist') {
    const status = String(a.status || 'all');
    let rows = scoped.map((s: any) => progress(s, photos));
    if (status !== 'all') rows = rows.filter((r: any) => r.status === status);
    return { rows };
  }

  if (name === 'get_site') {
    const code = String(a.code || '').toLowerCase();
    const site = sites.find((s: any) => String(s.code).toLowerCase() === code);
    if (!site) throw new Error('No site found with code ' + a.code);
    const mine = photos.filter((p) => p.siteId === site.id);
    const done = new Set(mine.map((p) => p.folderId));
    return {
      ...progress(site, photos),
      address: site.address,
      latlong: site.latlong,
      createdAt: site.createdAt,
      folderList: (site.subfolders || []).map((f: any) => ({
        id: f.id,
        name: f.name,
        uploaded: done.has(f.id),
      })),
      photos: mine.map((p: any) => ({
        id: p.id,
        folder: p.folderId,
        fileName: p.fileName,
        uploadedBy: p.uploadedBy,
        uploadedAt: p.uploadedAt,
        caption: p.caption,
      })),
    };
  }

  if (name === 'create_site') {
    const code = String(a.code || '').trim();
    if (!code) throw new Error('A site code is required');
    const layout = String(a.layout || '2DU2T');
    const rules = await loadRules(req);
    const folders = folderList(rules, layout);
    const row = {
      code,
      name: String(a.name || code),
      state: String(a.state || ''),
      region: String(a.region || 'east'),
      address: String(a.address || ''),
      latlong: String(a.latlong || ''),
      layout,
      subfolders: folders,
      created_at: nowMs(),
      updated_at: nowMs(),
    };
    const { data, error } = await sb().from('sites').insert(row).select().single();
    if (error) throw new Error(error.message);
    return {
      ok: true,
      siteCode: data.code,
      region: data.region,
      layout,
      folders: (data.subfolders || []).length,
    };
  }

  if (name === 'upload_photo') {
    const siteCode = String(a.site_code || '').toLowerCase();
    const site = sites.find((s: any) => String(s.code).toLowerCase() === siteCode);
    if (!site) throw new Error('No site found with code ' + a.site_code);
    const folderName = String(a.folder || '').trim();
    const folder = (site.subfolders || []).find(
      (f: any) => String(f.name).toLowerCase() === folderName.toLowerCase(),
    );
    if (!folder) {
      throw new Error(
        'No sub folder called ' + folderName + '. Available: ' +
          (site.subfolders || []).map((f: any) => f.name).join(', '),
      );
    }
    let buf: Buffer;
    let contentType = String(a.content_type || 'image/jpeg');
    if (a.file_url) {
      const res = await fetch(String(a.file_url));
      if (!res.ok) throw new Error('Could not fetch that URL (' + res.status + ')');
      buf = Buffer.from(await res.arrayBuffer());
      contentType = res.headers.get('content-type') || contentType;
    } else if (a.file_base64) {
      const raw = String(a.file_base64);
      buf = Buffer.from(raw.indexOf(',') >= 0 ? raw.split(',').pop() || '' : raw, 'base64');
    } else {
      throw new Error('Give either file_url or file_base64');
    }
    if (!buf.length) throw new Error('That file came back empty');
    const fileName = String(a.file_name || 'photo.jpg').replace(/[^A-Za-z0-9._-]/g, '_');
    const path =
      site.id + '/' + folder.id + '/' + Date.now() + '-' + Math.random().toString(36).slice(2, 8) + '-' + fileName;
    const up = await sb().storage.from('photos').upload(path, buf, { contentType, upsert: true });
    if (up.error) throw new Error(up.error.message);
    const { data, error } = await sb()
      .from('photos')
      .insert({
        site_id: site.id,
        folder_id: folder.id,
        file_id: path,
        file_name: String(a.file_name || 'photo.jpg'),
        mime: contentType,
        size_bytes: buf.length,
        uploaded_at: nowMs(),
        uploaded_by: 'claude-mcp',
        caption: '',
      })
      .select()
      .single();
    if (error) throw new Error(error.message);
    return {
      ok: true,
      siteCode: site.code,
      folder: folder.name,
      fileName: data.file_name,
      sizeBytes: data.size_bytes,
    };
  }

  if (name === 'list_photos') {
    const siteCode = String(a.site_code || '').toLowerCase();
    const site = sites.find((s: any) => String(s.code).toLowerCase() === siteCode);
    if (!site) throw new Error('No site found with code ' + a.site_code);
    let mine = photos.filter((p) => p.siteId === site.id);
    if (a.folder) {
      const folder = (site.subfolders || []).find(
        (f: any) => String(f.name).toLowerCase() === String(a.folder).toLowerCase(),
      );
      if (folder) mine = mine.filter((p) => p.folderId === folder.id);
    }
    const out = [];
    for (const p of mine) {
      const signed = await sb().storage.from('photos').createSignedUrl(p.fileId, 3600);
      out.push({
        folder: p.folderId,
        fileName: p.fileName,
        uploadedBy: p.uploadedBy,
        uploadedAt: p.uploadedAt,
        url: signed.data ? signed.data.signedUrl : null,
      });
    }
    return { siteCode: site.code, photos: out };
  }

  if (name === 'pending_uploads') {
    const limit = Number(a.limit || 25);
    const rows = scoped.map((s: any) => progress(s, photos));
    rows.sort((x: any, y: any) => x.progressPct - y.progressPct);
    return { rows: rows.slice(0, limit) };
  }

  if (name === 'export_excel') {
    const token = mcpToken();
    const url =
      baseUrl(req) + '/api/app?export=excel&region=' + encodeURIComponent(region) +
      (token ? '&token=' + encodeURIComponent(token) : '');
    return { url, region };
  }

  throw new Error('Unknown tool: ' + name);
}

export async function POST(req: NextRequest) {
  if (!authorized(req)) {
    return NextResponse.json(
      { jsonrpc: '2.0', id: null, error: { code: -32001, message: 'Unauthorized' } },
      { status: 401 },
    );
  }

  let body: any = null;
  try {
    body = await req.json();
  } catch {
    body = null;
  }
  if (!body) {
    return NextResponse.json(
      { jsonrpc: '2.0', id: null, error: { code: -32700, message: 'Parse error' } },
      { status: 400 },
    );
  }

  const id = body.id === undefined ? null : body.id;
  const method = String(body.method || '');

  if (method.indexOf('notifications/') === 0) {
    return new NextResponse(null, { status: 202 });
  }

  try {
    let result: any = null;
    if (method === 'initialize') {
      result = {
        protocolVersion: PROTOCOL_VERSION,
        capabilities: { tools: { listChanged: false } },
        serverInfo: SERVER_INFO,
      };
    } else if (method === 'ping') {
      result = {};
    } else if (method === 'tools/list') {
      result = { tools: TOOLS };
    } else if (method === 'tools/call') {
      const name = String((body.params && body.params.name) || '');
      const args = (body.params && body.params.arguments) || {};
      const data = await callTool(name, args, req);
      result = {
        content: [{ type: 'text', text: JSON.stringify(data, null, 2) }],
        isError: false,
      };
    } else {
      return NextResponse.json(
        { jsonrpc: '2.0', id, error: { code: -32601, message: 'Method not found: ' + method } },
        { status: 200 },
      );
    }
    return NextResponse.json({ jsonrpc: '2.0', id, result });
  } catch (e: any) {
    return NextResponse.json({
      jsonrpc: '2.0',
      id,
      result: {
        content: [{ type: 'text', text: String(e?.message || 'Tool failed') }],
        isError: true,
      },
    });
  }
}

export async function GET() {
  return NextResponse.json(
    {
      name: SERVER_INFO.name,
      version: SERVER_INFO.version,
      protocolVersion: PROTOCOL_VERSION,
      transport: 'streamable-http',
      endpoint: '/api/mcp',
      tools: TOOLS.map((t) => t.name),
    },
    { status: 200 },
  );
}
