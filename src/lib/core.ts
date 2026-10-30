import { createHmac, timingSafeEqual } from 'node:crypto';
import { createClient, type SupabaseClient } from '@supabase/supabase-js';

export const APP_COOKIE = 'cai_session';

export type ZoneId = 'north' | 'east';

export interface Zone {
  id: ZoneId;
  label: string;
  short: string;
  states: string[];
}

export const ZONES: Record<string, Zone> = {
  north: {
    id: 'north',
    label: 'North Zone',
    short: 'North',
    states: [
      'Punjab',
      'Haryana',
      'Himachal Pradesh',
      'Jammu & Kashmir',
      'Ladakh',
      'Rajasthan',
      'Uttar Pradesh',
      'Uttarakhand',
      'Delhi',
      'Chandigarh',
      'Madhya Pradesh',
      'Other',
    ],
  },
  east: {
    id: 'east',
    label: 'East Zone',
    short: 'East',
    states: [
      'Odisha',
      'West Bengal',
      'Bihar',
      'Jharkhand',
      'Assam',
      'Meghalaya',
      'Tripura',
      'Arunachal Pradesh',
      'Nagaland',
      'Manipur',
      'Mizoram',
      'Sikkim',
      'Other',
    ],
  },
};

export function zoneLabel(id: string): string {
  if (!id || id === 'all') return 'All zones';
  return ZONES[id] ? ZONES[id].label : String(id);
}

export function adminPassword(): string {
  return process.env.ADMIN_PASSWORD || 'Admin@2026';
}

export function sessionSecret(): string {
  return process.env.SESSION_SECRET || 'cai-platform-default-secret-change-me';
}

export function signSession(data: Record<string, unknown>): string {
  const payload = Buffer.from(JSON.stringify(data)).toString('base64url');
  const sig = createHmac('sha256', sessionSecret()).update(payload).digest('base64url');
  return payload + '.' + sig;
}

export function verifySession(token: string | undefined | null): any {
  if (!token) return null;
  const parts = token.split('.');
  if (parts.length !== 2) return null;
  const expected = createHmac('sha256', sessionSecret()).update(parts[0]).digest('base64url');
  const a = Buffer.from(parts[1]);
  const b = Buffer.from(expected);
  if (a.length !== b.length || !timingSafeEqual(a, b)) return null;
  try {
    const data = JSON.parse(Buffer.from(parts[0], 'base64url').toString('utf8'));
    if (data && data.exp && Date.now() > data.exp) return null;
    return data;
  } catch {
    return null;
  }
}

let client: SupabaseClient | null = null;

export function sb(): SupabaseClient {
  if (client) return client;
  const url = process.env.SUPABASE_URL;
  const key = process.env.SUPABASE_SERVICE_ROLE_KEY;
  if (!url || !key) {
    throw new Error('SUPABASE_URL and SUPABASE_SERVICE_ROLE_KEY are not configured');
  }
  client = createClient(url, key, {
    auth: { persistSession: false, autoRefreshToken: false },
  });
  return client;
}

export function nowMs(): number {
  return Date.now();
}

export function baseUrl(req: Request): string {
  if (process.env.APP_URL) return process.env.APP_URL.replace(/\/$/, '');
  const host = req.headers.get('x-forwarded-host') || req.headers.get('host') || 'localhost:3000';
  const proto = req.headers.get('x-forwarded-proto') || (host.indexOf('localhost') === 0 ? 'http' : 'https');
  return proto + '://' + host;
}

let rulesCache: any = null;
let templateCache: any = null;

export async function loadRules(req: Request): Promise<any> {
  if (rulesCache) return rulesCache;
  const res = await fetch(baseUrl(req) + '/assets/folders.json', { cache: 'no-store' });
  if (!res.ok) throw new Error('Folder rules are not available yet');
  rulesCache = await res.json();
  return rulesCache;
}

export async function loadTemplate(req: Request): Promise<any> {
  if (templateCache) return templateCache;
  const res = await fetch(baseUrl(req) + '/assets/bhf141.template.json', { cache: 'no-store' });
  if (!res.ok) throw new Error('Slide template is not available yet');
  templateCache = await res.json();
  return templateCache;
}

export function folderList(rules: any, layout: string): any[] {
  if (!rules) return [];
  const key = String(layout || '2DU2T');
  const entry = rules[key] || rules['2DU2T'];
  if (!entry) return [];
  return entry.folders || [];
}
