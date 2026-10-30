# PPT Creation Photo Upload - Multi Zone Platform

One platform for North Zone and East Zone site photo packs. This replaces the Google Sheets and Google Drive setup behind the original East Region app.

## What is different now

- One platform, two zones. A zone switcher sits on every page. North Zone and East Zone data stay separate, and the dashboard, site list, checklist, PPT pack and Excel report all behave the same way in both.
- A real database. Supabase Postgres holds the sites, sub folders and photo records. Supabase Storage holds the photos themselves. No more Google Sheets or Google Drive.
- Shared sign in. Every user signs in with their own login ID and the shared platform password. The ID is recorded against every upload so you always know who sent what.
- Claude connector. The platform publishes an MCP server, so Claude can read progress, create site folders, upload photos and pull the Excel report.
- More visible. A zone bar across the top, live site counts, and a sharper sign in card.

## Sign in

| Field | Value |
| --- | --- |
| Login ID | `Admin` (any ID works, it is used to record who uploaded) |
| Password | `Admin@2026` |

The shared password can be changed with the `ADMIN_PASSWORD` environment variable.

## Stack

- Next.js 15, deployed on Vercel
- Supabase Postgres and Supabase Storage
- MCP server at `/api/mcp`

## Repository layout

| Path | What it is |
| --- | --- |
| `scripts/build-front-end.mjs` | Builds the front end from the original app, pulls out the slide template and folder rules, swaps the Apps Script calls for the platform API and adds the zone switcher |
| `src/app/api/app/route.ts` | The platform API. Sign in, site list, uploads, photo files, Excel export |
| `src/app/api/mcp/route.ts` | The MCP server that Claude connects to |
| `src/lib/core.ts` | Zones, sessions, Supabase client, template and folder rules loader |
| `public/assets/platform.css` | The zone bar and the theme layer |
| `supabase/schema.sql` | Database schema and the private photo bucket |

The original app was one HTML file with the BHF141 slide template inlined, so the build downloads it, extracts the template and the sub folder rules and writes the patched result into `public/`. The source app lives at `raju-bose/PPT-Creation--East-Region-NRO_12345`.

## Setup

### 1. Supabase

1. Create a project at [supabase.com](https://supabase.com).
2. Open the SQL editor and run `supabase/schema.sql`. It creates the `sites`, `photos` and `app_users` tables plus the private `photos` storage bucket.
3. Copy the project URL and the service role key from Project Settings, API.

### 2. Environment variables

Copy `.env.example` to `.env.local` for local work, or add the same values in the Vercel project settings.

| Variable | What it holds |
| --- | --- |
| `APP_URL` | The public address of the deployment, used by the MCP tools |
| `SUPABASE_URL` | Supabase project URL |
| `SUPABASE_SERVICE_ROLE_KEY` | Supabase service role key, server only |
| `ADMIN_PASSWORD` | Shared platform password, defaults to `Admin@2026` |
| `SESSION_SECRET` | Long random string that signs login sessions |
| `MCP_TOKEN` | Long random string that Claude sends as `x-mcp-token` |

### 3. Vercel

1. Import the repository from GitHub.
2. Add the environment variables above.
3. Deploy. The build step downloads the original app, so the first build needs a network connection to GitHub.

## Claude MCP connector

The platform speaks MCP over Streamable HTTP.

| Setting | Value |
| --- | --- |
| Endpoint | `https://YOUR-APP.vercel.app/api/mcp` |
| Header | `x-mcp-token: YOUR_MCP_TOKEN` |

In Claude Desktop, open Settings, Connectors, Add a custom connector and use:

```json
{
  "mcpServers": {
    "cai-photo-platform": {
      "type": "http",
      "url": "https://YOUR-APP.vercel.app/api/mcp",
      "headers": {
        "x-mcp-token": "YOUR-MCP-TOKEN"
      }
    }
  }
}
```

The server exposes these tools:

| Tool | What it does |
| --- | --- |
| `list_zones` | The zones and the states under each one |
| `list_sites` | Every site with its upload progress |
| `get_site` | One site in full: every sub folder, who uploaded it, what is still open |
| `get_checklist` | Checklist rows for the PPT pack, filterable by zone and status |
| `create_site` | Creates a site folder with the standard 34 sub folders (45 for a 3DU3T layout) |
| `upload_photo` | Stores a photo into a sub folder from a URL or base64 |
| `list_photos` | Photos for a site with signed links so they can be viewed |
| `pending_uploads` | Site folders still missing photos, least complete first |
| `export_excel` | Builds the Excel progress report and returns a download link |

## Zone model

| Zone | States |
| --- | --- |
| North Zone | Punjab, Haryana, Himachal Pradesh, Jammu & Kashmir, Ladakh, Rajasthan, Uttar Pradesh, Uttarakhand, Delhi, Chandigarh, Madhya Pradesh, Other |
| East Zone | Odisha, West Bengal, Bihar, Jharkhand, Assam, Meghalaya, Tripura, Arunachal Pradesh, Nagaland, Manipur, Mizoram, Sikkim, Other |

The state list lives in `src/lib/core.ts` and is easy to adjust.
