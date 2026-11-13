# PPT Creation Photo Upload - Multi Zone Platform

One platform for North Zone and East Zone site photo packs. This replaces the Google Sheets and Google Drive setup behind the original East Region app.

## What is different now

- One platform, two zones. A zone switcher sits on every page. North Zone and East Zone data stay separate, and the dashboard, site list, checklist, PPT pack and Excel report all behave the same way in both.
- A real database. Supabase Postgres holds the sites, sub folders and photo records. Supabase Storage holds the photos themselves. No more Google Sheets or Google Drive.
- Shared sign in. Every user signs in with their own login ID and the shared platform password. The ID is recorded against every upload so you always know who sent what.
- Claude connector. The platform publishes an MCP server, so Claude can read progress, create site folders, upload photos and pull the Excel report.
- More visible. A zone bar across the top, live site counts, and a sharper sign in card.

## Sign in

| Who | Login ID | Password | Can do |
| --- | --- | --- | --- |
| PPT team | Their own name (Raju Bose, Mahadev, Avinash, Nayan Nimkar, Vinay Limaye) | Their own password, stored hashed in `app_users` | Create site folders, edit sites, generate the PPT, see the access history |
| Engineers | Their own name (`Admin` is not accepted, so the history shows who it was) | Shared password, `ADMIN_PASSWORD`, default `Admin@2026` | Create site folders, upload photos and files |

A PPT team name never works with the shared password. Members change their password from the account menu; the owner (Raju Bose) can reset one from the History page. The owner adds a member, or switches a login off, on the History page.

## PPT tables from the serial-number Excel

`scripts/autofill-client.js` reads the Excel uploaded to **14.1 Device Serial Numbers** (sheet with MATERIAL / MAKE / MODEL / SERIAL NO. / IP ADDRESS, and the CCTV sheet with Camera Name / Camera Sr No / Camera IP / CCTV Gateway / Subnet) and writes the values into every table page: instrument panel, IT rack, AMS, BOS PC, MMS, OPT, locks, EPS, sensors, CCTV inventory, both IP schemes, the consolidated serial list and the port details. A value missing from the Excel is shown as "-", never as the template value.

Page 2 carries the login name of the member who generated the PPT. The created date is the first DPR date (chosen in the Generate dialog, default: the first upload for the site) and the updated date is the day of the last generation.

## Access history

Sign-ins, refused sign-ins, website visits, sign-outs, PPT generations and site folder changes are written to `access_log` and shown to the PPT team under **History**.

## Stack

- Next.js 15, deployed on Vercel
- Supabase Postgres and Supabase Storage
- MCP server at `/api/mcp`

## Repository layout

| Path | What it is |
| --- | --- |
| `scripts/auto-fill.mjs`, `scripts/autofill-client.js` | Run last in the front end build: Excel auto fill of the PPT tables, page 2 history, individual logins and the History page |
| `scripts/four-du.mjs` | Runs after the front end build and adds the 4 DU layouts (4DU2T, 4DU3T): OPT 5 to 8, DU JB 3 and 4, DU 3 and 4 sump sensors |
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
| `create_site` | Creates a site folder with the standard 34 sub folders (45 for 3DU3T, 44 for 4DU2T, 50 for 4DU3T) |
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
