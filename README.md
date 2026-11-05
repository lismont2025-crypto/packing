# FabricQ ERP – Fabric Quality Chart Management (no login)
HTML + CSS + JS + Supabase (Postgres, Realtime, Storage). The site opens straight to the Dashboard; there is no login, username or password.
## Setup
1. Create a project at supabase.com.
2. SQL Editor → paste `supabase/schema.sql` → Run (safe to re-run; removes the old auth tables if present).
3. Project Settings → API → copy URL + anon key into `js/config.js`.
4. Open `index.html` through any static host.
## Deploy to GitHub Pages
Push the folder to a repo → Settings → Pages → Branch `main` / root → Save. Netlify / Vercel / Cloudflare Pages also work.
## Security note
Without login, anyone who has the site link (and can read the anon key in config.js) can read, edit and delete data. Keep the URL private and export Excel backups regularly.
