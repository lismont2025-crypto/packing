# FabricQ ERP – Fabric Quality Chart Management
Stack: HTML + CSS + JS + Supabase (Auth, Postgres, Realtime, Storage). No build step.
## Architecture
- `profiles` (id, email, role admin|user) · `qualities` (all grey/finish/additional fields, auto `quality_id` QLT-00001 via sequence)
- RLS: any logged-in user can read/add/edit; only **admin** can delete or change roles. First registered user = admin.
- Realtime subscription syncs changes across devices; files go to Storage bucket `quality-files`.
## Setup
1. Create a project at supabase.com.
2. SQL Editor → paste `supabase/schema.sql` → Run.
3. Project Settings → API → copy URL + anon key into `js/config.js`.
4. Authentication → Providers → Email (disable "Confirm email" for instant sign-up if desired).
5. Open `index.html` via any static server (or deploy), **Create Account** – the first account is Admin. Promote others in Users.
## Deploy to GitHub Pages
Push folder to a repo → Settings → Pages → Branch `main` / root → Save. Then add your Pages URL in Supabase → Authentication → URL Configuration → Site URL.
Also works on Netlify / Vercel / Cloudflare Pages (static).
