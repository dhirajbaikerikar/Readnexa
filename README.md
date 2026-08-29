# Readnexa — Setup & Deploy Guide

## Why you got the 404

You uploaded the raw `folio.jsx` file straight to GitHub. A browser cannot run a `.jsx`
file directly — it needs to be **built** first (turned into a real `index.html` +
bundled JavaScript) by a tool called Vite. GitHub Pages looked for an `index.html`
at the root of your site and found nothing → 404.

This folder is a complete, buildable project. Once it's set up, GitHub Actions
builds it automatically every time you push, and Pages serves the result.

---

## Part 1 — Supabase (your database + file storage), 10 minutes

1. Go to **supabase.com** → **Sign up** (use your GitHub account) → **New Project**.
2. Pick a name and a database password (save the password somewhere) → **Create**.
   Wait ~2 minutes while it provisions.
3. **Create the storage bucket** (this is where PDF/Word files live):
   - Left sidebar → **Storage** → **New bucket**
   - Name: `books` (must match exactly)
   - Toggle **Public bucket** → **ON**
   - Click **Create bucket**
4. **Create the database tables**:
   - Left sidebar → **SQL Editor** → **New query**
   - Open `supabase-setup.sql` from this folder, paste its contents in, click **Run**
5. **Allow uploads to the bucket** (important — skip this and uploads will fail):
   - Left sidebar → **Storage** → click the `books` bucket → **Policies** tab
   - Click **New policy** → choose **"For full customization"**
   - Policy name: `Allow all`
   - Allowed operation: check **SELECT, INSERT, UPDATE, DELETE**
   - Target roles: `anon`
   - USING expression / WITH CHECK expression: `true`
   - Save
6. **Get your API keys**:
   - Left sidebar → **Project Settings** (gear icon) → **API**
   - Copy the **Project URL** and the **anon public** key — you'll need both next.

---

## Part 2 — Get the code ready

1. Download this whole `readnexa` folder.
2. Open it in a code editor (VS Code is fine) or just upload it as-is to GitHub.
3. Create a file named `.env.local` in the root (copy `.env.example` and rename it),
   and paste in your real Supabase URL and key:
   ```
   VITE_SUPABASE_URL=https://xxxxxxxx.supabase.co
   VITE_SUPABASE_ANON_KEY=eyJhbGciOi...
   ```
   This file is only for testing on your own computer — **do not** upload `.env.local`
   to GitHub (see step 4 below, GitHub Actions uses Secrets instead).

4. (Optional) Test locally before deploying:
   ```
   npm install
   npm run dev
   ```
   Open the URL it gives you (usually `http://localhost:5173`).

---

## Part 3 — Push to GitHub

1. Go to **github.com** → your existing repo `Readnexa` (or create a new one named
   exactly `Readnexa` — capitalization matters, it must match `vite.config.js`).
2. Delete the old `folio.jsx` file from the repo if it's sitting there alone.
3. Upload **all files and folders** from this `readnexa` project to the repo root
   (so `package.json`, `index.html`, `src/`, `.github/` etc. are all at the top level).
   - Easiest way: on your computer, run `git add . && git commit -m "Readnexa" && git push`
   - Or use GitHub's "Add file → Upload files" in the browser (drag the whole folder in).

4. **Add your Supabase keys as GitHub Secrets** (so the automatic build can use them):
   - In your repo → **Settings** → **Secrets and variables** → **Actions**
   - Click **New repository secret** → Name: `VITE_SUPABASE_URL` → paste your URL → Add
   - Click **New repository secret** again → Name: `VITE_SUPABASE_ANON_KEY` → paste your key → Add

5. **Turn on GitHub Pages via Actions**:
   - Repo → **Settings** → **Pages**
   - Under "Build and deployment" → Source: choose **GitHub Actions**

6. Go to the **Actions** tab in your repo — you should see "Deploy Readnexa to GitHub
   Pages" running (triggered automatically by your push). Wait for the green checkmark
   (1–2 minutes).

7. Visit **https://dhirajbaikerikar.github.io/Readnexa/** — it should now load properly.

---

## Troubleshooting

- **Still 404 / blank page**: Check `vite.config.js` — the `base` must be
  `"/Readnexa/"` exactly matching your repo name and capitalization.
- **"Failed to fetch" or upload errors**: Double-check the bucket policy in Part 1,
  step 5 — this is the #1 cause of upload failures.
- **Books don't save / load**: Confirm the two GitHub Secrets are named exactly
  `VITE_SUPABASE_URL` and `VITE_SUPABASE_ANON_KEY`, and re-run the Action
  (Actions tab → select the failed run → "Re-run all jobs") after adding them.
- **Legacy `.doc` files won't preview**: This is expected — only `.pdf` and `.docx`
  can be read/edited in-browser. `.doc` files can still be uploaded and downloaded.
