# Fleet Ledger

Track trucks and the contracts booked for each one. Built with React + Vite,
deployable as a free static site on GitHub Pages, installable on your phone
as an app (PWA).

## 1. Put this on GitHub

1. Create a new **empty** repository on GitHub (no README/license), e.g. `fleet-ledger`.
2. In this folder, run:
   ```bash
   git init
   git add .
   git commit -m "Initial commit"
   git branch -M main
   git remote add origin https://github.com/<your-username>/<your-repo>.git
   git push -u origin main
   ```

## 2. Turn on GitHub Pages

1. On GitHub, open your repo → **Settings** → **Pages**.
2. Under "Build and deployment" → **Source**, choose **GitHub Actions**.
3. Go to the **Actions** tab — a "Deploy to GitHub Pages" run should already
   be in progress (it triggers automatically on every push to `main`). Wait
   for it to finish (green check).
4. Your app is now live at:
   ```
   https://<your-username>.github.io/<your-repo>/
   ```
   (Settings → Pages also shows this URL directly.)

Every time you push new changes to `main`, the site rebuilds and redeploys
automatically — no extra steps needed.

## 3. Install it on your phone

Open the URL above in your phone's browser, then:

- **iPhone (Safari):** tap the Share icon → **Add to Home Screen**.
- **Android (Chrome):** tap the ⋮ menu → **Install app** (or **Add to Home
  screen**).

It'll show up as its own icon and open full-screen like a normal app.

## About your data

This app stores everything **locally in the browser** on whichever device
you're using (no server, no account). That means:

- Data on your phone and data on your laptop are **separate** — they don't
  sync automatically.
- Use the **Export .xlsx** button to download a spreadsheet of everything,
  and **Import .xlsx** on another device to bring it in there.
- Clearing your browser's site data/history for this app will erase what's
  stored — export a backup periodically if the data matters.

## Local development

```bash
npm install
npm run dev       # runs at http://localhost:5173
npm run build     # production build, output in dist/
```
