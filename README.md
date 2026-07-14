# The Roll Ledger — Attendance Ledger

A React + Vite app for taking classroom attendance from an uploaded Excel roster
and exporting a dated attendance column back into a spreadsheet.

> **Note on storage:** the original component was written for Claude.ai's
> artifact sandbox, which provides a `window.storage` API for saving data.
> That API doesn't exist in a normal browser, so this project includes
> `src/storage-polyfill.js`, a drop-in replacement backed by `localStorage`.
> It keeps the app fully working, but "saved registers" will only persist on
> the same browser/device — they are **not** shared across users or devices.
> If you need real cross-device saving (e.g. an email really does pull up the
> same registers from any computer), swap the polyfill's `get`/`set`/`delete`/
> `list` functions for calls to your own backend (Firebase, Supabase, a small
> Node/Express API, etc.).

## 1. Test it locally

Requirements: [Node.js](https://nodejs.org) 18+ and npm.

```bash
npm install     # install dependencies
npm run dev     # start a local dev server (usually http://localhost:5173)
```

Open the printed URL, upload a sample `.xlsx`/`.xls` roster, and click through
the steps (Upload → Mapping → Configure → Take → Export) to confirm it behaves
as expected.

To verify the production build works (this is what actually gets deployed):

```bash
npm run build     # builds into dist/
npm run preview   # serves the dist/ build locally so you can double check it
```

## 2. Push it to GitHub

From inside this project folder:

```bash
git init
git add .
git commit -m "Initial commit: attendance ledger app"

# Create a new empty repo on github.com first (no README/license/gitignore,
# since you already have them), then:
git branch -M main
git remote add origin https://github.com/<your-username>/<your-repo>.git
git push -u origin main
```

(You can also create the repo directly from the GitHub CLI with
`gh repo create <your-repo> --public --source=. --push` if you have `gh`
installed and authenticated.)

## 3. Deploy it

Any static host works since this builds to plain HTML/JS/CSS. Two easy options:

### Option A — Vercel (recommended, zero config)
1. Go to https://vercel.com, sign in with GitHub.
2. "Add New Project" → import your repo.
3. Framework preset: **Vite**. Build command `npm run build`, output dir `dist`.
4. Deploy — you'll get a live URL (and every future push to `main` auto-deploys).

### Option B — Netlify
1. Go to https://app.netlify.com, "Add new site" → "Import an existing project".
2. Connect the GitHub repo.
3. Build command: `npm run build`, publish directory: `dist`.
4. Deploy.

### Option C — GitHub Pages
1. In `vite.config.js`, uncomment and set `base: '/<your-repo-name>/'`.
2. Install the pages helper: `npm install -D gh-pages`
3. Add to `package.json` scripts: `"deploy": "npm run build && npx gh-pages -d dist"`
4. Run `npm run deploy`, then enable Pages in your repo's Settings → Pages,
   pointing at the `gh-pages` branch.

## Project structure

```
attendance-ledger/
├── index.html
├── package.json
├── vite.config.js
├── src/
│   ├── main.jsx                # React entry point
│   ├── AttendanceLedger.jsx    # the app itself (your original component)
│   └── storage-polyfill.js     # localStorage shim for window.storage
└── README.md
```
