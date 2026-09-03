# Data Workbench

Clean, dedupe, and compare spreadsheets entirely in your browser. No APIs, no cloud
services, no AI calls — your data never leaves your device.

Built for the DevNetwork [API + Cloud + AI] Hackathon 2026 (Overall Winner track).

## Features

- **Pipeline** — chain cleaning steps (trim, standardize case, fill missing values,
  exact/fuzzy dedupe, split/merge columns, keep-only-columns) onto any CSV or Excel
  file, with a live before/after row-count preview at every step.
- **Save/Load pipeline** — export a pipeline as a reusable `.json` template and
  reapply it to a new file later, so recurring cleanups become one click.
- **Compare** — reconcile two files (e.g. two vendor price lists, two contact
  exports) with fuzzy key matching and automatic value deltas.
- **Explorer** — auto-detects column types (date, number, currency, email, phone,
  text) and shows null/unique counts for a quick data-quality overview.
- Fully accessible: keyboard navigable, screen-reader labeled, visible focus states.

## Run it locally

Requires [Node.js](https://nodejs.org) (v18+).

```bash
npm install
npm run dev
```

Then open the local URL it prints (usually `http://localhost:5173`).

## Build for production

```bash
npm run build
npm run preview   # to test the production build locally
```

## Deploy to GitHub Pages

1. Push this project to a public GitHub repo.
2. In `vite.config.js`, set `base: "/your-repo-name/"` to match your repo's name.
3. Run:
   ```bash
   npm install
   npm run build
   npm run deploy
   ```
4. In your repo's Settings → Pages, set the source branch to `gh-pages`.
5. Your live site will be at `https://<your-username>.github.io/<your-repo-name>/`.

## Tech stack

React, Vite, PapaParse (CSV parsing), SheetJS/xlsx (Excel parsing), lucide-react
(icons). No backend, no external API calls, no analytics, no tracking.
