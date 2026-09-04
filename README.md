# ARC Audience Package Configuration Form

A standalone React + Vite build of the package/filter/custom-question form.
No backend required — everything (including the Excel export on submit)
runs in the browser.

## What's in here

- `src/App.jsx` — the form itself
- `dist/` — a pre-built, ready-to-upload static site (already built for you)
- everything else — the Vite/React/Tailwind project that produced it

## Fastest path: drag-and-drop (no account setup beyond signing in)

1. Go to https://app.netlify.com/drop
2. Drag the `dist` folder onto the page
3. You get a public URL immediately (e.g. `random-name-123.netlify.app`)

## Recommended path: connect to GitHub (get automatic redeploys on future edits)

1. Push this folder to a new GitHub repository
2. Go to https://vercel.com (or https://netlify.com) → "New Project" → import the repo
3. Framework preset: **Vite**. Build command: `npm run build`. Output directory: `dist`
4. Deploy — you'll get a public `*.vercel.app` (or `*.netlify.app`) URL, and every
   future push to the repo redeploys automatically

## Local development

```
npm install
npm run dev       # local dev server with hot reload
npm run build      # produces dist/ for deployment
```

## Notes

- This is a fully static site — the "Submit Selections" button generates
  and downloads an Excel file directly in the browser (via SheetJS). Nothing
  is sent to a server. If you later want submissions saved centrally (e.g.
  emailed, logged to a database, posted to a CRM) that requires adding a
  small backend or a serverless function — the static hosting options below
  all support adding those later without switching platforms.
- If you want the public URL to not be guessable/indexed, Vercel and Netlify
  both offer password-protection on paid tiers; on free tiers, an unlisted
  URL (not linked anywhere public) is the default protection.
