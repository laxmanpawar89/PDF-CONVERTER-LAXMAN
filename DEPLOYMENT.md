# Deployment notes (Global website)

## What we have
- Frontend (static): `index.html`, `style.css`, `first.js`
- Backend (Node/Express): `docx-to-pdf-backend/server.js`

The UI freezes fix is already in `first.js`.

## Hosting options
### 1) Recommended: Deploy backend on a Node platform + frontend on a static host
- Deploy backend to something like: Render / Railway / Fly.io / (VPS)
- Deploy frontend to: Netlify / Vercel / GitHub Pages

## Backend setup (so the frontend can call it)
Frontend currently calls:
- `http://localhost:3001/convert-docx-to-pdf`
- `http://localhost:3001/merge-pdfs`

After backend is deployed, update `first.js`:
- replace `http://localhost:3001` with your backend base URL.

## DOCX→PDF dependency
`docx-to-pdf-backend/server.js` uses LibreOffice (`soffice`).
On the backend host you must ensure LibreOffice is installed and available in PATH.
- Render/Railway/Fly typically require a custom build/apt-get install.

## Quick checklist before going live
- [ ] Backend deploys and `/health` returns `{ ok: true }`
- [ ] LibreOffice present in backend runtime
- [ ] Frontend uses correct backend URL
- [ ] Open console: verify requests to `/convert-docx-to-pdf` / `/merge-pdfs`

## PR link containing the DOCX UI fix
- https://github.com/laxmanpawar89/PDF-CONVERTER-LAXMAN/pull/new/blackboxai/docx-freeze-fix

