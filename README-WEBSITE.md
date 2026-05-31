# PDF Converter (Local Website)

This project is a **static website** (HTML/CSS/JS) that runs in your browser.

## Run locally (works anywhere on your PC)
1. Open `index.html` in a browser.
2. For the DOCX → PDF feature, you also need to run the backend:
   - `docx-to-pdf-backend/server.js`
   - Requires LibreOffice (`soffice`) installed and available in PATH.

## Why this is “global”
Because it’s static files, you can copy this entire folder to any computer and it will work (browser-side PDF.js conversion).

## Deploy online (public URL)
You have two options:

### Option A: Static hosting (recommended for PDF→JPG and PDF→DOCX)
Host these files:
- `index.html`
- `style.css`
- `first.js`

The DOCX→PDF (server-side) will still require the backend, so it will only work if you also host the backend/API.

### Option B: Full hosting with backend
Host both:
- Frontend static files
- Backend Node service from `docx-to-pdf-backend/server.js`

Then update the frontend URL `http://localhost:3001/...` to your hosted backend URL.

## Backend endpoints used
Frontend calls:
- `POST /convert-docx-to-pdf`
- `POST /merge-pdfs`

## Local development hint
If you just want to test quickly, open `index.html` and use the buttons.

