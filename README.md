# FluxPDF Studio

FluxPDF Studio is a PDF tools project with a polished frontend and a lightweight backend for editing sessions.

The project is split into two main parts:

- A browser-based PDF tools website in the project root
- A Node.js backend in `backend/` for PDF editing sessions and exports

There is also a separate React sample in `mupdf-webviewer-react-sample/` for testing MuPDF WebViewer on its own.

## What this project does

The root frontend is designed like an online PDF tools product. It includes tools such as:

- Merge PDF
- Split PDF
- Extract Pages
- Delete Pages
- Rotate PDF
- Reorder Pages
- Watermark PDF
- Page Numbers
- Add Text
- Images to PDF
- PDF to Images
- Compress PDF
- Unlock PDF
- Protect PDF
- OCR PDF
- eSign PDF
- HTML to PDF
- PDF to Word
- Word to PDF
- Scan to PDF
- Compare PDFs

Some of these run directly in the browser. The editing flow can also talk to the backend at `http://localhost:4000/api`.

## Project structure

- `index.html`, `styles.css`, `app.js`
  The main website and browser-side PDF tools
- `backend/`
  Express backend for editor sessions, uploads, session storage, and export endpoints
- `mupdf-webviewer-react-sample/`
  Standalone React + Vite sample for MuPDF WebViewer

## Tech stack

- Frontend: HTML, CSS, JavaScript
- Backend: Node.js, Express
- PDF libraries and tools: PDF.js, PDFLib, jsPDF, MuPDF-related sample viewer
- Storage: local files in `backend/storage/`

## How to run the project

If you want the main website and backend working together, use 2 terminals.

### Terminal 1: Start the backend

```powershell
cd "c:\Users\hp\Documents\Project 3\backend"
npm install
npm run dev
```

The backend runs on:

```text
http://localhost:4000
```

You can test it here:

```text
http://localhost:4000/api/health
```

### Terminal 2: Start the main frontend

```powershell
cd "c:\Users\hp\Documents\Project 3"
python -m http.server 8000
```

Then open:

```text
http://localhost:8000
```

This matches the current backend CORS setting in `backend/.env`:

```env
FRONTEND_ORIGIN=http://localhost:8000
```

## Running the React MuPDF sample

The React sample is separate from the main root frontend.

```powershell
cd "c:\Users\hp\Documents\Project 3\mupdf-webviewer-react-sample"
npm install
npm run dev
```

Open the Vite URL shown in the terminal, usually:

```text
http://localhost:5173
```

Important:

- This sample is mainly for testing the MuPDF WebViewer setup
- It is not the same as the main website in the project root
- If you use a MuPDF license key later, keep that key out of GitHub

## Backend notes

The backend currently uses local file storage, not a database.

Data is stored in:

- `backend/storage/sessions/`
- `backend/storage/originals/`
- `backend/storage/exports/`

The backend also supports environment-based config through `backend/.env`.

## Files you should not push to GitHub

These should stay private or be ignored:

- `backend/.env`
- `backend/storage/`
- `backend/node_modules/`
- `mupdf-webviewer-react-sample/node_modules/`

You should keep `backend/.env.example` in the repo as the public template.

## Current status

This project is in a practical prototype stage:

- The UI is much more complete than a basic demo
- Many PDF tools already work in the browser
- The backend is a real scaffold for session-based editing
- Some advanced PDF editing behavior is still approximate and not production-grade

## Backend API

The backend currently exposes these routes:

- `POST /api/editor/session`
- `GET /api/editor/capabilities`
- `GET /api/editor/sessions`
- `GET /api/editor/session/:sessionId`
- `PATCH /api/editor/session/:sessionId/blocks/:blockId`
- `PATCH /api/editor/session/:sessionId/blocks/:blockId/properties`
- `POST /api/editor/session/:sessionId/export/pdf`

## Next improvement ideas

- Add a root `.gitignore`
- Connect the React sample to the backend if needed
- Improve backend export fidelity
- Replace local file storage with a real database if the project grows

## License

Add your preferred license here before publishing the project publicly.
