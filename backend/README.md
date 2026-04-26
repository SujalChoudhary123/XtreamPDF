# FluxPDF Backend

Backend scaffold for a real PDF editing architecture with persistent sessions, engine abstraction, and editor APIs.

This repo is now biased toward a zero-budget stack:
- MuPDF (`mutool`)
- Poppler (`pdftotext`)
- qpdf
- ImageMagick (`magick`)
- OCRmyPDF

## What this gives you

- Express API service for editor sessions
- Upload endpoint for source PDFs
- Session endpoint for extracted editable data
- Update endpoint for modified text blocks
- Export endpoint placeholder for final PDF generation
- Persistent JSON session storage on disk
- Engine abstraction layer for swapping in a real engine
- Open-source engine mode with local tool capability detection

## What is still stubbed

- True PDF text extraction
- Exact font extraction and replacement
- Content stream rewriting
- Embedded font reuse
- Exact export of edited PDFs

## Free stack recommendation

Install these local tools on Windows and add them to `PATH`:

- MuPDF (`mutool`)
- Poppler (`pdftotext`)
- qpdf
- ImageMagick (`magick`)
- OCRmyPDF

This backend can now report which of those tools are available through:

- `GET /api/editor/capabilities`

## Practical truth

This free stack is the best no-money path, but it still will not instantly become Adobe-level editing. The right strategy is:

1. Use MuPDF/Poppler/qpdf for extraction and structural work
2. Preserve untouched pages as locked background content
3. Rewrite only edited blocks where possible
4. Fall back to page-level rebuilds when exact low-level rewrite is not safe

## Run

```powershell
cd "c:\Users\hp\Documents\Project 3\backend"
npm install
npm run dev
```

## API shape

- `POST /api/editor/session`
- `GET /api/editor/capabilities`
- `GET /api/editor/sessions`
- `GET /api/editor/session/:sessionId`
- `PATCH /api/editor/session/:sessionId/blocks/:blockId`
- `PATCH /api/editor/session/:sessionId/blocks/:blockId/properties`
- `POST /api/editor/session/:sessionId/export/pdf`
