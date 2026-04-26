import path from "path";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

export const config = {
  port: Number(process.env.PORT || 4000),
  frontendOrigin: process.env.FRONTEND_ORIGIN || true,
  pdfEngine: process.env.PDF_ENGINE || "stub",
  storageDir: path.resolve(__dirname, "../storage"),
  sessionDir: path.resolve(__dirname, "../storage/sessions"),
  originalPdfDir: path.resolve(__dirname, "../storage/originals"),
  exportDir: path.resolve(__dirname, "../storage/exports")
};
