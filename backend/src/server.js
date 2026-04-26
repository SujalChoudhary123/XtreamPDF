import "dotenv/config";
import express from "express";
import cors from "cors";
import path from "path";
import { fileURLToPath } from "url";
import { editorRouter } from "./routes/editor.routes.js";
import { config } from "./config.js";
import { ensureStorage } from "./repositories/session.repository.js";

const app = express();
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

app.use(cors({
  origin: config.frontendOrigin
}));
app.use(express.json({ limit: "10mb" }));
app.use(express.urlencoded({ extended: true, limit: "10mb" }));

app.get("/api/health", (_req, res) => {
  res.json({
    ok: true,
    service: "fluxpdf-backend",
    engine: config.pdfEngine
  });
});

app.use("/api/editor", editorRouter);
app.use("/storage", express.static(path.resolve(__dirname, "../storage")));

await ensureStorage();

app.listen(config.port, () => {
  console.log(`FluxPDF backend listening on http://localhost:${config.port}`);
});
