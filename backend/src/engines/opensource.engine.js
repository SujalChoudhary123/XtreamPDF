import { randomUUID } from "crypto";
import fs from "fs/promises";
import { PDFDocument, StandardFonts, rgb } from "pdf-lib";
import { fileURLToPath } from "url";
import { extractStructuredText, runTool, toolPaths } from "../utils/pdf-tools.js";

const TOOL_CANDIDATES = [
  {
    id: "mutool",
    commands: [toolPaths.mutool, "mutool"],
    args: ["-v"],
    purpose: "MuPDF CLI for PDF inspection, drawing, and low-level document operations"
  },
  {
    id: "pdftotext",
    commands: ["pdftotext"],
    args: ["-v"],
    purpose: "Poppler utility for text extraction"
  },
  {
    id: "qpdf",
    commands: [toolPaths.qpdf, "qpdf"],
    args: ["--version"],
    purpose: "Structural PDF transformations and repair"
  },
  {
    id: "magick",
    commands: ["magick"],
    args: ["-version"],
    purpose: "Image conversions for page previews or OCR preprocessing"
  },
  {
    id: "ocrmypdf",
    commands: ["ocrmypdf"],
    args: ["--version"],
    purpose: "OCR pipeline for scanned PDFs"
  }
];

export class OpenSourcePdfEngine {
  constructor() {
    this.capabilitiesPromise = null;
  }

  getName() {
    return "open-source";
  }

  async getCapabilities() {
    if (!this.capabilitiesPromise) {
      this.capabilitiesPromise = detectInstalledTools();
    }
    return await this.capabilitiesPromise;
  }

  async createSessionModel({ sessionId, fileName, originalPdfUrl }) {
    const capabilities = await this.getCapabilities();
    const originalPdfPath = fileUrlToPath(originalPdfUrl);
    const extracted = await extractStructuredText(originalPdfPath);
    const notes = [
      "Open-source engine path is active.",
      "This mode is designed for zero-cost local tooling.",
      "For truly Adobe-grade editing fidelity, a commercial engine is still stronger, but this stack is the right free starting point."
    ];

    if (!capabilities.some((tool) => tool.available)) {
      notes.push("No open-source PDF tools were detected on this machine yet. Install MuPDF, Poppler, and qpdf to unlock stronger backend workflows.");
    }

    return {
      sessionId,
      fileName,
      engine: this.getName(),
      status: "draft",
      originalPdfUrl,
      pages: extracted.pages.map((page) => ({
        id: `page-${page.pageNumber}`,
        pageNumber: page.pageNumber,
        width: page.width,
        height: page.height,
        backgroundLocked: true
      })),
      editableBlocks: flattenEditableBlocks(extracted.pages),
      images: [],
      annotations: [],
      capabilities,
      notes
    };
  }

  async applyTextUpdate(session, blockId, text) {
    const block = session.editableBlocks.find((entry) => entry.id === blockId);
    if (!block) {
      return null;
    }

    block.text = text;
    block.modified = text !== block.originalText;
    block.updatedAt = new Date().toISOString();
    return block;
  }

  async applyBlockPatch(session, blockId, patch) {
    const block = session.editableBlocks.find((entry) => entry.id === blockId);
    if (!block) {
      return null;
    }

    Object.assign(block, patch);
    block.modified = true;
    block.updatedAt = new Date().toISOString();
    return block;
  }

  async exportPdf(session) {
    const capabilities = await this.getCapabilities();
    const modifiedBlocks = session.editableBlocks.filter((block) => block.modified);

    if (!modifiedBlocks.length) {
      return {
        sessionId: session.sessionId,
        status: "unchanged",
        downloadUrl: session.originalPdfUrl,
        capabilities,
        notes: [
          "No modified text blocks were found, so the original PDF is returned."
        ]
      };
    }

    const originalPdfPath = fileUrlToPath(session.originalPdfUrl);
    const pdfBytes = await fs.readFile(originalPdfPath);
    const pdfDoc = await PDFDocument.load(pdfBytes);
    const fontCache = new Map();

    for (const block of modifiedBlocks) {
      const pageIndex = Number(String(block.pageId || "").replace("page-", "")) - 1;
      const page = pdfDoc.getPage(pageIndex);
      if (!page) {
        continue;
      }

      const font = await getEmbeddedFont(pdfDoc, block.fontFamily, fontCache);
      const x = Number(block.x || 0);
      const width = Number(block.width || 10);
      const height = Number(block.height || block.fontSize || 12);
      const top = Number(block.y || 0);
      const pageHeight = page.getHeight();
      const fontSize = Math.max(6, Number(block.fontSize || 12));
      const lineHeight = Math.max(fontSize, Number(block.lineHeight || height || fontSize));
      const text = normalizeText(block.text);
      const lines = splitPreservingLines(block.text);
      const textBoxHeight = Math.max(height, lineHeight * Math.max(lines.length, 1));
      const baselineStart = Math.max(
        0,
        pageHeight - top - estimateFontAscent(fontSize, height)
      );

      page.drawRectangle({
        x: Math.max(0, x - 1),
        y: Math.max(0, pageHeight - top - textBoxHeight - 1),
        width: Math.max(4, width + 2),
        height: Math.max(4, textBoxHeight + 2),
        color: rgb(1, 1, 1)
      });

      lines.forEach((line, index) => {
        const y = Math.max(0, baselineStart - (index * lineHeight));
        const lineWidth = font.widthOfTextAtSize(line, fontSize);
        const drawX = resolveAlignedX(x, width, lineWidth, block.textAlign);

        page.drawText(line, {
          x: drawX,
          y,
          size: fontSize,
          font,
          color: hexToRgb(block.color || "#111111")
        });
      });
    }

    return {
      sessionId: session.sessionId,
      status: "exported",
      pdfBytes: await pdfDoc.save(),
      capabilities,
      notes: [
        "Open-source export wrote text changes back as page overlays on top of the original PDF."
      ]
    };
  }
}

async function detectInstalledTools() {
  const checks = TOOL_CANDIDATES.map(async (tool) => {
    for (const command of tool.commands) {
      try {
        return createCapabilityResult({
          tool,
          command,
          result: await runTool(command, tool.args)
        });
      } catch (error) {
        const output = `${error.stdout || ""}${error.stderr || ""}`.trim();
        if (output) {
          return createCapabilityResult({
            tool,
            command,
            result: {
              stdout: error.stdout,
              stderr: error.stderr
            }
          });
        }
      }
    }

    return {
      id: tool.id,
      command: tool.commands[0],
      available: false,
      purpose: tool.purpose,
      version: null
    };
  });

  return await Promise.all(checks);
}

function createCapabilityResult({ tool, command, result }) {
  return {
    id: tool.id,
    command,
    available: true,
    purpose: tool.purpose,
    version: `${result.stdout || ""}${result.stderr || ""}`.trim().split(/\r?\n/)[0] || "Detected"
  };
}

function flattenEditableBlocks(pages) {
  return pages.flatMap((page) => page.blocks
    .filter((block) => block.type === "text")
    .flatMap((block) => (block.lines || [])
      .filter((line) => normalizeText(line.text))
      .map((line) => ({
        id: randomUUID(),
        pageId: `page-${page.pageNumber}`,
        text: normalizeText(line.text),
        originalText: normalizeText(line.text),
        x: line.bbox?.x ?? block.bbox?.x ?? 0,
        y: line.bbox?.y ?? block.bbox?.y ?? 0,
        width: line.bbox?.w ?? block.bbox?.w ?? 0,
        height: line.bbox?.h ?? block.bbox?.h ?? (line.font?.size || 12),
        fontFamily: mapFontFamily(line.font?.family),
        fontSize: line.font?.size || 12,
        fontWeight: line.font?.weight === "bold" ? "700" : "400",
        fontStyle: line.font?.style === "italic" ? "italic" : "normal",
        color: "#111111",
        modified: false
      }))));
}

function normalizeText(value) {
  return String(value || "")
    .replace(/\u2028/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function splitPreservingLines(value) {
  return String(value || "")
    .replace(/\u2028/g, "\n")
    .split(/\r?\n/)
    .map((line) => line.replace(/\s+/g, " ").trimEnd())
    .filter((line, index, lines) => line || index < lines.length - 1);
}

function mapFontFamily(family) {
  const value = String(family || "").toLowerCase();
  if (value.includes("mono")) {
    return "Courier";
  }
  if (value.includes("serif")) {
    return "Times Roman";
  }
  return "Helvetica";
}

function fileUrlToPath(publicUrl) {
  const fileName = publicUrl.split("/").pop();
  const isExport = publicUrl.includes("/exports/");
  const folder = isExport ? "../../storage/exports" : "../../storage/originals";
  return fileURLToPath(new URL(`${folder}/${fileName}`, import.meta.url));
}

async function getEmbeddedFont(pdfDoc, fontFamily, cache) {
  const family = mapFontFamily(fontFamily);
  if (!cache.has(family)) {
    const fontName = family === "Courier"
      ? StandardFonts.Courier
      : family === "Times Roman"
        ? StandardFonts.TimesRoman
        : StandardFonts.Helvetica;
    cache.set(family, await pdfDoc.embedFont(fontName));
  }

  return cache.get(family);
}

function hexToRgb(hex) {
  const clean = String(hex || "#111111").replace("#", "");
  const normalized = clean.length === 3
    ? clean.split("").map((value) => value + value).join("")
    : clean;
  const num = Number.parseInt(normalized, 16);
  return rgb(((num >> 16) & 255) / 255, ((num >> 8) & 255) / 255, (num & 255) / 255);
}

function estimateFontAscent(fontSize, boxHeight) {
  const referenceHeight = Math.max(fontSize, Number(boxHeight || 0));
  return Math.min(referenceHeight, Math.max(fontSize * 0.78, referenceHeight * 0.72));
}

function resolveAlignedX(x, width, lineWidth, textAlign) {
  if (textAlign === "center") {
    return x + Math.max(0, (width - lineWidth) / 2);
  }
  if (textAlign === "right" || textAlign === "end") {
    return x + Math.max(0, width - lineWidth);
  }
  return x;
}
