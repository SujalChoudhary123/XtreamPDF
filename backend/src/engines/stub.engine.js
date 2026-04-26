import { randomUUID } from "crypto";

export class StubPdfEngine {
  getName() {
    return "stub";
  }

  async createSessionModel({ sessionId, fileName, originalPdfUrl }) {
    return {
      sessionId,
      fileName,
      engine: this.getName(),
      status: "draft",
      originalPdfUrl,
      pages: [
        {
          id: "page-1",
          pageNumber: 1,
          width: 595,
          height: 842,
          backgroundLocked: true
        }
      ],
      editableBlocks: [
        {
          id: randomUUID(),
          pageId: "page-1",
          text: "Stub editable text block",
          originalText: "Stub editable text block",
          x: 72,
          y: 120,
          width: 220,
          height: 24,
          fontFamily: "Arial, Helvetica, sans-serif",
          fontSize: 14,
          fontWeight: "400",
          fontStyle: "normal",
          color: "#111111",
          modified: false
        }
      ],
      images: [],
      annotations: [],
      notes: [
        "Stub engine is active.",
        "Replace this engine with Apryse, PSPDFKit, Foxit, or MuPDF integration for true PDF text extraction and rewrite."
      ]
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
    return {
      sessionId: session.sessionId,
      status: "stubbed",
      downloadUrl: session.originalPdfUrl,
      notes: [
        "Stub export returns the original PDF.",
        "Real export must be implemented by an engine that rewrites PDF content streams and embedded fonts."
      ]
    };
  }
}
