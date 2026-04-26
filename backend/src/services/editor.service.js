import { randomUUID } from "crypto";
import { createPdfEngine } from "../engines/index.js";
import {
  listSessions,
  readSession,
  saveExportPdf,
  saveOriginalPdf,
  saveSession
} from "../repositories/session.repository.js";

const engine = createPdfEngine();

export async function getEngineInfo() {
  return {
    name: engine.getName(),
    capabilities: typeof engine.getCapabilities === "function"
      ? await engine.getCapabilities()
      : []
  };
}

export async function createEditSession({ fileName, pdfBuffer }) {
  const sessionId = randomUUID();
  const storedPdf = await saveOriginalPdf(sessionId, pdfBuffer);
  const session = await engine.createSessionModel({
    sessionId,
    fileName,
    originalPdfUrl: storedPdf.publicUrl
  });

  session.createdAt = new Date().toISOString();
  session.updatedAt = session.createdAt;
  await saveSession(session);
  return session;
}

export async function getSession(sessionId) {
  return await readSession(sessionId);
}

export async function getSessions() {
  return await listSessions();
}

export async function updateTextBlock({ sessionId, blockId, text }) {
  const session = await readSession(sessionId);
  if (!session) {
    return null;
  }

  const block = await engine.applyTextUpdate(session, blockId, text);
  if (!block) {
    return null;
  }

  session.updatedAt = new Date().toISOString();
  await saveSession(session);
  return block;
}

export async function patchBlock({ sessionId, blockId, patch }) {
  const session = await readSession(sessionId);
  if (!session) {
    return null;
  }

  const block = await engine.applyBlockPatch(session, blockId, patch);
  if (!block) {
    return null;
  }

  session.updatedAt = new Date().toISOString();
  await saveSession(session);
  return block;
}

export async function exportEditedPdf(sessionId) {
  const session = await readSession(sessionId);
  if (!session) {
    return null;
  }

  const exported = await engine.exportPdf(session);
  if (exported?.pdfBytes) {
    const saved = await saveExportPdf(sessionId, exported.pdfBytes);
    delete exported.pdfBytes;
    exported.downloadUrl = saved.publicUrl;
    exported.fileName = saved.fileName;
  }

  return exported;
}
