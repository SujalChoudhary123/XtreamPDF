import fs from "fs/promises";
import path from "path";
import { config } from "../config.js";

export async function ensureStorage() {
  await Promise.all([
    fs.mkdir(config.storageDir, { recursive: true }),
    fs.mkdir(config.sessionDir, { recursive: true }),
    fs.mkdir(config.originalPdfDir, { recursive: true }),
    fs.mkdir(config.exportDir, { recursive: true })
  ]);
}

export async function saveOriginalPdf(sessionId, buffer) {
  await ensureStorage();
  const fileName = `${sessionId}.pdf`;
  const filePath = path.join(config.originalPdfDir, fileName);
  await fs.writeFile(filePath, buffer);
  return {
    fileName,
    filePath,
    publicUrl: `/storage/originals/${fileName}`
  };
}

export async function saveExportPdf(sessionId, buffer) {
  await ensureStorage();
  const fileName = `${sessionId}-edited.pdf`;
  const filePath = path.join(config.exportDir, fileName);
  await fs.writeFile(filePath, buffer);
  return {
    fileName,
    filePath,
    publicUrl: `/storage/exports/${fileName}`
  };
}

export async function saveSession(session) {
  await ensureStorage();
  const filePath = path.join(config.sessionDir, `${session.sessionId}.json`);
  await fs.writeFile(filePath, JSON.stringify(session, null, 2), "utf8");
  return session;
}

export async function readSession(sessionId) {
  try {
    const filePath = path.join(config.sessionDir, `${sessionId}.json`);
    const raw = await fs.readFile(filePath, "utf8");
    return JSON.parse(raw);
  } catch (error) {
    if (error.code === "ENOENT") {
      return null;
    }
    throw error;
  }
}

export async function listSessions() {
  await ensureStorage();
  const files = await fs.readdir(config.sessionDir);
  const sessions = await Promise.all(
    files
      .filter((name) => name.endsWith(".json"))
      .map((name) => readSession(path.basename(name, ".json")))
  );

  return sessions.filter(Boolean);
}
