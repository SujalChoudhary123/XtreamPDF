import express from "express";
import multer from "multer";
import {
  createEditSession,
  exportEditedPdf,
  getEngineInfo,
  getSessions,
  getSession,
  patchBlock,
  updateTextBlock
} from "../services/editor.service.js";

const upload = multer({ storage: multer.memoryStorage() });
export const editorRouter = express.Router();

editorRouter.get("/capabilities", async (_req, res) => {
  try {
    return res.json(await getEngineInfo());
  } catch (error) {
    console.error(error);
    return res.status(500).json({ error: "Could not inspect engine capabilities." });
  }
});

editorRouter.get("/sessions", async (_req, res) => {
  try {
    const sessions = await getSessions();
    return res.json(sessions);
  } catch (error) {
    console.error(error);
    return res.status(500).json({ error: "Could not list sessions." });
  }
});

editorRouter.post("/session", upload.single("file"), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: "A PDF file is required." });
    }

    const session = await createEditSession({
      fileName: req.file.originalname,
      pdfBuffer: req.file.buffer
    });

    return res.status(201).json(session);
  } catch (error) {
    console.error(error);
    return res.status(500).json({ error: "Could not create edit session." });
  }
});

editorRouter.get("/session/:sessionId", async (req, res) => {
  try {
    const session = await getSession(req.params.sessionId);
    if (!session) {
      return res.status(404).json({ error: "Session not found." });
    }
    return res.json(session);
  } catch (error) {
    console.error(error);
    return res.status(500).json({ error: "Could not load session." });
  }
});

editorRouter.patch("/session/:sessionId/blocks/:blockId", async (req, res) => {
  try {
    const block = await updateTextBlock({
      sessionId: req.params.sessionId,
      blockId: req.params.blockId,
      text: req.body.text ?? ""
    });

    if (!block) {
      return res.status(404).json({ error: "Editable block not found." });
    }

    return res.json(block);
  } catch (error) {
    console.error(error);
    return res.status(500).json({ error: "Could not update text block." });
  }
});

editorRouter.patch("/session/:sessionId/blocks/:blockId/properties", async (req, res) => {
  try {
    const block = await patchBlock({
      sessionId: req.params.sessionId,
      blockId: req.params.blockId,
      patch: req.body || {}
    });

    if (!block) {
      return res.status(404).json({ error: "Editable block not found." });
    }

    return res.json(block);
  } catch (error) {
    console.error(error);
    return res.status(500).json({ error: "Could not patch block properties." });
  }
});

editorRouter.post("/session/:sessionId/export/pdf", async (req, res) => {
  try {
    const exported = await exportEditedPdf(req.params.sessionId);
    if (!exported) {
      return res.status(404).json({ error: "Session not found." });
    }
    return res.json(exported);
  } catch (error) {
    console.error(error);
    return res.status(500).json({ error: "Could not export PDF." });
  }
});
