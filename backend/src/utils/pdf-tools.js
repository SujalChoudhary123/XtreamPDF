import fs from "fs/promises";
import path from "path";
import { execFile } from "child_process";
import { fileURLToPath } from "url";
import { promisify } from "util";

const execFileAsync = promisify(execFile);
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const backendRoot = path.resolve(__dirname, "..", "..");

export const toolPaths = {
  mutool: path.join(backendRoot, "tools", "bin", "mupdf", "mutool.exe"),
  qpdf: path.join(backendRoot, "tools", "bin", "qpdf", "qpdf-12.3.2-msvc64", "bin", "qpdf.exe")
};

export async function runTool(command, args, options = {}) {
  try {
    return await execFileAsync(command, args, {
      cwd: backendRoot,
      timeout: 15000,
      maxBuffer: 32 * 1024 * 1024,
      ...options
    });
  } catch (error) {
    const output = `${error.stdout || ""}${error.stderr || ""}`.trim();
    if (output) {
      return {
        stdout: error.stdout || "",
        stderr: error.stderr || ""
      };
    }

    throw error;
  }
}

export async function extractStructuredText(pdfPath) {
  const [pagesResult, textResult] = await Promise.all([
    runTool(toolPaths.mutool, ["pages", pdfPath]),
    runTool(toolPaths.mutool, ["draw", "-q", "-F", "stext.json", "-o", "-", pdfPath])
  ]);
  const pagesOutput = pagesResult.stdout || pagesResult.stderr || "";
  const textOutput = textResult.stdout || textResult.stderr || "";

  const pageSizes = parsePageSizes(pagesOutput);
  const parsed = JSON.parse(stripTrailingMutoolLogs(textOutput));

  return {
    pages: (parsed.pages || []).map((page, index) => ({
      pageNumber: index + 1,
      width: pageSizes[index]?.width || 612,
      height: pageSizes[index]?.height || 792,
      blocks: page.blocks || []
    }))
  };
}

function parsePageSizes(output) {
  const sizes = [];
  const pattern = /<MediaBox l="[^"]*" b="[^"]*" r="([^"]+)" t="([^"]+)" \/>/g;

  let match = pattern.exec(output);
  while (match) {
    sizes.push({
      width: Number(match[1]),
      height: Number(match[2])
    });
    match = pattern.exec(output);
  }

  return sizes;
}

function stripTrailingMutoolLogs(output) {
  const raw = String(output || "").trim();
  const lastJsonBrace = raw.lastIndexOf("}");
  return lastJsonBrace >= 0 ? raw.slice(0, lastJsonBrace + 1) : raw;
}

export async function fileExists(filePath) {
  try {
    await fs.access(filePath);
    return true;
  } catch {
    return false;
  }
}
