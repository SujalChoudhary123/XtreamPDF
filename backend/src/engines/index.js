import { StubPdfEngine } from "./stub.engine.js";
import { OpenSourcePdfEngine } from "./opensource.engine.js";
import { config } from "../config.js";

export function createPdfEngine() {
  switch (config.pdfEngine) {
    case "opensource":
    case "open-source":
      return new OpenSourcePdfEngine();
    case "stub":
    default:
      return new StubPdfEngine();
  }
}
