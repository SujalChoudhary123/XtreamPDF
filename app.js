import * as pdfjsLib from "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.3.136/pdf.min.mjs";

pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.3.136/pdf.worker.min.mjs";

const { PDFDocument, StandardFonts, degrees, rgb } = PDFLib;
const { jsPDF } = window.jspdf;
const BACKEND_CONFIG = {
  enabled: true,
  apiBaseUrl: "https://document-backend-5ick.onrender.com/api"
};

const toolRegistry = [
  {
    id: "pdf-editor",
    name: "PDF Editor",
    status: "live",
    category: "Edit",
    output: "Edited PDF",
    description: "Upload a PDF, edit it in a Word-style workspace on the website, then export the edited result as PDF.",
    fields: [
      { type: "file", name: "file", label: "Upload PDF", accept: ".pdf", required: true },
      { type: "select", name: "exportFormat", label: "Download format", options: ["pdf", "word+pdf", "word"] }
    ],
    handler: pdfEditor
  },
  {
    id: "merge",
    name: "Merge PDF",
    status: "live",
    category: "Organize",
    output: "Single merged PDF",
    description: "Combine multiple PDF files into one clean document.",
    fields: [
      { type: "file", name: "files", label: "Upload PDF files", accept: ".pdf", multiple: true, required: true, help: "Add two or more PDFs in the order you want to merge." }
    ],
    handler: mergePdf
  },
  {
    id: "split",
    name: "Split PDF",
    status: "live",
    category: "Organize",
    output: "One PDF per page",
    description: "Split a document into separate single-page PDF files.",
    fields: [
      { type: "file", name: "file", label: "Upload PDF", accept: ".pdf", required: true }
    ],
    handler: splitPdf
  },
  {
    id: "extract",
    name: "Extract Pages",
    status: "live",
    category: "Organize",
    output: "New PDF with selected pages",
    description: "Pull out specific pages using ranges like 1-3, 5, 8-10.",
    fields: [
      { type: "file", name: "file", label: "Upload PDF", accept: ".pdf", required: true },
      { type: "text", name: "pages", label: "Pages to extract", placeholder: "1-3, 5, 8", required: true, help: "Use page numbers starting at 1." }
    ],
    handler: extractPages
  },
  {
    id: "delete",
    name: "Delete Pages",
    status: "live",
    category: "Organize",
    output: "PDF without removed pages",
    description: "Remove unwanted pages before sharing or exporting.",
    fields: [
      { type: "file", name: "file", label: "Upload PDF", accept: ".pdf", required: true },
      { type: "text", name: "pages", label: "Pages to delete", placeholder: "2, 4-6", required: true }
    ],
    handler: deletePages
  },
  {
    id: "rotate",
    name: "Rotate PDF",
    status: "live",
    category: "Edit",
    output: "Rotated PDF",
    description: "Rotate all pages by 90, 180, or 270 degrees.",
    fields: [
      { type: "file", name: "file", label: "Upload PDF", accept: ".pdf", required: true },
      { type: "select", name: "angle", label: "Rotation", options: ["90", "180", "270"], required: true }
    ],
    handler: rotatePdf
  },
  {
    id: "reorder",
    name: "Reorder Pages",
    status: "live",
    category: "Organize",
    output: "PDF with new page order",
    description: "Create a fresh document using any page order you choose.",
    fields: [
      { type: "file", name: "file", label: "Upload PDF", accept: ".pdf", required: true },
      { type: "text", name: "order", label: "New order", placeholder: "3, 1, 2, 4", required: true, help: "Each page should appear once." }
    ],
    handler: reorderPages
  },
  {
    id: "watermark",
    name: "Watermark PDF",
    status: "live",
    category: "Brand",
    output: "Watermarked PDF",
    description: "Overlay brand text on every page.",
    fields: [
      { type: "file", name: "file", label: "Upload PDF", accept: ".pdf", required: true },
      { type: "text", name: "text", label: "Watermark text", placeholder: "CONFIDENTIAL", required: true },
      { type: "color", name: "color", label: "Text color", value: "#d85f3f" },
      { type: "select", name: "opacity", label: "Opacity", options: ["0.15", "0.25", "0.35", "0.5"] }
    ],
    handler: watermarkPdf
  },
  {
    id: "number",
    name: "Page Numbers",
    status: "live",
    category: "Edit",
    output: "PDF with page numbers",
    description: "Stamp every page with a consistent number label.",
    fields: [
      { type: "file", name: "file", label: "Upload PDF", accept: ".pdf", required: true },
      { type: "text", name: "format", label: "Number format", placeholder: "Page {n}", value: "Page {n}", required: true },
      { type: "select", name: "position", label: "Position", options: ["bottom-center", "bottom-right", "top-right"] }
    ],
    handler: numberPages
  },
  {
    id: "annotate",
    name: "Add Text",
    status: "live",
    category: "Edit",
    output: "Annotated PDF",
    description: "Add a text overlay to every page for quick editing and stamping.",
    fields: [
      { type: "file", name: "file", label: "Upload PDF", accept: ".pdf", required: true },
      { type: "text", name: "text", label: "Text to place", placeholder: "Reviewed by Finance Team", required: true },
      { type: "number", name: "x", label: "X position", value: "60", required: true },
      { type: "number", name: "y", label: "Y position", value: "60", required: true }
    ],
    handler: annotatePdf
  },
  {
    id: "images-to-pdf",
    name: "Images to PDF",
    status: "live",
    category: "Convert",
    output: "PDF generated from images",
    description: "Turn JPG or PNG files into a crisp multipage PDF.",
    fields: [
      { type: "file", name: "files", label: "Upload images", accept: ".png,.jpg,.jpeg", multiple: true, required: true },
      { type: "checkbox", name: "fit", label: "Fit images to page", checked: true }
    ],
    handler: imagesToPdf
  },
  {
    id: "pdf-to-images",
    name: "PDF to Images",
    status: "live",
    category: "Convert",
    output: "PNG images",
    description: "Render each PDF page as a downloadable PNG preview.",
    fields: [
      { type: "file", name: "file", label: "Upload PDF", accept: ".pdf", required: true }
    ],
    handler: pdfToImages
  },
  {
    id: "compress",
    name: "Compress PDF",
    status: "live",
    category: "Optimize",
    output: "Compressed PDF",
    description: "Rebuild the PDF from rendered page images to reduce file size.",
    fields: [
      { type: "file", name: "file", label: "Upload PDF", accept: ".pdf", required: true },
      { type: "select", name: "quality", label: "Image quality", options: ["0.45", "0.6", "0.75"] },
      { type: "select", name: "scale", label: "Render scale", options: ["1.1", "1.4", "1.8"] }
    ],
    handler: compressPdf
  },
  {
    id: "unlock",
    name: "Unlock PDF",
    status: "live",
    category: "Security",
    output: "Unlocked PDF",
    description: "Open a password-protected PDF and export an unlocked copy.",
    fields: [
      { type: "file", name: "file", label: "Upload locked PDF", accept: ".pdf", required: true },
      { type: "password", name: "password", label: "PDF password", required: true }
    ],
    handler: unlockPdf
  },
  {
    id: "protect",
    name: "Protect PDF",
    status: "live",
    category: "Security",
    output: "Password-protected PDF",
    description: "Export a password-protected version of your PDF.",
    fields: [
      { type: "file", name: "file", label: "Upload PDF", accept: ".pdf", required: true },
      { type: "password", name: "userPassword", label: "Open password", required: true },
      { type: "password", name: "ownerPassword", label: "Owner password", help: "Optional. If empty, the open password will also be used as owner password." }
    ],
    handler: protectPdf
  },
  {
    id: "ocr",
    name: "OCR PDF",
    status: "live",
    category: "AI",
    output: "Extracted OCR text",
    description: "Run OCR on a PDF or image and download the recognized text.",
    fields: [
      { type: "file", name: "file", label: "Upload PDF or image", accept: ".pdf,.png,.jpg,.jpeg", required: true },
      { type: "select", name: "language", label: "OCR language", options: ["eng"] },
      { type: "select", name: "format", label: "Download format", options: ["txt", "docx"] }
    ],
    handler: ocrPdf
  },
  {
    id: "esign",
    name: "eSign PDF",
    status: "live",
    category: "Workflow",
    output: "Signed PDF",
    description: "Place a signature image on a selected page and position.",
    fields: [
      { type: "file", name: "file", label: "Upload PDF", accept: ".pdf", required: true },
      { type: "file", name: "signature", label: "Upload signature image", accept: ".png,.jpg,.jpeg", required: true },
      { type: "number", name: "page", label: "Page number", value: "1", required: true },
      { type: "number", name: "x", label: "X position", value: "60", required: true },
      { type: "number", name: "y", label: "Y position", value: "80", required: true },
      { type: "number", name: "width", label: "Signature width", value: "140", required: true }
    ],
    handler: esignPdf
  },
  {
    id: "html-to-pdf",
    name: "HTML to PDF",
    status: "live",
    category: "Convert",
    output: "PDF export",
    description: "Paste HTML markup and export it as a PDF document.",
    fields: [
      { type: "textarea", name: "html", label: "HTML content", value: "<div style='padding:32px;font-family:Arial'><h1>Proposal</h1><p>Paste your HTML here.</p></div>", required: true }
    ],
    handler: htmlToPdf
  },
  {
    id: "pdf-to-word",
    name: "PDF to Word",
    status: "live",
    category: "Convert",
    output: "DOCX document",
    description: "Export the PDF into a layout-preserving Word document.",
    fields: [
      { type: "file", name: "file", label: "Upload PDF", accept: ".pdf", required: true },
      { type: "select", name: "mode", label: "Conversion mode", options: ["layout+editable", "layout-preserving", "editable-text"] },
      { type: "select", name: "exportFormat", label: "Edited export", options: ["word+pdf", "word", "pdf"] }
    ],
    handler: pdfToWord
  },
  {
    id: "word-to-pdf",
    name: "Word to PDF",
    status: "live",
    category: "Convert",
    output: "PDF document",
    description: "Convert a DOCX or TXT file into a readable text-based PDF.",
    fields: [
      { type: "file", name: "file", label: "Upload DOCX or TXT", accept: ".docx,.txt", required: true }
    ],
    handler: wordToPdf
  },
  {
    id: "scan",
    name: "Scan to PDF",
    status: "live",
    category: "Workflow",
    output: "Scanned PDF",
    description: "Capture document photos and turn them into a PDF.",
    fields: [
      { type: "file", name: "files", label: "Upload camera images", accept: "image/*", multiple: true, required: true, help: "On mobile, this will let you capture photos directly." },
      { type: "checkbox", name: "fit", label: "Fit scans to page", checked: true }
    ],
    handler: scanToPdf
  },
  {
    id: "compare",
    name: "Compare PDFs",
    status: "live",
    category: "Review",
    output: "Comparison report",
    description: "Compare the extracted text of two PDFs and download a change report.",
    fields: [
      { type: "file", name: "left", label: "First PDF", accept: ".pdf", required: true },
      { type: "file", name: "right", label: "Second PDF", accept: ".pdf", required: true }
    ],
    handler: comparePdfs
  }
];

const state = {
  currentTool: toolRegistry[0].id,
  downloadUrl: null,
  transientUrls: [],
  editorHistory: [],
  editorTool: "select"
};

const grid = document.querySelector("#tool-grid");
const toolSelector = document.querySelector("#tool-selector");
const toolDescription = document.querySelector("#tool-description");
const toolStatusBadge = document.querySelector("#tool-status-badge");
const toolOutput = document.querySelector("#tool-output");
const toolForm = document.querySelector("#tool-form");
const resultPanel = document.querySelector("#result-panel");
const resultSummary = document.querySelector("#result-summary");
const resultActions = document.querySelector("#result-actions");
const resultPreview = document.querySelector("#result-preview");
const workspaceSection = document.querySelector("#workspace");

bootstrap();

function bootstrap() {
  renderCatalog();
  renderToolSelector();
  renderActiveTool();
}

function renderCatalog() {
  grid.innerHTML = toolRegistry.map((tool) => `
    <button class="tool-card" type="button" data-tool-id="${tool.id}" aria-label="Open ${tool.name}">
      <div class="tool-chip-row">
        <span class="chip">${tool.category}</span>
        <span class="${tool.status === "live" ? "chip status-live" : "chip status-roadmap"}">${tool.status === "live" ? "Live now" : "Roadmap ready"}</span>
      </div>
      <h3>${tool.name}</h3>
      <p>${tool.description}</p>
      <div class="tool-chip-row">
        <span class="chip">${tool.output}</span>
      </div>
    </button>
  `).join("");

  [...grid.querySelectorAll(".tool-card[data-tool-id]")].forEach((card) => {
    card.addEventListener("click", () => {
      const { toolId } = card.dataset;
      if (!toolId) {
        return;
      }
      state.currentTool = toolId;
      renderActiveTool();
      workspaceSection?.scrollIntoView({ behavior: "smooth", block: "start" });
    });
  });
}

function renderToolSelector() {
  toolSelector.innerHTML = toolRegistry
    .map((tool) => `
      <button
        class="tool-selector-button"
        type="button"
        data-tool-id="${tool.id}"
        aria-pressed="${tool.id === state.currentTool ? "true" : "false"}"
      >${tool.name}</button>
    `)
    .join("");

  [...toolSelector.querySelectorAll(".tool-selector-button[data-tool-id]")].forEach((button) => {
    button.addEventListener("click", () => {
      const { toolId } = button.dataset;
      if (!toolId) {
        return;
      }
      state.currentTool = toolId;
      renderActiveTool();
      workspaceSection?.scrollIntoView({ behavior: "smooth", block: "start" });
    });
  });
}

function renderActiveTool() {
  clearResult();
  const tool = getTool();
  [...toolSelector.querySelectorAll(".tool-selector-button[data-tool-id]")].forEach((button) => {
    const isActive = button.dataset.toolId === tool.id;
    button.classList.toggle("is-selected", isActive);
    button.setAttribute("aria-pressed", isActive ? "true" : "false");
  });
  [...grid.querySelectorAll(".tool-card[data-tool-id]")].forEach((card) => {
    card.classList.toggle("is-selected", card.dataset.toolId === tool.id);
  });
  toolDescription.textContent = tool.description;
  toolStatusBadge.textContent = tool.status === "live" ? "Live in Browser" : "Roadmap Integration";
  toolStatusBadge.className = `status-badge ${tool.status === "live" ? "status-live" : "status-roadmap"}`;
  toolOutput.textContent = tool.output;

  if (tool.status !== "live") {
    toolForm.innerHTML = `
      <div class="inline-notice">
        This tool card is product-ready in the UI, but not wired to a processing engine yet. The homepage and architecture are ready for a backend/API integration.
      </div>
    `;
    toolForm.onsubmit = null;
    return;
  }

  toolForm.innerHTML = `
    ${tool.fields.map(renderField).join("")}
    <div class="action-row">
      <button class="button button-primary" type="submit">Process ${tool.name}</button>
      <button class="button button-secondary" type="button" id="reset-form">Reset</button>
    </div>
  `;

  enhanceCustomSelects(toolForm);
  toolForm.onsubmit = handleSubmit;
  document.querySelector("#reset-form").addEventListener("click", () => renderActiveTool());
}

function renderField(field) {
  if (field.type === "checkbox") {
    return `
      <label class="checkbox">
        <input type="checkbox" name="${field.name}" ${field.checked ? "checked" : ""}>
        <span>${field.label}</span>
      </label>
    `;
  }

  if (field.type === "select") {
    return `
      <div class="field">
        <label for="${field.name}">${field.label}</label>
        <div class="custom-select" data-name="${field.name}">
          <select class="custom-select-native" id="${field.name}" name="${field.name}" ${field.required ? "required" : ""}>
            ${field.options.map((option) => `<option value="${option}" ${option === (field.value || field.options[0]) ? "selected" : ""}>${option}</option>`).join("")}
          </select>
          <button class="custom-select-trigger" type="button" aria-haspopup="listbox" aria-expanded="false">
            <span class="custom-select-label">${field.value || field.options[0]}</span>
          </button>
          <div class="custom-select-menu" role="listbox">
            ${field.options.map((option) => `
              <button class="custom-select-option ${option === (field.value || field.options[0]) ? "is-selected" : ""}" type="button" data-value="${option}" role="option" aria-selected="${option === (field.value || field.options[0]) ? "true" : "false"}">${option}</button>
            `).join("")}
          </div>
        </div>
        ${field.help ? `<small>${field.help}</small>` : ""}
      </div>
    `;
  }

  if (field.type === "textarea") {
    return `
      <div class="field">
        <label for="${field.name}">${field.label}</label>
        <textarea
          id="${field.name}"
          name="${field.name}"
          ${field.required ? "required" : ""}
          ${field.placeholder ? `placeholder="${field.placeholder}"` : ""}
        >${field.value || ""}</textarea>
        ${field.help ? `<small>${field.help}</small>` : ""}
      </div>
    `;
  }

  return `
    <div class="field">
      <label for="${field.name}">${field.label}</label>
      <input
        id="${field.name}"
        name="${field.name}"
        type="${field.type}"
        ${field.accept ? `accept="${field.accept}"` : ""}
        ${field.multiple ? "multiple" : ""}
        ${field.required ? "required" : ""}
        ${field.placeholder ? `placeholder="${field.placeholder}"` : ""}
        ${field.value ? `value="${field.value}"` : ""}
      >
      ${field.help ? `<small>${field.help}</small>` : ""}
    </div>
  `;
}

function enhanceCustomSelects(root) {
  const selects = [...root.querySelectorAll(".custom-select")];
  const closeAll = () => {
    selects.forEach((select) => {
      select.classList.remove("is-open");
      select.querySelector(".custom-select-trigger")?.setAttribute("aria-expanded", "false");
    });
  };

  selects.forEach((select) => {
    const nativeSelect = select.querySelector(".custom-select-native");
    const trigger = select.querySelector(".custom-select-trigger");
    const label = select.querySelector(".custom-select-label");
    const options = [...select.querySelectorAll(".custom-select-option")];
    if (!nativeSelect || !trigger || !label) {
      return;
    }

    const syncValue = (value) => {
      nativeSelect.value = value;
      label.textContent = value;
      options.forEach((option) => {
        const isSelected = option.dataset.value === value;
        option.classList.toggle("is-selected", isSelected);
        option.setAttribute("aria-selected", isSelected ? "true" : "false");
      });
      nativeSelect.dispatchEvent(new Event("change", { bubbles: true }));
    };

    trigger.addEventListener("click", (event) => {
      event.preventDefault();
      const willOpen = !select.classList.contains("is-open");
      closeAll();
      select.classList.toggle("is-open", willOpen);
      trigger.setAttribute("aria-expanded", willOpen ? "true" : "false");
    });

    options.forEach((option) => {
      option.addEventListener("click", () => {
        syncValue(option.dataset.value || "");
        closeAll();
      });
    });
  });

  document.addEventListener("click", (event) => {
    if (!event.target.closest(".custom-select")) {
      closeAll();
    }
  }, { once: true });
}

async function handleSubmit(event) {
  event.preventDefault();
  const tool = getTool();
  setResult(`Processing ${tool.name}...`, []);
  try {
    await tool.handler(new FormData(toolForm));
  } catch (error) {
    console.error(error);
    setResult(error.message || "Something went wrong while processing the file.", []);
  }
}

function getTool() {
  return toolRegistry.find((tool) => tool.id === state.currentTool);
}

function parseRanges(input, pageCount) {
  const uniquePages = new Set();
  input.split(",").map((part) => part.trim()).filter(Boolean).forEach((token) => {
    if (token.includes("-")) {
      const [startText, endText] = token.split("-").map((value) => value.trim());
      const start = Number(startText);
      const end = Number(endText);
      if (!start || !end || start > end) {
        throw new Error(`Invalid range: ${token}`);
      }
      for (let page = start; page <= end; page += 1) {
        validatePage(page, pageCount);
        uniquePages.add(page);
      }
      return;
    }
    const page = Number(token);
    validatePage(page, pageCount);
    uniquePages.add(page);
  });
  return [...uniquePages].sort((a, b) => a - b);
}

function validatePage(page, pageCount) {
  if (!Number.isInteger(page) || page < 1 || page > pageCount) {
    throw new Error(`Page ${page} is out of range. This PDF has ${pageCount} page(s).`);
  }
}

async function fileToBytes(file) {
  return new Uint8Array(await file.arrayBuffer());
}

async function fileToArrayBuffer(file) {
  return await file.arrayBuffer();
}

async function createBackendEditSession(file) {
  if (!BACKEND_CONFIG.enabled) {
    return null;
  }

  const formData = new FormData();
  formData.append("file", file);
  const response = await fetch(`${BACKEND_CONFIG.apiBaseUrl}/editor/session`, {
    method: "POST",
    body: formData
  });

  if (!response.ok) {
    throw new Error("Backend session creation failed.");
  }

  return await response.json();
}

function backendAssetUrl(publicPath) {
  if (!publicPath) {
    return null;
  }

  const baseUrl = BACKEND_CONFIG.apiBaseUrl.replace(/\/api$/, "");
  return `${baseUrl}${publicPath}`;
}

async function updateBackendTextBlock(sessionId, blockId, text) {
  const response = await fetch(`${BACKEND_CONFIG.apiBaseUrl}/editor/session/${sessionId}/blocks/${blockId}`, {
    method: "PATCH",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify({ text })
  });

  if (!response.ok) {
    throw new Error("Failed to save an edited text block to the backend session.");
  }

  return await response.json();
}

async function updateBackendBlockProperties(sessionId, blockId, patch) {
  const response = await fetch(`${BACKEND_CONFIG.apiBaseUrl}/editor/session/${sessionId}/blocks/${blockId}/properties`, {
    method: "PATCH",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify(patch)
  });

  if (!response.ok) {
    throw new Error("Failed to sync edited text block properties to the backend session.");
  }

  return await response.json();
}

async function exportBackendSession(sessionId) {
  const response = await fetch(`${BACKEND_CONFIG.apiBaseUrl}/editor/session/${sessionId}/export/pdf`, {
    method: "POST"
  });

  if (!response.ok) {
    throw new Error("Backend PDF export failed.");
  }

  return await response.json();
}

function buildTextPagesFromBackendSession(session) {
  return session.pages.map((page) => ({
    pageNumber: page.pageNumber,
    width: page.width,
    height: page.height,
    items: session.editableBlocks
      .filter((block) => block.pageId === page.id)
      .map((block) => ({
        text: block.text,
        left: Number(block.x || 0),
        top: Number(block.y || 0),
        width: Number(block.width || 0),
        height: Number(block.height || block.fontSize || 12),
        fontSize: Number(block.fontSize || 12),
        fontFamily: block.fontFamily || "Arial, Helvetica, sans-serif",
        fontWeight: block.fontWeight || "400",
        fontStyle: block.fontStyle || "normal",
        dir: "ltr",
        angle: 0,
        scaleX: 1,
        blockId: block.id
      }))
  }));
}

function buildBackendBlockPatch(block) {
  const page = block.closest(".editor-page");
  const content = block.querySelector(".editor-text-content, .editor-text-input");
  const computedBlockStyle = window.getComputedStyle(block);
  const computedTextStyle = content ? window.getComputedStyle(content) : computedBlockStyle;
  const scale = Number.parseFloat(page?.dataset.pageScale || "1") || 1;
  const left = parseCssPixels(block.style.left) / scale;
  const top = parseCssPixels(block.style.top) / scale;
  const width = parseCssPixels(block.style.width || computedBlockStyle.width) / scale;
  const height = parseCssPixels(block.style.height || computedBlockStyle.height) / scale;
  const fontSize = parseCssPixels(computedTextStyle.fontSize, parseCssPixels(block.style.fontSize, 16)) / scale;
  const lineHeight = parseCssPixels(computedTextStyle.lineHeight, Math.max(fontSize, height || fontSize)) / scale;

  return {
    x: Number(left.toFixed(2)),
    y: Number(top.toFixed(2)),
    width: Number(width.toFixed(2)),
    height: Number(Math.max(height, fontSize).toFixed(2)),
    fontSize: Number(Math.max(6, fontSize).toFixed(2)),
    lineHeight: Number(Math.max(fontSize, lineHeight).toFixed(2)),
    fontFamily: computedTextStyle.fontFamily || block.dataset.fontFamily || "Arial, Helvetica, sans-serif",
    fontWeight: computedTextStyle.fontWeight || block.dataset.fontWeight || "400",
    fontStyle: computedTextStyle.fontStyle || block.dataset.fontStyle || "normal",
    color: computedTextStyle.color || "#111111",
    textAlign: computedTextStyle.textAlign || "left"
  };
}

async function syncBackendSessionFromPreview(previewNode, sessionId) {
  const blocks = [...previewNode.querySelectorAll(".editor-text-block[data-block-id][data-modified='true']")];
  for (const block of blocks) {
    await updateBackendTextBlock(sessionId, block.dataset.blockId, block.dataset.currentText || "");
    await updateBackendBlockProperties(sessionId, block.dataset.blockId, buildBackendBlockPatch(block));
    block.dataset.originalText = block.dataset.currentText || "";
    block.dataset.modified = "false";
    block.classList.remove("is-modified");
  }
}

function hasUnsupportedBackendOverlays(previewNode) {
  if (!previewNode) {
    return false;
  }

  return [...previewNode.querySelectorAll(".editor-overlay, .editor-text-block")].some((node) => {
    if (node.classList.contains("editor-text-block")) {
      return !node.dataset.blockId;
    }

    return true;
  });
}

async function blobToObjectUrl(blob) {
  const href = URL.createObjectURL(blob);
  state.transientUrls.push(href);
  return href;
}

async function downloadBlob(blob, filename, message, label = "Download File") {
  const href = await blobToObjectUrl(blob);
  setResult(message, [{ label, href, download: filename }]);
}

async function dataUrlToBytes(dataUrl) {
  const response = await fetch(dataUrl);
  return new Uint8Array(await response.arrayBuffer());
}

async function renderPdfPages(file, options = {}) {
  const data = new Uint8Array(await fileToArrayBuffer(file));
  const loadingTask = pdfjsLib.getDocument({
    data,
    password: options.password || undefined
  });
  const pdf = await loadingTask.promise;
  const scale = Number(options.scale || 1.5);
  const pages = [];

  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
    const page = await pdf.getPage(pageNumber);
    const viewport = page.getViewport({ scale });
    const canvas = document.createElement("canvas");
    const context = canvas.getContext("2d");
    canvas.width = viewport.width;
    canvas.height = viewport.height;
    await page.render({ canvasContext: context, viewport }).promise;
    pages.push({ pageNumber, canvas, width: viewport.width, height: viewport.height });
  }

  return pages;
}

async function detectPdfPageImages(file, password = "") {
  const data = new Uint8Array(await fileToArrayBuffer(file));
  const pdf = await pdfjsLib.getDocument({
    data,
    password: password || undefined
  }).promise;
  const pageFlags = [];
  const imageOps = new Set([
    pdfjsLib.OPS.paintImageXObject,
    pdfjsLib.OPS.paintInlineImageXObject,
    pdfjsLib.OPS.paintImageMaskXObject,
    pdfjsLib.OPS.paintJpegXObject
  ]);

  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
    const page = await pdf.getPage(pageNumber);
    const operatorList = await page.getOperatorList();
    const hasImages = operatorList.fnArray.some((fn) => imageOps.has(fn));
    pageFlags.push({ pageNumber, hasImages });
  }

  return pageFlags;
}

async function extractPdfText(file, password = "") {
  const data = new Uint8Array(await fileToArrayBuffer(file));
  const pdf = await pdfjsLib.getDocument({
    data,
    password: password || undefined
  }).promise;
  const pages = [];

  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
    const page = await pdf.getPage(pageNumber);
    const content = await page.getTextContent();
    const text = content.items.map((item) => item.str).join(" ").replace(/\s+/g, " ").trim();
    pages.push({ pageNumber, text });
  }

  return pages;
}

async function extractPdfTextLayout(file, password = "", scale = 1.5) {
  const data = new Uint8Array(await fileToArrayBuffer(file));
  const pdf = await pdfjsLib.getDocument({
    data,
    password: password || undefined
  }).promise;
  const pages = [];

  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
    const page = await pdf.getPage(pageNumber);
    const viewport = page.getViewport({ scale });
    const content = await page.getTextContent();
    const items = applyFormattingHeuristics(dedupeTextItems(content.items
      .filter((item) => item.str && item.str.trim())
      .map((item) => {
      const transform = pdfjsLib.Util.transform(viewport.transform, item.transform);
      const style = content.styles[item.fontName] || {};
      const angle = Math.atan2(transform[1], transform[0]);
      const fontHeight = Math.hypot(transform[2], transform[3]);
      const ascent = typeof style.ascent === "number"
        ? style.ascent
        : (typeof style.descent === "number" ? 1 + style.descent : 0.8);
      const left = transform[4];
      const top = transform[5] - (fontHeight * ascent);
      const traits = detectFontTraits(item.fontName || "", style.fontFamily || "");
      const fontFamily = sanitizeFontFamily(style.fontFamily || item.fontName || "Arial");
      const measuredWidth = Math.max(
        1,
        measureStyledTextWidth(item.str || "", fontHeight, fontFamily, traits.fontWeight, traits.fontStyle)
      );
      const pdfWidth = Math.max(item.width ? item.width * scale : 0, item.str.length * fontHeight * 0.25);
      const width = Math.max(pdfWidth, measuredWidth);
      const height = Math.max(fontHeight, item.height ? item.height * scale : fontHeight);

      return {
        text: item.str || "",
        left,
        top,
        width,
        height,
        fontSize: fontHeight,
        fontFamily,
        fontName: item.fontName || "",
        dir: item.dir || "ltr",
        angle,
        scaleX: item.str.trim() ? Math.max(0.85, Math.min(1.25, pdfWidth / measuredWidth)) : 1,
        ...traits
      };
    })));

    pages.push({
      pageNumber,
      width: viewport.width,
      height: viewport.height,
      items
    });
  }

  return pages;
}

function sanitizeFontFamily(fontFamily) {
  const rawValue = String(fontFamily || "Arial")
    .replaceAll("+", " ")
    .replace(/["']/g, "")
    .trim();

  const value = rawValue.toLowerCase();

  if (!value) {
    return '"Times New Roman", Georgia, serif';
  }

  if (
    value.includes("times") ||
    value.includes("roman") ||
    value.includes("serif") ||
    value.includes("garamond") ||
    value.includes("cambria") ||
    value.includes("georgia") ||
    value.includes("bookman")
  ) {
    return '"Times New Roman", Georgia, serif';
  }

  if (
    value.includes("arial") ||
    value.includes("helvetica") ||
    value.includes("sans") ||
    value.includes("calibri") ||
    value.includes("verdana") ||
    value.includes("tahoma") ||
    value.includes("franklin")
  ) {
    return 'Arial, Helvetica, sans-serif';
  }

  if (
    value.includes("courier") ||
    value.includes("mono") ||
    value.includes("consolas") ||
    value.includes("menlo")
  ) {
    return '"Courier New", Consolas, monospace';
  }

  return `"${rawValue}", "Times New Roman", Georgia, serif`;
}

function detectFontTraits(fontName = "", fontFamily = "") {
  const combined = `${fontName} ${fontFamily}`
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ");
  const boldLike = /\b(bold|black|heavy|extrabold|ultrabold|semibold|demibold|bd|sb)\b/.test(combined);
  const mediumLike = /\b(medium|demi|semi)\b/.test(combined);
  const italicLike = /\b(italic|oblique|it)\b/.test(combined);
  const fontWeight = boldLike ? "700" : (mediumLike ? "600" : "400");
  const fontStyle = italicLike
    ? "italic"
    : "normal";
  return { fontWeight, fontStyle };
}

function measureStyledTextWidth(text, fontSize, fontFamily, fontWeight = "400", fontStyle = "normal") {
  const canvas = measureStyledTextWidth.canvas || (measureStyledTextWidth.canvas = document.createElement("canvas"));
  const context = canvas.getContext("2d");
  context.font = `${fontStyle} ${fontWeight} ${fontSize}px ${fontFamily}`;
  return context.measureText(text).width;
}

function dedupeTextItems(items) {
  const deduped = [];

  items.forEach((item) => {
    const duplicate = deduped.find((existing) =>
      existing.text === item.text &&
      Math.abs(existing.left - item.left) < 1.5 &&
      Math.abs(existing.top - item.top) < 1.5 &&
      Math.abs(existing.width - item.width) < 2 &&
      Math.abs(existing.fontSize - item.fontSize) < 1
    );

    if (!duplicate) {
      deduped.push(item);
    }
  });

  return deduped;
}

function applyFormattingHeuristics(items) {
  if (!items.length) {
    return items;
  }

  const fontSizes = items
    .map((item) => item.fontSize)
    .filter((value) => Number.isFinite(value))
    .sort((left, right) => left - right);

  const medianFontSize = fontSizes[Math.floor(fontSizes.length / 2)] || 12;

  return items.map((item) => {
    const text = item.text.trim();
    const looksLikeHeading =
      /^q\d+\s*[:.]/i.test(text) ||
      /^(objective|task|sample|code|class|date|name|uid|section|output)\b[:]?$/i.test(text) ||
      (text.endsWith(":") && text.length <= 32) ||
      (/^[A-Z0-9 .:/_-]+$/.test(text) && text.length <= 40);

    const looksVisuallyProminent =
      item.fontSize >= medianFontSize * 1.16 &&
      text.length <= 100;

    const looksLikeShortLabel =
      text.length <= 18 &&
      (
        /^(objective|task|sample|code|class|date|name|uid|section|output)$/i.test(text.replace(/:$/, "")) ||
        /^[-*]$/.test(text) ||
        text.endsWith(":")
      );

    const shouldPromoteBold =
      item.fontWeight === "400" &&
      (looksLikeShortLabel || looksLikeHeading || (looksVisuallyProminent && text.length <= 40));

    if (!shouldPromoteBold) {
      return item;
    }

    return {
      ...item,
      fontWeight: "700"
    };
  });
}

function buildEditableTextPages(textPages) {
  return textPages.map((page) => ({
    ...page,
    items: groupItemsIntoEditableLines(page.items)
  }));
}

function snapPointToNearestTextBlock(page, x, y) {
  const blocks = [...page.querySelectorAll(".editor-text-block")];
  if (!blocks.length) {
    return { x, y };
  }

  let best = null;
  let bestDistance = Number.POSITIVE_INFINITY;

  blocks.forEach((block) => {
    const left = Number.parseFloat(block.style.left || "0");
    const top = Number.parseFloat(block.style.top || "0");
    const width = Number.parseFloat(block.style.width || "0");
    const height = Number.parseFloat(block.style.height || "0");
    const centerX = left + (width / 2);
    const centerY = top + (height / 2);
    const distance = Math.hypot(centerX - x, centerY - y);
    if (distance < bestDistance) {
      bestDistance = distance;
      best = {
        x: left,
        y: top,
        width,
        height,
        fontFamily: block.dataset.fontFamily || "Arial, sans-serif",
        fontWeight: block.dataset.fontWeight || "400",
        fontStyle: block.dataset.fontStyle || "normal",
        fontSize: Number.parseFloat(block.style.fontSize || "16"),
        lineHeight: block.style.lineHeight || ""
      };
    }
  });

  if (best && bestDistance <= 28) {
    return {
      x: best.x,
      y: best.y,
      width: best.width,
      height: best.height,
      fontFamily: best.fontFamily,
      fontWeight: best.fontWeight,
      fontStyle: best.fontStyle,
      fontSize: best.fontSize,
      lineHeight: best.lineHeight,
      snapped: true
    };
  }

  return { x, y };
}

function groupItemsIntoEditableLines(items) {
  const sorted = [...items].sort((a, b) => {
    const topDiff = a.top - b.top;
    if (Math.abs(topDiff) > Math.max(4, Math.min(a.fontSize, b.fontSize) * 0.45)) {
      return topDiff;
    }
    return a.left - b.left;
  });

  const lines = [];

  sorted.forEach((item) => {
    const last = lines[lines.length - 1];
    if (!last || Math.abs(last.top - item.top) > Math.max(5, item.fontSize * 0.5)) {
      lines.push({
        top: item.top,
        left: item.left,
        right: item.left + item.width,
        height: item.height,
        fontSize: item.fontSize,
        fontFamily: item.fontFamily,
        fontWeight: item.fontWeight,
        fontStyle: item.fontStyle,
        dir: item.dir,
        segments: [item]
      });
      return;
    }

    last.segments.push(item);
    last.left = Math.min(last.left, item.left);
    last.right = Math.max(last.right, item.left + item.width);
    last.height = Math.max(last.height, item.height);
    last.fontSize = Math.max(last.fontSize, item.fontSize);
    if (item.fontWeight === "700") {
      last.fontWeight = "700";
    }
    if (item.fontStyle === "italic") {
      last.fontStyle = "italic";
    }
  });

  return lines.map((line) => {
    const ordered = line.segments.sort((a, b) => a.left - b.left);
    let text = "";
    let cursor = line.left;

    ordered.forEach((segment, index) => {
      const gap = segment.left - cursor;
      const addSpace = index > 0 && gap > Math.max(segment.fontSize * 0.22, 6);
      if (addSpace && !text.endsWith(" ")) {
        text += " ";
      }
      text += segment.text;
      cursor = segment.left + segment.width;
    });

    const normalizedText = text
      .replace(/\s+([,.;:])/g, "$1")
      .replace(/([(:])\s+/g, "$1")
      .replace(/\s{2,}/g, " ")
      .trim();

    const width = Math.max(
      line.right - line.left,
      measureStyledTextWidth(normalizedText, line.fontSize, line.fontFamily, line.fontWeight, line.fontStyle) * 1.02
    );

    return {
      text: normalizedText,
      left: line.left,
      top: line.top,
      width,
      height: line.height,
      fontSize: line.fontSize,
      fontFamily: line.fontFamily,
      fontWeight: line.fontWeight,
      fontStyle: line.fontStyle,
      dir: line.dir,
      angle: 0
    };
  }).filter((line) => line.text);
}

async function buildPdfFromRenderedPages(renderedPages, options = {}) {
  const quality = Number(options.quality || 0.75);
  const encrypted = Boolean(options.encryption);
  let doc;

  renderedPages.forEach((page, index) => {
    const imageData = page.canvas.toDataURL("image/jpeg", quality);
    if (index === 0) {
      doc = new jsPDF({
        orientation: page.width >= page.height ? "landscape" : "portrait",
        unit: "pt",
        format: [page.width, page.height],
        compress: true,
        encryption: encrypted ? options.encryption : undefined
      });
    } else {
      doc.addPage([page.width, page.height], page.width >= page.height ? "landscape" : "portrait");
    }
    doc.addImage(imageData, "JPEG", 0, 0, page.width, page.height, undefined, "FAST");
  });

  return doc.output("blob");
}

async function buildTextPdf(text, title = "Document") {
  const pdfDoc = await PDFDocument.create();
  const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
  const bold = await pdfDoc.embedFont(StandardFonts.HelveticaBold);
  const pageSize = { width: 595, height: 842 };
  const margin = 48;
  const lineHeight = 16;
  const maxWidth = pageSize.width - margin * 2;
  const lines = [];

  lines.push({ text: title, font: bold, size: 18 });
  lines.push({ text: "", font, size: 12 });

  text.split(/\r?\n/).forEach((paragraph) => {
    wrapText(paragraph || " ", font, 11, maxWidth).forEach((line) => lines.push({ text: line, font, size: 11 }));
    lines.push({ text: "", font, size: 11 });
  });

  let page = pdfDoc.addPage([pageSize.width, pageSize.height]);
  let y = pageSize.height - margin;

  lines.forEach((line) => {
    if (y < margin) {
      page = pdfDoc.addPage([pageSize.width, pageSize.height]);
      y = pageSize.height - margin;
    }
    if (line.text) {
      page.drawText(line.text, {
        x: margin,
        y,
        size: line.size,
        font: line.font,
        color: rgb(0.15, 0.15, 0.15)
      });
    }
    y -= lineHeight;
  });

  return new Blob([await pdfDoc.save()], { type: "application/pdf" });
}

function wrapText(text, font, size, maxWidth) {
  const words = text.split(/\s+/).filter(Boolean);
  if (!words.length) {
    return [""];
  }
  const lines = [];
  let current = words[0];

  for (let index = 1; index < words.length; index += 1) {
    const test = `${current} ${words[index]}`;
    if (font.widthOfTextAtSize(test, size) <= maxWidth) {
      current = test;
    } else {
      lines.push(current);
      current = words[index];
    }
  }

  lines.push(current);
  return lines;
}

function wireOriginalPreservingTextBlock(block) {
  const content = block.querySelector(".editor-text-content, .editor-text-input");
  if (!content) {
    return;
  }

  const resizeBlockToFitText = () => {
    const text = "value" in content ? content.value : (content.textContent || "");
    const fontSize = Number.parseFloat(content.style.fontSize || block.style.fontSize || "16");
    const fontFamily = block.dataset.fontFamily || "Arial, sans-serif";
    const fontWeight = block.dataset.fontWeight || "400";
    const fontStyle = block.dataset.fontStyle || "normal";
    const scaleX = Number.parseFloat(block.dataset.scaleX || "1") || 1;
    const minWidth = Number.parseFloat(block.style.minWidth || block.style.width || "10");
    const measuredWidth = measureStyledTextWidth(text || " ", fontSize, fontFamily, fontWeight, fontStyle);
    const desiredWidth = Math.max(minWidth, (measuredWidth * scaleX) + Math.max(12, fontSize * 0.45));
    const page = block.closest(".editor-page");
    const blockLeft = Number.parseFloat(block.style.left || "0");
    const maxWidth = page
      ? Math.max(minWidth, (page.clientWidth || page.offsetWidth || desiredWidth) - blockLeft - 12)
      : desiredWidth;

    block.style.width = `${Math.min(desiredWidth, maxWidth)}px`;
  };

  if (typeof block.dataset.currentText !== "string") {
    block.dataset.currentText = "value" in content
      ? content.value
      : (content.textContent || block.dataset.originalText || "");
  }

  const syncState = () => {
    const rawCurrent = "value" in content ? content.value : (content.textContent || "");
    block.dataset.currentText = rawCurrent;
    const current = rawCurrent.replace(/\u00a0/g, " ").replace(/\s+/g, " ").trim();
    const original = (block.dataset.originalText || "").replace(/\u00a0/g, " ").replace(/\s+/g, " ").trim();
    const modified = current !== original;
    block.dataset.modified = modified ? "true" : "false";
    block.classList.toggle("is-modified", modified);
    resizeBlockToFitText();
  };

  syncState();

  content.addEventListener("focus", () => {
    if ("value" in content) {
      content.value = block.dataset.currentText || "";
    } else {
      content.textContent = block.dataset.currentText || "";
    }
    block.classList.add("is-active-edit");
  });

  content.addEventListener("blur", () => {
    syncState();
    block.classList.remove("is-active-edit");
    if ("value" in content) {
      content.value = block.dataset.currentText || "";
    } else {
      content.textContent = block.dataset.currentText || "";
    }
  });

  content.addEventListener("input", syncState);
  content.addEventListener("keyup", syncState);
  content.addEventListener("focus", resizeBlockToFitText);
  content.addEventListener("paste", () => requestAnimationFrame(syncState));
  content.addEventListener("cut", () => requestAnimationFrame(syncState));
}

function getEditorBlockText(block) {
  const content = block?.querySelector(".editor-text-content, .editor-text-input");
  if (!content) {
    return block?.dataset.currentText || "";
  }
  return block.dataset.currentText || ("value" in content ? content.value : (content.textContent || ""));
}

function refreshEditorBlockState(block) {
  const content = block?.querySelector(".editor-text-content, .editor-text-input");
  if (!block || !content) {
    return;
  }

  const rawCurrent = "value" in content ? content.value : (content.textContent || "");
  block.dataset.currentText = rawCurrent;

  const current = rawCurrent.replace(/\u00a0/g, " ").replace(/\s+/g, " ").trim();
  const original = (block.dataset.originalText || "").replace(/\u00a0/g, " ").replace(/\s+/g, " ").trim();
  const modified = current !== original;

  block.dataset.modified = modified ? "true" : "false";
  block.classList.toggle("is-modified", modified);
}

function refreshPreviewEditorState(previewNode) {
  if (!previewNode) {
    return;
  }

  previewNode.querySelectorAll(".editor-text-block").forEach((block) => {
    refreshEditorBlockState(block);
  });
}

function setEditorBlockText(block, text) {
  const content = block?.querySelector(".editor-text-content, .editor-text-input");
  block.dataset.currentText = text;
  block.dataset.modified = "true";
  block.classList.add("is-modified");
  if (!content) {
    return;
  }
  if ("value" in content) {
    content.value = text;
  } else {
    content.textContent = text;
  }
}

async function buildDocxFromText(title, pages) {
  const paragraphs = [
    new window.docx.Paragraph({
      text: title,
      heading: window.docx.HeadingLevel.TITLE
    })
  ];

  pages.forEach((page, index) => {
    paragraphs.push(new window.docx.Paragraph({ text: `Page ${page.pageNumber || index + 1}`, heading: window.docx.HeadingLevel.HEADING_1 }));
    paragraphs.push(new window.docx.Paragraph({ text: page.text || "" }));
  });

  const document = new window.docx.Document({
    sections: [{ properties: {}, children: paragraphs }]
  });

  return await window.docx.Packer.toBlob(document);
}

async function buildDocxFromRenderedPages(title, renderedPages) {
  const pageWidthTwips = 11906;
  const pageHeightTwips = 16838;
  const marginTwips = 360;
  const usableWidthPx = 720;
  const usableHeightPx = 1040;
  const children = [
    new window.docx.Paragraph({
      text: title,
      heading: window.docx.HeadingLevel.TITLE
    }),
    new window.docx.Paragraph({ text: "" })
  ];

  for (let index = 0; index < renderedPages.length; index += 1) {
    const renderedPage = renderedPages[index];
    const imageData = renderedPage.canvas.toDataURL("image/png");
    const imageBytes = await dataUrlToBytes(imageData);
    const scale = Math.min(usableWidthPx / renderedPage.width, usableHeightPx / renderedPage.height);
    const width = Math.round(renderedPage.width * scale);
    const height = Math.round(renderedPage.height * scale);
    children.push(
      new window.docx.Paragraph({
        spacing: { after: 0, before: 0 },
        children: [
          new window.docx.ImageRun({
            data: imageBytes,
            transformation: {
              width,
              height
            }
          })
        ]
      })
    );

    if (index < renderedPages.length - 1) {
      children.push(new window.docx.Paragraph({ children: [new window.docx.PageBreak()] }));
    }
  }

  const document = new window.docx.Document({
    sections: [{
      properties: {
        page: {
          size: { width: pageWidthTwips, height: pageHeightTwips },
          margin: { top: marginTwips, right: marginTwips, bottom: marginTwips, left: marginTwips }
        }
      },
      children
    }]
  });

  return await window.docx.Packer.toBlob(document);
}

async function buildWordHtmlFromPdf(file) {
  const [renderedPages, textPages] = await Promise.all([
    renderPdfPages(file, { scale: 1.8 }),
    extractPdfTextLayout(file, "", 1.8)
  ]);

  const pageMarkup = await Promise.all(renderedPages.map(async (renderedPage, index) => {
    const pageText = textPages[index];
    const imageData = renderedPage.canvas.toDataURL("image/png");
    const spans = pageText.items
      .filter((item) => item.text && item.text.trim())
      .map((item) => {
        const safeText = escapeHtml(item.text);
        const left = item.left.toFixed(2);
        const top = item.top.toFixed(2);
        const fontSize = Math.max(8, item.fontSize).toFixed(2);
        const width = Math.max(8, item.width).toFixed(2);
        const angle = item.angle.toFixed(5);
        return `<div class="word-span" style="left:${left}px;top:${top}px;font-size:${fontSize}px;width:${width}px;font-family:${item.fontFamily};font-weight:${item.fontWeight};font-style:${item.fontStyle};transform:rotate(${angle}rad);">${safeText}</div>`;
      })
      .join("");

    return `
      <section class="word-page" style="width:${pageText.width}px;height:${pageText.height}px;">
        <img class="word-bg" src="${imageData}" alt="Page ${index + 1}">
        <div class="word-text-layer">${spans}</div>
      </section>
    `;
  }));

  const html = `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="utf-8">
      <title>PDF to Word Export</title>
      <style>
        @page { size: A4; margin: 0.35in; }
        body { font-family: Arial, sans-serif; background: #f4f1ea; margin: 0; }
        .word-page {
          position: relative;
          margin: 0 auto 24px;
          background: white;
          box-shadow: 0 8px 30px rgba(0,0,0,0.08);
          page-break-after: always;
          overflow: hidden;
        }
        .word-page:last-child { page-break-after: auto; }
        .word-bg {
          position: absolute;
          inset: 0;
          width: 100%;
          height: 100%;
          object-fit: fill;
          opacity: 0.22;
        }
        .word-text-layer {
          position: absolute;
          inset: 0;
        }
        .word-span {
          position: absolute;
          color: #111;
          line-height: 1;
          white-space: pre-wrap;
          overflow: visible;
        }
        .word-image {
          position: absolute;
          object-fit: contain;
        }
        .word-overlay.editor-whiteout {
          position: absolute;
          background: #ffffff;
          border: 1px solid rgba(24,34,31,0.1);
        }
        .word-overlay.editor-annotate {
          position: absolute;
          background: #fff5b3;
          border: 1px solid #d9bf4b;
          padding: 8px;
          white-space: pre-wrap;
        }
        .word-overlay.editor-shapes {
          position: absolute;
          border: 2px solid #0f6be8;
          background: transparent;
        }
      </style>
    </head>
    <body>
      ${pageMarkup.join("\n")}
    </body>
    </html>
  `;

  return new Blob([html], { type: "application/msword" });
}

function buildEditableWordPreview(renderedPages, textPages, imageFlags = [], options = {}) {
  const editablePages = options.advancedEditor ? textPages : buildEditableTextPages(textPages);
  const wrapper = document.createElement("div");
  wrapper.className = "editor-workspace-shell";
  const toolbar = document.createElement("div");
  toolbar.className = "editor-toolbar";
  toolbar.innerHTML = "<strong>PDF to Word Workspace</strong>";

  const toggle = document.createElement("div");
  toggle.className = "editor-toggle";
  const exactButton = document.createElement("button");
  exactButton.type = "button";
  exactButton.className = "is-active";
  exactButton.textContent = "Preview";
  const editButton = document.createElement("button");
  editButton.type = "button";
  editButton.textContent = "Enable Editing";
  toggle.append(exactButton, editButton);
  toolbar.appendChild(toggle);
  wrapper.appendChild(toolbar);

  const pagesHost = document.createElement("div");
  pagesHost.className = "preview-host";

  if (options.advancedEditor) {
    const mergeButton = document.createElement("button");
    mergeButton.type = "button";
    mergeButton.className = "editor-floating-merge-button";
    mergeButton.textContent = "Merge Boxes";
    mergeButton.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopPropagation();
      pagesHost._setEditorTool?.("merge");
    });
    wrapper.appendChild(mergeButton);
  }

  renderedPages.forEach((renderedPage, index) => {
    const textPage = editablePages[index];
    const imageFlag = imageFlags[index];
    const scale = Math.min(820 / textPage.width, 1);
    const pageShell = document.createElement("div");
    pageShell.className = "editor-page-shell";
    const page = document.createElement("section");
    page.className = "editor-page is-exact";
    if (options.advancedEditor) {
      page.classList.add("editor-preserve-background");
      page.dataset.preserveBackground = "true";
    }
    page.dataset.pageScale = String(scale);
    page.dataset.pdfWidth = String(textPage.width);
    page.dataset.pdfHeight = String(textPage.height);
    page.dataset.hasSourceImage = imageFlag?.hasImages ? "true" : "false";
    page.style.height = `${Math.round(textPage.height * scale)}px`;
    page.style.width = `${Math.round(textPage.width * scale)}px`;
    page.style.maxWidth = `${Math.round(textPage.width * scale)}px`;

    const image = document.createElement("img");
    image.src = renderedPage.canvas.toDataURL("image/png");
    image.alt = `Page ${index + 1}`;
    image.classList.add("editor-page-guide");
    page.appendChild(image);

    const layer = document.createElement("div");
    layer.className = "editor-text-layer";

    textPage.items
      .filter((item) => item.text && item.text.trim())
      .forEach((item) => {
        const block = document.createElement("div");
        block.className = "editor-text-block";
        block.dataset.page = String(index + 1);
        block.dataset.left = String(item.left);
        block.dataset.top = String(item.top);
        block.dataset.width = String(item.width || 0);
        block.dataset.height = String(item.height || item.fontSize);
        block.dataset.fontSize = String(item.fontSize);
        block.dataset.fontFamily = item.fontFamily;
        block.dataset.fontWeight = item.fontWeight;
        block.dataset.fontStyle = item.fontStyle;
        block.dataset.angle = String(item.angle || 0);
        block.dataset.scaleX = String(item.scaleX || 1);
        if (item.blockId) {
          block.dataset.blockId = item.blockId;
        }
        block.style.left = `${item.left * scale}px`;
        block.style.top = `${item.top * scale}px`;
        block.style.fontSize = `${Math.max(8, item.fontSize * scale)}px`;
        block.style.width = `${Math.max(10, item.width * scale)}px`;
        block.style.minWidth = `${Math.max(10, item.width * scale)}px`;
        block.style.height = `${Math.max(item.height * scale, item.fontSize * scale * 1.08)}px`;
        block.style.minHeight = `${Math.max(item.height * scale, item.fontSize * scale * 1.08)}px`;
        const content = options.advancedEditor
          ? document.createElement("textarea")
          : document.createElement("div");
        content.className = options.advancedEditor ? "editor-text-input" : "editor-text-content";
        if (!options.advancedEditor) {
          content.contentEditable = "plaintext-only";
        }
        content.spellcheck = false;
        content.style.lineHeight = `${Math.max(item.height * scale, item.fontSize * scale * 1.02)}px`;
        content.style.fontFamily = item.fontFamily;
        content.style.fontWeight = item.fontWeight;
        content.style.fontStyle = item.fontStyle;
        content.style.direction = item.dir;
        content.style.transform = `scaleX(${item.scaleX || 1}) rotate(${item.angle || 0}rad)`;
        if (options.advancedEditor) {
          content.value = item.text;
          content.wrap = "off";
        } else {
          content.textContent = item.text;
        }
        block.dataset.originalText = item.text;
        block.appendChild(content);
        if (options.advancedEditor) {
          wireOriginalPreservingTextBlock(block);
        }
        content.addEventListener("click", (event) => {
          event.stopPropagation();
        });
        block.addEventListener("click", (event) => {
          event.stopPropagation();
          content.focus();
        });
        layer.appendChild(block);
    });

    page.appendChild(layer);
    pageShell.appendChild(page);
    pagesHost.appendChild(pageShell);
  });

  exactButton.addEventListener("click", () => {
    exactButton.classList.add("is-active");
    editButton.classList.remove("is-active");
    pagesHost.querySelectorAll(".editor-page").forEach((page) => {
      page.classList.add("is-exact");
      page.classList.remove("is-edit");
    });
  });

  editButton.addEventListener("click", () => {
    editButton.classList.add("is-active");
    exactButton.classList.remove("is-active");
    pagesHost.querySelectorAll(".editor-page").forEach((page) => {
      page.classList.add("is-edit");
      page.classList.remove("is-exact");
    });
  });

  if (options.advancedEditor) {
    const statusNode = document.createElement("div");
    wrapper.appendChild(buildPdfEditorToolbar(pagesHost, statusNode));
  }

  wrapper.appendChild(pagesHost);
  return wrapper;
}

function buildWordHtmlFromPreviewBlocks() {
  const pageNodes = [...resultPreview.querySelectorAll(".editor-page")];
  const markup = pageNodes.map((pageNode) => {
    const image = pageNode.querySelector("img")?.src || "";
    const blocks = serializeEditorNodes(pageNode);

    return `
      <section class="word-page" style="width:${pageNode.offsetWidth}px;height:${pageNode.offsetHeight}px;">
        <img class="word-bg" src="${image}" alt="">
        <div class="word-text-layer">${blocks}</div>
      </section>
    `;
  }).join("\n");

  return `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="utf-8">
      <title>Editable PDF to Word Export</title>
      <style>
        @page { size: A4; margin: 0.35in; }
        body { font-family: Arial, sans-serif; margin: 0; background: #ffffff; }
        .word-page {
          position: relative;
          margin: 0 auto 24px;
          background: white;
          page-break-after: always;
          overflow: hidden;
        }
        .word-page:last-child { page-break-after: auto; }
        .word-bg {
          position: absolute;
          inset: 0;
          width: 100%;
          height: 100%;
          object-fit: fill;
          opacity: 0.2;
        }
        .word-text-layer { position: absolute; inset: 0; }
        .word-span {
          position: absolute;
          color: #111;
          line-height: 1;
          white-space: pre-wrap;
          overflow: visible;
        }
      </style>
    </head>
    <body>${markup}</body>
    </html>
  `;
}

function buildExportMarkupFromPreviewBlocks() {
  const pageNodes = [...resultPreview.querySelectorAll(".editor-page")];
  return pageNodes.map((pageNode) => {
    const image = pageNode.querySelector("img")?.src || "";
    const blocks = serializeEditorNodes(pageNode);

    return `
      <section class="word-page" style="width:${pageNode.offsetWidth}px;height:${pageNode.offsetHeight}px;">
        <img class="word-bg" src="${image}" alt="">
        <div class="word-text-layer">${blocks}</div>
      </section>
    `;
  }).join("\n");
}

function serializeEditorNodes(pageNode) {
  const nodes = [...pageNode.querySelectorAll(".editor-text-block, .editor-overlay")];
  return nodes.map((node) => {
    if (node.classList.contains("editor-text-block") && pageNode.dataset.preserveBackground === "true" && node.dataset.modified !== "true") {
      return "";
    }

    const textContentNode = node.classList?.contains("editor-text-block")
      ? node.querySelector(".editor-text-content, .editor-text-input")
      : null;

    if (node.tagName === "IMG") {
      return `
        <img
          class="word-image ${node.classList.contains("editor-sign") ? "word-sign" : ""}"
          src="${node.getAttribute("src") || ""}"
          style="left:${node.style.left};top:${node.style.top};width:${node.style.width};height:${node.style.height || "auto"};"
          alt=""
        >
      `;
    }

    if (node.classList.contains("editor-whiteout") || node.classList.contains("editor-shapes") || node.classList.contains("editor-annotate")) {
      return `
        <div class="word-overlay ${node.className.replace(/\s+/g, " ")}" style="
          left:${node.style.left};
          top:${node.style.top};
          width:${node.style.width};
          height:${node.style.height};
        ">${escapeHtml(node.textContent || "")}</div>
      `;
    }

    const needsMask = pageNode.dataset.preserveBackground === "true" && node.dataset.modified === "true";
    const computedNodeStyle = window.getComputedStyle(node);
    const computedTextStyle = textContentNode ? window.getComputedStyle(textContentNode) : computedNodeStyle;
    const fontSizeValue = computedTextStyle.fontSize || computedNodeStyle.fontSize || node.style.fontSize || "16px";
    const blockHeightValue = node.style.height || computedTextStyle.lineHeight || fontSizeValue;
    const baselineOffset = needsMask ? "0.2em" : "0";
    const maskMarkup = needsMask
      ? `
        <div class="word-text-mask" style="
          left:calc(${node.style.left} - 1px);
          top:calc(${node.style.top} - 1px);
          width:calc(${node.style.width} + 2px);
          height:calc(${blockHeightValue} + 2px);
        "></div>
      `
      : "";

    return `
      ${maskMarkup}
      <div class="word-span" style="
        left:${node.style.left};
        top:calc(${node.style.top} - ${baselineOffset});
        font-size:${fontSizeValue};
        width:${node.style.width};
        height:${blockHeightValue};
        font-family:${computedTextStyle.fontFamily || node.dataset.fontFamily || "Arial, sans-serif"};
        font-weight:${computedTextStyle.fontWeight || node.dataset.fontWeight || "400"};
        font-style:${computedTextStyle.fontStyle || node.dataset.fontStyle || "normal"};
        letter-spacing:${computedTextStyle.letterSpacing || "normal"};
        text-align:${computedTextStyle.textAlign || "left"};
        direction:${computedTextStyle.direction || node.style.direction || "ltr"};
        transform:${computedTextStyle.transform !== "none" ? computedTextStyle.transform : (node.style.transform || "none")};
        line-height:${computedTextStyle.lineHeight || "1"};
      ">${escapeHtml(node.dataset.currentText || ("value" in (textContentNode || {}) ? textContentNode?.value : textContentNode?.textContent) || node.textContent || "")}</div>
    `;
  }).join("");
}

function createPdfExportContainer() {
  const markup = buildExportMarkupFromPreviewBlocks();
  const container = document.createElement("div");
  container.style.position = "fixed";
  container.style.left = "-99999px";
  container.style.top = "0";
  container.style.width = "900px";
  container.style.background = "#ffffff";
  container.innerHTML = `
    <style>
      .word-page {
        position: relative;
        margin: 0 auto 20px;
        background: white;
        overflow: hidden;
        page-break-after: always;
        -webkit-text-size-adjust: 100%;
        text-size-adjust: 100%;
      }
      .word-page:last-child { page-break-after: auto; }
      .word-bg {
        position: absolute;
        inset: 0;
        width: 100%;
        height: 100%;
        object-fit: fill;
        opacity: 1;
      }
      .word-text-layer { position: absolute; inset: 0; }
      .word-text-mask {
        position: absolute;
        background: #ffffff;
      }
      .word-span {
        position: absolute;
        color: #111;
        display: block;
        margin: 0;
        padding: 0;
        box-sizing: border-box;
        line-height: 1;
        white-space: pre;
        overflow: hidden;
        transform-origin: top left;
        text-rendering: geometricPrecision;
        -webkit-font-smoothing: antialiased;
        -webkit-text-size-adjust: 100%;
        text-size-adjust: 100%;
        font-synthesis: none;
      }
      .word-image {
        position: absolute;
        object-fit: contain;
      }
      .word-overlay.editor-whiteout {
        position: absolute;
        background: #ffffff;
        border: 1px solid rgba(24,34,31,0.1);
      }
      .word-overlay.editor-annotate {
        position: absolute;
        background: #fff5b3;
        border: 1px solid #d9bf4b;
        padding: 8px;
        white-space: pre-wrap;
      }
      .word-overlay.editor-shapes {
        position: absolute;
        border: 2px solid #0f6be8;
        background: transparent;
      }
    </style>
    ${markup}
  `;
  document.body.appendChild(container);
  return container;
}

function parseCssPixels(value, fallback = 0) {
  const parsed = Number.parseFloat(value);
  return Number.isFinite(parsed) ? parsed : fallback;
}

async function ensureImageReady(image) {
  if (!image) {
    return;
  }

  if (image.complete && image.naturalWidth) {
    return;
  }

  if (typeof image.decode === "function") {
    try {
      await image.decode();
      return;
    } catch (error) {
      // Fall back to load/error events below.
    }
  }

  await new Promise((resolve) => {
    const cleanup = () => {
      image.removeEventListener("load", onLoad);
      image.removeEventListener("error", onError);
    };
    const onLoad = () => {
      cleanup();
      resolve();
    };
    const onError = () => {
      cleanup();
      resolve();
    };

    image.addEventListener("load", onLoad, { once: true });
    image.addEventListener("error", onError, { once: true });
  });
}

function drawCanvasTextLine(context, text, x, y, letterSpacing = 0) {
  const value = String(text || "");
  if (!letterSpacing) {
    context.fillText(value, x, y);
    return;
  }

  let offsetX = x;
  for (const character of value) {
    context.fillText(character, offsetX, y);
    offsetX += context.measureText(character).width + letterSpacing;
  }
}

function measureCanvasFontMetrics(context, sampleText = "Hg") {
  const metrics = context.measureText(sampleText);
  const ascent = Number.isFinite(metrics.actualBoundingBoxAscent) && metrics.actualBoundingBoxAscent > 0
    ? metrics.actualBoundingBoxAscent
    : parseCssPixels(context.font, 16) * 0.8;
  const descent = Number.isFinite(metrics.actualBoundingBoxDescent) && metrics.actualBoundingBoxDescent >= 0
    ? metrics.actualBoundingBoxDescent
    : parseCssPixels(context.font, 16) * 0.2;

  return { ascent, descent };
}

async function renderPreviewPageToCanvas(pageNode, scale = 2) {
  const pageWidth = pageNode.clientWidth || pageNode.offsetWidth || 595;
  const pageHeight = pageNode.clientHeight || pageNode.offsetHeight || 842;
  const canvas = document.createElement("canvas");
  canvas.width = Math.max(1, Math.round(pageWidth * scale));
  canvas.height = Math.max(1, Math.round(pageHeight * scale));

  const context = canvas.getContext("2d");
  context.scale(scale, scale);
  context.fillStyle = "#ffffff";
  context.fillRect(0, 0, pageWidth, pageHeight);

  if (document.fonts?.ready) {
    try {
      await document.fonts.ready;
    } catch (error) {
      console.warn("Fonts were not fully ready before PDF export.", error);
    }
  }

  const backgroundImage = pageNode.querySelector("img.editor-page-guide");
  if (backgroundImage?.src) {
    await ensureImageReady(backgroundImage);
    if (backgroundImage.naturalWidth) {
      context.drawImage(backgroundImage, 0, 0, pageWidth, pageHeight);
    }
  }

  const nodes = [...pageNode.querySelectorAll(".editor-text-block, .editor-overlay")];
  const preserveBackground = pageNode.dataset.preserveBackground === "true";

  for (const node of nodes) {
    const left = parseCssPixels(node.style.left);
    const top = parseCssPixels(node.style.top);
    const width = parseCssPixels(node.style.width);
    const height = parseCssPixels(node.style.height);

    if (node.tagName === "IMG") {
      await ensureImageReady(node);
      if (node.naturalWidth) {
        context.drawImage(node, left, top, width || node.naturalWidth, height || node.naturalHeight);
      }
      continue;
    }

    if (node.classList.contains("editor-whiteout")) {
      context.fillStyle = "#ffffff";
      context.fillRect(left, top, width, height);
      context.strokeStyle = "rgba(24, 34, 31, 0.1)";
      context.lineWidth = 1;
      context.strokeRect(left, top, width, height);
      continue;
    }

    if (node.classList.contains("editor-shapes")) {
      context.strokeStyle = "#0f6be8";
      context.lineWidth = 2;
      context.strokeRect(left, top, width, height);
      continue;
    }

    if (node.classList.contains("editor-annotate")) {
      context.fillStyle = "#fff5b3";
      context.fillRect(left, top, width, height);
      context.strokeStyle = "#d9bf4b";
      context.lineWidth = 1;
      context.strokeRect(left, top, width, height);
      context.fillStyle = "#111111";
      context.font = '16px "Space Grotesk", sans-serif';
      context.textBaseline = "top";
      const noteText = (node.textContent || "").split(/\r?\n/);
      noteText.forEach((line, index) => {
        context.fillText(line, left + 8, top + 8 + (index * 18));
      });
      continue;
    }

    if (!node.classList.contains("editor-text-block")) {
      continue;
    }

    if (preserveBackground && node.dataset.modified !== "true") {
      continue;
    }

    const textContentNode = node.querySelector(".editor-text-content, .editor-text-input");
    const computedNodeStyle = window.getComputedStyle(node);
    const computedTextStyle = textContentNode ? window.getComputedStyle(textContentNode) : computedNodeStyle;
    const text = node.dataset.currentText
      || ("value" in (textContentNode || {}) ? textContentNode?.value : textContentNode?.textContent)
      || "";
    const fontSize = parseCssPixels(computedTextStyle.fontSize, parseCssPixels(node.style.fontSize, 16));
    const lineHeight = parseCssPixels(computedTextStyle.lineHeight, Math.max(fontSize, height || fontSize));
    const letterSpacing = computedTextStyle.letterSpacing === "normal" ? 0 : parseCssPixels(computedTextStyle.letterSpacing);
    const fontFamily = computedTextStyle.fontFamily || node.dataset.fontFamily || "Arial, sans-serif";
    const fontWeight = computedTextStyle.fontWeight || node.dataset.fontWeight || "400";
    const fontStyle = computedTextStyle.fontStyle || node.dataset.fontStyle || "normal";
    const scaleX = Number.parseFloat(node.dataset.scaleX || "1") || 1;
    const angle = Number.parseFloat(node.dataset.angle || "0") || 0;
    const needsMask = preserveBackground && node.dataset.modified === "true";
    const baselineOffset = needsMask ? fontSize * 0.2 : 0;

    if (needsMask) {
      context.fillStyle = "#ffffff";
      context.fillRect(left - 1, top - 1, width + 2, Math.max(height, lineHeight) + 2);
    }

    context.save();
    context.beginPath();
    context.rect(left - 1, top - 1, Math.max(width + 2, 1), Math.max(height + 2, lineHeight + 2));
    context.clip();
    context.font = `${fontStyle} ${fontWeight} ${fontSize}px ${fontFamily}`;
    const fontMetrics = measureCanvasFontMetrics(context);
    const textBoxHeight = Math.max(height || 0, lineHeight, fontMetrics.ascent + fontMetrics.descent);
    const lineBoxOffset = Math.max(0, (textBoxHeight - (fontMetrics.ascent + fontMetrics.descent)) / 2);
    const mobileNudgeDown = fontSize * 0.16;
    context.translate(left, top - baselineOffset + lineBoxOffset + mobileNudgeDown);
    if (angle) {
      context.rotate(angle);
    }
    if (scaleX !== 1) {
      context.scale(scaleX, 1);
    }
    context.fillStyle = computedTextStyle.color && computedTextStyle.color !== "transparent"
      ? computedTextStyle.color
      : "#111111";
    context.textBaseline = "alphabetic";
    context.textAlign = computedTextStyle.textAlign === "center" || computedTextStyle.textAlign === "right"
      ? computedTextStyle.textAlign
      : "left";

    const lines = String(text).split(/\r?\n/);
    lines.forEach((line, index) => {
      const y = lineBoxOffset + fontMetrics.ascent + (index * lineHeight);
      let x = 0;
      if (context.textAlign === "center") {
        x = (width || 0) / 2;
      } else if (context.textAlign === "right") {
        x = width || 0;
      }
      drawCanvasTextLine(context, line, x, y, letterSpacing);
    });
    context.restore();
  }

  return canvas;
}

async function buildEditedPdfFromPreviewBlocks() {
  const sourcePages = [...resultPreview.querySelectorAll(".editor-page")];
  if (!sourcePages.length) {
    throw new Error("No edited pages are available to export.");
  }
  const pdfDoc = await PDFDocument.create();

  for (const pageNode of sourcePages) {
    const pageWidth = pageNode.clientWidth || pageNode.offsetWidth || 595;
    const pageHeight = pageNode.clientHeight || pageNode.offsetHeight || 842;
    const canvas = await renderPreviewPageToCanvas(pageNode, 2);
    const imageBytes = await dataUrlToBytes(canvas.toDataURL("image/png"));
    const image = await pdfDoc.embedPng(imageBytes);
    const pdfPage = pdfDoc.addPage([pageWidth, pageHeight]);
    pdfPage.drawImage(image, {
      x: 0,
      y: 0,
      width: pageWidth,
      height: pageHeight
    });
  }

  return new Blob([await pdfDoc.save()], { type: "application/pdf" });
}

function wrapTextToWidth(text, font, size, maxWidth) {
  const paragraphs = String(text || "").split(/\r?\n/);
  const lines = [];

  paragraphs.forEach((paragraph) => {
    const words = paragraph.split(/\s+/).filter(Boolean);
    if (!words.length) {
      lines.push("");
      return;
    }

    let current = words[0];
    for (let index = 1; index < words.length; index += 1) {
      const candidate = `${current} ${words[index]}`;
      if (font.widthOfTextAtSize(candidate, size) <= maxWidth) {
        current = candidate;
      } else {
        lines.push(current);
        current = words[index];
      }
    }
    lines.push(current);
  });

  return lines;
}

function resetEditorHistory() {
  state.editorHistory = [];
  state.editorTool = "select";
}

function pushEditorHistory(pagesHost) {
  if (!pagesHost) {
    return;
  }
  state.editorHistory.push(pagesHost.innerHTML);
  if (state.editorHistory.length > 40) {
    state.editorHistory.shift();
  }
}

function undoEditorChange(pagesHost) {
  if (!pagesHost || state.editorHistory.length < 2) {
    return;
  }
  state.editorHistory.pop();
  pagesHost.innerHTML = state.editorHistory[state.editorHistory.length - 1];
}

function createEditorControlButton(label, toolId, variant = "secondary") {
  const button = document.createElement("button");
  button.type = "button";
  button.className = `editor-tool-button ${variant === "primary" ? "is-primary" : ""}`;
  button.dataset.tool = toolId;
  button.textContent = label;
  return button;
}

function buildPdfEditorToolbar(pagesHost, statusNode) {
  const shell = document.createElement("div");
  shell.className = "editor-toolbar-shell";

  const mergeRow = document.createElement("div");
  mergeRow.className = "editor-merge-row";
  mergeRow.appendChild(createEditorControlButton("Merge Boxes", "merge", "primary"));

  const toolRow = document.createElement("div");
  toolRow.className = "editor-tool-row";
  const controls = [
    ["Text", "text"],
    ["Links", "links"],
    ["Forms", "forms"],
    ["Images", "images"],
    ["Sign", "sign"],
    ["Whiteout", "whiteout"],
    ["Annotate", "annotate"],
    ["Shapes", "shapes"],
    ["Undo", "undo"]
  ];

  controls.forEach(([label, toolId]) => {
    toolRow.appendChild(createEditorControlButton(label, toolId));
  });

  const pageRow = document.createElement("div");
  pageRow.className = "editor-page-row";
  pageRow.innerHTML = `
    <span class="editor-page-indicator">1</span>
    <button type="button" class="editor-mini-button" data-action="delete-page">Delete Page</button>
    <button type="button" class="editor-mini-button" data-action="zoom-in">Zoom In</button>
    <button type="button" class="editor-mini-button" data-action="zoom-out">Zoom Out</button>
    <button type="button" class="editor-mini-button" data-action="rotate-left">Rotate Left</button>
    <button type="button" class="editor-mini-button" data-action="rotate-right">Rotate Right</button>
    <button type="button" class="editor-mini-button is-insert" data-action="insert-page">Insert Page Here</button>
  `;

  statusNode.classList.add("editor-status");
  const hasLockedImages = [...pagesHost.querySelectorAll(".editor-page")].some((page) => page.dataset.hasSourceImage === "true");
  statusNode.textContent = hasLockedImages
    ? "Choose a tool, switch to Editable Layer, then click on the page to place it. Original PDF images are detected and kept locked."
    : "Choose a tool, switch to Editable Layer, then click on the page to place it.";

  shell.append(mergeRow, toolRow, pageRow, statusNode);
  wirePdfEditorToolbar(shell, pagesHost, statusNode);
  return shell;
}

function wirePdfEditorToolbar(shell, pagesHost, statusNode) {
  const toolButtons = [...shell.querySelectorAll(".editor-tool-button")];
  const pageButtons = [...shell.querySelectorAll(".editor-mini-button")];
  const clearMergeSelections = () => {
    pagesHost.querySelectorAll(".editor-text-block.is-merge-selected").forEach((block) => {
      block.classList.remove("is-merge-selected");
    });
  };

  const mergeEditorBlocks = (blocks) => {
    if (blocks.length !== 2) {
      return null;
    }

    const [first, second] = [...blocks].sort((left, right) => {
      const topDiff = Number.parseFloat(left.style.top || "0") - Number.parseFloat(right.style.top || "0");
      return Math.abs(topDiff) > 8
        ? topDiff
        : Number.parseFloat(left.style.left || "0") - Number.parseFloat(right.style.left || "0");
    });

    const firstLeft = Number.parseFloat(first.style.left || "0");
    const secondLeft = Number.parseFloat(second.style.left || "0");
    const firstTop = Number.parseFloat(first.style.top || "0");
    const secondTop = Number.parseFloat(second.style.top || "0");
    const firstWidth = Number.parseFloat(first.style.width || "0");
    const secondWidth = Number.parseFloat(second.style.width || "0");
    const firstHeight = Number.parseFloat(first.style.height || "0");
    const secondHeight = Number.parseFloat(second.style.height || "0");
    const firstRight = firstLeft + firstWidth;
    const secondRight = secondLeft + secondWidth;
    const firstBottom = firstTop + firstHeight;
    const secondBottom = secondTop + secondHeight;
    const gap = secondLeft - firstRight;
    const fontSize = Math.max(
      Number.parseFloat(first.style.fontSize || "16"),
      Number.parseFloat(second.style.fontSize || "16")
    );
    const left = Math.min(firstLeft, secondLeft);
    const top = Math.min(firstTop, secondTop);
    const width = Math.max(firstRight, secondRight) - left;
    const height = Math.max(firstBottom, secondBottom) - top;

    let leftText = getEditorBlockText(first).trimEnd();
    const rightText = getEditorBlockText(second).trimStart();
    const needsSpace = gap > Math.max(fontSize * 0.12, 4) && !leftText.endsWith("-") && !/^[,.;:!?)]/.test(rightText);
    if (leftText.endsWith("-")) {
      leftText = leftText.slice(0, -1);
    }
    const mergedText = `${leftText}${needsSpace ? " " : ""}${rightText}`.trim();

    pushEditorHistory(pagesHost);
    first.style.left = `${left}px`;
    first.style.top = `${top}px`;
    first.style.width = `${Math.max(width, Number.parseFloat(first.style.minWidth || "10"))}px`;
    first.style.height = `${Math.max(height, Number.parseFloat(first.style.minHeight || "10"))}px`;
    first.style.minWidth = first.style.width;
    first.style.minHeight = first.style.height;
    first.dataset.blockId = "";
    setEditorBlockText(first, mergedText);
    second.remove();
    clearMergeSelections();
    return first;
  };

  const setActiveTool = (toolId) => {
    state.editorTool = toolId;
    if (toolId !== "merge") {
      clearMergeSelections();
    }
    toolButtons.forEach((button) => button.classList.toggle("is-active", button.dataset.tool === toolId));
    const messages = {
      text: "Click on the page to add a text box.",
      merge: "Click two neighboring text boxes to merge them into one editable box.",
      images: "Click on the page to place an uploaded image.",
      sign: "Click on the page to place a signature image.",
      whiteout: "Click on the page to add a whiteout box.",
      annotate: "Click on the page to add an annotation note.",
      shapes: "Click on the page to add a shape box.",
      links: "Links UI added. Link targets are not exported yet.",
      forms: "Forms UI added. Interactive form export is not wired yet.",
      select: "Editable Layer is active. Click existing text to edit it."
    };
    statusNode.textContent = messages[toolId] || "Tool selected.";
  };

  pagesHost._setEditorTool = setActiveTool;

  toolButtons.forEach((button) => {
    button.addEventListener("click", async () => {
      const toolId = button.dataset.tool;
      if (toolId === "undo") {
        undoEditorChange(pagesHost);
        statusNode.textContent = "Undid the last editor action.";
        return;
      }
      setActiveTool(toolId);
    });
  });

  pagesHost.addEventListener("click", (event) => {
    const block = event.target.closest(".editor-text-block");
    if (!block || state.editorTool !== "merge") {
      return;
    }

    const page = block.closest(".editor-page");
    const selected = [...pagesHost.querySelectorAll(".editor-text-block.is-merge-selected")];
    const selectedOnOtherPages = selected.filter((entry) => entry.closest(".editor-page") !== page);
    if (selectedOnOtherPages.length) {
      clearMergeSelections();
    }

    event.preventDefault();
    event.stopPropagation();

    if (block.classList.contains("is-merge-selected")) {
      block.classList.remove("is-merge-selected");
      statusNode.textContent = "Select one more text box to merge.";
      return;
    }

    if (pagesHost.querySelectorAll(".editor-text-block.is-merge-selected").length >= 2) {
      clearMergeSelections();
    }

    block.classList.add("is-merge-selected");
    const selectedBlocks = [...pagesHost.querySelectorAll(".editor-text-block.is-merge-selected")];
    if (selectedBlocks.length < 2) {
      statusNode.textContent = "Select one more text box to merge.";
      return;
    }

    const mergedBlock = mergeEditorBlocks(selectedBlocks);
    if (mergedBlock) {
      statusNode.textContent = "Merged the selected text boxes.";
      setActiveTool("select");
      const content = mergedBlock.querySelector(".editor-text-content, .editor-text-input");
      content?.focus();
    }
  }, true);

  pageButtons.forEach((button) => {
    button.addEventListener("click", () => {
      const action = button.dataset.action;
      const pages = [...pagesHost.querySelectorAll(".editor-page")];
      if (!pages.length) {
        return;
      }
      const firstPage = pages[0];

      if (action === "zoom-in" || action === "zoom-out") {
        const delta = action === "zoom-in" ? 0.1 : -0.1;
        pages.forEach((page) => {
          const current = Number.parseFloat(page.dataset.zoom || "1");
          const next = Math.max(0.6, Math.min(1.8, current + delta));
          page.dataset.zoom = String(next);
          page.style.transform = `scale(${next})`;
          page.style.transformOrigin = "top center";
        });
        statusNode.textContent = action === "zoom-in" ? "Zoomed in." : "Zoomed out.";
        return;
      }

      if (action === "rotate-left" || action === "rotate-right") {
        const current = Number.parseFloat(firstPage.dataset.rotation || "0");
        const next = current + (action === "rotate-left" ? -90 : 90);
        firstPage.dataset.rotation = String(next);
        firstPage.style.rotate = `${next}deg`;
        statusNode.textContent = action === "rotate-left" ? "Rotated page left." : "Rotated page right.";
        return;
      }

      if (action === "delete-page" && pages.length > 1) {
        pushEditorHistory(pagesHost);
        pages[pages.length - 1].remove();
        statusNode.textContent = "Removed the last page from the editor preview.";
        return;
      }

      if (action === "insert-page") {
        pushEditorHistory(pagesHost);
        const clone = firstPage.cloneNode(true);
        clone.querySelectorAll(".editor-text-block, .editor-overlay").forEach((node) => node.remove());
        pagesHost.appendChild(clone);
        statusNode.textContent = "Inserted a blank editable page at the end.";
      }
    });
  });

  pagesHost.addEventListener("click", async (event) => {
    const page = event.target.closest(".editor-page");
    if (!page || !page.classList.contains("is-edit")) {
      return;
    }
    if (event.target.closest(".editor-text-block, .editor-overlay")) {
      return;
    }
    const pageRect = page.getBoundingClientRect();
    const rawX = event.clientX - pageRect.left;
    const rawY = event.clientY - pageRect.top;
    const snapped = snapPointToNearestTextBlock(page, rawX, rawY);
    const x = snapped.x;
    const y = snapped.y;

    if (state.editorTool === "text") {
      pushEditorHistory(pagesHost);
      const block = document.createElement("div");
      block.className = "editor-overlay editor-text-block editor-free-text";
      block.contentEditable = "plaintext-only";
      block.spellcheck = false;
      block.dataset.fontFamily = snapped.fontFamily || "Arial, sans-serif";
      block.dataset.fontWeight = snapped.fontWeight || "400";
      block.dataset.fontStyle = snapped.fontStyle || "normal";
      block.dataset.modified = "true";
      block.dataset.currentText = "New text";
      block.style.left = `${x}px`;
      block.style.top = `${y}px`;
      block.style.width = `${snapped.width ? Math.max(120, snapped.width) : 180}px`;
      block.style.fontSize = `${snapped.fontSize || 18}px`;
      block.style.fontFamily = snapped.fontFamily || "Arial, sans-serif";
      block.style.fontWeight = snapped.fontWeight || "400";
      block.style.fontStyle = snapped.fontStyle || "normal";
      block.style.height = `${snapped.height ? Math.max(28, snapped.height) : 32}px`;
      block.style.lineHeight = snapped.lineHeight || `${snapped.height ? Math.max(28, snapped.height) : 32}px`;
      const input = document.createElement("textarea");
      input.className = "editor-text-input";
      input.value = "New text";
      input.wrap = "off";
      input.style.fontFamily = snapped.fontFamily || "Arial, sans-serif";
      input.style.fontWeight = snapped.fontWeight || "400";
      input.style.fontStyle = snapped.fontStyle || "normal";
      input.style.lineHeight = block.style.lineHeight;
      block.appendChild(input);
      wireOriginalPreservingTextBlock(block);
      page.querySelector(".editor-text-layer")?.appendChild(block);
      input.focus();
      statusNode.textContent = "Added a text box.";
      return;
    }

    if (state.editorTool === "whiteout" || state.editorTool === "annotate" || state.editorTool === "shapes") {
      pushEditorHistory(pagesHost);
      const overlay = document.createElement("div");
      overlay.className = `editor-overlay editor-${state.editorTool}`;
      overlay.style.left = `${x}px`;
      overlay.style.top = `${y}px`;
      overlay.style.width = state.editorTool === "annotate" ? "180px" : "140px";
      overlay.style.height = state.editorTool === "annotate" ? "70px" : "48px";
      if (state.editorTool === "annotate") {
        overlay.contentEditable = "plaintext-only";
        overlay.textContent = "Note";
      }
      page.querySelector(".editor-text-layer")?.appendChild(overlay);
      statusNode.textContent = `Added ${state.editorTool}.`;
      return;
    }

    if (state.editorTool === "images" || state.editorTool === "sign") {
      const input = document.createElement("input");
      input.type = "file";
      input.accept = "image/*";
      input.addEventListener("change", async () => {
        const file = input.files?.[0];
        if (!file) {
          return;
        }
        pushEditorHistory(pagesHost);
        const dataUrl = await fileToDataUrl(file);
        const overlay = document.createElement("img");
        overlay.className = `editor-overlay editor-image ${state.editorTool === "sign" ? "editor-sign" : ""}`;
        overlay.src = dataUrl;
        overlay.dataset.src = dataUrl;
        overlay.style.left = `${x}px`;
        overlay.style.top = `${y}px`;
        overlay.style.width = state.editorTool === "sign" ? "140px" : "180px";
        overlay.style.height = "auto";
        page.querySelector(".editor-text-layer")?.appendChild(overlay);
        statusNode.textContent = state.editorTool === "sign" ? "Placed signature image." : "Placed image.";
      });
      input.click();
    }
  });
}

async function fileToDataUrl(file) {
  return await new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || ""));
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

async function launchEditablePdfWorkspace(file, exportFormat = "word+pdf", summary = "Your PDF is shown below in Exact Preview mode. Switch to Editable Layer when you want to edit extracted text, then download the file.", options = {}) {
  let backendSession = null;
  try {
    backendSession = await createBackendEditSession(file);
  } catch (error) {
    console.warn("Backend editor session is unavailable, using frontend fallback.", error);
  }

  const [renderedPages, fallbackTextPages, imageFlags] = await Promise.all([
    renderPdfPages(file, { scale: 1.8 }),
    extractPdfTextLayout(file, "", 1.8),
    detectPdfPageImages(file)
  ]);
  const textPages = backendSession
    ? buildTextPagesFromBackendSession(backendSession)
    : fallbackTextPages;
  resetEditorHistory();
  const previewNode = buildEditableWordPreview(renderedPages, textPages, imageFlags, options);
  if (backendSession?.sessionId) {
    previewNode.dataset.backendSessionId = backendSession.sessionId;
  }
  const backendNote = options.showStatusMessage === false
    ? ""
    : backendSession
      ? ` Backend session ${backendSession.sessionId} is ready and text edits will export through the server when possible.`
      : " Running in frontend fallback mode until the real PDF engine is connected.";
  setResult(`${summary}${backendNote}`.trim(), []);
  setPreview(previewNode);
  const pagesHost = previewNode.querySelector(".preview-host");
  if (pagesHost) {
    pushEditorHistory(pagesHost);
  }

  if (exportFormat === "word+pdf" || exportFormat === "word") {
    const wordAction = document.createElement("button");
    wordAction.type = "button";
    wordAction.className = exportFormat === "word" ? "button button-primary" : "button button-primary";
    wordAction.textContent = "Download Word File";
    wordAction.addEventListener("click", async () => {
      const html = buildWordHtmlFromPreviewBlocks();
      const blob = new Blob([html], { type: "application/msword" });
      const href = await blobToObjectUrl(blob);
      const link = document.createElement("a");
      link.href = href;
      link.download = "pdf-to-word-layout-editable.doc";
      link.click();
    });
    resultActions.appendChild(wordAction);
  }

    if (exportFormat === "word+pdf" || exportFormat === "pdf") {
      const pdfAction = document.createElement("button");
      pdfAction.type = "button";
      pdfAction.className = exportFormat === "pdf" ? "button button-primary" : "button button-secondary";
      pdfAction.textContent = "Download Edited PDF";
      pdfAction.addEventListener("click", async () => {
        try {
          refreshPreviewEditorState(previewNode);
          const sessionId = previewNode.dataset.backendSessionId;
          const prefersVisualExport = Boolean(previewNode.querySelector(".editor-page.editor-preserve-background"));
          if (sessionId && !prefersVisualExport && !hasUnsupportedBackendOverlays(previewNode)) {
            await syncBackendSessionFromPreview(previewNode, sessionId);
            const exported = await exportBackendSession(sessionId);
            const href = backendAssetUrl(exported.downloadUrl);
            if (!href) {
              throw new Error("Backend export completed, but no download URL was returned.");
            }
            const link = document.createElement("a");
            link.href = href;
            link.download = exported.fileName || "edited-document.pdf";
            link.click();
            return;
          }

          const blob = await buildEditedPdfFromPreviewBlocks();
          const href = await blobToObjectUrl(blob);
          const link = document.createElement("a");
          link.href = href;
          link.download = "edited-document.pdf";
          link.click();
        } catch (error) {
          console.error(error);
          setResult(error.message || "The edited PDF could not be downloaded.", []);
          setPreview(previewNode);
        }
      });
      resultActions.appendChild(pdfAction);
    }
}

function escapeHtml(value) {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll("\"", "&quot;")
    .replaceAll("'", "&#39;");
}

async function recognizeImageText(imageSource, language = "eng") {
  const result = await window.Tesseract.recognize(imageSource, language, {});
  return result.data.text.trim();
}

function summarizeDifferences(leftPages, rightPages) {
  const total = Math.max(leftPages.length, rightPages.length);
  const changedPages = [];

  for (let index = 0; index < total; index += 1) {
    const left = leftPages[index]?.text || "";
    const right = rightPages[index]?.text || "";
    if (left !== right) {
      changedPages.push({
        page: index + 1,
        leftPreview: left.slice(0, 180),
        rightPreview: right.slice(0, 180)
      });
    }
  }

  const lines = [
    "FluxPDF Comparison Report",
    `First document pages: ${leftPages.length}`,
    `Second document pages: ${rightPages.length}`,
    `Changed pages: ${changedPages.length}`,
    ""
  ];

  changedPages.forEach((entry) => {
    lines.push(`Page ${entry.page}`);
    lines.push(`First: ${entry.leftPreview || "[no text found]"}`);
    lines.push(`Second: ${entry.rightPreview || "[no text found]"}`);
    lines.push("");
  });

  if (!changedPages.length) {
    lines.push("No text differences were detected in the extracted page content.");
  }

  return lines.join("\n");
}

function hexToRgb(hex) {
  const clean = hex.replace("#", "");
  const normalized = clean.length === 3
    ? clean.split("").map((value) => value + value).join("")
    : clean;
  const num = Number.parseInt(normalized, 16);
  return rgb(((num >> 16) & 255) / 255, ((num >> 8) & 255) / 255, (num & 255) / 255);
}

function setResult(message, actions = []) {
  resultSummary.textContent = message;
  resultSummary.style.display = message ? "" : "none";
  resultActions.innerHTML = "";
  resultPreview.innerHTML = "";
  actions.forEach((action) => {
    const element = document.createElement("a");
    element.className = "button button-primary";
    element.href = action.href;
    element.download = action.download || "";
    element.textContent = action.label;
    resultActions.appendChild(element);
  });
}

function setPreview(node) {
  resultPreview.innerHTML = "";
  if (node) {
    resultPreview.appendChild(node);
    resultPanel?.scrollIntoView({ behavior: "smooth", block: "start" });
  }
}

function clearResult() {
  if (state.downloadUrl) {
    URL.revokeObjectURL(state.downloadUrl);
    state.downloadUrl = null;
  }
  state.transientUrls.forEach((url) => URL.revokeObjectURL(url));
  state.transientUrls = [];
  setResult("Choose a tool and upload your files to generate an edited PDF or exported images.", []);
}

async function downloadPdf(pdfDoc, filename, message) {
  const pdfBytes = await pdfDoc.save();
  if (state.downloadUrl) {
    URL.revokeObjectURL(state.downloadUrl);
  }
  state.downloadUrl = URL.createObjectURL(new Blob([pdfBytes], { type: "application/pdf" }));
  setResult(message, [{ label: "Download PDF", href: state.downloadUrl, download: filename }]);
}

async function mergePdf(formData) {
  const files = [...formData.getAll("files")].filter(Boolean);
  if (files.length < 2) {
    throw new Error("Please upload at least two PDF files to merge.");
  }
  const merged = await PDFDocument.create();
  for (const file of files) {
    const pdf = await PDFDocument.load(await fileToBytes(file));
    const copiedPages = await merged.copyPages(pdf, pdf.getPageIndices());
    copiedPages.forEach((page) => merged.addPage(page));
  }
  await downloadPdf(merged, "merged-document.pdf", `Merged ${files.length} PDF files into one document.`);
}

async function extractPages(formData) {
  const file = formData.get("file");
  const source = await PDFDocument.load(await fileToBytes(file));
  const pages = parseRanges(formData.get("pages"), source.getPageCount()).map((page) => page - 1);
  const result = await PDFDocument.create();
  const copied = await result.copyPages(source, pages);
  copied.forEach((page) => result.addPage(page));
  await downloadPdf(result, "extracted-pages.pdf", `Extracted ${pages.length} page(s) into a new PDF.`);
}

async function splitPdf(formData) {
  const file = formData.get("file");
  const source = await PDFDocument.load(await fileToBytes(file));
  const totalPages = source.getPageCount();
  const actions = [];

  for (let index = 0; index < totalPages; index += 1) {
    const result = await PDFDocument.create();
    const [page] = await result.copyPages(source, [index]);
    result.addPage(page);
    const pdfBytes = await result.save();
    const href = URL.createObjectURL(new Blob([pdfBytes], { type: "application/pdf" }));
    actions.push({
      label: `Download Page ${index + 1}`,
      href,
      download: `page-${index + 1}.pdf`
    });
    state.transientUrls.push(href);
  }

  setResult(`Split the PDF into ${totalPages} single-page files. Download each page below.`, actions);
}

async function deletePages(formData) {
  const file = formData.get("file");
  const source = await PDFDocument.load(await fileToBytes(file));
  const totalPages = source.getPageCount();
  const pagesToDelete = new Set(parseRanges(formData.get("pages"), totalPages).map((page) => page - 1));
  const result = await PDFDocument.create();
  const keptPages = Array.from({ length: totalPages }, (_, index) => index).filter((index) => !pagesToDelete.has(index));
  if (!keptPages.length) {
    throw new Error("You deleted every page. Keep at least one page in the document.");
  }
  const copied = await result.copyPages(source, keptPages);
  copied.forEach((page) => result.addPage(page));
  await downloadPdf(result, "trimmed-document.pdf", `Removed ${pagesToDelete.size} page(s) from the PDF.`);
}

async function rotatePdf(formData) {
  const file = formData.get("file");
  const angle = Number(formData.get("angle"));
  const pdfDoc = await PDFDocument.load(await fileToBytes(file));
  pdfDoc.getPages().forEach((page) => page.setRotation(degrees(angle)));
  await downloadPdf(pdfDoc, "rotated-document.pdf", `Rotated all pages by ${angle} degrees.`);
}

async function reorderPages(formData) {
  const file = formData.get("file");
  const source = await PDFDocument.load(await fileToBytes(file));
  const totalPages = source.getPageCount();
  const order = formData.get("order").split(",").map((value) => Number(value.trim()));
  if (order.length !== totalPages) {
    throw new Error(`Please provide exactly ${totalPages} page numbers for the new order.`);
  }
  const unique = new Set(order);
  if (unique.size !== totalPages) {
    throw new Error("Each page must appear exactly once in the new order.");
  }
  order.forEach((page) => validatePage(page, totalPages));
  const result = await PDFDocument.create();
  const copied = await result.copyPages(source, order.map((page) => page - 1));
  copied.forEach((page) => result.addPage(page));
  await downloadPdf(result, "reordered-document.pdf", "Rebuilt the PDF with the page order you selected.");
}

async function watermarkPdf(formData) {
  const file = formData.get("file");
  const text = formData.get("text").trim();
  const color = hexToRgb(formData.get("color"));
  const opacity = Number(formData.get("opacity"));
  const pdfDoc = await PDFDocument.load(await fileToBytes(file));
  const font = await pdfDoc.embedFont(StandardFonts.HelveticaBold);
  pdfDoc.getPages().forEach((page) => {
    const { width, height } = page.getSize();
    page.drawText(text, {
      x: width * 0.18,
      y: height * 0.45,
      size: Math.max(28, width * 0.06),
      rotate: degrees(35),
      opacity,
      color,
      font
    });
  });
  await downloadPdf(pdfDoc, "watermarked-document.pdf", "Applied your watermark across every page.");
}

async function numberPages(formData) {
  const file = formData.get("file");
  const format = formData.get("format").trim() || "Page {n}";
  const position = formData.get("position");
  const pdfDoc = await PDFDocument.load(await fileToBytes(file));
  const font = await pdfDoc.embedFont(StandardFonts.Helvetica);

  pdfDoc.getPages().forEach((page, index) => {
    const label = format.replaceAll("{n}", index + 1);
    const size = 11;
    const width = font.widthOfTextAtSize(label, size);
    const pageSize = page.getSize();
    let x = (pageSize.width - width) / 2;
    let y = 22;

    if (position === "bottom-right") {
      x = pageSize.width - width - 24;
    }
    if (position === "top-right") {
      x = pageSize.width - width - 24;
      y = pageSize.height - 24;
    }

    page.drawText(label, {
      x,
      y,
      size,
      color: rgb(0.25, 0.25, 0.25),
      font
    });
  });

  await downloadPdf(pdfDoc, "numbered-document.pdf", "Inserted page numbers into the PDF.");
}

async function annotatePdf(formData) {
  const file = formData.get("file");
  const text = formData.get("text").trim();
  const x = Number(formData.get("x"));
  const y = Number(formData.get("y"));
  const pdfDoc = await PDFDocument.load(await fileToBytes(file));
  const font = await pdfDoc.embedFont(StandardFonts.HelveticaBold);
  pdfDoc.getPages().forEach((page) => {
    page.drawText(text, {
      x,
      y,
      size: 18,
      color: rgb(0.12, 0.24, 0.38),
      font
    });
  });
  await downloadPdf(pdfDoc, "annotated-document.pdf", "Placed your text overlay on every page.");
}

async function imagesToPdf(formData) {
  const files = [...formData.getAll("files")].filter(Boolean);
  if (!files.length) {
    throw new Error("Please upload at least one image.");
  }
  const fit = formData.get("fit") === "on";
  const pdfDoc = await PDFDocument.create();

  for (const file of files) {
    const bytes = await file.arrayBuffer();
    const isPng = file.type === "image/png" || file.name.toLowerCase().endsWith(".png");
    const image = isPng ? await pdfDoc.embedPng(bytes) : await pdfDoc.embedJpg(bytes);
    const page = pdfDoc.addPage([595, 842]);
    const dimensions = image.scale(1);

    if (fit) {
      const scale = Math.min(515 / dimensions.width, 762 / dimensions.height);
      const drawnWidth = dimensions.width * scale;
      const drawnHeight = dimensions.height * scale;
      page.drawImage(image, {
        x: (595 - drawnWidth) / 2,
        y: (842 - drawnHeight) / 2,
        width: drawnWidth,
        height: drawnHeight
      });
    } else {
      page.setSize(dimensions.width, dimensions.height);
      page.drawImage(image, { x: 0, y: 0, width: dimensions.width, height: dimensions.height });
    }
  }

  await downloadPdf(pdfDoc, "images-to-pdf.pdf", `Converted ${files.length} image(s) into a PDF.`);
}

async function scanToPdf(formData) {
  await imagesToPdf(formData);
}

async function pdfToImages(formData) {
  const file = formData.get("file");
  const arrayBuffer = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: new Uint8Array(arrayBuffer) }).promise;
  const actions = [];

  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
    const page = await pdf.getPage(pageNumber);
    const viewport = page.getViewport({ scale: 1.7 });
    const canvas = document.createElement("canvas");
    const context = canvas.getContext("2d");
    canvas.width = viewport.width;
    canvas.height = viewport.height;
    await page.render({ canvasContext: context, viewport }).promise;
    actions.push({
      label: `Download Page ${pageNumber}`,
      href: canvas.toDataURL("image/png"),
      download: `page-${pageNumber}.png`
    });
  }

  setResult(`Rendered ${pdf.numPages} page(s) as PNG images. Download each page below.`, actions);
}

async function compressPdf(formData) {
  const file = formData.get("file");
  const renderedPages = await renderPdfPages(file, { scale: formData.get("scale") });
  const blob = await buildPdfFromRenderedPages(renderedPages, { quality: formData.get("quality") });
  await downloadBlob(blob, "compressed-document.pdf", "Compressed the PDF by rebuilding it from optimized page images.", "Download Compressed PDF");
}

async function unlockPdf(formData) {
  const file = formData.get("file");
  const password = formData.get("password");
  const renderedPages = await renderPdfPages(file, { password, scale: 1.6 });
  const blob = await buildPdfFromRenderedPages(renderedPages, { quality: 0.9 });
  await downloadBlob(blob, "unlocked-document.pdf", "Exported an unlocked copy of the PDF. This browser-based workflow rebuilds pages from the unlocked render.", "Download Unlocked PDF");
}

async function protectPdf(formData) {
  const file = formData.get("file");
  const userPassword = String(formData.get("userPassword") || "").trim();
  const ownerPassword = String(formData.get("ownerPassword") || "").trim() || userPassword;

  if (!userPassword) {
    throw new Error("Please enter an open password for the protected PDF.");
  }

  const renderedPages = await renderPdfPages(file, { scale: 1.6 });
  const blob = await buildPdfFromRenderedPages(renderedPages, {
    quality: 0.92,
    encryption: {
      userPassword,
      ownerPassword,
      userPermissions: ["print"]
    }
  });
  await downloadBlob(blob, "protected-document.pdf", "Created a password-protected PDF export.", "Download Protected PDF");
}

async function ocrPdf(formData) {
  const file = formData.get("file");
  const language = formData.get("language");
  const format = formData.get("format");
  let pages = [];

  if (file.type === "application/pdf" || file.name.toLowerCase().endsWith(".pdf")) {
    const renderedPages = await renderPdfPages(file, { scale: 1.6 });
    for (const page of renderedPages) {
      const text = await recognizeImageText(page.canvas, language);
      pages.push({ pageNumber: page.pageNumber, text });
    }
  } else {
    const text = await recognizeImageText(file, language);
    pages = [{ pageNumber: 1, text }];
  }

  if (format === "docx") {
    const blob = await buildDocxFromText("OCR Export", pages);
    await downloadBlob(blob, "ocr-export.docx", "OCR finished. Download the extracted text as a DOCX file.", "Download DOCX");
    return;
  }

  const textBlob = new Blob([pages.map((page) => `Page ${page.pageNumber}\n${page.text}`).join("\n\n")], { type: "text/plain" });
  await downloadBlob(textBlob, "ocr-export.txt", "OCR finished. Download the extracted text below.", "Download TXT");
}

async function esignPdf(formData) {
  const file = formData.get("file");
  const signature = formData.get("signature");
  const pageNumber = Number(formData.get("page"));
  const x = Number(formData.get("x"));
  const y = Number(formData.get("y"));
  const width = Number(formData.get("width"));
  const pdfDoc = await PDFDocument.load(await fileToBytes(file));
  validatePage(pageNumber, pdfDoc.getPageCount());
  const targetPage = pdfDoc.getPage(pageNumber - 1);
  const signatureBytes = await signature.arrayBuffer();
  const isPng = signature.type === "image/png" || signature.name.toLowerCase().endsWith(".png");
  const image = isPng ? await pdfDoc.embedPng(signatureBytes) : await pdfDoc.embedJpg(signatureBytes);
  const dimensions = image.scale(1);
  const height = width * (dimensions.height / dimensions.width);

  targetPage.drawImage(image, {
    x,
    y,
    width,
    height
  });

  await downloadPdf(pdfDoc, "signed-document.pdf", `Placed the signature on page ${pageNumber}.`);
}

async function htmlToPdf(formData) {
  const html = String(formData.get("html") || "").trim();
  if (!html) {
    throw new Error("Please paste some HTML to convert.");
  }

  const container = document.createElement("div");
  container.style.position = "fixed";
  container.style.left = "-99999px";
  container.style.top = "0";
  container.style.width = "800px";
  container.style.background = "#ffffff";
  container.style.padding = "24px";
  container.innerHTML = html;
  document.body.appendChild(container);

  try {
    const blob = await window.html2pdf().set({
      margin: 12,
      filename: "html-export.pdf",
      html2canvas: { scale: 2, backgroundColor: "#ffffff" },
      jsPDF: { unit: "pt", format: "a4", orientation: "portrait" }
    }).from(container).outputPdf("blob");
    await downloadBlob(blob, "html-export.pdf", "Converted your HTML markup into a PDF.", "Download PDF");
  } finally {
    container.remove();
  }
}

async function pdfToWord(formData) {
  const file = formData.get("file");
  const mode = formData.get("mode");
  const exportFormat = formData.get("exportFormat") || "word+pdf";

  if (mode === "layout+editable") {
    await launchEditablePdfWorkspace(
      file,
      exportFormat,
      "Your PDF is shown below in Exact Preview mode. Switch to Editable Layer when you want to edit extracted text, then download the Word file.",
      { advancedEditor: false }
    );
    return;
  }

  if (mode === "editable-text") {
    const pages = await extractPdfText(file);
    const blob = await buildDocxFromText("PDF to Word Export", pages);
    await downloadBlob(blob, "pdf-to-word.docx", "Converted the PDF into an editable text DOCX. Layout may differ from the original PDF.", "Download DOCX");
    return;
  }

  const renderedPages = await renderPdfPages(file, { scale: 2 });
  const blob = await buildDocxFromRenderedPages("PDF to Word Export", renderedPages);
  await downloadBlob(blob, "pdf-to-word.docx", "Created a layout-preserving Word file by embedding each PDF page as a full-page image.", "Download DOCX");
}

async function wordToPdf(formData) {
  const file = formData.get("file");
  let text = "";

  if (file.name.toLowerCase().endsWith(".txt")) {
    text = await file.text();
  } else {
    const result = await window.mammoth.extractRawText({ arrayBuffer: await fileToArrayBuffer(file) });
    text = result.value;
  }

  const blob = await buildTextPdf(text, "Word to PDF Export");
  await downloadBlob(blob, "word-to-pdf.pdf", "Converted the document into a readable text-based PDF.", "Download PDF");
}

async function pdfEditor(formData) {
  const file = formData.get("file");
  const exportFormat = formData.get("exportFormat") || "pdf";
  await launchEditablePdfWorkspace(
    file,
    exportFormat,
    "",
    { advancedEditor: true, showStatusMessage: false }
  );
}

async function comparePdfs(formData) {
  const left = formData.get("left");
  const right = formData.get("right");
  const leftPages = await extractPdfText(left);
  const rightPages = await extractPdfText(right);
  const report = summarizeDifferences(leftPages, rightPages);
  const blob = new Blob([report], { type: "text/plain" });
  await downloadBlob(blob, "pdf-comparison-report.txt", report.split("\n").slice(0, 4).join(" | "), "Download Report");
}

function populateDefaults(event) {
  event.preventDefault();
  const tool = getTool();
  if (tool.status !== "live") {
    return;
  }
  const defaults = {
    extract: { pages: "1-2" },
    delete: { pages: "2" },
    reorder: { order: "1" },
    watermark: { text: "CONFIDENTIAL", color: "#d85f3f", opacity: "0.25" },
    number: { format: "Page {n}", position: "bottom-center" },
    annotate: { text: "Approved", x: "60", y: "60" },
    compress: { quality: "0.6", scale: "1.4" },
    ocr: { language: "eng", format: "txt" },
    esign: { page: "1", x: "60", y: "80", width: "140" },
    "html-to-pdf": { html: "<div style='padding:32px;font-family:Arial'><h1>Proposal</h1><p>This PDF was generated from custom HTML.</p></div>" },
    "pdf-to-word": { mode: "layout+editable", exportFormat: "word+pdf" },
    "pdf-editor": { exportFormat: "pdf" }
  };
  const toolDefaults = defaults[tool.id];
  if (!toolDefaults) {
    return;
  }
  Object.entries(toolDefaults).forEach(([name, value]) => {
    const field = toolForm.elements.namedItem(name);
    if (field) {
      field.value = value;
    }
  });
}
