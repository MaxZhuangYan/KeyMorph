import { mkdtemp, stat } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { createDemoDeck } from "../src/demo/createDemoDeck.ts";
import { validateIR } from "../src/ir/index.ts";
import { exportIrToPptx } from "../src/pptx/index.ts";
import { renderHtmlDocument } from "../src/runtime/index.ts";

const deck = createDemoDeck();
const validation = validateIR(deck);
if (!validation.valid) {
  throw new Error(`Demo IR is invalid: ${validation.errors.map((error) => error.message).join("; ")}`);
}

const html = renderHtmlDocument(deck);
if (!html.includes("window.__KEYMORPH_DECK__")) {
  throw new Error("Runtime HTML smoke check failed.");
}

const outDir = await mkdtemp(path.join(tmpdir(), "keymorph-check-"));
const pptxPath = path.join(outDir, "demo.pptx");
await exportIrToPptx(deck, pptxPath);
const pptxStats = await stat(pptxPath);
if (pptxStats.size <= 0) {
  throw new Error("PPTX smoke check failed.");
}

console.log("KeyMorph build smoke check passed.");
