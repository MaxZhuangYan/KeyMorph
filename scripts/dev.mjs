import { createReadStream } from "node:fs";
import { mkdir, stat } from "node:fs/promises";
import http from "node:http";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { createDemoDeck } from "../src/demo/createDemoDeck.ts";
import { parsePptxToIr, exportIrToPptx } from "../src/pptx/index.ts";
import { scoreConversion } from "../src/report/index.ts";
import { renderHtmlDocument } from "../src/runtime/index.ts";

import { writeFile } from "node:fs/promises";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const host = "127.0.0.1";
const port = Number(process.env.PORT ?? 4173);

await generateDemo();

const server = http.createServer(async (request, response) => {
  const url = new URL(request.url ?? "/", `http://${host}:${port}`);
  const pathname = decodeURIComponent(url.pathname === "/" ? "/demo/out/runtime.html" : url.pathname);
  const filePath = path.resolve(root, `.${pathname}`);

  if (!filePath.startsWith(root)) {
    response.writeHead(403);
    response.end("Forbidden");
    return;
  }

  try {
    const stats = await stat(filePath);
    if (!stats.isFile()) throw new Error("Not a file");
    response.writeHead(200, { "Content-Type": contentType(filePath) });
    createReadStream(filePath).pipe(response);
  } catch {
    response.writeHead(404);
    response.end("Not found");
  }
});

server.listen(port, host, () => {
  console.log(`KeyMorph dev server: http://${host}:${port}/demo/out/runtime.html`);
});

async function generateDemo() {
  const outDir = path.join(root, "demo/out");
  await mkdir(outDir, { recursive: true });

  const sourceDeck = createDemoDeck();
  const sourceIrPath = path.join(outDir, "source.ir.json");
  const originalPptxPath = path.join(outDir, "original.pptx");
  const importedIrPath = path.join(outDir, "imported.ir.json");
  const htmlPath = path.join(outDir, "runtime.html");
  const pptxPath = path.join(outDir, "rebuilt.pptx");
  const reportPath = path.join(outDir, "conversion-report.json");

  await writeJson(sourceIrPath, sourceDeck);
  await exportIrToPptx(sourceDeck, originalPptxPath);
  const importedDeck = await parsePptxToIr(originalPptxPath);
  await writeJson(importedIrPath, importedDeck);
  await writeFile(htmlPath, renderHtmlDocument(importedDeck), "utf8");
  await exportIrToPptx(importedDeck, pptxPath);
  await writeJson(reportPath, scoreConversion(importedDeck.conversion ?? { status: "success", messages: [] }));
}

async function writeJson(filePath, value) {
  await writeFile(filePath, `${JSON.stringify(value, null, 2)}\n`, "utf8");
}

function contentType(filePath) {
  if (filePath.endsWith(".html")) return "text/html; charset=utf-8";
  if (filePath.endsWith(".json")) return "application/json; charset=utf-8";
  if (filePath.endsWith(".pptx")) {
    return "application/vnd.openxmlformats-officedocument.presentationml.presentation";
  }
  return "application/octet-stream";
}
