import { createReadStream } from "node:fs";
import { mkdir, readFile, stat, writeFile } from "node:fs/promises";
import http from "node:http";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { randomUUID } from "node:crypto";

import { createDemoDeck } from "../src/demo/createDemoDeck.ts";
import { parsePptxToIr, exportIrToPptx } from "../src/pptx/index.ts";
import { createLossReport, scoreConversion } from "../src/report/index.ts";
import { renderHtmlDocument } from "../src/runtime/index.ts";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const host = "127.0.0.1";
const preferredPort = Number(process.env.PORT ?? 4173);
let activePort = preferredPort;
const outRoot = path.join(root, "demo/out");
const jobsRoot = path.join(outRoot, "jobs");

await generateDemo();
await mkdir(jobsRoot, { recursive: true });

const server = http.createServer(async (request, response) => {
  try {
    const url = new URL(request.url ?? "/", `http://${host}:${activePort}`);

    if ((request.method === "GET" || request.method === "HEAD") && url.pathname === "/") {
      if (request.method === "HEAD") {
        response.writeHead(200, { "Content-Type": "text/html; charset=utf-8" });
        response.end();
        return;
      }
      send(response, 200, renderAppHtml(), "text/html; charset=utf-8");
      return;
    }

    if (request.method === "POST" && url.pathname === "/api/convert") {
      await handleConvert(request, response);
      return;
    }

    if (request.method === "GET" && url.pathname.startsWith("/demo/")) {
      await serveStatic(url.pathname, response);
      return;
    }

    send(response, 404, "Not found", "text/plain; charset=utf-8");
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    send(response, 500, JSON.stringify({ error: message }), "application/json; charset=utf-8");
  }
});

await listenWithFallback(server, preferredPort);

async function handleConvert(request, response) {
  const upload = await readUpload(request);
  const jobId = randomUUID();
  const jobDir = path.join(jobsRoot, jobId);
  await mkdir(jobDir, { recursive: true });

  const safeName = sanitizeFileName(upload.fileName || "deck");
  const sourcePath = path.join(jobDir, safeName);
  await writeFile(sourcePath, upload.data);

  const deck = await inputToIr(sourcePath, safeName, upload.data);
  const importedIrPath = path.join(jobDir, "deck.ir.json");
  const htmlPath = path.join(jobDir, "runtime.html");
  const pptxPath = path.join(jobDir, "rebuilt.pptx");
  const reportPath = path.join(jobDir, "loss-report.json");

  await writeJson(importedIrPath, deck);
  await writeFile(htmlPath, renderHtmlDocument(deck), "utf8");
  await exportIrToPptx(deck, pptxPath);
  await writeJson(reportPath, createLossReport(deck.conversion ?? { status: "success", messages: [] }));

  const report = createLossReport(deck.conversion ?? { status: "success", messages: [] });
  send(
    response,
    200,
    JSON.stringify({
      jobId,
      sourceName: safeName,
      slideCount: deck.deck.slides.length,
      objectCount: deck.conversion?.statistics?.objectCount ?? deck.deck.slides.reduce((sum, slide) => sum + slide.objects.length, 0),
      fidelityScore: report.fidelityScore,
      animationLostCount: report.animationLostCount,
      degradedAnimationCount: report.degradedAnimationCount,
      uncertainMappingCount: report.uncertainMappingCount,
      recommendedFixes: report.recommendedFixes,
      previewUrl: `/demo/out/jobs/${jobId}/runtime.html`,
      downloads: {
        ir: `/demo/out/jobs/${jobId}/deck.ir.json`,
        pptx: `/demo/out/jobs/${jobId}/rebuilt.pptx`,
        report: `/demo/out/jobs/${jobId}/loss-report.json`
      }
    }),
    "application/json; charset=utf-8"
  );
}

function listenWithFallback(server, startPort) {
  return new Promise((resolve, reject) => {
    const tryListen = (port) => {
      const onError = (error) => {
        server.off("listening", onListening);
        if (error.code === "EADDRINUSE" && port < startPort + 20) {
          tryListen(port + 1);
        } else {
          reject(error);
        }
      };
      const onListening = () => {
        server.off("error", onError);
        activePort = port;
        console.log(`KeyMorph product UI: http://${host}:${port}/`);
        resolve();
      };
      server.once("error", onError);
      server.once("listening", onListening);
      server.listen(port, host);
    };
    tryListen(startPort);
  });
}

async function inputToIr(sourcePath, fileName, data) {
  const lowerName = fileName.toLowerCase();
  if (lowerName.endsWith(".pptx")) return parsePptxToIr(sourcePath);
  if (lowerName.endsWith(".json") || lowerName.endsWith(".ir.json")) {
    return JSON.parse(new TextDecoder().decode(data));
  }
  if (lowerName.endsWith(".key")) {
    throw new Error("Native .key import needs a Keynote-exported PPTX bridge in this MVP. Export the Keynote deck as PPTX and drop that file here.");
  }
  throw new Error("Unsupported file type. Drop a .pptx or .ir.json file.");
}

async function readUpload(request) {
  const chunks = [];
  for await (const chunk of request) chunks.push(chunk);
  const body = Buffer.concat(chunks);
  const contentType = request.headers["content-type"] ?? "";

  if (contentType.startsWith("application/json")) {
    const payload = JSON.parse(body.toString("utf8"));
    return {
      fileName: payload.fileName || "deck.ir.json",
      data: Buffer.from(payload.dataBase64, "base64")
    };
  }

  const fileName = decodeURIComponent(String(request.headers["x-keymorph-filename"] ?? "upload.pptx"));
  return { fileName, data: body };
}

async function serveStatic(pathname, response) {
  const filePath = path.resolve(root, `.${decodeURIComponent(pathname)}`);
  if (!filePath.startsWith(root)) {
    send(response, 403, "Forbidden", "text/plain; charset=utf-8");
    return;
  }

  try {
    const stats = await stat(filePath);
    if (!stats.isFile()) throw new Error("Not a file");
    response.writeHead(200, { "Content-Type": contentType(filePath) });
    createReadStream(filePath).pipe(response);
  } catch {
    send(response, 404, "Not found", "text/plain; charset=utf-8");
  }
}

async function generateDemo() {
  await mkdir(outRoot, { recursive: true });

  const sourceDeck = createDemoDeck();
  const sourceIrPath = path.join(outRoot, "source.ir.json");
  const originalPptxPath = path.join(outRoot, "original.pptx");
  const importedIrPath = path.join(outRoot, "imported.ir.json");
  const htmlPath = path.join(outRoot, "runtime.html");
  const pptxPath = path.join(outRoot, "rebuilt.pptx");
  const reportPath = path.join(outRoot, "conversion-report.json");

  await writeJson(sourceIrPath, sourceDeck);
  await exportIrToPptx(sourceDeck, originalPptxPath);
  const importedDeck = await parsePptxToIr(originalPptxPath);
  await writeJson(importedIrPath, importedDeck);
  await writeFile(htmlPath, renderHtmlDocument(importedDeck), "utf8");
  await exportIrToPptx(importedDeck, pptxPath);
  await writeJson(reportPath, scoreConversion(importedDeck.conversion ?? { status: "success", messages: [] }));
}

function renderAppHtml() {
  return `<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>KeyMorph</title>
  <style>
    :root { color-scheme: light; font-family: Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; }
    body { margin: 0; min-height: 100vh; background: #f4f7fb; color: #111827; display: grid; grid-template-columns: minmax(320px, 420px) 1fr; }
    aside { background: #ffffff; border-right: 1px solid #d8dee8; padding: 22px; display: flex; flex-direction: column; gap: 18px; }
    main { min-width: 0; display: grid; grid-template-rows: auto 1fr; }
    h1 { margin: 0; font-size: 22px; line-height: 1.15; }
    .sub { color: #64748b; font-size: 13px; line-height: 1.45; }
    .drop { border: 2px dashed #94a3b8; border-radius: 8px; background: #f8fafc; min-height: 180px; padding: 18px; display: grid; place-items: center; text-align: center; transition: .16s border-color, .16s background; }
    .drop[data-active="true"] { border-color: #0f766e; background: #ecfdf5; }
    .drop strong { display: block; margin-bottom: 6px; font-size: 17px; }
    input[type=file] { display: none; }
    button, .file-button, a.download { border: 1px solid #cbd5e1; border-radius: 6px; background: #fff; color: #0f172a; padding: 9px 11px; font: inherit; cursor: pointer; text-decoration: none; display: inline-flex; align-items: center; justify-content: center; gap: 8px; }
    button.primary, .file-button { background: #0f766e; color: #fff; border-color: #0f766e; }
    .row { display: flex; gap: 10px; flex-wrap: wrap; }
    .status { border: 1px solid #d8dee8; border-radius: 8px; padding: 12px; background: #fff; font-size: 13px; color: #334155; }
    .metric { display: grid; grid-template-columns: 1fr auto; gap: 8px; padding: 9px 0; border-bottom: 1px solid #e2e8f0; }
    .metric:last-child { border-bottom: 0; }
    .metric span:first-child { color: #64748b; }
    .downloads { display: grid; gap: 8px; }
    .topbar { height: 56px; background: #ffffff; border-bottom: 1px solid #d8dee8; display: flex; align-items: center; justify-content: space-between; padding: 0 18px; }
    iframe { width: 100%; height: 100%; border: 0; background: #202124; }
    .empty { height: 100%; display: grid; place-items: center; color: #64748b; }
    @media (max-width: 900px) { body { grid-template-columns: 1fr; grid-template-rows: auto 1fr; } aside { border-right: 0; border-bottom: 1px solid #d8dee8; } }
  </style>
</head>
<body>
  <aside>
    <div>
      <h1>KeyMorph</h1>
      <div class="sub">Drop a PowerPoint deck or KeyMorph IR. Conversion runs locally in this dev server and returns preview, PPTX, IR, and loss report.</div>
    </div>
    <label id="drop" class="drop">
      <input id="file" type="file" accept=".pptx,.json,.key">
      <span><strong>Drop deck here</strong><span class="sub">or click to choose .pptx / .ir.json</span></span>
    </label>
    <div class="row">
      <label class="file-button" for="file">Choose file</label>
      <button id="loadDemo" type="button">Load sample</button>
    </div>
    <div id="status" class="status">Waiting for a deck.</div>
    <div id="metrics" class="status" hidden></div>
    <div id="downloads" class="downloads" hidden></div>
  </aside>
  <main>
    <div class="topbar">
      <strong id="previewTitle">Preview</strong>
      <a id="openPreview" class="download" href="/demo/out/runtime.html" target="_blank">Open preview</a>
    </div>
    <div id="previewPane" class="empty">Drop a deck to generate an interactive HTML runtime.</div>
  </main>
  <script>
    const drop = document.getElementById('drop');
    const fileInput = document.getElementById('file');
    const statusEl = document.getElementById('status');
    const metricsEl = document.getElementById('metrics');
    const downloadsEl = document.getElementById('downloads');
    const previewPane = document.getElementById('previewPane');
    const previewTitle = document.getElementById('previewTitle');
    const openPreview = document.getElementById('openPreview');
    const loadDemo = document.getElementById('loadDemo');

    drop.addEventListener('dragover', (event) => { event.preventDefault(); drop.dataset.active = 'true'; });
    drop.addEventListener('dragleave', () => { drop.dataset.active = 'false'; });
    drop.addEventListener('drop', async (event) => {
      event.preventDefault();
      drop.dataset.active = 'false';
      const file = event.dataTransfer.files[0];
      if (file) await convertFile(file);
    });
    fileInput.addEventListener('change', async () => {
      const file = fileInput.files[0];
      if (file) await convertFile(file);
    });
    loadDemo.addEventListener('click', async () => {
      const res = await fetch('/demo/out/original.pptx');
      const blob = await res.blob();
      await convertFile(new File([blob], 'sample.pptx', { type: blob.type }));
    });

    async function convertFile(file) {
      statusEl.textContent = 'Converting ' + file.name + '...';
      metricsEl.hidden = true;
      downloadsEl.hidden = true;
      const response = await fetch('/api/convert', {
        method: 'POST',
        headers: { 'Content-Type': 'application/octet-stream', 'X-KeyMorph-Filename': encodeURIComponent(file.name) },
        body: await file.arrayBuffer()
      });
      const payload = await response.json();
      if (!response.ok) throw new Error(payload.error || 'Conversion failed');
      renderResult(payload);
    }

    function renderResult(result) {
      statusEl.textContent = 'Converted ' + result.sourceName + '.';
      metricsEl.hidden = false;
      metricsEl.innerHTML = [
        ['Slides', result.slideCount],
        ['Objects', result.objectCount],
        ['Fidelity score', result.fidelityScore],
        ['Lost animations', result.animationLostCount],
        ['Degraded animations', result.degradedAnimationCount],
        ['Uncertain mappings', result.uncertainMappingCount]
      ].map(([label, value]) => '<div class="metric"><span>' + label + '</span><strong>' + value + '</strong></div>').join('');
      downloadsEl.hidden = false;
      downloadsEl.innerHTML = '<a class="download" href="' + result.downloads.pptx + '">Download PPTX</a><a class="download" href="' + result.downloads.ir + '">Download IR</a><a class="download" href="' + result.downloads.report + '">Download report</a>';
      previewPane.className = '';
      previewPane.innerHTML = '<iframe title="KeyMorph runtime preview" src="' + result.previewUrl + '"></iframe>';
      previewTitle.textContent = result.sourceName;
      openPreview.href = result.previewUrl;
    }
  </script>
</body>
</html>`;
}

async function writeJson(filePath, value) {
  await writeFile(filePath, `${JSON.stringify(value, null, 2)}\n`, "utf8");
}

function sanitizeFileName(fileName) {
  return path.basename(fileName).replace(/[^a-zA-Z0-9._-]/g, "_");
}

function send(response, status, body, type) {
  response.writeHead(status, { "Content-Type": type });
  response.end(body);
}

function contentType(filePath) {
  if (filePath.endsWith(".html")) return "text/html; charset=utf-8";
  if (filePath.endsWith(".json")) return "application/json; charset=utf-8";
  if (filePath.endsWith(".pptx")) {
    return "application/vnd.openxmlformats-officedocument.presentationml.presentation";
  }
  return "application/octet-stream";
}
