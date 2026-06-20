import { createReadStream } from "node:fs";
import { mkdir, readFile, stat, writeFile } from "node:fs/promises";
import http from "node:http";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { randomUUID } from "node:crypto";

import { createDemoDeck } from "../src/demo/createDemoDeck.ts";
import { parsePptxToIr, exportIrToPptx } from "../src/pptx/index.ts";
import { exportIrToKeynote, parseKeynoteToIr } from "../src/keynote/index.ts";
import { createLossReport, scoreConversion } from "../src/report/index.ts";
import { renderHtmlDocument } from "../src/runtime/index.ts";
import { createVideoExportPlan, exportIrToVideo, VideoExportDependencyError } from "../src/video/index.ts";

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

    const videoMatch = /^\/api\/jobs\/([a-f0-9-]+)\/video$/.exec(url.pathname);
    if (request.method === "POST" && videoMatch) {
      await handleVideoExport(videoMatch[1], url, response);
      return;
    }

    const keynoteMatch = /^\/api\/jobs\/([a-f0-9-]+)\/keynote$/.exec(url.pathname);
    if (request.method === "POST" && keynoteMatch) {
      await handleKeynoteExport(keynoteMatch[1], response);
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

  const report = createLossReport(deck.conversion ?? { status: "success", messages: [] });
  const videoPlan = summarizeVideoPlan(createVideoExportPlan(deck));
  await writeJson(reportPath, report);
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
      videoPlan,
      videoEndpoint: `/api/jobs/${jobId}/video`,
      keynoteEndpoint: `/api/jobs/${jobId}/keynote`,
      downloads: {
        html: `/demo/out/jobs/${jobId}/runtime.html`,
        ir: `/demo/out/jobs/${jobId}/deck.ir.json`,
        pptx: `/demo/out/jobs/${jobId}/rebuilt.pptx`,
        key: null,
        report: `/demo/out/jobs/${jobId}/loss-report.json`,
        video: null
      },
      keynoteAvailable: null,
      keynoteMessage: "Keynote export is available on demand and may ask macOS for Keynote automation permission."
    }),
    "application/json; charset=utf-8"
  );
}

async function handleKeynoteExport(jobId, response) {
  const jobDir = safeJobDir(jobId);
  const deck = await readJson(path.join(jobDir, "deck.ir.json"));
  const keynotePath = path.join(jobDir, "rebuilt.key");
  const result = await tryExportKeynote(deck, keynotePath, path.join(jobDir, "rebuilt.key-bridge.pptx"));
  send(
    response,
    result.available ? 200 : 424,
    JSON.stringify({
      status: result.available ? "ready" : "dependency-missing",
      keyUrl: result.available ? `/demo/out/jobs/${jobId}/rebuilt.key` : null,
      message: result.message
    }),
    "application/json; charset=utf-8"
  );
}

async function handleVideoExport(jobId, url, response) {
  const jobDir = safeJobDir(jobId);
  const deck = await readJson(path.join(jobDir, "deck.ir.json"));
  const outputPath = path.join(jobDir, "render.mp4");
  const options = {
    fps: parsePositiveNumber(url.searchParams.get("fps")) ?? 30,
    scale: parsePositiveNumber(url.searchParams.get("scale")) ?? 4
  };
  const plan = createVideoExportPlan(deck, options);

  try {
    await exportIrToVideo(deck, outputPath, options);
    send(
      response,
      200,
      JSON.stringify({
        status: "ready",
        videoUrl: `/demo/out/jobs/${jobId}/render.mp4`,
        plan: summarizeVideoPlan(plan)
      }),
      "application/json; charset=utf-8"
    );
  } catch (error) {
    if (error instanceof VideoExportDependencyError) {
      send(
        response,
        424,
        JSON.stringify({
          status: "dependency-missing",
          missing: error.missing,
          message: error.message,
          plan: summarizeVideoPlan(plan)
        }),
        "application/json; charset=utf-8"
      );
      return;
    }
    throw error;
  }
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
    return parseKeynoteToIr(sourcePath, { workDir: path.dirname(sourcePath) });
  }
  throw new Error("Unsupported file type. Drop a .pptx, .key, or .ir.json file.");
}

async function tryExportKeynote(deck, keynotePath, intermediatePptxPath) {
  try {
    await exportIrToKeynote(deck, keynotePath, { intermediatePptxPath });
    return { available: true, message: "Keynote file generated." };
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    return {
      available: false,
      message: `Keynote export unavailable: ${message}`
    };
  }
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
    button:disabled { opacity: .58; cursor: wait; }
    button.primary, .file-button { background: #0f766e; color: #fff; border-color: #0f766e; }
    .row { display: flex; gap: 10px; flex-wrap: wrap; }
    .status { border: 1px solid #d8dee8; border-radius: 8px; padding: 12px; background: #fff; font-size: 13px; color: #334155; }
    .status.error { border-color: #fecaca; background: #fef2f2; color: #991b1b; }
    .status.warn { border-color: #fde68a; background: #fffbeb; color: #92400e; }
    .metric { display: grid; grid-template-columns: 1fr auto; gap: 8px; padding: 9px 0; border-bottom: 1px solid #e2e8f0; }
    .metric:last-child { border-bottom: 0; }
    .metric span:first-child { color: #64748b; }
    .downloads { display: grid; gap: 8px; }
    .fixes { margin: 0; padding-left: 18px; color: #475569; }
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
      <div class="sub">Drop a PowerPoint, Keynote, or KeyMorph IR deck. Conversion runs locally and returns runtime preview, editable exports, video export, IR, and report.</div>
    </div>
    <label id="drop" class="drop">
      <input id="file" type="file" accept=".pptx,.json,.key">
      <span><strong>Drop deck here</strong><span class="sub">or click to choose .pptx / .key / .ir.json</span></span>
    </label>
    <div class="row">
      <label class="file-button" for="file">Choose file</label>
      <button id="loadDemo" type="button">Load sample</button>
    </div>
    <div id="status" class="status">Waiting for a deck.</div>
    <div id="metrics" class="status" hidden></div>
    <div id="recommendations" class="status" hidden></div>
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
    const recommendationsEl = document.getElementById('recommendations');
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
      statusEl.className = 'status';
      metricsEl.hidden = true;
      recommendationsEl.hidden = true;
      downloadsEl.hidden = true;
      previewPane.className = 'empty';
      previewPane.textContent = 'Converting deck...';
      try {
        const response = await fetch('/api/convert', {
          method: 'POST',
          headers: { 'Content-Type': 'application/octet-stream', 'X-KeyMorph-Filename': encodeURIComponent(file.name) },
          body: await file.arrayBuffer()
        });
        const payload = await response.json();
        if (!response.ok) throw new Error(payload.error || 'Conversion failed');
        renderResult(payload);
      } catch (error) {
        statusEl.className = 'status error';
        statusEl.textContent = error instanceof Error ? error.message : String(error);
        previewPane.className = 'empty';
        previewPane.textContent = 'Conversion failed.';
      }
    }

    function renderResult(result) {
      statusEl.textContent = 'Converted ' + result.sourceName + '.';
      statusEl.className = 'status';
      metricsEl.hidden = false;
      metricsEl.innerHTML = [
        ['Slides', result.slideCount],
        ['Objects', result.objectCount],
        ['Fidelity score', result.fidelityScore],
        ['Lost animations', result.animationLostCount],
        ['Degraded animations', result.degradedAnimationCount],
        ['Uncertain mappings', result.uncertainMappingCount],
        ['Video frames', result.videoPlan.totalFrames + ' @ ' + result.videoPlan.fps + ' fps']
      ].map(([label, value]) => '<div class="metric"><span>' + label + '</span><strong>' + value + '</strong></div>').join('');
      if (result.recommendedFixes && result.recommendedFixes.length) {
        recommendationsEl.hidden = false;
        recommendationsEl.innerHTML = '<strong>Recommended fixes</strong><ul class="fixes">' + result.recommendedFixes.map((fix) => '<li>' + escapeHtml(fix) + '</li>').join('') + '</ul>';
      } else {
        recommendationsEl.hidden = true;
      }
      downloadsEl.hidden = false;
      downloadsEl.innerHTML =
        '<a class="download" href="' + result.downloads.html + '" target="_blank">Download HTML Runtime</a>' +
        '<a class="download" href="' + result.downloads.pptx + '">Download PPTX</a>' +
        '<button id="exportKeynote" type="button">Export Keynote</button>' +
        '<a class="download" href="' + result.downloads.ir + '">Download IR</a>' +
        '<a class="download" href="' + result.downloads.report + '">Download report</a>' +
        '<button id="renderVideo" class="primary" type="button">Render MP4</button><div id="keynoteStatus" class="status warn">' + escapeHtml(result.keynoteMessage || '') + '</div><div id="videoStatus" class="status" hidden></div>';
      document.getElementById('exportKeynote')?.addEventListener('click', () => exportKeynote(result));
      document.getElementById('renderVideo')?.addEventListener('click', () => renderVideo(result));
      previewPane.className = '';
      previewPane.innerHTML = '<iframe title="KeyMorph runtime preview" src="' + result.previewUrl + '"></iframe>';
      previewTitle.textContent = result.sourceName;
      openPreview.href = result.previewUrl;
    }

    async function exportKeynote(result) {
      const button = document.getElementById('exportKeynote');
      const keynoteStatus = document.getElementById('keynoteStatus');
      button.disabled = true;
      keynoteStatus.className = 'status';
      keynoteStatus.textContent = 'Exporting Keynote via local Keynote automation...';
      try {
        const response = await fetch(result.keynoteEndpoint, { method: 'POST' });
        const payload = await response.json();
        if (!response.ok) {
          keynoteStatus.className = 'status warn';
          keynoteStatus.textContent = payload.message || 'Keynote export is unavailable.';
          return;
        }
        keynoteStatus.className = 'status';
        keynoteStatus.innerHTML = 'Keynote ready. <a class="download" href="' + payload.keyUrl + '">Download Keynote</a>';
      } catch (error) {
        keynoteStatus.className = 'status error';
        keynoteStatus.textContent = error instanceof Error ? error.message : String(error);
      } finally {
        button.disabled = false;
      }
    }

    async function renderVideo(result) {
      const button = document.getElementById('renderVideo');
      const videoStatus = document.getElementById('videoStatus');
      button.disabled = true;
      videoStatus.hidden = false;
      videoStatus.className = 'status';
      videoStatus.textContent = 'Rendering ' + result.videoPlan.totalFrames + ' frames at 4x resolution...';
      try {
        const response = await fetch(result.videoEndpoint, { method: 'POST' });
        const payload = await response.json();
        if (!response.ok) {
          videoStatus.className = response.status === 424 ? 'status warn' : 'status error';
          videoStatus.textContent = payload.message || 'Video export failed.';
          if (payload.missing?.length) {
            videoStatus.textContent += ' Missing: ' + payload.missing.join(', ') + '.';
          }
          return;
        }
        videoStatus.className = 'status';
        videoStatus.innerHTML = 'MP4 ready. <a class="download" href="' + payload.videoUrl + '">Download video</a>';
      } catch (error) {
        videoStatus.className = 'status error';
        videoStatus.textContent = error instanceof Error ? error.message : String(error);
      } finally {
        button.disabled = false;
      }
    }

    function escapeHtml(value) {
      return String(value ?? '').replace(/[&<>"']/g, (char) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[char]));
    }
  </script>
</body>
</html>`;
}

async function writeJson(filePath, value) {
  await writeFile(filePath, `${JSON.stringify(value, null, 2)}\n`, "utf8");
}

async function readJson(filePath) {
  return JSON.parse(await readFile(filePath, "utf8"));
}

function sanitizeFileName(fileName) {
  return path.basename(fileName).replace(/[^a-zA-Z0-9._-]/g, "_");
}

function safeJobDir(jobId) {
  if (!/^[a-f0-9-]+$/.test(jobId)) throw new Error("Invalid job id.");
  const jobDir = path.resolve(jobsRoot, jobId);
  if (!jobDir.startsWith(jobsRoot)) throw new Error("Invalid job path.");
  return jobDir;
}

function parsePositiveNumber(value) {
  if (!value) return undefined;
  const number = Number(value);
  if (!Number.isFinite(number) || number <= 0) return undefined;
  return number;
}

function summarizeVideoPlan(plan) {
  return {
    width: plan.width,
    height: plan.height,
    scale: plan.scale,
    fps: plan.fps,
    durationMs: plan.durationMs,
    totalFrames: plan.totalFrames,
    outputWidth: plan.outputWidth,
    outputHeight: plan.outputHeight
  };
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
  if (filePath.endsWith(".key")) return "application/vnd.apple.keynote";
  if (filePath.endsWith(".mp4")) return "video/mp4";
  return "application/octet-stream";
}
