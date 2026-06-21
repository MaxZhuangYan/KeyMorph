import { createReadStream } from "node:fs";
import { mkdir, readFile, stat, writeFile } from "node:fs/promises";
import http from "node:http";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { randomUUID } from "node:crypto";

import { createDemoDeck } from "../src/demo/createDemoDeck.ts";
import { parsePptxToIr, exportIrToPptx } from "../src/pptx/index.ts";
import { scoreConversion } from "../src/report/index.ts";
import { renderHtmlDocument } from "../src/runtime/index.ts";
import {
  createProductApiResponse,
  createProductBundle,
  createProductFrameFidelityInsights,
  exportProductBundleBaseline,
  exportProductBundleKeynote,
  exportProductBundleVideo
} from "../src/cli.ts";
import { VideoExportDependencyError } from "../src/video/index.ts";

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

    const baselineMatch = /^\/api\/jobs\/([a-f0-9-]+)\/baseline$/.exec(url.pathname);
    if (request.method === "POST" && baselineMatch) {
      await handleBaselineExport(baselineMatch[1], url, response);
      return;
    }

    const keynoteMatch = /^\/api\/jobs\/([a-f0-9-]+)\/keynote$/.exec(url.pathname);
    if (request.method === "POST" && keynoteMatch) {
      await handleKeynoteExport(keynoteMatch[1], url, response);
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

  const bundle = await createProductBundle(sourcePath, jobDir, { sourceName: safeName, jobId });
  send(
    response,
    200,
    JSON.stringify(createProductApiResponse(bundle, `/demo/out/jobs/${jobId}`)),
    "application/json; charset=utf-8"
  );
}

async function handleKeynoteExport(jobId, url, response) {
  const jobDir = safeJobDir(jobId);
  const result = await exportProductBundleKeynote(jobDir, {
    allowKeynoteAutomation: url.searchParams.get("allowKeynote") === "1" || process.env.KEYMORPH_ALLOW_KEYNOTE_AUTOMATION === "1"
  });
  send(
    response,
    result.status === "ready" ? 200 : 424,
    JSON.stringify({
      status: result.status,
      keyUrl: result.status === "ready" ? `/demo/out/jobs/${jobId}/rebuilt.key` : null,
      message: result.message
    }),
    "application/json; charset=utf-8"
  );
}

async function handleVideoExport(jobId, url, response) {
  const jobDir = safeJobDir(jobId);
  const options = {
    fps: parsePositiveNumber(url.searchParams.get("fps")) ?? 30,
    scale: parsePositiveNumber(url.searchParams.get("scale")) ?? 4
  };

  try {
    const result = await exportProductBundleVideo(jobDir, options);
    send(
      response,
      200,
      JSON.stringify({
        status: "ready",
        videoUrl: `/demo/out/jobs/${jobId}/render.mp4`,
        plan: result.plan,
        framesUrl: `/demo/out/jobs/${jobId}/frames/latest/`,
        frameFidelity: result.frameFidelity
          ? {
              reportUrl: `/demo/out/jobs/${jobId}/frame-fidelity.json`,
              diffUrl: `/demo/out/jobs/${jobId}/frame-diffs/`,
              summary: result.frameFidelity.summary
            }
          : null
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
          guidance: error.guidance,
          plan: null
        }),
        "application/json; charset=utf-8"
      );
      return;
    }
    throw error;
  }
}

async function handleBaselineExport(jobId, url, response) {
  const jobDir = safeJobDir(jobId);
  const options = {
    allowKeynoteAutomation: url.searchParams.get("allowKeynote") === "1" || process.env.KEYMORPH_ALLOW_KEYNOTE_AUTOMATION === "1",
    video: {
      fps: parsePositiveNumber(url.searchParams.get("fps")) ?? 30,
      scale: parsePositiveNumber(url.searchParams.get("scale")) ?? 4
    }
  };

  const result = await exportProductBundleBaseline(jobDir, options);
  if (result.status !== "ready") {
    send(
      response,
      result.status === "unsupported" ? 400 : 424,
      JSON.stringify({
        status: result.status,
        message: result.message,
        missing: result.missing,
        guidance: result.guidance
      }),
      "application/json; charset=utf-8"
    );
    return;
  }
  const insightReport = JSON.parse(await readFile(path.join(jobDir, "baseline-fidelity.json"), "utf8"));
  const insights = createProductFrameFidelityInsights(insightReport, { frameLimit: 5, slideLimit: 5 });
  const toJobUrl = (filePath) => filePath ? `/demo/out/jobs/${jobId}/${path.relative(jobDir, filePath).split(path.sep).join('/')}` : null;

  send(
    response,
    200,
    JSON.stringify({
      status: "ready",
      source: result.source,
      plan: result.plan,
      baselineMovieUrl: result.referenceMoviePath ? `/demo/out/jobs/${jobId}/baseline/keynote-reference.m4v` : null,
      baselineFramesUrl: `/demo/out/jobs/${jobId}/frames/baseline/`,
      actualFramesUrl: `/demo/out/jobs/${jobId}/frames/keymorph-baseline/`,
      reportUrl: `/demo/out/jobs/${jobId}/baseline-fidelity.json`,
      diffUrl: `/demo/out/jobs/${jobId}/baseline-diffs/`,
      insights: {
        worstFrames: insights.worstFrames.map((frame) => ({
          ...frame,
          referenceUrl: toJobUrl(frame.referencePath),
          actualUrl: toJobUrl(frame.actualPath),
          diffUrl: toJobUrl(frame.diffPath)
        })),
        worstSlides: insights.worstSlides
      },
      summary: result.summary
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
    if (stats.isDirectory()) {
      send(response, 200, JSON.stringify({ path: pathname, message: "Directory listing is not enabled." }), "application/json; charset=utf-8");
      return;
    }
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
    select.lang-select { border: 1px solid #cbd5e1; border-radius: 6px; background: #fff; color: #0f172a; padding: 8px 10px; font: inherit; }
    button:disabled { opacity: .58; cursor: not-allowed; }
    button.primary, .file-button { background: #0f766e; color: #fff; border-color: #0f766e; }
    .row { display: flex; gap: 10px; flex-wrap: wrap; }
    .language-row { align-items: center; justify-content: space-between; }
    .status { border: 1px solid #d8dee8; border-radius: 8px; padding: 12px; background: #fff; font-size: 13px; color: #334155; }
    .status.error { border-color: #fecaca; background: #fef2f2; color: #991b1b; }
    .status.warn { border-color: #fde68a; background: #fffbeb; color: #92400e; }
    .metric { display: grid; grid-template-columns: 1fr auto; gap: 8px; padding: 9px 0; border-bottom: 1px solid #e2e8f0; }
    .metric:last-child { border-bottom: 0; }
    .metric span:first-child { color: #64748b; }
    .downloads { display: grid; gap: 10px; }
    .action-group { border: 1px solid #d8dee8; border-radius: 8px; background: #fff; padding: 10px; display: grid; gap: 8px; }
    .action-title { display: flex; justify-content: space-between; align-items: center; gap: 8px; font-weight: 650; font-size: 13px; }
    .link-row { display: flex; gap: 8px; flex-wrap: wrap; }
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
      <div class="sub" data-i18n="appSubtitle">Drop a PowerPoint, Keynote, or KeyMorph IR deck. Conversion runs locally and returns runtime preview, editable exports, video export, IR, and report.</div>
    </div>
    <div class="row language-row">
      <label class="sub" for="language" data-i18n="language">Language</label>
      <select id="language" class="lang-select" aria-label="Language">
        <option value="en">English</option>
        <option value="zh-Hans">简体中文</option>
        <option value="zh-Hant">繁體中文</option>
      </select>
    </div>
    <label id="drop" class="drop">
      <input id="file" type="file" accept=".pptx,.json,.key">
      <span><strong data-i18n="dropTitle">Drop deck here</strong><span class="sub" data-i18n="dropHint">or click to choose .pptx / .key / .ir.json</span></span>
    </label>
    <div class="row">
      <label class="file-button" for="file" data-i18n="chooseFile">Choose file</label>
      <button id="loadDemo" type="button" data-i18n="loadSample">Load sample</button>
    </div>
    <div id="status" class="status" data-i18n="waiting">Waiting for a deck.</div>
    <div id="metrics" class="status" hidden></div>
    <div id="recommendations" class="status" hidden></div>
    <div id="downloads" class="downloads" hidden></div>
  </aside>
  <main>
    <div class="topbar">
      <strong id="previewTitle" data-i18n="preview">Preview</strong>
      <a id="openPreview" class="download" href="/demo/out/runtime.html" target="_blank" data-i18n="openPreview">Open preview</a>
    </div>
    <div id="previewPane" class="empty" data-i18n="emptyPreview">Drop a deck to generate an interactive HTML runtime.</div>
  </main>
  <script>
    const translations = {
      en: {
        appSubtitle: 'Drop a PowerPoint, Keynote, or KeyMorph IR deck. Conversion runs locally and returns runtime preview, editable exports, video export, IR, and report.',
        language: 'Language',
        dropTitle: 'Drop deck here',
        dropHint: 'or click to choose .pptx / .key / .ir.json',
        chooseFile: 'Choose file',
        loadSample: 'Load sample',
        waiting: 'Waiting for a deck.',
        preview: 'Preview',
        openPreview: 'Open preview',
        emptyPreview: 'Drop a deck to generate an interactive HTML runtime.',
        converting: 'Converting {name}...',
        convertingDeck: 'Converting locally. Preview and export actions will appear when ready.',
        unsupportedFile: 'Unsupported file. Use .pptx, .key, or .ir.json.',
        conversionFailed: 'Conversion failed.',
        converted: 'Converted {name}.',
        slides: 'Slides',
        objects: 'Objects',
        conversionFidelity: 'Conversion fidelity',
        runtimeMode: 'Runtime mode',
        lostAnimations: 'Lost animations',
        degradedAnimations: 'Degraded animations',
        uncertainMappings: 'Uncertain mappings',
        videoFrames: 'Video frames',
        frameCapture: 'Frame capture',
        mp4Encoder: 'MP4 encoder',
        recommendedFixes: 'Recommended fixes',
        downloadsTitle: 'Downloads',
        runtimeIr: 'KeyMorph IR reconstruction',
        runtimeKeynoteHtml: 'Keynote native HTML',
        runtimeKeynoteMovie: 'Keynote rendered movie',
        downloadHtml: 'Download HTML Runtime',
        downloadPptx: 'Download PPTX',
        exportKeynote: 'Export Keynote',
        downloadIr: 'Download IR',
        downloadReport: 'Download report',
        renderMp4: 'Render MP4',
        downloadVideoStatus: 'Download video status',
        openRenderedFrames: 'Open rendered frames',
        openFrameDiffs: 'Open frame diffs',
        renderBaseline: 'Run Keynote baseline',
        baselineHint: 'Compare Keynote reference frames against the HTML runtime. macOS may ask for Keynote automation permission.',
        baselineRequiresKeynote: 'Keynote baseline requires an original .key source deck.',
        baselineRunning: 'Rendering Keynote reference frames and comparing the HTML runtime...',
        baselineFailed: 'Keynote baseline comparison failed.',
        baselineReady: 'Baseline fidelity ready.',
        baselineUnavailable: 'Keynote baseline is unavailable.',
        downloadBaselineReport: 'Download baseline report',
        downloadBaselineMovie: 'Download reference movie',
        openBaselineDiffs: 'Open baseline diffs',
        openReferenceFrames: 'Open reference frames',
        openRuntimeFrames: 'Open runtime frames',
        baselineSource: 'Source: {source}.',
        baselineFidelityMean: 'Baseline mean score: {score}.',
        worstFramesTitle: 'Worst frames',
        worstSlidesTitle: 'Worst slides',
        worstFrameLabel: 'Frame {frame} · slide {slide} · score {score}',
        worstSlideLabel: 'Slide {slide} · mean {score} · worst frame {frame}',
        openDiff: 'Diff',
        openReference: 'Reference',
        openRuntime: 'Runtime',
        keynoteExporting: 'Exporting Keynote via local Keynote automation...',
        keynoteUnavailable: 'Keynote export is unavailable.',
        keynoteReady: 'Keynote ready.',
        downloadKeynote: 'Download Keynote',
        renderingVideo: 'Rendering {frames} frames at 4x resolution...',
        videoReadyToRender: 'Ready to render {frames} frames at {fps} fps ({width}x{height}).',
        videoMissingDependencies: 'MP4 render is unavailable. Missing: {items}.',
        videoFailed: 'Video export failed.',
        missing: 'Missing: {items}.',
        mp4Ready: 'MP4 ready.',
        downloadVideo: 'Download video',
        frameReportPending: 'Frame fidelity report will be created after a second render of this job.',
        frameFidelityMean: 'Frame fidelity mean score: {score}.',
        downloadFrameReport: 'Download frame report',
        ready: 'ready',
        missingValue: 'missing',
        unknown: 'unknown'
      },
      'zh-Hans': {
        appSubtitle: '拖入 PowerPoint、Keynote 或 KeyMorph IR 文件。本地完成转换，并生成运行时预览、可编辑导出、视频导出、IR 和报告。',
        language: '语言',
        dropTitle: '把演示文稿拖到这里',
        dropHint: '或点击选择 .pptx / .key / .ir.json',
        chooseFile: '选择文件',
        loadSample: '加载示例',
        waiting: '等待演示文稿。',
        preview: '预览',
        openPreview: '打开预览',
        emptyPreview: '拖入演示文稿后生成可交互 HTML 运行时。',
        converting: '正在转换 {name}...',
        convertingDeck: '正在本地转换。预览和导出动作会在完成后出现。',
        unsupportedFile: '不支持的文件。请使用 .pptx、.key 或 .ir.json。',
        conversionFailed: '转换失败。',
        converted: '已转换 {name}。',
        slides: '幻灯片',
        objects: '对象',
        conversionFidelity: '转换保真度',
        runtimeMode: '运行时模式',
        lostAnimations: '丢失动画',
        degradedAnimations: '降级动画',
        uncertainMappings: '不确定映射',
        videoFrames: '视频帧',
        frameCapture: '帧捕获',
        mp4Encoder: 'MP4 编码器',
        recommendedFixes: '建议修复',
        downloadsTitle: '下载',
        runtimeIr: 'KeyMorph IR 重建',
        runtimeKeynoteHtml: 'Keynote 原生 HTML',
        runtimeKeynoteMovie: 'Keynote 渲染影片',
        downloadHtml: '下载 HTML 运行时',
        downloadPptx: '下载 PPTX',
        exportKeynote: '导出 Keynote',
        downloadIr: '下载 IR',
        downloadReport: '下载报告',
        renderMp4: '渲染 MP4',
        downloadVideoStatus: '下载视频状态',
        openRenderedFrames: '打开渲染帧',
        openFrameDiffs: '打开帧差异',
        renderBaseline: '运行 Keynote 基准对比',
        baselineHint: '将 Keynote 官方参考帧与 HTML 运行时逐帧对比。macOS 可能会请求 Keynote 自动化权限。',
        baselineRequiresKeynote: 'Keynote 基准对比需要原始 .key 源文件。',
        baselineRunning: '正在渲染 Keynote 官方基准帧，并与 HTML 运行时逐帧对比...',
        baselineFailed: 'Keynote 基准对比失败。',
        baselineReady: '基准保真度已生成。',
        baselineUnavailable: 'Keynote 基准对比不可用。',
        downloadBaselineReport: '下载基准报告',
        downloadBaselineMovie: '下载参考影片',
        openBaselineDiffs: '打开基准差异帧',
        openReferenceFrames: '打开参考帧',
        openRuntimeFrames: '打开运行时帧',
        baselineSource: '来源：{source}。',
        baselineFidelityMean: '基准平均分：{score}。',
        worstFramesTitle: '最差帧',
        worstSlidesTitle: '最差页',
        worstFrameLabel: '第 {frame} 帧 · 第 {slide} 页 · 分数 {score}',
        worstSlideLabel: '第 {slide} 页 · 平均 {score} · 最差帧 {frame}',
        openDiff: '差异图',
        openReference: '参考帧',
        openRuntime: '运行时帧',
        keynoteExporting: '正在通过本地 Keynote 自动化导出...',
        keynoteUnavailable: 'Keynote 导出不可用。',
        keynoteReady: 'Keynote 已就绪。',
        downloadKeynote: '下载 Keynote',
        renderingVideo: '正在以 4x 分辨率渲染 {frames} 帧...',
        videoReadyToRender: '可渲染 {frames} 帧，{fps} fps（{width}x{height}）。',
        videoMissingDependencies: 'MP4 渲染不可用。缺少：{items}。',
        videoFailed: '视频导出失败。',
        missing: '缺少：{items}。',
        mp4Ready: 'MP4 已就绪。',
        downloadVideo: '下载视频',
        frameReportPending: '第二次渲染同一个任务后会生成逐帧保真度报告。',
        frameFidelityMean: '逐帧保真度平均分：{score}。',
        downloadFrameReport: '下载逐帧报告',
        ready: '就绪',
        missingValue: '缺失',
        unknown: '未知'
      },
      'zh-Hant': {
        appSubtitle: '拖入 PowerPoint、Keynote 或 KeyMorph IR 檔案。本機完成轉換，並產生執行時預覽、可編輯匯出、影片匯出、IR 和報告。',
        language: '語言',
        dropTitle: '把簡報拖到這裡',
        dropHint: '或點擊選擇 .pptx / .key / .ir.json',
        chooseFile: '選擇檔案',
        loadSample: '載入範例',
        waiting: '等待簡報。',
        preview: '預覽',
        openPreview: '開啟預覽',
        emptyPreview: '拖入簡報後產生可互動 HTML 執行時。',
        converting: '正在轉換 {name}...',
        convertingDeck: '正在本機轉換。預覽和匯出動作會在完成後出現。',
        unsupportedFile: '不支援的檔案。請使用 .pptx、.key 或 .ir.json。',
        conversionFailed: '轉換失敗。',
        converted: '已轉換 {name}。',
        slides: '投影片',
        objects: '物件',
        conversionFidelity: '轉換保真度',
        runtimeMode: '執行時模式',
        lostAnimations: '遺失動畫',
        degradedAnimations: '降級動畫',
        uncertainMappings: '不確定映射',
        videoFrames: '影片影格',
        frameCapture: '影格擷取',
        mp4Encoder: 'MP4 編碼器',
        recommendedFixes: '建議修復',
        downloadsTitle: '下載',
        runtimeIr: 'KeyMorph IR 重建',
        runtimeKeynoteHtml: 'Keynote 原生 HTML',
        runtimeKeynoteMovie: 'Keynote 渲染影片',
        downloadHtml: '下載 HTML 執行時',
        downloadPptx: '下載 PPTX',
        exportKeynote: '匯出 Keynote',
        downloadIr: '下載 IR',
        downloadReport: '下載報告',
        renderMp4: '渲染 MP4',
        downloadVideoStatus: '下載影片狀態',
        openRenderedFrames: '開啟渲染影格',
        openFrameDiffs: '開啟影格差異',
        renderBaseline: '執行 Keynote 基準對比',
        baselineHint: '將 Keynote 官方參考影格與 HTML 執行時逐幀對比。macOS 可能會要求 Keynote 自動化權限。',
        baselineRequiresKeynote: 'Keynote 基準對比需要原始 .key 來源檔。',
        baselineRunning: '正在渲染 Keynote 官方基準影格，並與 HTML 執行時逐幀對比...',
        baselineFailed: 'Keynote 基準對比失敗。',
        baselineReady: '基準保真度已產生。',
        baselineUnavailable: 'Keynote 基準對比不可用。',
        downloadBaselineReport: '下載基準報告',
        downloadBaselineMovie: '下載參考影片',
        openBaselineDiffs: '開啟基準差異影格',
        openReferenceFrames: '開啟參考影格',
        openRuntimeFrames: '開啟執行時影格',
        baselineSource: '來源：{source}。',
        baselineFidelityMean: '基準平均分：{score}。',
        worstFramesTitle: '最差影格',
        worstSlidesTitle: '最差頁',
        worstFrameLabel: '第 {frame} 幀 · 第 {slide} 頁 · 分數 {score}',
        worstSlideLabel: '第 {slide} 頁 · 平均 {score} · 最差幀 {frame}',
        openDiff: '差異圖',
        openReference: '參考影格',
        openRuntime: '執行時影格',
        keynoteExporting: '正在透過本機 Keynote 自動化匯出...',
        keynoteUnavailable: 'Keynote 匯出不可用。',
        keynoteReady: 'Keynote 已就緒。',
        downloadKeynote: '下載 Keynote',
        renderingVideo: '正在以 4x 解析度渲染 {frames} 影格...',
        videoReadyToRender: '可渲染 {frames} 影格，{fps} fps（{width}x{height}）。',
        videoMissingDependencies: 'MP4 渲染不可用。缺少：{items}。',
        videoFailed: '影片匯出失敗。',
        missing: '缺少：{items}。',
        mp4Ready: 'MP4 已就緒。',
        downloadVideo: '下載影片',
        frameReportPending: '第二次渲染同一個任務後會產生逐幀保真度報告。',
        frameFidelityMean: '逐幀保真度平均分：{score}。',
        downloadFrameReport: '下載逐幀報告',
        ready: '就緒',
        missingValue: '缺失',
        unknown: '未知'
      }
    };

    const drop = document.getElementById('drop');
    const fileInput = document.getElementById('file');
    const languageSelect = document.getElementById('language');
    const statusEl = document.getElementById('status');
    const metricsEl = document.getElementById('metrics');
    const recommendationsEl = document.getElementById('recommendations');
    const downloadsEl = document.getElementById('downloads');
    const previewPane = document.getElementById('previewPane');
    const previewTitle = document.getElementById('previewTitle');
    const openPreview = document.getElementById('openPreview');
    const loadDemo = document.getElementById('loadDemo');
    let currentLang = resolveInitialLanguage();
    let lastResult = null;

    languageSelect.value = currentLang;
    translateStatic();
    languageSelect.addEventListener('change', () => {
      currentLang = languageSelect.value;
      localStorage.setItem('keymorph-language', currentLang);
      translateStatic();
      if (lastResult) {
        renderResult(lastResult);
      } else {
        statusEl.textContent = t('waiting');
        previewTitle.textContent = t('preview');
        previewPane.textContent = t('emptyPreview');
      }
    });

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
      if (!isSupportedFile(file.name)) {
        lastResult = null;
        statusEl.className = 'status error';
        statusEl.textContent = t('unsupportedFile');
        metricsEl.hidden = true;
        recommendationsEl.hidden = true;
        downloadsEl.hidden = true;
        previewPane.className = 'empty';
        previewPane.textContent = t('emptyPreview');
        return;
      }
      lastResult = null;
      statusEl.textContent = t('converting', { name: file.name });
      statusEl.className = 'status';
      metricsEl.hidden = true;
      recommendationsEl.hidden = true;
      downloadsEl.hidden = true;
      previewPane.className = 'empty';
      previewPane.textContent = t('convertingDeck');
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
        previewPane.textContent = t('conversionFailed');
      }
    }

    function renderResult(result) {
      lastResult = result;
      statusEl.textContent = t('converted', { name: result.sourceName });
      statusEl.className = 'status';
      metricsEl.hidden = false;
      metricsEl.innerHTML = [
        [t('slides'), result.slideCount],
        [t('objects'), result.objectCount],
        [t('conversionFidelity'), result.fidelityScore],
        [t('runtimeMode'), runtimeLabel(result.runtimeMode || result.runtime?.mode)],
        [t('lostAnimations'), result.animationLostCount],
        [t('degradedAnimations'), result.degradedAnimationCount],
        [t('uncertainMappings'), result.uncertainMappingCount],
        [t('videoFrames'), result.videoPlan.totalFrames + ' @ ' + result.videoPlan.fps + ' fps'],
        [t('frameCapture'), dependencyLabel(result.videoDependencies, ['playwright', 'browser'])],
        [t('mp4Encoder'), dependencyLabel(result.videoDependencies, ['ffmpeg'])]
      ].map(([label, value]) => '<div class="metric"><span>' + label + '</span><strong>' + value + '</strong></div>').join('');
      if (result.recommendedFixes && result.recommendedFixes.length) {
        recommendationsEl.hidden = false;
        recommendationsEl.innerHTML = '<strong>' + t('recommendedFixes') + '</strong><ul class="fixes">' + result.recommendedFixes.map((fix) => '<li>' + escapeHtml(fix) + '</li>').join('') + '</ul>';
      } else {
        recommendationsEl.hidden = true;
      }
      const baselineCanRun = result.baselineCanRun ?? result.sourceKind === 'keynote';
      const baselineMessage = result.baselineMessage || (baselineCanRun ? t('baselineHint') : t('baselineRequiresKeynote'));
      const baselineStatusClass = baselineCanRun ? 'status' : 'status warn';
      const baselineStatusLinks = result.downloads.baselineFidelity
        ? '<div class="link-row">' +
          downloadLink(result.downloads.baselineFidelity, t('downloadBaselineReport')) +
          downloadLink(result.downloads.baselineDiffs, t('openBaselineDiffs'), true) +
          '</div>'
        : '';
      const videoMessage = result.videoDependencies?.canExportVideo
        ? t('videoReadyToRender', {
            frames: result.videoPlan.totalFrames,
            fps: result.videoPlan.fps,
            width: result.videoPlan.outputWidth,
            height: result.videoPlan.outputHeight
          })
        : t('videoMissingDependencies', { items: result.videoDependencies?.missing?.join(', ') || t('unknown') });
      downloadsEl.hidden = false;
      downloadsEl.innerHTML =
        '<div class="action-group">' +
          '<div class="action-title"><span>' + t('downloadsTitle') + '</span></div>' +
          '<div class="link-row">' +
            downloadLink(result.downloads.html, t('downloadHtml'), true) +
            downloadLink(result.downloads.pptx, t('downloadPptx')) +
            downloadLink(result.downloads.ir, t('downloadIr')) +
            downloadLink(result.downloads.report, t('downloadReport')) +
          '</div>' +
        '</div>' +
        '<div class="action-group">' +
          '<div class="action-title"><span>' + t('renderMp4') + '</span><button id="renderVideo" class="primary" type="button">' + t('renderMp4') + '</button></div>' +
          '<div id="videoStatus" class="' + (result.videoDependencies?.canExportVideo ? 'status' : 'status warn') + '">' + escapeHtml(videoMessage) + '</div>' +
        '</div>' +
        '<div class="action-group">' +
          '<div class="action-title"><span>' + t('exportKeynote') + '</span><button id="exportKeynote" type="button">' + t('exportKeynote') + '</button></div>' +
          '<div id="keynoteStatus" class="status warn">' + escapeHtml(result.keynoteMessage || '') + '</div>' +
        '</div>' +
        '<div class="action-group">' +
          '<div class="action-title"><span>' + t('renderBaseline') + '</span><button id="renderBaseline" type="button"' + (baselineCanRun ? '' : ' disabled') + '>' + t('renderBaseline') + '</button></div>' +
          '<div id="baselineStatus" class="' + baselineStatusClass + '">' + escapeHtml(baselineMessage) + baselineStatusLinks + '</div>' +
        '</div>';
      document.getElementById('exportKeynote')?.addEventListener('click', () => exportKeynote(result));
      document.getElementById('renderVideo')?.addEventListener('click', () => renderVideo(result));
      document.getElementById('renderBaseline')?.addEventListener('click', () => renderBaseline(result));
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
      keynoteStatus.textContent = t('keynoteExporting');
      try {
        const response = await fetch(result.keynoteEndpoint + '?allowKeynote=1', { method: 'POST' });
        const payload = await response.json();
        if (!response.ok) {
          keynoteStatus.className = 'status warn';
          keynoteStatus.textContent = (payload.message || t('keynoteUnavailable')) + guidanceText(payload.guidance);
          return;
        }
        keynoteStatus.className = 'status';
        keynoteStatus.innerHTML = t('keynoteReady') + ' <a class="download" href="' + payload.keyUrl + '">' + t('downloadKeynote') + '</a>';
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
      videoStatus.className = 'status';
      videoStatus.textContent = t('renderingVideo', { frames: result.videoPlan.totalFrames });
      try {
        const response = await fetch(result.videoEndpoint, { method: 'POST' });
        const payload = await response.json();
        if (!response.ok) {
          videoStatus.className = response.status === 424 ? 'status warn' : 'status error';
          videoStatus.textContent = (payload.message || t('videoFailed')) + guidanceText(payload.guidance);
          if (payload.missing?.length) {
            videoStatus.textContent += ' ' + t('missing', { items: payload.missing.join(', ') });
          }
          return;
        }
        videoStatus.className = 'status';
        videoStatus.innerHTML =
          t('mp4Ready') +
          '<div class="link-row">' +
            downloadLink(payload.videoUrl, t('downloadVideo')) +
            downloadLink(result.downloads.videoStatus, t('downloadVideoStatus')) +
            downloadLink(payload.framesUrl, t('openRenderedFrames'), true) +
            downloadLink(payload.frameFidelity?.diffUrl, t('openFrameDiffs'), true) +
          '</div>' +
          fidelityLinks(payload.frameFidelity);
      } catch (error) {
        videoStatus.className = 'status error';
        videoStatus.textContent = error instanceof Error ? error.message : String(error);
      } finally {
        button.disabled = false;
      }
    }

    async function renderBaseline(result) {
      const button = document.getElementById('renderBaseline');
      const baselineStatus = document.getElementById('baselineStatus');
      if (!(result.baselineCanRun ?? result.sourceKind === 'keynote')) {
        baselineStatus.className = 'status warn';
        baselineStatus.textContent = result.baselineMessage || t('baselineRequiresKeynote');
        return;
      }
      button.disabled = true;
      baselineStatus.className = 'status';
      baselineStatus.textContent = t('baselineRunning');
      try {
        const response = await fetch(result.baselineEndpoint + '?allowKeynote=1', { method: 'POST' });
        const payload = await response.json();
        if (!response.ok) {
          baselineStatus.className = response.status === 424 || response.status === 400 ? 'status warn' : 'status error';
          baselineStatus.textContent = (payload.message || t('baselineFailed')) + guidanceText(payload.guidance);
          if (payload.missing?.length) {
            baselineStatus.textContent += ' ' + t('missing', { items: payload.missing.join(', ') });
          }
          return;
        }
        baselineStatus.className = 'status';
        baselineStatus.innerHTML =
          t('baselineReady') + ' ' +
          (payload.source ? t('baselineSource', { source: payload.source }) + ' ' : '') +
          t('baselineFidelityMean', { score: payload.summary?.meanPixelFidelityScore ?? 'n/a' }) +
          '<div class="link-row">' +
            downloadLink(payload.reportUrl, t('downloadBaselineReport')) +
            downloadLink(payload.diffUrl, t('openBaselineDiffs'), true) +
            downloadLink(payload.baselineMovieUrl, t('downloadBaselineMovie')) +
            downloadLink(payload.baselineFramesUrl, t('openReferenceFrames'), true) +
            downloadLink(payload.actualFramesUrl, t('openRuntimeFrames'), true) +
          '</div>' +
          fidelityInsightsHtml(payload.insights);
      } catch (error) {
        baselineStatus.className = 'status error';
        baselineStatus.textContent = error instanceof Error ? error.message : String(error);
      } finally {
        button.disabled = false;
      }
    }

    function escapeHtml(value) {
      return String(value ?? '').replace(/[&<>"']/g, (char) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[char]));
    }
    function isSupportedFile(fileName) {
      return /\.(pptx|key|ir\.json|json)$/i.test(fileName);
    }
    function resolveInitialLanguage() {
      const stored = localStorage.getItem('keymorph-language');
      if (stored && translations[stored]) return stored;
      const browserLangs = navigator.languages?.length ? navigator.languages : [navigator.language || 'en'];
      const match = browserLangs.find((lang) => /^zh/i.test(lang));
      if (/zh-(tw|hk|mo|hant)/i.test(match || '')) return 'zh-Hant';
      if (match) return 'zh-Hans';
      return 'en';
    }
    function translateStatic() {
      document.documentElement.lang = currentLang;
      document.querySelectorAll('[data-i18n]').forEach((node) => {
        node.textContent = t(node.dataset.i18n);
      });
    }
    function t(key, replacements = {}) {
      const value = translations[currentLang]?.[key] ?? translations.en[key] ?? key;
      return Object.entries(replacements).reduce((text, [name, replacement]) => {
        return text.replaceAll('{' + name + '}', String(replacement));
      }, value);
    }
    function dependencyLabel(status, keys) {
      if (!status?.available) return t('unknown');
      const ready = keys.every((key) => status.available[key]);
      return ready ? t('ready') : t('missingValue');
    }
    function runtimeLabel(mode) {
      if (mode === 'keynote-html') return t('runtimeKeynoteHtml');
      if (mode === 'keynote-movie') return t('runtimeKeynoteMovie');
      return t('runtimeIr');
    }
    function downloadLink(href, label, newTab = false) {
      if (!href) return '';
      const target = newTab ? ' target="_blank"' : '';
      return '<a class="download" href="' + href + '"' + target + '>' + label + '</a>';
    }
    function guidanceText(guidance) {
      return Array.isArray(guidance) && guidance.length ? ' ' + guidance.join(' ') : '';
    }
    function fidelityInsightsHtml(insights) {
      const frames = insights?.worstFrames || [];
      const slides = insights?.worstSlides || [];
      if (!frames.length && !slides.length) return '';
      const frameItems = frames.slice(0, 5).map((frame) => {
        return '<li>' +
          escapeHtml(t('worstFrameLabel', { frame: frame.frame, slide: Number(frame.slideIndex || 0) + 1, score: frame.pixelFidelityScore })) +
          '<div class="link-row">' +
            downloadLink(frame.diffUrl, t('openDiff'), true) +
            downloadLink(frame.referenceUrl, t('openReference'), true) +
            downloadLink(frame.actualUrl, t('openRuntime'), true) +
          '</div>' +
        '</li>';
      }).join('');
      const slideItems = slides.slice(0, 5).map((slide) => {
        return '<li>' + escapeHtml(t('worstSlideLabel', {
          slide: Number(slide.slideIndex || 0) + 1,
          score: slide.meanPixelFidelityScore,
          frame: slide.worstFrame ?? 'n/a'
        })) + '</li>';
      }).join('');
      return '<div class="sub"><strong>' + t('worstFramesTitle') + '</strong><ul class="fixes">' + frameItems + '</ul><strong>' + t('worstSlidesTitle') + '</strong><ul class="fixes">' + slideItems + '</ul></div>';
    }
    function fidelityLinks(frameFidelity) {
      if (!frameFidelity) return '<div class="sub">' + t('frameReportPending') + '</div>';
      const score = frameFidelity.summary?.meanPixelFidelityScore ?? 'n/a';
      return '<div class="sub">' + t('frameFidelityMean', { score }) + ' ' + downloadLink(frameFidelity.reportUrl, t('downloadFrameReport')) + '</div>';
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

function send(response, status, body, type) {
  response.writeHead(status, { "Content-Type": type });
  response.end(body);
}

function contentType(filePath) {
  if (filePath.endsWith(".html")) return "text/html; charset=utf-8";
  if (filePath.endsWith(".json")) return "application/json; charset=utf-8";
  if (filePath.endsWith(".png")) return "image/png";
  if (filePath.endsWith(".pptx")) {
    return "application/vnd.openxmlformats-officedocument.presentationml.presentation";
  }
  if (filePath.endsWith(".key")) return "application/vnd.apple.keynote";
  if (filePath.endsWith(".mp4")) return "video/mp4";
  if (filePath.endsWith(".m4v")) return "video/mp4";
  return "application/octet-stream";
}
