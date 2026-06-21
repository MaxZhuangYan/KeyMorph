# KeyMorph

KeyMorph is a local-first prototype for bidirectional presentation conversion built around a single animation intermediate representation.

The current MVP target is:

```text
PPTX -> IR -> HTML runtime -> IR -> PPTX
```

The system optimizes for visual fidelity and records any unsupported animation or layout mapping in a conversion report.

## Architecture

- `src/ir`: shared JSON IR, validation, loss report primitives.
- `src/pptx`: PPTX parsing and PPTX export.
- `src/keynote`: Keynote bridge through PPTX export/import behavior.
- `src/runtime`: browser HTML renderer and timeline player.
- `src/video`: Playwright frame capture and ffmpeg encoding pipeline.
- `src/report`: fidelity scoring and recommendations.
- `demo`: local sample IR and generated demo artifacts.

## Commands

```bash
npm run build
npm test
npm run dev
npm run demo
npm run convert -- deck.pptx demo/out/jobs/manual
npm run inspect -- demo/out/source.ir.json
npm run benchmark:key -- deck.key demo/out/benchmarks/deck --allow-keynote --fps 30 --scale 4
npm run bundle:key -- demo/out/jobs/manual
npm run bundle:video -- demo/out/jobs/manual
npm run bundle:baseline -- demo/out/jobs/manual --allow-keynote
npm run keyhtml:to-ir -- keynote-html-export deck.ir.json
npm run ir:report -- demo/out/imported.ir.json demo/out/conversion-report.json
npm run png:fidelity -- reference.png actual.png fidelity-report.json
```

`npm run demo` generates a local round-trip under `demo/out`: source IR, original PPTX, imported IR, HTML runtime, rebuilt PPTX, and conversion report.

`npm run convert -- <input.pptx|input.key|input.ir.json> <output-dir>` is the product bundle path used by the CLI and local API. It writes:

- the source file or package copy
- `deck.ir.json`
- `runtime.html`
- `rebuilt.pptx`
- `loss-report.json`
- `video-plan.json`
- `video-status.json`
- `baseline-status.json`
- `manifest.json`

`npm run inspect -- <input>` parses the input and prints validation, conversion-risk, and video dependency status without writing a bundle. `npm run bundle:key -- <output-dir>`, `npm run bundle:video -- <output-dir>`, and `npm run bundle:baseline -- <output-dir>` run the deferred Keynote, MP4, and Keynote golden-baseline exports for an existing bundle. Keynote automation is disabled by default; pass `--allow-keynote` or set `KEYMORPH_ALLOW_KEYNOTE_AUTOMATION=1` when you intentionally want AppleScript automation.

`npm run benchmark:key -- <input.key> <output-dir> --allow-keynote --fps 30 --scale 4` copies the original `.key` into `<output-dir>/source-copy`, builds the product bundle from that copy, then writes `benchmark-summary.json`. With `--allow-keynote`, it also runs the Keynote golden baseline and lists the lowest-scoring frames and slides so fidelity work can start from concrete PNG diffs. Without `--allow-keynote`, it still creates the copied-source bundle and records that the baseline has not run.

`npm run bundle:baseline -- <bundle-dir> --allow-keynote --fps 30 --scale 4` is for original `.key` bundles. It uses the copied `.key` inside the bundle, exports a high-resolution Keynote reference movie, extracts `frames/baseline`, captures the KeyMorph HTML runtime into `frames/keymorph-baseline`, and writes `baseline-fidelity.json` plus `baseline-diffs`.

`npm run dev` starts the local product UI for drag-and-drop conversion. Drop a `.pptx`, `.key`, or `.ir.json`, and it returns an HTML runtime preview plus downloadable HTML, PPTX, Keynote when available, IR, loss report, and an MP4 render action. The server starts at `http://127.0.0.1:4173/` or the next available port.
The `POST /api/convert` route calls the same bundle workflow as `npm run convert`, and `/api/jobs/:id/keynote`, `/api/jobs/:id/video`, plus `/api/jobs/:id/baseline` perform the optional on-demand exports against that bundle.
On macOS with Keynote installed, `.key` input and `.key` export are handled through local Keynote automation as a PPTX bridge. If the bridge is unavailable, KeyMorph can inspect native directory-style or ZIP-backed `.key` packages and recover approximate text slides with explicit loss-report warnings. macOS may prompt for automation permission the first time.

## Current conversion coverage

- PPTX animation XML: imports and exports fade effects, visibility set events, and simple motion paths as IR timeline events. Unsupported timing nodes are recorded in the report.
- HTML runtime: plays a deck-level timeline with global scrubbing, slide stepping, basic keyframes, visibility/property events, and best-effort transitions/morph-style numeric interpolation.
- Video export: renders the HTML runtime through Playwright and encodes with `ffmpeg` when both are available. Missing local dependencies are reported in the UI instead of failing silently.
- Pixel fidelity: compares rendered PNG frames against reference frames and reports mismatch ratio, error metrics, and a pixel fidelity score.
- Keynote golden baseline: for original `.key` bundles, compares Keynote-rendered reference frames against KeyMorph HTML runtime frames and writes a per-frame fidelity report plus diff PNGs.

This checkpoint intentionally has no external npm runtime dependencies. It uses Node 24's built-in TypeScript transform.

## Non-goals

- Full Office or Keynote spec compatibility.
- Cloud service execution.
- Treating HTML as a conversion interchange format.

IR is the source of truth; HTML is only a runtime renderer.
