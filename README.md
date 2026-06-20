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
npm run ir:report -- demo/out/imported.ir.json demo/out/conversion-report.json
```

`npm run demo` generates a local round-trip under `demo/out`: source IR, original PPTX, imported IR, HTML runtime, rebuilt PPTX, and conversion report.
`npm run dev` starts the local product UI for drag-and-drop conversion. Drop a `.pptx`, `.key`, or `.ir.json`, and it returns an HTML runtime preview plus downloadable HTML, PPTX, Keynote when available, IR, loss report, and an MP4 render action. The server starts at `http://127.0.0.1:4173/` or the next available port.
On macOS with Keynote installed, `.key` input and `.key` export are handled through local Keynote automation as a PPTX bridge. If the bridge is unavailable, KeyMorph can inspect native directory-style or ZIP-backed `.key` packages and recover approximate text slides with explicit loss-report warnings. macOS may prompt for automation permission the first time.

## Current conversion coverage

- PPTX animation XML: imports and exports fade effects, visibility set events, and simple motion paths as IR timeline events. Unsupported timing nodes are recorded in the report.
- HTML runtime: plays a deck-level timeline with global scrubbing, slide stepping, basic keyframes, visibility/property events, and best-effort transitions/morph-style numeric interpolation.
- Video export: renders the HTML runtime through Playwright and encodes with `ffmpeg` when both are available. Missing local dependencies are reported in the UI instead of failing silently.

This checkpoint intentionally has no external npm runtime dependencies. It uses Node 24's built-in TypeScript transform.

## Non-goals

- Full Office or Keynote spec compatibility.
- Cloud service execution.
- Treating HTML as a conversion interchange format.

IR is the source of truth; HTML is only a runtime renderer.
