import { mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";

import { createDemoDeck } from "./demo/createDemoDeck.ts";
import { renderHtmlDocument } from "./runtime/index.ts";
import { createLossReport, scoreConversion } from "./report/index.ts";
import { exportIrToPptx, parsePptxToIr } from "./pptx/index.ts";
import { exportIrToKeynote, parseKeynoteToIr } from "./keynote/index.ts";
import { exportIrToVideo } from "./video/index.ts";
import type { DeckIR } from "./ir/index.ts";

async function readJson<T>(filePath: string): Promise<T> {
  const source = await readFile(filePath, "utf8");
  return JSON.parse(source) as T;
}

async function writeJson(filePath: string, value: unknown): Promise<void> {
  await mkdir(path.dirname(filePath), { recursive: true });
  await writeFile(filePath, `${JSON.stringify(value, null, 2)}\n`);
}

async function runDemo(): Promise<void> {
  const outDir = path.resolve("demo/out");
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

  console.log("Demo generated:");
  console.log(`  source IR: ${sourceIrPath}`);
  console.log(`  original PPTX: ${originalPptxPath}`);
  console.log(`  imported IR: ${importedIrPath}`);
  console.log(`  HTML runtime: ${pathToFileURL(htmlPath).toString()}`);
  console.log(`  rebuilt PPTX: ${pptxPath}`);
  console.log(`  report: ${reportPath}`);
}

async function main(): Promise<void> {
  const [command, input, output] = process.argv.slice(2);

  switch (command) {
    case "demo":
      await runDemo();
      return;
    case "pptx-to-ir": {
      if (!input || !output) throw new Error("Usage: pptx-to-ir <input.pptx> <output.ir.json>");
      await writeJson(output, await parsePptxToIr(input));
      return;
    }
    case "key-to-ir": {
      if (!input || !output) throw new Error("Usage: key-to-ir <input.key> <output.ir.json>");
      await writeJson(output, await parseKeynoteToIr(input, { workDir: path.dirname(path.resolve(output)) }));
      return;
    }
    case "ir-to-html": {
      if (!input || !output) throw new Error("Usage: ir-to-html <input.ir.json> <output.html>");
      const deck = await readJson<DeckIR>(input);
      await mkdir(path.dirname(path.resolve(output)), { recursive: true });
      await writeFile(output, renderHtmlDocument(deck), "utf8");
      return;
    }
    case "ir-to-pptx": {
      if (!input || !output) throw new Error("Usage: ir-to-pptx <input.ir.json> <output.pptx>");
      await exportIrToPptx(await readJson<DeckIR>(input), output);
      return;
    }
    case "ir-to-key": {
      if (!input || !output) throw new Error("Usage: ir-to-key <input.ir.json> <output.key>");
      await exportIrToKeynote(await readJson<DeckIR>(input), output);
      return;
    }
    case "ir-to-video": {
      if (!input || !output) throw new Error("Usage: ir-to-video <input.ir.json> <output.mp4>");
      await exportIrToVideo(await readJson<DeckIR>(input), output);
      return;
    }
    case "ir-report": {
      if (!input || !output) throw new Error("Usage: ir-report <input.ir.json> <output.report.json>");
      const deck = await readJson<DeckIR>(input);
      await writeJson(output, createLossReport(deck.conversion ?? { status: "success", messages: [] }));
      return;
    }
    default:
      throw new Error(
        "Usage: keymorph <demo|pptx-to-ir|key-to-ir|ir-to-html|ir-to-pptx|ir-to-key|ir-to-video|ir-report> [input] [output]"
      );
  }
}

main().catch((error: unknown) => {
  console.error(error instanceof Error ? error.message : error);
  process.exitCode = 1;
});
