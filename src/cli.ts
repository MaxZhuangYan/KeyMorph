import { mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";

import { createDemoDeck } from "./demo/createDemoDeck.ts";
import { renderHtmlDocument } from "./runtime/index.ts";
import { createLossReport, scoreConversion } from "./report/index.ts";
import { exportIrToPptx, parsePptxToIr } from "./pptx/index.ts";
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

  const deck = createDemoDeck();
  const irPath = path.join(outDir, "demo.ir.json");
  const htmlPath = path.join(outDir, "runtime.html");
  const pptxPath = path.join(outDir, "rebuilt.pptx");
  const reportPath = path.join(outDir, "conversion-report.json");

  await writeJson(irPath, deck);
  await writeFile(htmlPath, renderHtmlDocument(deck), "utf8");
  await exportIrToPptx(deck, pptxPath);
  await writeJson(reportPath, scoreConversion(deck.conversion ?? { status: "success", messages: [] }));

  console.log("Demo generated:");
  console.log(`  IR: ${irPath}`);
  console.log(`  HTML runtime: ${pathToFileURL(htmlPath).toString()}`);
  console.log(`  PPTX: ${pptxPath}`);
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
        "Usage: keymorph <demo|pptx-to-ir|ir-to-html|ir-to-pptx|ir-to-video|ir-report> [input] [output]"
      );
  }
}

main().catch((error: unknown) => {
  console.error(error instanceof Error ? error.message : error);
  process.exitCode = 1;
});
