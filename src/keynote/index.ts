import type { DeckIR } from "../ir/index.ts";
import { parsePptxToIr } from "../pptx/index.ts";

export interface KeynoteImportOptions {
  exportedPptxPath?: string;
}

export async function parseKeynoteToIr(
  keynotePath: string,
  options: KeynoteImportOptions = {}
): Promise<DeckIR> {
  if (!options.exportedPptxPath) {
    throw new Error(
      `Keynote import for ${keynotePath} requires a PPTX exported by Keynote in this local-first MVP.`
    );
  }

  const deck = await parsePptxToIr(options.exportedPptxPath);
  deck.conversion ??= {
    status: "partial",
    messages: []
  };
  deck.conversion.uncertainMappings ??= [];
  deck.conversion.uncertainMappings.push({
    code: "keynote-pptx-bridge",
    description: "Imported through Keynote's PPTX export behavior; native Keynote animation mappings are uncertain.",
    severity: "warning",
    confidence: 0.6
  });
  return deck;
}

export interface KeynoteExportResult {
  pptxPath: string;
  unsupportedCount: number;
}

export async function exportIrToKeynoteBridge(
  deck: DeckIR,
  pptxPath: string,
  exportPptx: (deck: DeckIR, pptxPath: string) => Promise<void>
): Promise<KeynoteExportResult> {
  await exportPptx(deck, pptxPath);
  return {
    pptxPath,
    unsupportedCount:
      (deck.conversion?.unsupportedFeatures?.length ?? 0) + (deck.conversion?.degradedFeatures?.length ?? 0)
  };
}
