import { readFile, writeFile } from "node:fs/promises";
import { deflateRawSync } from "node:zlib";

import { IR_VERSION, type ConversionReport, type DeckIR, type IRObject, type Slide } from "../ir/index.ts";

const EMU_PER_INCH = 914400;
const WIDE_LAYOUT = { width: 13.333333, height: 7.5 };

export async function parsePptxToIr(filePath: string): Promise<DeckIR> {
  const data = await readFile(filePath);
  const report: ConversionReport = {
    source: { kind: "pptx", uri: filePath },
    status: "partial",
    generatedAt: new Date().toISOString(),
    tool: "keymorph-pptx-parser-mvp",
    messages: [
      {
        severity: "warning",
        code: "pptx-import-mvp",
        message:
          "This checkpoint verifies that the PPTX package is readable but does not yet extract full PresentationML."
      }
    ],
    unsupportedFeatures: [
      {
        code: "presentationml-static-extraction",
        severity: "warning",
        area: "layout",
        description: "PPTX shape, text, image, and animation XML extraction is deferred to the next checkpoint.",
        fallback: "Create a placeholder IR slide with a loss report."
      }
    ],
    degradedFeatures: [],
    uncertainMappings: [],
    statistics: {
      slideCount: 1,
      objectCount: 1,
      animationCount: 0,
      unsupportedFeatureCount: 1
    },
    metadata: {
      byteLength: data.byteLength
    }
  };

  return {
    irVersion: IR_VERSION,
    metadata: { title: "Imported PPTX", sourceApplication: "PowerPoint" },
    deck: {
      id: "pptx-import",
      title: "Imported PPTX",
      size: { width: 1280, height: 720, unit: "px" },
      slides: [
        {
          id: "slide-1",
          index: 0,
          name: "Import placeholder",
          background: { type: "solid", color: "#ffffff" },
          objects: [
            {
              id: "import-placeholder",
              type: "text",
              name: "Import placeholder",
              bounds: { x: 96, y: 96, width: 960, height: 96 },
              opacity: 1,
              text: {
                plainText: "PPTX import placeholder",
                runs: [
                  {
                    text: "PPTX import placeholder",
                    style: { fontFamily: "Arial", fontSize: 36, color: "#111827" }
                  }
                ]
              }
            }
          ],
          timeline: { durationMs: 2500, events: [], dependencyGraph: { edges: [] } }
        }
      ]
    },
    conversion: report
  };
}

export async function exportIrToPptx(deck: DeckIR, outputPath: string): Promise<void> {
  const files = createPptxFiles(deck);
  await writeFile(outputPath, createZip(files));
}

function createPptxFiles(deck: DeckIR): Map<string, string | Uint8Array> {
  const files = new Map<string, string | Uint8Array>();
  const slideCount = deck.deck.slides.length;

  files.set("[Content_Types].xml", contentTypes(slideCount));
  files.set("_rels/.rels", rootRels());
  files.set("docProps/core.xml", coreProps(deck));
  files.set("docProps/app.xml", appProps(slideCount));
  files.set("ppt/presentation.xml", presentationXml(deck));
  files.set("ppt/_rels/presentation.xml.rels", presentationRels(slideCount));
  files.set("ppt/theme/theme1.xml", themeXml());
  files.set("ppt/slideMasters/slideMaster1.xml", slideMasterXml());
  files.set("ppt/slideMasters/_rels/slideMaster1.xml.rels", slideMasterRels());
  files.set("ppt/slideLayouts/slideLayout1.xml", slideLayoutXml());
  files.set("ppt/slideLayouts/_rels/slideLayout1.xml.rels", slideLayoutRels());

  deck.deck.slides.forEach((slide, index) => {
    const slideNumber = index + 1;
    files.set(`ppt/slides/slide${slideNumber}.xml`, slideXml(deck, slide));
    files.set(`ppt/slides/_rels/slide${slideNumber}.xml.rels`, slideRels());
  });

  return files;
}

function presentationXml(deck: DeckIR): string {
  const ids = deck.deck.slides
    .map((_, index) => `<p:sldId id="${256 + index}" r:id="rId${index + 2}"/>`)
    .join("");
  return xml(`\
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>
  <p:sldIdLst>${ids}</p:sldIdLst>
  <p:sldSz cx="${inchesToEmu(WIDE_LAYOUT.width)}" cy="${inchesToEmu(WIDE_LAYOUT.height)}" type="wide"/>
  <p:notesSz cx="${inchesToEmu(10)}" cy="${inchesToEmu(7.5)}"/>
</p:presentation>`);
}

function presentationRels(slideCount: number): string {
  const slideRelsXml = Array.from({ length: slideCount }, (_, index) => {
    const id = index + 2;
    return `<Relationship Id="rId${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${index + 1}.xml"/>`;
  }).join("");
  return xml(`\
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  ${slideRelsXml}
  <Relationship Id="rId${slideCount + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>`);
}

function slideXml(deck: DeckIR, slide: Slide): string {
  const bg = solidFill(slide.background) ?? "#ffffff";
  const shapes = slide.objects.map((object, index) => objectToShape(deck, object, index + 2)).join("");
  return xml(`\
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:bg><p:bgPr><a:solidFill><a:srgbClr val="${stripHash(bg)}"/></a:solidFill></p:bgPr></p:bg>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x="0" y="0"/><a:ext cx="0" cy="0"/>
          <a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/>
        </a:xfrm>
      </p:grpSpPr>
      ${shapes}
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>`);
}

function objectToShape(deck: DeckIR, object: IRObject, id: number): string {
  if (object.type === "group") {
    return object.children.map((child, index) => objectToShape(deck, child, id + index)).join("");
  }

  const bounds = object.bounds ?? { x: 0, y: 0, width: 100, height: 100 };
  const x = pxToEmu(bounds.x, deck.deck.size.width, WIDE_LAYOUT.width);
  const y = pxToEmu(bounds.y, deck.deck.size.height, WIDE_LAYOUT.height);
  const cx = pxToEmu(bounds.width, deck.deck.size.width, WIDE_LAYOUT.width);
  const cy = pxToEmu(bounds.height, deck.deck.size.height, WIDE_LAYOUT.height);
  const name = escapeXml(object.name ?? object.id);

  if (object.type === "text") {
    const text = escapeXml(object.text.plainText ?? object.text.runs?.map((run) => run.text).join("") ?? "");
    const style = object.text.runs?.[0]?.style ?? object.style?.textStyle ?? {};
    return `\
<p:sp>
  <p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:noFill/>
  </p:spPr>
  <p:txBody>
    <a:bodyPr wrap="square"/>
    <a:lstStyle/>
    <a:p><a:r><a:rPr lang="en-US" sz="${fontSizeToOpenXml(style.fontSize ?? 28)}"><a:solidFill><a:srgbClr val="${stripHash(colorToString(style.color) ?? "#111827")}"/></a:solidFill></a:rPr><a:t>${text}</a:t></a:r><a:endParaRPr lang="en-US"/></a:p>
  </p:txBody>
</p:sp>`;
  }

  const fill = object.type === "shape" ? solidFill(object.style?.fill) ?? "#e2e8f0" : "#e2e8f0";
  const shape = object.type === "shape" ? shapePreset(object.shape) : "rect";
  return `\
<p:sp>
  <p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>
    <a:prstGeom prst="${shape}"><a:avLst/></a:prstGeom>
    <a:solidFill><a:srgbClr val="${stripHash(fill)}"/></a:solidFill>
    <a:ln><a:noFill/></a:ln>
  </p:spPr>
</p:sp>`;
}

function contentTypes(slideCount: number): string {
  const slideOverrides = Array.from({ length: slideCount }, (_, index) => {
    return `<Override PartName="/ppt/slides/slide${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`;
  }).join("");
  return xml(`\
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
  ${slideOverrides}
</Types>`);
}

function rootRels(): string {
  return xml(`\
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`);
}

function slideRels(): string {
  return xml(`\
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>`);
}

function slideMasterRels(): string {
  return xml(`\
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>`);
}

function slideLayoutRels(): string {
  return xml(`\
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>`);
}

function slideMasterXml(): string {
  return xml(`\
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>
  <p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles>
</p:sldMaster>`);
}

function slideLayoutXml(): string {
  return xml(`\
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank" preserve="1">
  <p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>`);
}

function themeXml(): string {
  return xml(`\
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="KeyMorph">
  <a:themeElements>
    <a:clrScheme name="KeyMorph">
      <a:dk1><a:srgbClr val="111827"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="1F2937"/></a:dk2><a:lt2><a:srgbClr val="F8FAFC"/></a:lt2>
      <a:accent1><a:srgbClr val="0F766E"/></a:accent1><a:accent2><a:srgbClr val="2563EB"/></a:accent2><a:accent3><a:srgbClr val="DC2626"/></a:accent3>
      <a:accent4><a:srgbClr val="CA8A04"/></a:accent4><a:accent5><a:srgbClr val="7C3AED"/></a:accent5><a:accent6><a:srgbClr val="0891B2"/></a:accent6>
      <a:hlink><a:srgbClr val="2563EB"/></a:hlink><a:folHlink><a:srgbClr val="7C3AED"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="KeyMorph"><a:majorFont><a:latin typeface="Arial"/></a:majorFont><a:minorFont><a:latin typeface="Arial"/></a:minorFont></a:fontScheme>
    <a:fmtScheme name="KeyMorph"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme>
  </a:themeElements>
</a:theme>`);
}

function coreProps(deck: DeckIR): string {
  return xml(`\
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>${escapeXml(deck.deck.title ?? deck.metadata?.title ?? "KeyMorph deck")}</dc:title>
  <dc:creator>${escapeXml(deck.metadata?.author ?? "KeyMorph")}</dc:creator>
  <cp:lastModifiedBy>KeyMorph</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">2026-01-01T00:00:00Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2026-01-01T00:00:00Z</dcterms:modified>
</cp:coreProperties>`);
}

function appProps(slideCount: number): string {
  return xml(`\
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>KeyMorph</Application>
  <PresentationFormat>On-screen Show (16:9)</PresentationFormat>
  <Slides>${slideCount}</Slides>
</Properties>`);
}

function createZip(files: Map<string, string | Uint8Array>): Uint8Array {
  const chunks: Uint8Array[] = [];
  const centralDirectory: Uint8Array[] = [];
  let offset = 0;

  for (const [name, content] of files) {
    const nameBytes = textBytes(name);
    const data = typeof content === "string" ? textBytes(content) : content;
    const compressed = deflateRawSync(data);
    const crc = crc32(data);
    const local = concat([
      u32(0x04034b50),
      u16(20),
      u16(0),
      u16(8),
      u16(0),
      u16(0),
      u32(crc),
      u32(compressed.length),
      u32(data.length),
      u16(nameBytes.length),
      u16(0),
      nameBytes,
      compressed
    ]);
    chunks.push(local);
    centralDirectory.push(
      concat([
        u32(0x02014b50),
        u16(20),
        u16(20),
        u16(0),
        u16(8),
        u16(0),
        u16(0),
        u32(crc),
        u32(compressed.length),
        u32(data.length),
        u16(nameBytes.length),
        u16(0),
        u16(0),
        u16(0),
        u16(0),
        u32(0),
        u32(offset),
        nameBytes
      ])
    );
    offset += local.length;
  }

  const centralStart = offset;
  const central = concat(centralDirectory);
  const end = concat([
    u32(0x06054b50),
    u16(0),
    u16(0),
    u16(files.size),
    u16(files.size),
    u32(central.length),
    u32(centralStart),
    u16(0)
  ]);

  return concat([...chunks, central, end]);
}

function xml(body: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${body.replaceAll("\r\n", "\n")}`;
}

function inchesToEmu(inches: number): number {
  return Math.round(inches * EMU_PER_INCH);
}

function pxToEmu(px: number, deckPixels: number, inches: number): number {
  return Math.round((px / deckPixels) * inches * EMU_PER_INCH);
}

function fontSizeToOpenXml(px: number): number {
  return Math.round(px * 0.75 * 100);
}

function shapePreset(shape: string): string {
  if (shape === "ellipse") return "ellipse";
  if (shape === "roundRect") return "roundRect";
  if (shape === "triangle") return "triangle";
  return "rect";
}

function solidFill(fill: unknown): string | undefined {
  if (typeof fill === "string") return fill;
  if (fill && typeof fill === "object" && "type" in fill && (fill as { type: string }).type === "solid") {
    return colorToString((fill as { color?: unknown }).color);
  }
  return undefined;
}

function colorToString(color: unknown): string | undefined {
  if (typeof color === "string") return color;
  if (color && typeof color === "object" && "value" in color) return String((color as { value: unknown }).value);
  return undefined;
}

function stripHash(color: string): string {
  return color.replace(/^#/, "").slice(0, 6);
}

function escapeXml(value: unknown): string {
  return String(value ?? "").replace(/[<>&'"]/g, (char) => {
    switch (char) {
      case "<":
        return "&lt;";
      case ">":
        return "&gt;";
      case "&":
        return "&amp;";
      case "'":
        return "&apos;";
      default:
        return "&quot;";
    }
  });
}

function textBytes(value: string): Uint8Array {
  return new TextEncoder().encode(value);
}

function concat(chunks: Uint8Array[]): Uint8Array {
  const length = chunks.reduce((total, chunk) => total + chunk.length, 0);
  const out = new Uint8Array(length);
  let offset = 0;
  for (const chunk of chunks) {
    out.set(chunk, offset);
    offset += chunk.length;
  }
  return out;
}

function u16(value: number): Uint8Array {
  const out = new Uint8Array(2);
  new DataView(out.buffer).setUint16(0, value, true);
  return out;
}

function u32(value: number): Uint8Array {
  const out = new Uint8Array(4);
  new DataView(out.buffer).setUint32(0, value >>> 0, true);
  return out;
}

const CRC_TABLE = new Uint32Array(256).map((_, index) => {
  let value = index;
  for (let bit = 0; bit < 8; bit += 1) {
    value = value & 1 ? 0xedb88320 ^ (value >>> 1) : value >>> 1;
  }
  return value >>> 0;
});

function crc32(data: Uint8Array): number {
  let crc = 0xffffffff;
  for (const byte of data) {
    crc = CRC_TABLE[(crc ^ byte) & 0xff] ^ (crc >>> 8);
  }
  return (crc ^ 0xffffffff) >>> 0;
}
