import { describe, test } from "node:test";
import assert from "node:assert/strict";
import { mkdir, mkdtemp, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { deflateRawSync } from "node:zlib";

import { detectNativeKeynotePackage, exportKeynoteToPptx, parseKeynoteToIr, parseNativeKeynoteToIr } from "../../src/keynote/index.ts";
import { validateIR } from "../../src/ir/index.ts";

describe("Keynote bridge", () => {
  test("fails with a clear local automation error for missing input or unavailable Keynote", async () => {
    await assert.rejects(
      () => exportKeynoteToPptx("/tmp/missing.key", "/tmp/missing.pptx"),
      /Keynote automation failed|Keynote conversion requires macOS/
    );
  });

  test("detects and parses a native directory-style .key package", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "keymorph-keynote-native-"));
    const keyPath = path.join(dir, "sample.key");
    await mkdir(path.join(keyPath, "Metadata"), { recursive: true });
    await writeFile(path.join(keyPath, "Metadata", "Properties.plist"), plist({ title: "Native Fixture", slideWidth: 1024, slideHeight: 768 }));
    await writeFile(
      path.join(keyPath, "Index.zip"),
      zip(
        new Map([
          ["Document.iwa", protoString("Document shell")],
          ["Slide-1.iwa", concat([protoString("Native Title"), protoString("First bullet")])],
          ["Slide-2.iwa", snappyFramed(protoString("Second slide copy"))]
        ])
      )
    );

    const detection = await detectNativeKeynotePackage(keyPath);
    assert.equal(detection.isNative, true);
    assert.equal(detection.container, "directory");
    assert.equal(detection.hasIndexZip, true);
    assert.deepEqual(detection.iwaPaths, ["Index/Document.iwa", "Index/Slide-1.iwa", "Index/Slide-2.iwa"]);

    const deck = await parseNativeKeynoteToIr(keyPath);
    assert.equal(validateIR(deck).valid, true);
    assert.equal(deck.deck.title, "Native Fixture");
    assert.equal(deck.deck.size.width, 1024);
    assert.equal(deck.deck.size.height, 768);
    assert.equal(deck.deck.slides.length, 2);
    assert.equal(deck.deck.slides[0]?.objects[0]?.type, "text");
    assert.equal(deck.deck.slides[0]?.objects[0]?.metadata?.nativeSourcePath, "Index/Slide-1.iwa");
    assert.deepEqual(
      deck.deck.slides.map((slide) => slide.objects.map((object) => (object.type === "text" ? object.text.plainText : "")).join("|")),
      ["Native Title|First bullet", "Second slide copy"]
    );
    assert.equal(deck.conversion?.status, "partial");
    assert.equal(deck.conversion?.tool, "keymorph-keynote-native-probe");
  });

  test("parses a ZIP-backed .key package with loose Index IWA entries", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "keymorph-keynote-zip-"));
    const keyPath = path.join(dir, "sample.key");
    await writeFile(
      keyPath,
      zip(
        new Map([
          ["Metadata/Properties.plist", plist({ title: "Zipped Native Fixture" })],
          ["Index/Slide1.iwa", protoString("Zipped slide text")]
        ])
      )
    );

    const detection = await detectNativeKeynotePackage(keyPath);
    assert.equal(detection.isNative, true);
    assert.equal(detection.container, "zip");
    assert.equal(detection.hasIndexZip, false);

    const deck = await parseKeynoteToIr(keyPath, { preferNative: true });
    assert.equal(validateIR(deck).valid, true);
    assert.equal(deck.deck.slides.length, 1);
    assert.equal(deck.deck.slides[0]?.objects[0]?.type, "text");
    assert.equal(deck.deck.slides[0]?.objects[0]?.type === "text" ? deck.deck.slides[0].objects[0].text.plainText : undefined, "Zipped slide text");
  });

  test("falls back to native parsing when the Keynote bridge cannot run", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "keymorph-keynote-fallback-"));
    const keyPath = path.join(dir, "fallback.key");
    await mkdir(path.join(keyPath, "Index"), { recursive: true });
    await writeFile(path.join(keyPath, "Index", "Slide-1.iwa"), protoString("Fallback text"));

    const deck = await parseKeynoteToIr(keyPath);
    assert.equal(validateIR(deck).valid, true);
    assert.equal(deck.deck.slides[0]?.objects[0]?.type, "text");
    assert.equal(deck.conversion?.messages[0]?.code, "keynote-bridge-fallback-native");
  });
});

function protoString(value: string, fieldNumber = 1): Uint8Array {
  const encoded = new TextEncoder().encode(value);
  return concat([varint((fieldNumber << 3) | 2), varint(encoded.length), encoded]);
}

function plist(values: Record<string, string | number>): Uint8Array {
  const body = Object.entries(values)
    .map(([key, value]) => {
      const element = typeof value === "number" ? `<real>${value}</real>` : `<string>${escapeXml(value)}</string>`;
      return `<key>${escapeXml(key)}</key>${element}`;
    })
    .join("");

  return new TextEncoder().encode(`<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0"><dict>${body}</dict></plist>`);
}

function zip(files: Map<string, Uint8Array>): Uint8Array {
  const chunks: Uint8Array[] = [];
  for (const [name, data] of files) {
    const nameBytes = new TextEncoder().encode(name);
    const compressed = deflateRawSync(data);
    chunks.push(
      u32(0x04034b50),
      u16(20),
      u16(0),
      u16(8),
      u16(0),
      u16(0),
      u32(crc32(data)),
      u32(compressed.byteLength),
      u32(data.byteLength),
      u16(nameBytes.byteLength),
      u16(0),
      nameBytes,
      compressed
    );
  }
  return concat(chunks);
}

function snappyFramed(payload: Uint8Array): Uint8Array {
  const compressed = snappyLiteralBlock(payload);
  const checksum = u32(0);
  return concat([
    new Uint8Array([0xff, 0x06, 0x00, 0x00]),
    new TextEncoder().encode("sNaPpY"),
    new Uint8Array([0x00]),
    u24(checksum.byteLength + compressed.byteLength),
    checksum,
    compressed
  ]);
}

function snappyLiteralBlock(payload: Uint8Array): Uint8Array {
  const chunks: Uint8Array[] = [varint(payload.byteLength)];
  if (payload.byteLength < 61) {
    chunks.push(new Uint8Array([((payload.byteLength - 1) << 2) | 0]));
  } else if (payload.byteLength < 256) {
    chunks.push(new Uint8Array([(60 << 2) | 0, payload.byteLength - 1]));
  } else {
    throw new Error("Synthetic Snappy helper only supports short payloads.");
  }
  chunks.push(payload);
  return concat(chunks);
}

function varint(value: number): Uint8Array {
  const bytes: number[] = [];
  let remaining = value;
  while (remaining > 0x7f) {
    bytes.push((remaining & 0x7f) | 0x80);
    remaining = Math.floor(remaining / 128);
  }
  bytes.push(remaining);
  return new Uint8Array(bytes);
}

function concat(chunks: Uint8Array[]): Uint8Array {
  const length = chunks.reduce((total, chunk) => total + chunk.byteLength, 0);
  const out = new Uint8Array(length);
  let offset = 0;
  for (const chunk of chunks) {
    out.set(chunk, offset);
    offset += chunk.byteLength;
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

function u24(value: number): Uint8Array {
  return new Uint8Array([value & 0xff, (value >>> 8) & 0xff, (value >>> 16) & 0xff]);
}

function escapeXml(value: string): string {
  return value.replace(/[<>&'"]/g, (char) => {
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
