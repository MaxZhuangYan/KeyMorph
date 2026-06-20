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
    const dir = await mkdtemp(path.join(tmpdir(), "keymorph-keynote-missing-"));
    const missingKeyPath = path.join(dir, "missing.key");
    await assert.rejects(
      () => exportKeynoteToPptx(missingKeyPath, path.join(dir, "missing.pptx")),
      /Input Keynote file does not exist/
    );

    const keyPath = path.join(dir, "present.key");
    await writeFile(keyPath, new Uint8Array([0x50, 0x4b, 0x03, 0x04]));
    await withAutomationDisabled(async () => {
      await assert.rejects(
        () => exportKeynoteToPptx(keyPath, path.join(dir, "present.pptx")),
        /Keynote GUI automation is disabled by default/
      );
    });
  });

  test("detects and parses a native directory-style .key package", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "keymorph-keynote-native-"));
    const keyPath = path.join(dir, "sample.key");
    await mkdir(path.join(keyPath, "Metadata"), { recursive: true });
    await mkdir(path.join(keyPath, "QuickLook"), { recursive: true });
    await writeFile(
      path.join(keyPath, "Metadata", "Properties.plist"),
      plist({
        title: "Native Fixture",
        author: "KeyMorph Test",
        CreationDate: "2026-01-02T03:04:05Z",
        ModificationDate: "2026-01-03T03:04:05Z",
        DocumentIdentifier: "fixture-document-id",
        slideWidth: 1024,
        slideHeight: 768
      })
    );
    await writeFile(path.join(keyPath, "QuickLook", "Thumbnail.jpg"), new Uint8Array([0xff, 0xd8, 0xff, 0xd9]));
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
    assert.equal(detection.packageFormat, "directory-index-zip");
    assert.equal(detection.hasLooseIndexDirectory, false);
    assert.equal(detection.hasQuickLookPreview, true);
    assert.deepEqual(detection.quickLookPaths, ["QuickLook/Thumbnail.jpg"]);
    assert.equal(detection.hasIndexZip, true);
    assert.deepEqual(detection.iwaPaths, ["Index/Document.iwa", "Index/Slide-1.iwa", "Index/Slide-2.iwa"]);
    assert.equal(detection.iwaStreams?.find((stream) => stream.path === "Index/Slide-1.iwa")?.role, "slide");
    assert.ok(detection.iwaStreams?.find((stream) => stream.path === "Index/Slide-2.iwa")?.compression.includes("snappy-framed"));
    assert.ok((detection.iwaStreams?.find((stream) => stream.path === "Index/Slide-1.iwa")?.protobufFieldCount ?? 0) >= 2);

    const deck = await parseNativeKeynoteToIr(keyPath);
    assert.equal(validateIR(deck).valid, true);
    assert.equal(deck.deck.title, "Native Fixture");
    assert.equal(deck.metadata?.author, "KeyMorph Test");
    assert.equal(deck.metadata?.createdAt, "2026-01-02T03:04:05Z");
    assert.equal(deck.metadata?.updatedAt, "2026-01-03T03:04:05Z");
    assert.equal(deck.metadata?.custom?.nativePackageFormat, "directory-index-zip");
    assert.equal(deck.metadata?.custom?.nativeDocumentIdentifier, "fixture-document-id");
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
    assert.equal(deck.conversion?.messages.some((message) => message.code === "keynote-native-iwa-fields-scanned"), true);
    assert.equal(deck.conversion?.metadata?.packageFormat, "directory-index-zip");
  });

  test("scans nested protobuf fields and records uncertain field evidence", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "keymorph-keynote-nested-iwa-"));
    const keyPath = path.join(dir, "nested.key");
    await mkdir(path.join(keyPath, "Data"), { recursive: true });
    await mkdir(path.join(keyPath, "Index"), { recursive: true });
    await writeFile(path.join(keyPath, "Data", "hero image.png"), pngBytes());

    const nestedShapeMessage = concat([
      protoString("Nested field title", 1),
      protoString("Data/hero%20image.png?checksum=ignored", 2)
    ]);
    await writeFile(
      path.join(keyPath, "Index", "Slide-1.iwa"),
      concat([
        protoString("Top-level title", 1),
        protoBytes(nestedShapeMessage, 7),
        protoVarint(3, 42)
      ])
    );

    const detection = await detectNativeKeynotePackage(keyPath);
    const stream = detection.iwaStreams?.[0];
    assert.equal(detection.packageFormat, "directory-loose-index");
    assert.equal(detection.hasLooseIndexDirectory, true);
    assert.ok((stream?.protobufFieldCount ?? 0) >= 4);
    assert.ok((stream?.nestedMessageCount ?? 0) >= 1);
    assert.equal(stream?.fieldSummaries.some((summary) => summary.fieldPath === "7.1" && summary.sampleText === "Nested field title"), true);
    assert.equal(stream?.assetReferenceCount, 1);

    const deck = await parseNativeKeynoteToIr(keyPath);
    assert.equal(validateIR(deck).valid, true);
    assert.equal(deck.deck.slides[0]?.metadata?.nativeNestedMessageCount, 1);
    assert.equal(deck.deck.slides[0]?.metadata?.nativeFieldSummaries?.some((summary) => summary.fieldPath === "7.1"), true);
    assert.equal(deck.deck.slides[0]?.objects.some((object) => object.type === "image" && object.name === "hero image.png"), true);
    assert.equal(
      deck.deck.slides[0]?.objects.some(
        (object) => object.type === "text" && object.metadata?.nativeFieldPath === "7.1" && object.text.plainText === "Nested field title"
      ),
      true
    );
    assert.equal(deck.conversion?.unsupportedFeatures?.some((feature) => feature.code === "keynote-native-protobuf-schema-private"), true);
    assert.equal(deck.conversion?.uncertainMappings?.some((mapping) => mapping.code === "keynote-native-container-detection"), true);
    assert.equal(deck.conversion?.metadata?.totalNestedMessageCount, 1);
  });

  test("extracts package asset references into IR assets and approximate objects", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "keymorph-keynote-assets-"));
    const keyPath = path.join(dir, "assets.key");
    await mkdir(path.join(keyPath, "Data"), { recursive: true });
    await mkdir(path.join(keyPath, "Index"), { recursive: true });
    await writeFile(path.join(keyPath, "Data", "hero.png"), pngBytes());
    await writeFile(path.join(keyPath, "Data", "clip.mov"), new Uint8Array([0, 0, 0, 20, 102, 116, 121, 112]));
    await writeFile(
      path.join(keyPath, "Index", "Slide-1.iwa"),
      concat([protoString("Asset slide"), protoString("Data/hero.png"), protoString("clip.mov")])
    );

    const detection = await detectNativeKeynotePackage(keyPath);
    assert.deepEqual(detection.assetPaths, ["Data/clip.mov", "Data/hero.png"]);
    assert.equal(detection.packageFormat, "directory-loose-index");
    assert.equal(detection.iwaStreams?.[0]?.assetReferenceCount, 2);
    assert.ok((detection.iwaStreams?.[0]?.referenceCandidateCount ?? 0) >= 2);

    const deck = await parseNativeKeynoteToIr(keyPath);
    assert.equal(validateIR(deck).valid, true);
    assert.equal(deck.deck.assets?.length, 2);
    assert.equal(deck.deck.assets?.find((asset) => asset.name === "hero.png")?.kind, "image");
    assert.equal(deck.deck.assets?.find((asset) => asset.name === "clip.mov")?.kind, "video");
    assert.equal(deck.deck.slides[0]?.objects.some((object) => object.type === "image"), true);
    assert.equal(deck.deck.slides[0]?.objects.some((object) => object.type === "media" && object.mediaType === "video"), true);
    assert.equal(deck.conversion?.messages.some((message) => message.code === "keynote-native-assets-referenced"), true);
    assert.equal(deck.conversion?.unsupportedFeatures?.some((feature) => feature.code === "keynote-native-asset-layout-loss"), true);
    assert.equal(deck.conversion?.metadata?.recoveredAssetObjectCount, 2);
    assert.equal(deck.conversion?.metadata?.unrecoveredAssetCount, 0);
  });

  test("reports detected but unplaced package assets separately from recovered text", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "keymorph-keynote-unplaced-assets-"));
    const keyPath = path.join(dir, "unplaced.key");
    await mkdir(path.join(keyPath, "Data"), { recursive: true });
    await mkdir(path.join(keyPath, "Index"), { recursive: true });
    await writeFile(path.join(keyPath, "Data", "unused.jpg"), new Uint8Array([0xff, 0xd8, 0xff, 0xd9]));
    await writeFile(path.join(keyPath, "Index", "Slide-1.iwa"), protoString("Only recovered text"));

    const deck = await parseNativeKeynoteToIr(keyPath);
    assert.equal(validateIR(deck).valid, true);
    assert.equal(deck.deck.assets?.length, 1);
    assert.equal(deck.deck.slides[0]?.objects.some((object) => object.type === "text"), true);
    assert.equal(deck.deck.slides[0]?.objects.some((object) => object.type === "image" || object.type === "media"), false);
    assert.equal(deck.conversion?.messages.some((message) => message.code === "keynote-native-text-recovered"), true);
    assert.equal(deck.conversion?.messages.some((message) => message.code === "keynote-native-assets-unplaced"), true);
    assert.equal(deck.conversion?.metadata?.recoveredTextObjectCount, 1);
    assert.equal(deck.conversion?.metadata?.recoveredAssetObjectCount, 0);
    assert.equal(deck.conversion?.metadata?.unrecoveredAssetCount, 1);
  });

  test("uses a full-slide preview image instead of low-confidence raw native text", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "keymorph-keynote-preview-fallback-"));
    const keyPath = path.join(dir, "preview.key");
    await mkdir(path.join(keyPath, "Index"), { recursive: true });
    await mkdir(path.join(keyPath, "QuickLook"), { recursive: true });
    await writeFile(path.join(keyPath, "QuickLook", "Preview.jpg"), jpegBytes(1280, 720));
    await writeFile(path.join(keyPath, "Index", "Slide-1.iwa"), new TextEncoder().encode("Low confidence raw caption"));

    const deck = await parseNativeKeynoteToIr(keyPath);
    assert.equal(validateIR(deck).valid, true);
    assert.equal(deck.deck.assets?.some((asset) => asset.name === "Preview.jpg" && asset.metadata?.nativePreviewAsset === true), true);
    const object = deck.deck.slides[0]?.objects[0];
    assert.equal(object?.type, "image");
    assert.equal(object?.metadata?.nativeFallback, "full-slide-preview");
    assert.deepEqual(object?.bounds, { x: 0, y: 0, width: 1280, height: 720 });
    assert.equal(object?.type === "image" ? object.source.assetId : undefined, deck.deck.assets?.find((asset) => asset.name === "Preview.jpg")?.id);
    assert.equal(deck.deck.slides[0]?.objects.some((slideObject) => slideObject.type === "text"), false);
    assert.equal(deck.conversion?.messages.some((message) => message.code === "keynote-native-preview-fallback-used"), true);
    assert.equal(deck.conversion?.degradedFeatures?.some((feature) => feature.code === "keynote-native-preview-fallback"), true);
    assert.equal(deck.conversion?.metadata?.recoveredTextObjectCount, 0);
    assert.equal(deck.conversion?.metadata?.previewFallbackObjectCount, 1);
    assert.equal(deck.conversion?.metadata?.previewAssetCount, 1);
  });

  test("uses package snapshot assets as visual-first native fallback", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "keymorph-keynote-snapshot-assets-"));
    const keyPath = path.join(dir, "snapshot.key");
    await mkdir(path.join(keyPath, "Data"), { recursive: true });
    await mkdir(path.join(keyPath, "Index"), { recursive: true });
    await writeFile(path.join(keyPath, "Data", "st-12345678-1234-1234-1234-123456789abc.png"), pngBytes(1280, 720));
    await writeFile(path.join(keyPath, "Index", "Slide-1.iwa"), protoString("Confident recovered title"));

    const deck = await parseNativeKeynoteToIr(keyPath);
    assert.equal(validateIR(deck).valid, true);
    const snapshot = deck.deck.assets?.find((asset) => asset.name === "st-12345678-1234-1234-1234-123456789abc.png");
    assert.equal(snapshot?.metadata?.nativePreviewAsset, true);
    assert.equal(snapshot?.metadata?.nativePreviewRole, "snapshot");
    assert.equal(deck.deck.slides[0]?.objects[0]?.type, "image");
    assert.equal(deck.deck.slides[0]?.objects[0]?.metadata?.nativeFallback, "full-slide-preview");
    assert.equal(deck.conversion?.messages.some((message) => message.code === "keynote-native-preview-fallback-used"), true);
    assert.equal(deck.conversion?.metadata?.previewAssetCount, 1);
    assert.equal(deck.conversion?.metadata?.previewFallbackObjectCount, 1);
  });

  test("extracts numeric geometry, grouping, image dimensions, QuickLook metadata, and animation uncertainty", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "keymorph-keynote-geometry-"));
    const keyPath = path.join(dir, "geometry.key");
    await mkdir(path.join(keyPath, "Data"), { recursive: true });
    await mkdir(path.join(keyPath, "Index"), { recursive: true });
    await mkdir(path.join(keyPath, "QuickLook"), { recursive: true });
    await writeFile(path.join(keyPath, "Data", "hero.png"), pngBytes(320, 180));
    await writeFile(path.join(keyPath, "QuickLook", "Thumbnail.jpg"), jpegBytes(640, 360));

    const objectGroup = concat([
      protoString("Hero title", 1),
      protoString("Data/hero.png", 2),
      protoFixed32(3, 100),
      protoFixed32(4, 120),
      protoFixed32(5, 320),
      protoFixed32(6, 180),
      protoString("Magic Move morph transition", 7)
    ]);
    await writeFile(path.join(keyPath, "Index", "Slide-1.iwa"), protoBytes(objectGroup, 9));

    const detection = await detectNativeKeynotePackage(keyPath);
    assert.equal(detection.quickLookPreviews[0]?.path, "QuickLook/Thumbnail.jpg");
    assert.equal(detection.quickLookPreviews[0]?.width, 640);
    assert.equal(detection.quickLookPreviews[0]?.height, 360);
    const stream = detection.iwaStreams?.[0];
    assert.ok((stream?.numericCandidateCount ?? 0) >= 4);
    assert.ok((stream?.geometryCandidateCount ?? 0) >= 1);
    assert.ok((stream?.groupingHintCount ?? 0) >= 1);
    assert.ok((stream?.animationHintCount ?? 0) >= 2);
    assert.equal(stream?.magicMoveHintCount, 1);
    assert.equal(stream?.morphHintCount, 1);
    assert.deepEqual(stream?.geometryCandidates[0]?.bounds, { x: 100, y: 120, width: 320, height: 180 });
    assert.equal(stream?.groupingHints[0]?.groupPath, "9");

    const deck = await parseNativeKeynoteToIr(keyPath);
    assert.equal(validateIR(deck).valid, true);
    const imageAsset = deck.deck.assets?.find((asset) => asset.name === "hero.png");
    assert.equal(imageAsset?.width, 320);
    assert.equal(imageAsset?.height, 180);
    assert.equal(imageAsset?.metadata?.nativeImageDimensionSource, "png-ihdr");
    const textObject = deck.deck.slides[0]?.objects.find((object) => object.type === "text");
    assert.deepEqual(textObject?.metadata?.nativeGeometryCandidate?.bounds, { x: 100, y: 120, width: 320, height: 180 });
    assert.equal(deck.deck.slides[0]?.metadata?.nativeGeometryCandidateCount, 1);
    assert.equal(deck.deck.slides[0]?.metadata?.nativeGroupingHintCount, 1);
    assert.ok((deck.deck.slides[0]?.metadata?.nativeAnimationHintCount as number) >= 2);
    assert.equal(deck.conversion?.messages.some((message) => message.code === "keynote-native-geometry-candidates-detected"), true);
    assert.equal(deck.conversion?.messages.some((message) => message.code === "keynote-native-animation-hints-detected"), true);
    assert.equal(deck.conversion?.uncertainMappings?.some((mapping) => mapping.code === "keynote-native-geometry-candidate-scan"), true);
    assert.equal(deck.conversion?.uncertainMappings?.some((mapping) => mapping.code === "keynote-native-object-grouping-hints"), true);
    assert.equal(deck.conversion?.uncertainMappings?.some((mapping) => mapping.code === "keynote-native-magic-move-morph-hints"), true);
    assert.equal(deck.conversion?.metadata?.quickLookPreviewWithDimensionsCount, 1);
    assert.equal(deck.conversion?.metadata?.imageAssetWithDimensionsCount, 1);
    assert.equal(deck.conversion?.metadata?.lossReport?.automationUsed, false);
    assert.equal(deck.conversion?.metadata?.lossReport?.evidenceCounts?.geometryCandidateCount, 1);
    assert.equal(deck.conversion?.metadata?.lossReport?.evidenceCounts?.quickLookPreviewWithDimensionsCount, 1);
  });

  test("filters Keynote internal tokens and binary residue out of recovered slide text", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "keymorph-keynote-noise-filter-"));
    const keyPath = path.join(dir, "noise.key");
    await mkdir(path.join(keyPath, "Index"), { recursive: true });
    await writeFile(
      path.join(keyPath, "Index", "Slide-1.iwa"),
      concat([
        protoString("Transition"),
        protoString("apple:magic-move-implied-motion-path"),
        protoString("XBO Transition\u0012$a"),
        protoString("AI Agent 社会模拟游戏"),
        protoString("图片 8"),
        protoString("E.l"),
        protoString("decimal"),
        protoString("$apple:magic-move-implied-motion-path"),
        protoString("*b.S")
      ])
    );

    const deck = await parseNativeKeynoteToIr(keyPath);
    assert.equal(validateIR(deck).valid, true);
    const text = deck.deck.slides[0]?.objects.filter((object) => object.type === "text").map((object) => object.text.plainText);

    assert.deepEqual(text, ["AI Agent 社会模拟游戏"]);
  });

  test("uses native snapshot images instead of low-confidence text fallback", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "keymorph-keynote-snapshot-fallback-"));
    const keyPath = path.join(dir, "snapshot.key");
    await mkdir(path.join(keyPath, "Data"), { recursive: true });
    await mkdir(path.join(keyPath, "Index"), { recursive: true });
    await writeFile(path.join(keyPath, "Data", "st-97F332C4-B975-4E29-AE94-C56E183A29CC-16103.png"), pngBytes(356, 200));
    await writeFile(path.join(keyPath, "Index", "Slide-1.iwa"), protoString("1B/ Transition"));

    const deck = await parseNativeKeynoteToIr(keyPath);
    assert.equal(validateIR(deck).valid, true);
    const object = deck.deck.slides[0]?.objects.find((candidate) => candidate.type === "image");
    assert.equal(object?.type, "image");
    assert.equal(object?.metadata?.nativeFallback, "full-slide-preview");
    assert.equal(object?.type === "image" ? object.source.assetId : undefined, deck.deck.assets?.[0]?.id);
    assert.match(deck.deck.assets?.[0]?.uri ?? "", /^data:image\/png;base64,/);
    assert.equal(deck.conversion?.messages.some((message) => message.code === "keynote-native-preview-fallback-used"), true);
  });

  test("uses typed IWA archive message data-id and geometry evidence for native image placement", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "keymorph-keynote-typed-iwa-"));
    const keyPath = path.join(dir, "typed.key");
    await mkdir(path.join(keyPath, "Data"), { recursive: true });
    await mkdir(path.join(keyPath, "Index"), { recursive: true });
    await writeFile(path.join(keyPath, "Data", "hero-42.png"), pngBytes(320, 180));

    const imagePayload = concat([
      protoBytes(
        concat([
          protoBytes(
            concat([
              protoBytes(concat([protoFixed32(1, 100), protoFixed32(2, 120)]), 1),
              protoBytes(concat([protoFixed32(1, 320), protoFixed32(2, 180)]), 2)
            ]),
            1
          )
        ]),
        1
      ),
      protoBytes(protoVarint(1, 42), 11)
    ]);
    const buildPayload = concat([protoVarint(1, 1234), protoString("apple:bc-appear", 4)]);
    await writeFile(
      path.join(keyPath, "Index", "Slide-1.iwa"),
      iwaArchiveRecord(9001, [
        { type: 3005, payload: imagePayload, dataReferences: [42], objectReferences: [777] },
        { type: 8, payload: buildPayload, objectReferences: [777] },
        { type: 3097, payload: new Uint8Array() }
      ])
    );

    const detection = await detectNativeKeynotePackage(keyPath);
    const stream = detection.iwaStreams?.[0];
    assert.equal(stream?.archiveRecordCount, 1);
    assert.deepEqual(stream?.archiveRecords[0]?.messageTypes, [3005, 8, 3097]);
    assert.equal(stream?.typedArchiveMessageCount, 3);
    assert.equal(stream?.typedArchiveMessages.some((message) => message.type === 3005 && message.dataReferences.includes("42")), true);
    assert.equal(stream?.typedArchiveMessages.some((message) => message.type === 3097 && message.payloadLength === 0), true);

    const deck = await parseNativeKeynoteToIr(keyPath);
    assert.equal(validateIR(deck).valid, true);
    const object = deck.deck.slides[0]?.objects[0];
    assert.equal(object?.type, "image");
    assert.deepEqual(object?.bounds, { x: 100, y: 120, width: 320, height: 180 });
    assert.equal(object?.metadata?.nativeExtraction, "asset-archive-info-data-reference");
    assert.equal(object?.metadata?.nativeArchiveMessageType, 3005);
    assert.equal(object?.metadata?.nativeAssetDataId, "42");
    assert.deepEqual(object?.metadata?.nativeArchiveObjectReferences, ["777"]);
    assert.equal(deck.deck.slides[0]?.metadata?.nativeTypedArchiveMessageCount, 3);
    assert.equal(
      deck.deck.slides[0]?.metadata?.nativeTypedArchiveMessages?.some(
        (message) => message.type === 8 && message.textCandidates.includes("apple:bc-appear")
      ),
      true
    );
    assert.equal(deck.deck.slides[0]?.timeline?.events.length, 0);
    assert.equal(deck.conversion?.metadata?.totalTypedArchiveMessageCount, 3);
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
    assert.equal(detection.packageFormat, "zip-loose-index");
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

    const deck = await parseKeynoteToIr(keyPath, {
      bridgeExport: async () => {
        throw new Error("Synthetic Keynote bridge failure");
      }
    });
    assert.equal(validateIR(deck).valid, true);
    assert.equal(deck.deck.slides[0]?.objects[0]?.type, "text");
    assert.equal(deck.conversion?.messages[0]?.code, "keynote-bridge-fallback-native");
  });

  test("default bridge path does not trigger GUI automation before native fallback", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "keymorph-keynote-safe-fallback-"));
    const keyPath = path.join(dir, "safe.key");
    await mkdir(path.join(keyPath, "Index"), { recursive: true });
    await writeFile(path.join(keyPath, "Index", "Slide-1.iwa"), protoString("Safe native fallback"));

    const deck = await withAutomationDisabled(() => parseKeynoteToIr(keyPath));
    assert.equal(validateIR(deck).valid, true);
    assert.equal(deck.deck.slides[0]?.objects[0]?.type, "text");
    assert.match(deck.conversion?.messages[0]?.message ?? "", /Keynote GUI automation is disabled by default/);
  });
});

async function withAutomationDisabled<T>(callback: () => Promise<T>): Promise<T> {
  const previous = process.env.KEYMORPH_ALLOW_KEYNOTE_AUTOMATION;
  delete process.env.KEYMORPH_ALLOW_KEYNOTE_AUTOMATION;
  try {
    return await callback();
  } finally {
    if (previous === undefined) {
      delete process.env.KEYMORPH_ALLOW_KEYNOTE_AUTOMATION;
    } else {
      process.env.KEYMORPH_ALLOW_KEYNOTE_AUTOMATION = previous;
    }
  }
}

function protoString(value: string, fieldNumber = 1): Uint8Array {
  const encoded = new TextEncoder().encode(value);
  return protoBytes(encoded, fieldNumber);
}

function protoBytes(value: Uint8Array, fieldNumber: number): Uint8Array {
  return concat([varint((fieldNumber << 3) | 2), varint(value.length), value]);
}

function protoVarint(fieldNumber: number, value: number): Uint8Array {
  return concat([varint(fieldNumber << 3), varint(value)]);
}

function protoFixed32(fieldNumber: number, value: number): Uint8Array {
  const out = new Uint8Array(4);
  new DataView(out.buffer).setFloat32(0, value, true);
  return concat([varint((fieldNumber << 3) | 5), out]);
}

function iwaArchiveRecord(
  identifier: number,
  messages: Array<{ type: number; payload: Uint8Array; dataReferences?: number[]; objectReferences?: number[] }>
): Uint8Array {
  const archiveInfo = concat([
    protoVarint(1, identifier),
    ...messages.map((message) =>
      protoBytes(
        concat([
          protoVarint(1, message.type),
          protoVarint(3, message.payload.byteLength),
          ...(message.objectReferences ?? []).map((reference) => protoVarint(5, reference)),
          ...(message.dataReferences ?? []).map((reference) => protoVarint(6, reference))
        ]),
        2
      )
    )
  ]);
  return concat([varint(archiveInfo.byteLength), archiveInfo, ...messages.map((message) => message.payload)]);
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

function pngBytes(width = 1, height = 1): Uint8Array {
  const bytes = new Uint8Array([
    0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
    0x89, 0x00, 0x00, 0x00, 0x0a, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0x00, 0x01, 0x00, 0x00,
    0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae,
    0x42, 0x60, 0x82
  ]);
  new DataView(bytes.buffer).setUint32(16, width, false);
  new DataView(bytes.buffer).setUint32(20, height, false);
  return bytes;
}

function jpegBytes(width: number, height: number): Uint8Array {
  return new Uint8Array([
    0xff,
    0xd8,
    0xff,
    0xc0,
    0x00,
    0x11,
    0x08,
    (height >> 8) & 0xff,
    height & 0xff,
    (width >> 8) & 0xff,
    width & 0xff,
    0x03,
    0x01,
    0x11,
    0x00,
    0x02,
    0x11,
    0x00,
    0x03,
    0x11,
    0x00,
    0xff,
    0xd9
  ]);
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
