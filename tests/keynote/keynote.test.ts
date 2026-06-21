import { describe, test } from "node:test";
import assert from "node:assert/strict";
import { mkdir, mkdtemp, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { deflateRawSync } from "node:zlib";

import { detectNativeKeynotePackage, exportKeynoteToPptx, parseKeynoteToIr, parseNativeKeynoteToIr } from "../../src/keynote/index.ts";
import { validateIR } from "../../src/ir/index.ts";
import { createSlideTimingPlan } from "../../src/runtime/index.ts";

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

  test("reads native slide size from Document.iwa when plist dimensions are unavailable", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "keymorph-keynote-document-size-"));
    const keyPath = path.join(dir, "document-size.key");
    await mkdir(path.join(keyPath, "Index"), { recursive: true });
    await writeFile(
      path.join(keyPath, "Index", "Document.iwa"),
      iwaArchiveRecord(700, [{ type: 2, payload: nativeDocumentSizePayload(1920, 1080) }])
    );
    await writeFile(path.join(keyPath, "Index", "Slide-1.iwa"), protoString("Native size slide"));

    const deck = await parseNativeKeynoteToIr(keyPath);
    assert.equal(validateIR(deck).valid, true);
    assert.deepEqual(deck.deck.size, { width: 1920, height: 1080, unit: "px" });
    assert.equal(deck.metadata?.custom?.nativeDeckSizeSource, "document-iwa");
    assert.equal(deck.deck.slides[0]?.objects[0]?.bounds?.width, 1632);
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
    assert.equal(deck.deck.slides[0]?.transition, undefined);
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

  test("maps native Magic Move hints to conservative slide transitions", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "keymorph-keynote-magic-move-transition-"));
    const keyPath = path.join(dir, "magic.key");
    await mkdir(path.join(keyPath, "Index"), { recursive: true });
    await writeFile(path.join(keyPath, "Index", "Slide-1.iwa"), protoString("Intro slide"));
    await writeFile(path.join(keyPath, "Index", "Slide-2.iwa"), protoString("Magic Move morph transition"));

    const deck = await parseNativeKeynoteToIr(keyPath);
    assert.equal(validateIR(deck).valid, true);
    assert.equal(deck.deck.slides[0]?.transition, undefined);
    const transition = deck.deck.slides[1]?.transition;
    assert.equal(transition?.type, "magicMove");
    assert.equal(transition?.fromSlideId, "slide-1");
    assert.equal(transition?.toSlideId, "slide-2");
    assert.equal(transition?.morph?.strategy, "magicMove");
    assert.deepEqual(transition?.morph?.matching?.matchBy, ["morphKey", "objectId", "name", "geometry"]);
    assert.deepEqual(transition?.morph?.properties, ["bounds", "transform", "opacity"]);
    assert.equal(transition?.metadata?.nativeAnimationHintKind, "magicMove");
    assert.equal(deck.deck.slides[1]?.metadata?.nativeMagicMoveHintCount, 1);
    assert.equal(deck.deck.slides[1]?.objects.some((object) => object.type === "text" && /magic|morph|transition/i.test(object.text.plainText)), false);
  });

  test("keeps generic native transition hints without inventing a slide transition", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "keymorph-keynote-transition-hint-"));
    const keyPath = path.join(dir, "transition.key");
    await mkdir(path.join(keyPath, "Index"), { recursive: true });
    await writeFile(path.join(keyPath, "Index", "Slide-1.iwa"), protoString("Intro slide"));
    await writeFile(path.join(keyPath, "Index", "Slide-2.iwa"), protoString("XBO Transition"));

    const deck = await parseNativeKeynoteToIr(keyPath);
    assert.equal(validateIR(deck).valid, true);
    assert.equal(deck.deck.slides[1]?.transition, undefined);
    assert.equal(deck.deck.slides[1]?.metadata?.nativeAnimationHintCount, 1);
    assert.equal(deck.deck.slides[1]?.objects.some((object) => object.type === "text" && /transition/i.test(object.text.plainText)), false);
  });

  test("filters Keynote internal tokens and binary residue out of recovered slide text", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "keymorph-keynote-noise-filter-"));
    const keyPath = path.join(dir, "noise.key");
    await mkdir(path.join(keyPath, "Index"), { recursive: true });
    await writeFile(
      path.join(keyPath, "Index", "Slide-1.iwa"),
      concat([
        protoString("Transition"),
        protoString("path"),
        protoString("1B/ Transition"),
        protoString("zh-Hans zh-Hans"),
        protoString("S5=D+WC C"),
        protoString("T DBD"),
        protoString("9PD ĒDDD &"),
        protoString("C & VCBVC"),
        protoString("apple:magic-move-implied-motion-path"),
        protoString("XBO Transition\u0012$a"),
        protoString("AI Agent 社会模拟游戏"),
        protoString("Epoch 1 $NOM 上链 Agent 链上身份 资产可交易"),
        protoString("9Epoch 1 $NOM 上链 Agent 链上身份 资产可交易"),
        protoString("ODePIN 网络 玩家设备贡献算力"),
        protoString("DV1 链下原型 可运行的Demo"),
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

    assert.deepEqual(text, [
      "AI Agent 社会模拟游戏",
      "Epoch 1 $NOM 上链 Agent 链上身份 资产可交易",
      "DePIN 网络 玩家设备贡献算力",
      "V1 链下原型 可运行的Demo"
    ]);
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

    const imagePayload = nativeImagePlacementPayload(42, { x: 100, y: 120, width: 320, height: 180 });
    await writeFile(
      path.join(keyPath, "Index", "Slide-1.iwa"),
      concat([
        iwaArchiveRecord(777, [{ type: 3005, payload: imagePayload, dataReferences: [42], objectReferences: [777] }]),
        iwaArchiveRecord(9001, [
          {
            type: 8,
            payload: nativeBuildPayload({
              targetId: 777,
              direction: "In",
              effect: "apple:bc-appear",
              durationSeconds: 0.5
            })
          }
        ]),
        iwaArchiveRecord(9002, [{ type: 153, payload: nativeBuildTimingPayload({ buildId: 9001, durationSeconds: 0.5 }) }]),
        iwaArchiveRecord(9003, [{ type: 3097, payload: new Uint8Array() }])
      ])
    );

    const detection = await detectNativeKeynotePackage(keyPath);
    const stream = detection.iwaStreams?.[0];
    assert.equal(stream?.archiveRecordCount, 4);
    assert.equal(stream?.archiveRecords.some((record) => record.messageTypes.includes(3005)), true);
    assert.equal(stream?.archiveRecords.some((record) => record.messageTypes.includes(8)), true);
    assert.equal(stream?.archiveRecords.some((record) => record.messageTypes.includes(153)), true);
    assert.equal(stream?.archiveRecords.some((record) => record.messageTypes.includes(3097)), true);
    assert.equal(stream?.typedArchiveMessageCount, 4);
    assert.equal(stream?.typedArchiveMessages.some((message) => message.type === 3005 && message.dataReferences.includes("42")), true);
    assert.equal(stream?.typedArchiveMessages.some((message) => message.type === 3097 && message.payloadLength === 0), true);
    assert.equal(stream?.buildRecordCount, 1);
    assert.equal(stream?.buildTimingRecordCount, 1);

    const deck = await parseNativeKeynoteToIr(keyPath);
    assert.equal(validateIR(deck).valid, true);
    const object = deck.deck.slides[0]?.objects[0];
    assert.equal(object?.type, "image");
    assert.deepEqual(object?.bounds, { x: 100, y: 120, width: 320, height: 180 });
    assert.equal(object?.metadata?.nativeExtraction, "asset-archive-info-data-reference");
    assert.equal(object?.metadata?.nativeArchiveMessageType, 3005);
    assert.equal(object?.metadata?.nativeAssetDataId, "42");
    assert.deepEqual(object?.metadata?.nativeTypedVisualLayout, {
      kind: "image",
      frame: { x: 100, y: 120, width: 320, height: 180 },
      frameFieldPaths: ["1.1.1.1", "1.1.1.2", "1.1.2.1", "1.1.2.2"],
      schema: "typed-visual-frame-1.1",
      confidence: 0.97
    });
    assert.equal(object?.metadata?.nativeTypedVisualSchema, "typed-visual-frame-1.1");
    assert.deepEqual(object?.metadata?.nativeArchiveObjectReferences, ["777"]);
    assert.equal(deck.deck.slides[0]?.metadata?.nativeTypedArchiveMessageCount, 4);
    assert.equal(deck.deck.slides[0]?.metadata?.nativeTypedArchiveMessagesCount, 4);
    assert.equal(deck.deck.slides[0]?.metadata?.nativeTypedArchiveMessagesSampleCount, 4);
    assert.equal(deck.deck.slides[0]?.metadata?.nativeTypedArchiveMessagesSampleLimit, 24);
    assert.equal(deck.deck.slides[0]?.metadata?.nativeTypedArchiveMessagesTruncated, false);
    assert.equal(
      deck.deck.slides[0]?.metadata?.nativeTypedArchiveMessages?.some(
        (message) => message.type === 8 && message.build?.effect === "apple:bc-appear" && message.build.targetNativeId === "777"
      ),
      true
    );
    const event = deck.deck.slides[0]?.timeline?.events[0];
    assert.equal(event?.kind, "keyframes");
    assert.equal(event?.kind === "keyframes" ? event.targetId : undefined, object?.id);
    assert.equal(event?.durationMs, 500);
    assert.equal(event?.metadata?.nativeBuildEffect, "apple:bc-appear");
    assert.equal(event?.metadata?.nativeBuildFallback, "appear-in");
    assert.equal(event?.easing && typeof event.easing === "object" ? event.easing.type : undefined, "steps");
    assert.deepEqual(event?.kind === "keyframes" ? event.tracks[0]?.keyframes : undefined, [
      { offset: 0, value: 0 },
      { offset: 1, value: 1 }
    ]);
    assert.equal(event?.kind === "keyframes" ? event.tracks[0]?.interpolation : undefined, "discrete");
    assert.equal(deck.deck.slides[0]?.metadata?.nativeBuildAnimationRecoveredCount, 1);
    assert.equal(deck.conversion?.statistics.animationCount, 1);
    assert.equal(deck.conversion?.metadata?.totalTypedArchiveMessageCount, 4);
    assert.equal(deck.conversion?.metadata?.totalNativeBuildRecordCount, 1);
    assert.equal(deck.conversion?.metadata?.totalNativeBuildTimingRecordCount, 1);
    assert.equal(deck.conversion?.messages.some((message) => message.code === "keynote-native-build-animations-recovered"), true);
  });

  test("assigns stable native morph keys for unique repeated text and assets", async () => {
    const textDir = await mkdtemp(path.join(tmpdir(), "keymorph-keynote-native-text-morph-keys-"));
    const textKeyPath = path.join(textDir, "morph-text.key");
    await mkdir(path.join(textKeyPath, "Index"), { recursive: true });
    await writeFile(path.join(textKeyPath, "Index", "Slide-1.iwa"), protoString("Shared Title"));
    await writeFile(path.join(textKeyPath, "Index", "Slide-2.iwa"), protoString("Shared Title"));

    const textDeck = await parseNativeKeynoteToIr(textKeyPath);
    assert.equal(validateIR(textDeck).valid, true);
    const firstText = textDeck.deck.slides[0]?.objects.find((object) => object.type === "text" && object.text.plainText === "Shared Title");
    const secondText = textDeck.deck.slides[1]?.objects.find((object) => object.type === "text" && object.text.plainText === "Shared Title");
    assert.match(firstText?.morphKey ?? "", /^native:text:/);
    assert.equal(secondText?.morphKey, firstText?.morphKey);
    assert.equal(firstText?.metadata?.nativeMorphKeySource, "text-content");
    assert.equal(textDeck.deck.slides[0]?.metadata?.nativeMorphKeyCount, 1);
    assert.equal(textDeck.deck.slides[1]?.metadata?.nativeMorphKeyCount, 1);

    const assetDir = await mkdtemp(path.join(tmpdir(), "keymorph-keynote-native-asset-morph-keys-"));
    const assetKeyPath = path.join(assetDir, "morph-asset.key");
    await mkdir(path.join(assetKeyPath, "Data"), { recursive: true });
    await mkdir(path.join(assetKeyPath, "Index"), { recursive: true });
    await writeFile(path.join(assetKeyPath, "Data", "hero-42.png"), pngBytes(320, 180));
    await writeFile(
      path.join(assetKeyPath, "Index", "Slide-1.iwa"),
      iwaArchiveRecord(111, [
        { type: 3005, payload: nativeImagePlacementPayload(42, { x: 10, y: 20, width: 100, height: 80 }), dataReferences: [42] }
      ])
    );
    await writeFile(
      path.join(assetKeyPath, "Index", "Slide-2.iwa"),
      iwaArchiveRecord(222, [
        { type: 3005, payload: nativeImagePlacementPayload(42, { x: 210, y: 120, width: 180, height: 140 }), dataReferences: [42] }
      ])
    );

    const assetDeck = await parseNativeKeynoteToIr(assetKeyPath);
    assert.equal(validateIR(assetDeck).valid, true);
    const firstImage = assetDeck.deck.slides[0]?.objects.find((object) => object.type === "image");
    const secondImage = assetDeck.deck.slides[1]?.objects.find((object) => object.type === "image");
    assert.equal(firstImage?.morphKey, "native:asset-data:42");
    assert.equal(secondImage?.morphKey, firstImage?.morphKey);
    assert.equal(firstImage?.metadata?.nativeMorphKeySource, "asset-data-id");
    assert.equal(assetDeck.deck.slides[0]?.metadata?.nativeMorphKeyCount, 1);
    assert.equal(assetDeck.deck.slides[1]?.metadata?.nativeMorphKeyCount, 1);

    const checksumDir = await mkdtemp(path.join(tmpdir(), "keymorph-keynote-native-checksum-morph-keys-"));
    const checksumKeyPath = path.join(checksumDir, "morph-checksum.key");
    await mkdir(path.join(checksumKeyPath, "Data"), { recursive: true });
    await mkdir(path.join(checksumKeyPath, "Index"), { recursive: true });
    const sharedImageBytes = pngBytes(320, 180);
    await writeFile(path.join(checksumKeyPath, "Data", "hero.png"), sharedImageBytes);
    await writeFile(path.join(checksumKeyPath, "Data", "copy.png"), sharedImageBytes);
    await writeFile(path.join(checksumKeyPath, "Index", "Slide-1.iwa"), protoString("Data/hero.png"));
    await writeFile(path.join(checksumKeyPath, "Index", "Slide-2.iwa"), protoString("Data/copy.png"));

    const checksumDeck = await parseNativeKeynoteToIr(checksumKeyPath);
    assert.equal(validateIR(checksumDeck).valid, true);
    const firstChecksumImage = checksumDeck.deck.slides[0]?.objects.find((object) => object.type === "image");
    const secondChecksumImage = checksumDeck.deck.slides[1]?.objects.find((object) => object.type === "image");
    assert.match(firstChecksumImage?.morphKey ?? "", /^native:asset-checksum:/);
    assert.equal(secondChecksumImage?.morphKey, firstChecksumImage?.morphKey);
    assert.equal(firstChecksumImage?.metadata?.nativeMorphKeySource, "asset-checksum");
  });

  test("preserves repeated native image placements so build targets can resolve", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "keymorph-keynote-repeated-placement-"));
    const keyPath = path.join(dir, "repeat.key");
    await mkdir(path.join(keyPath, "Data"), { recursive: true });
    await mkdir(path.join(keyPath, "Index"), { recursive: true });
    await writeFile(path.join(keyPath, "Data", "hero-42.png"), pngBytes(320, 180));

    const firstPayload = nativeImagePlacementPayload(42, { x: 10, y: 20, width: 100, height: 80 });
    const secondPayload = nativeImagePlacementPayload(42, { x: 210, y: 20, width: 100, height: 80 });
    await writeFile(
      path.join(keyPath, "Index", "Slide-1.iwa"),
      concat([
        iwaArchiveRecord(111, [{ type: 3005, payload: firstPayload, dataReferences: [42] }]),
        iwaArchiveRecord(222, [{ type: 3005, payload: secondPayload, dataReferences: [42] }]),
        iwaArchiveRecord(9001, [
          {
            type: 8,
            payload: nativeBuildPayload({ targetId: 222, direction: "In", effect: "com.apple.iWork.Keynote.Blur", durationSeconds: 0.5 })
          }
        ]),
        iwaArchiveRecord(9002, [{ type: 153, payload: nativeBuildTimingPayload({ buildId: 9001, durationSeconds: 0.5 }) }])
      ])
    );

    const deck = await parseNativeKeynoteToIr(keyPath);
    assert.equal(validateIR(deck).valid, true);
    const imageObjects = deck.deck.slides[0]?.objects.filter((object) => object.type === "image") ?? [];
    assert.equal(imageObjects.length, 2);
    assert.deepEqual(
      imageObjects.map((object) => object.morphKey),
      [undefined, undefined]
    );
    assert.deepEqual(
      imageObjects.map((object) => object.metadata?.nativeMorphKeySuppressedReason),
      ["duplicate-within-slide", "duplicate-within-slide"]
    );
    assert.deepEqual(
      imageObjects.map((object) => object.metadata?.nativeArchiveIdentifier).sort(),
      ["111", "222"]
    );
    const event = deck.deck.slides[0]?.timeline?.events[0];
    assert.equal(event?.kind, "keyframes");
    assert.equal(event?.targetId, imageObjects.find((object) => object.metadata?.nativeArchiveIdentifier === "222")?.id);
    assert.equal(event?.metadata?.nativeBuildTargetId, "222");
    assert.equal(event?.metadata?.nativeBuildFallback, "blur-in");
    assert.equal(typeof event?.metadata?.nativeBuildDegradation, "string");
    assert.deepEqual(event?.kind === "keyframes" ? event.tracks.find((track) => track.property === "filter.blurPx")?.keyframes : undefined, [
      { offset: 0, value: 18 },
      { offset: 1, value: 0 }
    ]);
  });

  test("marks sampled native diagnostics as truncated when evidence exceeds the metadata sample limit", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "keymorph-keynote-sampled-diagnostics-"));
    const keyPath = path.join(dir, "sampled.key");
    await mkdir(path.join(keyPath, "Index"), { recursive: true });

    await writeFile(
      path.join(keyPath, "Index", "Slide-1.iwa"),
      concat(
        Array.from({ length: 30 }, (_, index) =>
          iwaArchiveRecord(1000 + index, [
            {
              type: 3097,
              payload: new Uint8Array()
            }
          ])
        )
      )
    );

    const deck = await parseNativeKeynoteToIr(keyPath);
    assert.equal(validateIR(deck).valid, true);
    assert.equal(deck.deck.slides[0]?.metadata?.nativeTypedArchiveMessageCount, 30);
    assert.equal(deck.deck.slides[0]?.metadata?.nativeTypedArchiveMessagesCount, 30);
    assert.equal(deck.deck.slides[0]?.metadata?.nativeTypedArchiveMessagesSampleCount, 24);
    assert.equal(deck.deck.slides[0]?.metadata?.nativeTypedArchiveMessagesSampleLimit, 24);
    assert.equal(deck.deck.slides[0]?.metadata?.nativeTypedArchiveMessagesTruncated, true);
    assert.equal(deck.deck.slides[0]?.metadata?.nativeTypedArchiveMessages?.length, 24);
  });

  test("marks near-identical native placements as a group without collapsing build targets", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "keymorph-keynote-placement-groups-"));
    const keyPath = path.join(dir, "placement-groups.key");
    await mkdir(path.join(keyPath, "Data"), { recursive: true });
    await mkdir(path.join(keyPath, "Index"), { recursive: true });
    await writeFile(path.join(keyPath, "Data", "hero-42.png"), pngBytes(320, 180));

    await writeFile(
      path.join(keyPath, "Index", "Slide-1.iwa"),
      concat([
        iwaArchiveRecord(111, [
          { type: 3005, payload: nativeImagePlacementPayload(42, { x: 100.1, y: 120.2, width: 320.3, height: 180.4 }), dataReferences: [42] }
        ]),
        iwaArchiveRecord(222, [
          { type: 3005, payload: nativeImagePlacementPayload(42, { x: 100.4, y: 120.3, width: 320.1, height: 180.2 }), dataReferences: [42] }
        ]),
        iwaArchiveRecord(9001, [
          { type: 8, payload: nativeBuildPayload({ targetId: 222, direction: "In", effect: "com.apple.iWork.Keynote.Blur", durationSeconds: 0.5 }) }
        ]),
        iwaArchiveRecord(9002, [{ type: 153, payload: nativeBuildTimingPayload({ buildId: 9001, durationSeconds: 0.5 }) }])
      ])
    );

    const deck = await parseNativeKeynoteToIr(keyPath);
    assert.equal(validateIR(deck).valid, true);
    const imageObjects = deck.deck.slides[0]?.objects.filter((object) => object.type === "image") ?? [];
    assert.equal(imageObjects.length, 2);
    assert.equal(imageObjects[0]?.metadata?.nativePlacementGroupSize, 2);
    assert.equal(imageObjects[0]?.metadata?.nativePlacementGroupKey, imageObjects[1]?.metadata?.nativePlacementGroupKey);
    assert.deepEqual(imageObjects[0]?.metadata?.nativePlacementGroupObjectIds, imageObjects.map((object) => object.id));
    const event = deck.deck.slides[0]?.timeline?.events[0];
    assert.equal(event?.targetId, imageObjects.find((object) => object.metadata?.nativeArchiveIdentifier === "222")?.id);
  });

  test("coalesces native asset variants for the same archive object", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "keymorph-keynote-asset-variants-"));
    const keyPath = path.join(dir, "variants.key");
    await mkdir(path.join(keyPath, "Data"), { recursive: true });
    await mkdir(path.join(keyPath, "Index"), { recursive: true });
    await writeFile(path.join(keyPath, "Data", "hero-42.png"), pngBytes(640, 360));
    await writeFile(path.join(keyPath, "Data", "hero-small-43.png"), pngBytes(160, 90));

    await writeFile(
      path.join(keyPath, "Index", "Slide-1.iwa"),
      concat([
        iwaArchiveRecord(777, [
          {
            type: 3005,
            payload: nativeImagePlacementPayload(42, { x: 100, y: 120, width: 320, height: 180 }),
            dataReferences: [42, 43],
            objectReferences: [777]
          }
        ]),
        iwaArchiveRecord(9001, [
          { type: 8, payload: nativeBuildPayload({ targetId: 777, direction: "In", effect: "com.apple.iWork.Keynote.Blur", durationSeconds: 0.5 }) }
        ]),
        iwaArchiveRecord(9002, [{ type: 153, payload: nativeBuildTimingPayload({ buildId: 9001, durationSeconds: 0.5 }) }])
      ])
    );

    const deck = await parseNativeKeynoteToIr(keyPath);
    assert.equal(validateIR(deck).valid, true);
    const imageObjects = deck.deck.slides[0]?.objects.filter((object) => object.type === "image") ?? [];
    assert.equal(imageObjects.length, 1);
    assert.equal(imageObjects[0]?.metadata?.nativeArchiveIdentifier, "777");
    assert.equal(imageObjects[0]?.metadata?.nativeAssetPath, "Data/hero-42.png");
    assert.equal(imageObjects[0]?.metadata?.nativeSuppressedAssetVariantCount, 1);
    assert.deepEqual(imageObjects[0]?.metadata?.nativeSuppressedAssetVariants, [
      { path: "Data/hero-small-43.png", mimeType: "image/png", width: 160, height: 90, embedded: true }
    ]);
    const event = deck.deck.slides[0]?.timeline?.events[0];
    assert.equal(event?.kind, "keyframes");
    assert.equal(event?.targetId, imageObjects[0]?.id);
    assert.equal(event?.metadata?.nativeBuildFallback, "blur-in");
  });

  test("creates a conservative media object for a unique movie-start 3007 target", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "keymorph-keynote-movie-start-"));
    const keyPath = path.join(dir, "movie.key");
    await mkdir(path.join(keyPath, "Data"), { recursive: true });
    await mkdir(path.join(keyPath, "Index"), { recursive: true });
    await writeFile(path.join(keyPath, "Data", "clip-77.mp4"), new Uint8Array([0, 0, 0, 20, 102, 116, 121, 112]));

    const moviePayload = nativeImagePlacementPayload(0, { x: 95, y: 58, width: 320, height: 180 });
    await writeFile(
      path.join(keyPath, "Index", "Slide-1.iwa"),
      concat([
        iwaArchiveRecord(333, [{ type: 3007, payload: moviePayload }]),
        iwaArchiveRecord(9001, [
          {
            type: 8,
            payload: nativeBuildPayload({ targetId: 333, direction: "In", effect: "apple:movie-start", durationSeconds: 0.5 })
          }
        ]),
        iwaArchiveRecord(9002, [{ type: 153, payload: nativeBuildTimingPayload({ buildId: 9001, durationSeconds: 0.5 }) }])
      ])
    );

    const deck = await parseNativeKeynoteToIr(keyPath);
    assert.equal(validateIR(deck).valid, true);
    const mediaObject = deck.deck.slides[0]?.objects.find((object) => object.type === "media");
    assert.equal(mediaObject?.type, "media");
    assert.equal(mediaObject?.mediaType, "video");
    assert.deepEqual(mediaObject?.bounds, { x: 95, y: 58, width: 320, height: 180 });
    assert.equal(mediaObject?.metadata?.nativeArchiveIdentifier, "333");
    assert.equal(mediaObject?.metadata?.nativeGeometryCandidate?.reason, "typed Keynote media geometry from archive message type 3007");
    const event = deck.deck.slides[0]?.timeline?.events[0];
    assert.equal(event?.kind, "media");
    assert.equal(event?.kind === "media" ? event.action : undefined, "play");
    assert.equal(event?.targetId, mediaObject?.id);
    assert.equal(event?.metadata?.nativeBuildEffect, "apple:movie-start");
  });

  test("recovers native 2011 text drawable build targets and simple motion paths", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "keymorph-keynote-text-drawable-"));
    const keyPath = path.join(dir, "text-drawable.key");
    await mkdir(path.join(keyPath, "Index"), { recursive: true });

    await writeFile(
      path.join(keyPath, "Index", "Slide-1.iwa"),
      concat([
        iwaArchiveRecord(1200, [{ type: 2001, payload: nativeTextContentPayload("不,这些还是你的!*") }]),
        iwaArchiveRecord(1100, [
          {
            type: 2011,
            payload: nativeTextDrawablePayload({
              slideId: 900,
              textId: 1200,
              bounds: { x: 240, y: 180, width: 420, height: 96 }
            })
          }
        ]),
        iwaArchiveRecord(1300, [
          {
            type: 8,
            payload: nativeBuildPayload({
              targetId: 1100,
              direction: "In",
              effect: "apple:bc-appear",
              durationSeconds: 0.5
            })
          }
        ]),
        iwaArchiveRecord(1301, [{ type: 153, payload: nativeBuildTimingPayload({ buildId: 1300, durationSeconds: 0.5 }) }]),
        iwaArchiveRecord(1400, [
          {
            type: 8,
            payload: nativeBuildPayload({
              targetId: 1100,
              direction: "Action",
              effect: "apple:action-motion-path",
              durationSeconds: 1,
              motionPath: [
                { x: 40, y: 0 },
                { x: 50, y: -252 }
              ]
            })
          }
        ]),
        iwaArchiveRecord(1401, [{ type: 153, payload: nativeBuildTimingPayload({ buildId: 1400, durationSeconds: 1 }) }])
      ])
    );

    const deck = await parseNativeKeynoteToIr(keyPath);
    assert.equal(validateIR(deck).valid, true);
    const object = deck.deck.slides[0]?.objects.find((candidate) => candidate.metadata?.nativeExtraction === "typed-iwa-text-drawable");
    assert.equal(object?.type, "text");
    if (object?.type !== "text") throw new Error("Expected native text drawable.");
    assert.equal(object.text.plainText, "不,这些还是你的!");
    assert.deepEqual(object.bounds, { x: 240, y: 180, width: 420, height: 96 });
    assert.equal(object.metadata?.nativeArchiveIdentifier, "1100");
    assert.equal(object.metadata?.nativeTextArchiveIdentifier, "1200");
    assert.deepEqual(object.metadata?.nativeTextDrawableLayout, {
      kind: "text",
      frame: { x: 240, y: 180, width: 420, height: 96 },
      frameFieldPaths: ["1.1.1.1.1", "1.1.1.1.2", "1.3.5.2.1", "1.3.5.2.2"],
      schema: "typed-text-drawable-frame-1.1",
      textArchiveIds: ["1200"],
      slideArchiveId: "900",
      confidence: 0.9
    });
    assert.equal(object.metadata?.nativeTextDrawableSchema, "typed-text-drawable-frame-1.1");
    assert.equal(object.metadata?.nativeTextDrawableParentSlideArchiveId, "900");

    const events = deck.deck.slides[0]?.timeline?.events ?? [];
    const appear = events.find((event) => event.metadata?.nativeBuildId === "1300");
    assert.equal(appear?.kind, "keyframes");
    assert.equal(appear?.targetId, object.id);
    const motion = events.find((event) => event.metadata?.nativeBuildId === "1400");
    assert.equal(motion?.kind, "keyframes");
    assert.equal(motion?.targetId, object.id);
    if (motion?.kind !== "keyframes") throw new Error("Expected motion path keyframes.");
    assert.deepEqual(motion.tracks.find((track) => track.property === "transform.translateX")?.keyframes, [
      { offset: 0, value: 0 },
      { offset: 0.5, value: 40 },
      { offset: 1, value: 50 }
    ]);
    assert.deepEqual(motion.tracks.find((track) => track.property === "transform.translateY")?.keyframes, [
      { offset: 0, value: 0 },
      { offset: 0.5, value: 0 },
      { offset: 1, value: -252 }
    ]);
    assert.deepEqual(motion.metadata?.nativeMotionPathPoints, [
      { x: 0, y: 0 },
      { x: 40, y: 0 },
      { x: 50, y: -252 }
    ]);
    assert.equal(motion.metadata?.nativeMotionPathRelative, true);
    assert.deepEqual(motion.metadata?.nativeMotionPathExtentPoints, [{ x: 50, y: 252 }]);
    assert.deepEqual(motion.metadata?.nativeMotionPathFieldPaths, [
      "4.22",
      "4.22.8.1.1.1.1",
      "4.22.8.1.1.1.2",
      "4.22.8.1.1.2.1",
      "4.22.8.1.1.2.2",
      "4.22.8.1.1.3.1",
      "4.22.8.1.1.3.2",
      "4.22.8.2.1",
      "4.22.8.2.2"
    ]);
    assert.equal(deck.deck.slides[0]?.metadata?.nativeBuildAnimationUnresolvedCount, 0);
  });

  test("preserves character-level dissolve build semantics when degrading to object opacity", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "keymorph-keynote-character-dissolve-"));
    const keyPath = path.join(dir, "character-dissolve.key");
    await mkdir(path.join(keyPath, "Index"), { recursive: true });

    await writeFile(
      path.join(keyPath, "Index", "Slide-1.iwa"),
      concat([
        iwaArchiveRecord(1200, [{ type: 2001, payload: nativeTextContentPayload("逐字溶解") }]),
        iwaArchiveRecord(1100, [
          {
            type: 2011,
            payload: nativeTextDrawablePayload({
              slideId: 900,
              textId: 1200,
              bounds: { x: 240, y: 180, width: 420, height: 96 }
            })
          }
        ]),
        iwaArchiveRecord(1300, [
          {
            type: 8,
            payload: nativeBuildPayload({
              targetId: 1100,
              direction: "In",
              effect: "apple:dissolve character",
              durationSeconds: 0.75
            })
          }
        ]),
        iwaArchiveRecord(1301, [{ type: 153, payload: nativeBuildTimingPayload({ buildId: 1300, durationSeconds: 0.75 }) }])
      ])
    );

    const deck = await parseNativeKeynoteToIr(keyPath);
    assert.equal(validateIR(deck).valid, true);
    const object = deck.deck.slides[0]?.objects.find((candidate) => candidate.metadata?.nativeExtraction === "typed-iwa-text-drawable");
    assert.equal(object?.type, "text");
    const event = deck.deck.slides[0]?.timeline?.events[0];
    assert.equal(event?.kind, "keyframes");
    assert.equal(event?.targetId, object?.id);
    assert.equal(event?.metadata?.nativeBuildFallback, "dissolve-in");
    assert.equal(event?.metadata?.nativeBuildGranularity, "character");
    assert.match(String(event?.metadata?.nativeBuildDegradation), /Per-character dissolve/);
    assert.equal(deck.conversion?.degradedFeatures?.some((feature) => feature.code === "keynote-native-character-build-degraded"), true);
    assert.equal(deck.conversion?.metadata?.recoveredCharacterBuildAnimationCount, 1);
  });

  test("preserves native 3006 typed visual auxiliary records without applying crop", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "keymorph-keynote-typed-visual-3006-"));
    const keyPath = path.join(dir, "typed-visual.key");
    await mkdir(path.join(keyPath, "Index"), { recursive: true });

    await writeFile(
      path.join(keyPath, "Index", "Slide-1.iwa"),
      iwaArchiveRecord(4040267, [
        {
          type: 3006,
          payload: concat([
            nativeImagePlacementPayload(0, { x: 217.714, y: 63.516, width: 378.466, height: 439.634 }),
            protoFixed32(5, 0.5),
            protoVarint(3, 3),
            protoVarint(4, 0),
            protoVarint(12, 0),
            protoVarint(13, 0)
          ]),
          objectReferences: [4040267]
        }
      ])
    );

    const deck = await parseNativeKeynoteToIr(keyPath);
    assert.equal(validateIR(deck).valid, true);
    assert.equal(deck.deck.slides[0]?.metadata?.nativeTypedVisualRecordCount, 1);
    const record = deck.deck.slides[0]?.metadata?.nativeTypedVisualRecords?.[0];
    assert.equal(record?.type, 3006);
    assert.equal(record?.archiveIdentifier, "4040267");
    assert.deepEqual(record?.layout, {
      kind: "image",
      frame: { x: 217.714, y: 63.516, width: 378.466, height: 439.634 },
      frameFieldPaths: ["1.1.1.1", "1.1.1.2", "1.1.2.1", "1.1.2.2"],
      schema: "typed-visual-frame-1.1",
      confidence: 0.97
    });
    assert.deepEqual(record?.geometryCandidates?.[0]?.bounds, { x: 217.714, y: 63.516, width: 378.466, height: 439.634 });
    assert.equal(record?.numericCandidates?.some((candidate) => candidate.fieldNumber === 5 && candidate.value === 0.5), true);
    assert.equal(deck.deck.slides[0]?.objects.some((object) => object.metadata?.nativeArchiveIdentifier === "4040267"), false);
    assert.equal(deck.conversion?.metadata?.totalNativeTypedVisualRecordCount, 1);
  });

  test("starts native Keynote builds together when timing startsWithPrevious is set", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "keymorph-keynote-timing-groups-"));
    const keyPath = path.join(dir, "timing.key");
    await mkdir(path.join(keyPath, "Data"), { recursive: true });
    await mkdir(path.join(keyPath, "Index"), { recursive: true });
    await writeFile(path.join(keyPath, "Data", "hero-42.png"), pngBytes(320, 180));

    await writeFile(
      path.join(keyPath, "Index", "Slide-1.iwa"),
      concat([
        iwaArchiveRecord(111, [
          { type: 3005, payload: nativeImagePlacementPayload(42, { x: 10, y: 20, width: 100, height: 80 }), dataReferences: [42] }
        ]),
        iwaArchiveRecord(222, [
          { type: 3005, payload: nativeImagePlacementPayload(42, { x: 140, y: 20, width: 100, height: 80 }), dataReferences: [42] }
        ]),
        iwaArchiveRecord(9001, [
          { type: 8, payload: nativeBuildPayload({ targetId: 111, direction: "In", effect: "apple:bc-appear", durationSeconds: 1 }) }
        ]),
        iwaArchiveRecord(9002, [
          { type: 153, payload: nativeBuildTimingPayload({ buildId: 9001, durationSeconds: 1, startsWithPrevious: false }) }
        ]),
        iwaArchiveRecord(9003, [
          { type: 8, payload: nativeBuildPayload({ targetId: 222, direction: "In", effect: "apple:bc-appear", durationSeconds: 1 }) }
        ]),
        iwaArchiveRecord(9004, [
          { type: 153, payload: nativeBuildTimingPayload({ buildId: 9003, durationSeconds: 1, startsWithPrevious: true }) }
        ])
      ])
    );

    const deck = await parseNativeKeynoteToIr(keyPath);
    assert.equal(validateIR(deck).valid, true);
    const events = deck.deck.slides[0]?.timeline?.events ?? [];
    assert.equal(events.length, 2);
    assert.deepEqual(
      events.map((event) => event.start),
      [
        { type: "absolute", atMs: 0 },
        { type: "absolute", atMs: 0 }
      ]
    );
    assert.equal(events[1]?.metadata?.nativeBuildStartsWithPrevious, true);
    const edges = deck.deck.slides[0]?.timeline?.dependencyGraph?.edges ?? [];
    assert.deepEqual(
      edges.map((edge) => ({ from: edge.from, to: edge.to, relation: edge.relation, offsetMs: edge.offsetMs })),
      [
        {
          from: events[0]!.id,
          to: events[1]!.id,
          relation: "with",
          offsetMs: undefined
        }
      ]
    );
    assert.deepEqual(createSlideTimingPlan(deck.deck.slides[0]).starts, {
      [events[0]!.id]: 0,
      [events[1]!.id]: 0
    });
    assert.equal(deck.deck.slides[0]?.timeline?.metadata?.nativeBuildTimingDependencyCount, 1);
    assert.equal(deck.deck.slides[0]?.metadata?.nativeBuildTimingDependencyCount, 1);
    assert.equal(deck.deck.slides[0]?.timeline?.durationMs, 2500);
  });

  test("preserves native timing cursor when an unresolved build has timing", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "keymorph-keynote-unresolved-timing-"));
    const keyPath = path.join(dir, "timing.key");
    await mkdir(path.join(keyPath, "Data"), { recursive: true });
    await mkdir(path.join(keyPath, "Index"), { recursive: true });
    await writeFile(path.join(keyPath, "Data", "hero-42.png"), pngBytes(320, 180));

    await writeFile(
      path.join(keyPath, "Index", "Slide-1.iwa"),
      concat([
        iwaArchiveRecord(111, [
          { type: 3005, payload: nativeImagePlacementPayload(42, { x: 10, y: 20, width: 100, height: 80 }), dataReferences: [42] }
        ]),
        iwaArchiveRecord(9001, [
          { type: 8, payload: nativeBuildPayload({ targetId: 999, direction: "In", effect: "apple:bc-appear", durationSeconds: 1 }) }
        ]),
        iwaArchiveRecord(9002, [
          { type: 153, payload: nativeBuildTimingPayload({ buildId: 9001, durationSeconds: 1, afterPrevious: true }) }
        ]),
        iwaArchiveRecord(9003, [
          { type: 8, payload: nativeBuildPayload({ targetId: 111, direction: "In", effect: "apple:bc-appear", durationSeconds: 1 }) }
        ]),
        iwaArchiveRecord(9004, [
          { type: 153, payload: nativeBuildTimingPayload({ buildId: 9003, durationSeconds: 1, afterPrevious: true }) }
        ])
      ])
    );

    const deck = await parseNativeKeynoteToIr(keyPath);
    assert.equal(validateIR(deck).valid, true);
    const events = deck.deck.slides[0]?.timeline?.events ?? [];
    assert.equal(events.length, 1);
    assert.deepEqual(events[0]?.start, { type: "absolute", atMs: 1000 });
    assert.equal(events[0]?.metadata?.nativeBuildStartRelation, "afterPrevious");
    assert.equal(events[0]?.metadata?.nativeBuildTimingRawField6, 1);
    assert.deepEqual(deck.deck.slides[0]?.timeline?.dependencyGraph?.edges ?? [], []);
    assert.equal(deck.deck.slides[0]?.metadata?.nativeBuildAnimationUnresolvedCount, 1);
  });

  test("derives native timing start relation from raw type 153 flags", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "keymorph-keynote-timing-relation-"));
    const keyPath = path.join(dir, "timing.key");
    await mkdir(path.join(keyPath, "Data"), { recursive: true });
    await mkdir(path.join(keyPath, "Index"), { recursive: true });
    await writeFile(path.join(keyPath, "Data", "hero-42.png"), pngBytes(320, 180));

    await writeFile(
      path.join(keyPath, "Index", "Slide-1.iwa"),
      concat([
        iwaArchiveRecord(111, [
          { type: 3005, payload: nativeImagePlacementPayload(42, { x: 10, y: 20, width: 100, height: 80 }), dataReferences: [42] }
        ]),
        iwaArchiveRecord(9001, [
          { type: 8, payload: nativeBuildPayload({ targetId: 111, direction: "In", effect: "apple:bc-appear", durationSeconds: 1 }) }
        ]),
        iwaArchiveRecord(9002, [
          { type: 153, payload: nativeBuildTimingPayload({ buildId: 9001, durationSeconds: 1, startsWithPrevious: false, afterPrevious: true }) }
        ]),
        iwaArchiveRecord(9003, [
          { type: 8, payload: nativeBuildPayload({ targetId: 111, direction: "In", effect: "apple:bc-appear", durationSeconds: 1 }) }
        ]),
        iwaArchiveRecord(9004, [
          { type: 153, payload: nativeBuildTimingPayload({ buildId: 9003, durationSeconds: 1, startsWithPrevious: true, afterPrevious: true }) }
        ]),
        iwaArchiveRecord(9005, [
          { type: 8, payload: nativeBuildPayload({ targetId: 111, direction: "In", effect: "apple:bc-appear", durationSeconds: 1 }) }
        ]),
        iwaArchiveRecord(9006, [
          { type: 153, payload: nativeBuildTimingPayload({ buildId: 9005, durationSeconds: 1, startsWithPrevious: true, afterPrevious: false }) }
        ]),
        iwaArchiveRecord(9007, [
          { type: 8, payload: nativeBuildPayload({ targetId: 111, direction: "In", effect: "apple:bc-appear", durationSeconds: 1 }) }
        ]),
        iwaArchiveRecord(9008, [
          { type: 153, payload: nativeBuildTimingPayload({ buildId: 9007, durationSeconds: 1, startsWithPrevious: false, afterPrevious: true }) }
        ])
      ])
    );

    const deck = await parseNativeKeynoteToIr(keyPath);
    assert.equal(validateIR(deck).valid, true);
    const events = deck.deck.slides[0]?.timeline?.events ?? [];
    assert.deepEqual(
      events.map((event) => event.start),
      [
        { type: "absolute", atMs: 0 },
        { type: "absolute", atMs: 0 },
        { type: "absolute", atMs: 0 },
        { type: "absolute", atMs: 1000 }
      ]
    );
    assert.deepEqual(
      events.map((event) => event.metadata?.nativeBuildStartRelation),
      ["afterPrevious", "withPrevious", "withPrevious", "afterPrevious"]
    );
    assert.equal(events[1]?.metadata?.nativeBuildStartsWithPrevious, true);
    assert.equal(events[1]?.metadata?.nativeBuildAfterPrevious, false);
    assert.equal(events[1]?.metadata?.nativeBuildTimingRawField5, 1);
    assert.equal(events[1]?.metadata?.nativeBuildTimingRawField6, 1);
    const edges = deck.deck.slides[0]?.timeline?.dependencyGraph?.edges ?? [];
    assert.deepEqual(
      edges.map((edge) => ({ from: edge.from, to: edge.to, relation: edge.relation, offsetMs: edge.offsetMs })),
      [
        { from: events[0]!.id, to: events[1]!.id, relation: "with", offsetMs: undefined },
        { from: events[0]!.id, to: events[2]!.id, relation: "with", offsetMs: undefined },
        { from: events[0]!.id, to: events[3]!.id, relation: "after", offsetMs: undefined }
      ]
    );
    assert.deepEqual(createSlideTimingPlan(deck.deck.slides[0]).starts, {
      [events[0]!.id]: 0,
      [events[1]!.id]: 0,
      [events[2]!.id]: 0,
      [events[3]!.id]: 1000
    });
    assert.equal(deck.deck.slides[0]?.timeline?.metadata?.nativeBuildTimingDependencyCount, 3);
  });

  test("maps native timing trigger groups into IR custom triggers", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "keymorph-keynote-trigger-group-"));
    const keyPath = path.join(dir, "trigger.key");
    await mkdir(path.join(keyPath, "Data"), { recursive: true });
    await mkdir(path.join(keyPath, "Index"), { recursive: true });
    await writeFile(path.join(keyPath, "Data", "hero-42.png"), pngBytes(320, 180));

    await writeFile(
      path.join(keyPath, "Index", "Slide-1.iwa"),
      concat([
        iwaArchiveRecord(111, [
          { type: 3005, payload: nativeImagePlacementPayload(42, { x: 10, y: 20, width: 100, height: 80 }), dataReferences: [42] }
        ]),
        iwaArchiveRecord(9001, [
          { type: 8, payload: nativeBuildPayload({ targetId: 111, direction: "In", effect: "apple:bc-appear", durationSeconds: 1 }) }
        ]),
        iwaArchiveRecord(9002, [
          {
            type: 153,
            payload: nativeBuildTimingPayload({ buildId: 9001, durationSeconds: 1, delaySeconds: 0.25, triggerGroup: 42 })
          }
        ])
      ])
    );

    const deck = await parseNativeKeynoteToIr(keyPath);
    assert.equal(validateIR(deck).valid, true);
    const slide = deck.deck.slides[0];
    const event = slide?.timeline?.events[0];

    assert.deepEqual(slide?.timeline?.triggers, [
      {
        id: "native-build-trigger-42",
        type: "custom",
        name: "Keynote build trigger group 42",
        metadata: {
          nativeSource: "keynote-iwa-build-timing",
          nativeBuildTimingTriggerGroupRaw: 42,
          nativeBuildTimingId: "9002"
        }
      }
    ]);
    assert.deepEqual(event?.start, { type: "trigger", triggerId: "native-build-trigger-42", offsetMs: 250 });
    assert.equal(event?.metadata?.nativeBuildTriggerId, "native-build-trigger-42");
    assert.equal(event?.metadata?.nativeBuildRecoveredAbsoluteStartMs, 250);
    assert.equal(slide?.timeline?.metadata?.nativeBuildTimingTriggerCount, 1);
    assert.deepEqual(createSlideTimingPlan(slide).starts, { [event!.id]: 250 });
  });

  test("links multi-target native builds as same-time timing dependencies", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "keymorph-keynote-multi-target-timing-"));
    const keyPath = path.join(dir, "timing.key");
    await mkdir(path.join(keyPath, "Data"), { recursive: true });
    await mkdir(path.join(keyPath, "Index"), { recursive: true });
    await writeFile(path.join(keyPath, "Data", "hero-42.png"), pngBytes(320, 180));

    await writeFile(
      path.join(keyPath, "Index", "Slide-1.iwa"),
      concat([
        iwaArchiveRecord(111, [
          {
            type: 3005,
            payload: nativeImagePlacementPayload(42, { x: 10, y: 20, width: 100, height: 80 }),
            dataReferences: [42],
            objectReferences: [777]
          }
        ]),
        iwaArchiveRecord(222, [
          {
            type: 3005,
            payload: nativeImagePlacementPayload(42, { x: 140, y: 20, width: 100, height: 80 }),
            dataReferences: [42],
            objectReferences: [777]
          }
        ]),
        iwaArchiveRecord(9001, [
          {
            type: 8,
            payload: nativeBuildPayload({ targetId: 777, direction: "In", effect: "apple:bc-appear", durationSeconds: 1 })
          }
        ]),
        iwaArchiveRecord(9002, [
          { type: 153, payload: nativeBuildTimingPayload({ buildId: 9001, durationSeconds: 1, startsWithPrevious: false }) }
        ])
      ])
    );

    const deck = await parseNativeKeynoteToIr(keyPath);
    assert.equal(validateIR(deck).valid, true);
    const slide = deck.deck.slides[0];
    const events = slide?.timeline?.events ?? [];
    assert.equal(events.length, 2);
    assert.deepEqual(
      events.map((event) => event.start),
      [
        { type: "absolute", atMs: 0 },
        { type: "absolute", atMs: 0 }
      ]
    );
    assert.deepEqual(slide?.timeline?.dependencyGraph?.edges, [
      {
        id: `native-timing-${events[0]!.id}-${events[1]!.id}-with`,
        from: events[0]!.id,
        to: events[1]!.id,
        relation: "with"
      }
    ]);
    assert.deepEqual(createSlideTimingPlan(slide).starts, {
      [events[0]!.id]: 0,
      [events[1]!.id]: 0
    });
  });

  test("approximates native Keynote wipe as a bounds reveal", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "keymorph-keynote-wipe-"));
    const keyPath = path.join(dir, "wipe.key");
    await mkdir(path.join(keyPath, "Data"), { recursive: true });
    await mkdir(path.join(keyPath, "Index"), { recursive: true });
    await writeFile(path.join(keyPath, "Data", "hero-42.png"), pngBytes(320, 180));

    await writeFile(
      path.join(keyPath, "Index", "Slide-1.iwa"),
      concat([
        iwaArchiveRecord(111, [
          { type: 3005, payload: nativeImagePlacementPayload(42, { x: 10, y: 20, width: 100, height: 80 }), dataReferences: [42] }
        ]),
        iwaArchiveRecord(9001, [
          { type: 8, payload: nativeBuildPayload({ targetId: 111, direction: "In", effect: "apple:wipe", durationSeconds: 1 }) }
        ]),
        iwaArchiveRecord(9002, [{ type: 153, payload: nativeBuildTimingPayload({ buildId: 9001, durationSeconds: 1 }) }])
      ])
    );

    const deck = await parseNativeKeynoteToIr(keyPath);
    assert.equal(validateIR(deck).valid, true);
    const event = deck.deck.slides[0]?.timeline?.events[0];
    assert.equal(event?.kind, "keyframes");
    if (event?.kind !== "keyframes") throw new Error("Expected wipe keyframes.");
    assert.equal(event.metadata?.nativeBuildFallback, "wipe-in");
    assert.equal(typeof event.metadata?.nativeBuildDegradation, "string");
    assert.deepEqual(event.tracks.find((track) => track.property === "bounds")?.keyframes, [
      { offset: 0, value: { x: 10, y: 20, width: 1, height: 80 } },
      { offset: 1, value: { x: 10, y: 20, width: 100, height: 80 } }
    ]);
  });

  test("maps native Keynote anvil and crumble to object-level motion fallbacks", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "keymorph-keynote-physical-builds-"));
    const keyPath = path.join(dir, "physical.key");
    await mkdir(path.join(keyPath, "Data"), { recursive: true });
    await mkdir(path.join(keyPath, "Index"), { recursive: true });
    await writeFile(path.join(keyPath, "Data", "hero-42.png"), pngBytes(320, 180));

    await writeFile(
      path.join(keyPath, "Index", "Slide-1.iwa"),
      concat([
        iwaArchiveRecord(111, [
          { type: 3005, payload: nativeImagePlacementPayload(42, { x: 10, y: 20, width: 100, height: 80 }), dataReferences: [42] }
        ]),
        iwaArchiveRecord(9001, [
          { type: 8, payload: nativeBuildPayload({ targetId: 111, direction: "In", effect: "com.apple.iWork.Keynote.BUKAnvil", durationSeconds: 1 }) }
        ]),
        iwaArchiveRecord(9002, [{ type: 153, payload: nativeBuildTimingPayload({ buildId: 9001, durationSeconds: 1 }) }]),
        iwaArchiveRecord(9003, [
          { type: 8, payload: nativeBuildPayload({ targetId: 111, direction: "Out", effect: "com.apple.iWork.Keynote.Crumble", durationSeconds: 1 }) }
        ]),
        iwaArchiveRecord(9004, [{ type: 153, payload: nativeBuildTimingPayload({ buildId: 9003, durationSeconds: 1 }) }])
      ])
    );

    const deck = await parseNativeKeynoteToIr(keyPath);
    assert.equal(validateIR(deck).valid, true);
    const events = deck.deck.slides[0]?.timeline?.events ?? [];
    assert.equal(events.length, 2);
    assert.equal(events[0]?.metadata?.nativeBuildFallback, "anvil-in");
    assert.equal(events[1]?.metadata?.nativeBuildFallback, "crumble-out");
    assert.equal(typeof events[0]?.metadata?.nativeBuildDegradation, "string");
    assert.equal(typeof events[1]?.metadata?.nativeBuildDegradation, "string");
    assert.deepEqual(events[0]?.kind === "keyframes" ? events[0].tracks.find((track) => track.property === "transform.translateY")?.keyframes : undefined, [
      { offset: 0, value: -48 },
      { offset: 1, value: 0 }
    ]);
    assert.deepEqual(events[1]?.kind === "keyframes" ? events[1].tracks.find((track) => track.property === "transform.rotateDeg")?.keyframes : undefined, [
      { offset: 0, value: 0 },
      { offset: 1, value: 8 }
    ]);
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

function protoFixed64(fieldNumber: number, value: number): Uint8Array {
  const out = new Uint8Array(8);
  new DataView(out.buffer).setFloat64(0, value, true);
  return concat([varint((fieldNumber << 3) | 1), out]);
}

function nativeImagePlacementPayload(
  dataId: number,
  bounds: { x: number; y: number; width: number; height: number }
): Uint8Array {
  return concat([
    protoBytes(
      concat([
        protoBytes(
          concat([
            protoBytes(concat([protoFixed32(1, bounds.x), protoFixed32(2, bounds.y)]), 1),
            protoBytes(concat([protoFixed32(1, bounds.width), protoFixed32(2, bounds.height)]), 2)
          ]),
          1
        )
      ]),
      1
    ),
    ...(dataId > 0 ? [protoBytes(protoVarint(1, dataId), 11)] : [])
  ]);
}

function nativeDocumentSizePayload(width: number, height: number): Uint8Array {
  return concat([protoFixed32(1, width), protoFixed32(2, height)]);
}

function nativeBuildPayload(values: {
  targetId: number;
  direction: string;
  effect: string;
  durationSeconds: number;
  delaySeconds?: number;
  motionPath?: { x: number; y: number } | Array<{ x: number; y: number }>;
}): Uint8Array {
  return concat([
    protoBytes(protoVarint(1, values.targetId), 1),
    protoString("All at Once", 2),
    protoFixed64(3, values.delaySeconds ?? 0),
    protoBytes(
      concat([
        protoFixed64(17, 60),
        protoBytes(
          concat([
            protoString(values.direction, 1),
            protoString(values.effect, 2),
            protoFixed64(3, values.durationSeconds),
            protoFixed64(5, values.delaySeconds ?? 0)
          ]),
          18
        ),
        ...(values.motionPath ? [protoBytes(nativeMotionPathPayload(values.motionPath), 22)] : [])
      ]),
      4
    )
  ]);
}

function nativeTextContentPayload(text: string): Uint8Array {
  return protoString(text, 3);
}

function nativeTextDrawablePayload(values: {
  slideId: number;
  textId: number;
  bounds: { x: number; y: number; width: number; height: number };
}): Uint8Array {
  return concat([
    protoBytes(
      concat([
        protoBytes(
          concat([
            protoBytes(
              concat([
                protoBytes(concat([protoFixed32(1, values.bounds.x), protoFixed32(2, values.bounds.y)]), 1)
              ]),
              1
            ),
            protoBytes(protoVarint(1, values.slideId), 2)
          ]),
          1
        ),
        protoBytes(
          protoBytes(
            protoBytes(concat([protoFixed32(1, values.bounds.width), protoFixed32(2, values.bounds.height)]), 2),
            5
          ),
          3
        )
      ]),
      1
    ),
    protoBytes(protoVarint(1, values.textId), 2),
    protoBytes(protoVarint(1, values.textId), 4)
  ]);
}

function nativeMotionPathPayload(pointOrPoints: { x: number; y: number } | Array<{ x: number; y: number }>): Uint8Array {
  const points = Array.isArray(pointOrPoints) ? pointOrPoints : [pointOrPoints];
  const last = points[points.length - 1] ?? { x: 0, y: 0 };
  const zero = nativeMotionPathPointCluster({ x: 0, y: 0 });
  const pathPoints = points.map(nativeMotionPathPointCluster);
  const extent = nativeMotionPathPoint({ x: Math.abs(last.x), y: Math.abs(last.y) });
  return protoBytes(concat([protoBytes(concat([zero, ...pathPoints]), 1), protoBytes(extent, 2)]), 8);
}

function nativeMotionPathPointCluster(point: { x: number; y: number }): Uint8Array {
  const control = nativeMotionPathPoint(point);
  return protoBytes(concat([protoBytes(control, 1), protoBytes(control, 2), protoBytes(control, 3)]), 1);
}

function nativeMotionPathPoint(point: { x: number; y: number }): Uint8Array {
  return concat([protoFixed32(1, point.x), protoFixed32(2, point.y)]);
}

function nativeBuildTimingPayload(values: {
  buildId: number;
  durationSeconds: number;
  delaySeconds?: number;
  startsWithPrevious?: boolean;
  afterPrevious?: boolean;
  triggerGroup?: number;
}): Uint8Array {
  return concat([
    protoBytes(protoVarint(1, values.buildId), 1),
    protoFixed64(3, values.delaySeconds ?? 0),
    protoFixed64(4, values.durationSeconds),
    protoVarint(5, values.startsWithPrevious ? 1 : 0),
    protoVarint(6, values.afterPrevious === false ? 0 : 1),
    protoBytes(protoVarint(2, values.triggerGroup ?? 1), 7)
  ]);
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
