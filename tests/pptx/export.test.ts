import { Buffer } from "node:buffer";
import { mkdtemp, readFile, stat, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { deflateRawSync, inflateRawSync } from "node:zlib";
import { describe, test } from "node:test";
import assert from "node:assert/strict";

import { createDemoDeck } from "../../src/demo/createDemoDeck.ts";
import { exportIrToPptx, parsePptxToIr } from "../../src/pptx/index.ts";
import { IR_VERSION, type DeckIR } from "../../src/ir/index.ts";

describe("PPTX export", () => {
  test("writes a PPTX file for the demo deck", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "keymorph-pptx-test-"));
    const out = path.join(dir, "deck.pptx");

    await exportIrToPptx(createDemoDeck(), out);

    assert.ok((await stat(out)).size > 0);
  });

  test("parses static slides from a KeyMorph-exported PPTX", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "keymorph-pptx-test-"));
    const out = path.join(dir, "deck.pptx");

    await exportIrToPptx(createDemoDeck(), out);
    const parsed = await parsePptxToIr(out);

    assert.equal(parsed.deck.slides.length, 2);
    assert.equal(parsed.conversion?.status, "partial");
    assert.ok(parsed.deck.slides[0].objects.some((object) => object.type === "text"));
    assert.match(JSON.stringify(parsed), /KeyMorph/);
  });

  test("exports image objects as embedded PPTX media", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "keymorph-pptx-image-test-"));
    const out = path.join(dir, "image.pptx");
    const deck = createImageDeck();

    await exportIrToPptx(deck, out);
    const entries = readZipEntries(await readFile(out));
    const slideXml = new TextDecoder().decode(entries.get("ppt/slides/slide1.xml"));
    const relsXml = new TextDecoder().decode(entries.get("ppt/slides/_rels/slide1.xml.rels"));

    assert.ok(entries.has("ppt/media/image1.png"));
    assert.match(slideXml, /<p:pic>/);
    assert.match(slideXml, /<a:blip r:embed="rId2"/);
    assert.match(relsXml, /Target="\.\.\/media\/image1\.png"/);

    const parsed = await parsePptxToIr(out);
    assert.equal(parsed.deck.slides[0].objects.some((object) => object.type === "image"), true);
  });

  test("strips XML-invalid control characters from exported text", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "keymorph-pptx-xml-safety-"));
    const out = path.join(dir, "safe-text.pptx");
    const deck = createDemoDeck();
    const textObject = deck.deck.slides[0].objects.find((object) => object.type === "text");
    if (!textObject || textObject.type !== "text") throw new Error("Expected demo text object.");
    textObject.name = "Unsafe\u0012Name";
    textObject.text.plainText = "XBO Transition\u0012$a";
    textObject.text.runs = [{ text: "XBO Transition\u0012$a", style: { fontFamily: "Arial", fontSize: 28, color: "#111827" } }];

    await exportIrToPptx(deck, out);
    const slideXml = await readPptxXml(out, "ppt/slides/slide1.xml");

    assert.equal(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/.test(slideXml), false);
    assert.match(slideXml, /XBO Transition\$a/);
    assert.doesNotMatch(slideXml, /Unsafe\u0012Name/);
  });

  test("round-trips simple keyframe and visibility timing through PPTX XML", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "keymorph-pptx-test-"));
    const out = path.join(dir, "animated.pptx");

    await exportIrToPptx(createAnimatedDeck(), out);
    const parsed = await parsePptxToIr(out);
    const events = parsed.deck.slides[0].timeline?.events ?? [];

    assert.equal(parsed.conversion?.statistics?.animationCount, 5);
    assert.equal(parsed.conversion?.statistics?.unsupportedFeatureCount, 0);
    assert.ok(events.some((event) => event.kind === "visibility" && event.visible === false && event.start?.type === "absolute" && event.start.atMs === 1100));

    const fade = events.find((event) => event.kind === "keyframes" && event.tracks.some((track) => track.property === "opacity"));
    assert.equal(fade?.kind, "keyframes");
    if (fade?.kind !== "keyframes") throw new Error("Expected imported fade keyframes.");
    assert.equal(fade.durationMs, 700);
    assert.deepEqual(fade.tracks.find((track) => track.property === "opacity")?.keyframes, [
      { offset: 0, value: 0 },
      { offset: 1, value: 1 }
    ]);

    const motion = events.find(
      (event) => event.kind === "keyframes" && event.tracks.some((track) => track.property === "transform.translateX")
    );
    assert.equal(motion?.kind, "keyframes");
    if (motion?.kind !== "keyframes") throw new Error("Expected imported motion keyframes.");
    assert.equal(motion.durationMs, 600);
    assert.deepEqual(motion.tracks.find((track) => track.property === "transform.translateX")?.keyframes, [
      { offset: 0, value: 0 },
      { offset: 1, value: 120 }
    ]);
    assert.deepEqual(motion.tracks.find((track) => track.property === "transform.translateY")?.keyframes, [
      { offset: 0, value: 0 },
      { offset: 1, value: 45 }
    ]);

    const scale = events.find(
      (event) => event.kind === "keyframes" && event.tracks.some((track) => track.property === "transform.scaleX")
    );
    assert.equal(scale?.kind, "keyframes");
    if (scale?.kind !== "keyframes") throw new Error("Expected imported scale keyframes.");
    assert.deepEqual(scale.tracks.find((track) => track.property === "transform.scaleX")?.keyframes, [
      { offset: 0, value: 1 },
      { offset: 1, value: 1.25 }
    ]);
    assert.deepEqual(scale.tracks.find((track) => track.property === "transform.scaleY")?.keyframes, [
      { offset: 0, value: 1 },
      { offset: 1, value: 0.75 }
    ]);

    const rotation = events.find(
      (event) => event.kind === "keyframes" && event.tracks.some((track) => track.property === "transform.rotateDeg")
    );
    assert.equal(rotation?.kind, "keyframes");
    if (rotation?.kind !== "keyframes") throw new Error("Expected imported rotation keyframes.");
    assert.deepEqual(rotation.tracks.find((track) => track.property === "transform.rotateDeg")?.keyframes, [
      { offset: 0, value: 0 },
      { offset: 1, value: 45 }
    ]);
  });

  test("round-trips click and previous-event timing dependencies through PPTX XML", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "keymorph-pptx-test-"));
    const out = path.join(dir, "sequenced.pptx");

    await exportIrToPptx(createSequencedDeck(), out);
    const parsed = await parsePptxToIr(out);
    const events = parsed.deck.slides[0].timeline?.events ?? [];
    const triggers = parsed.deck.slides[0].timeline?.triggers ?? [];
    const edges = parsed.deck.slides[0].timeline?.dependencyGraph?.edges ?? [];

    assert.equal(events.length, 3);
    assert.equal(events[0]?.start?.type, "trigger");
    assert.equal(events[1]?.start?.type, "withPrevious");
    assert.equal(events[2]?.start?.type, "afterPrevious");
    assert.ok(triggers.some((trigger) => trigger.type === "onClick"));
    assert.ok(edges.some((edge) => edge.relation === "triggers" && edge.to === events[0]?.id));
    assert.ok(edges.some((edge) => edge.relation === "with" && edge.to === events[1]?.id));
    assert.ok(edges.some((edge) => edge.relation === "after" && edge.to === events[2]?.id));
  });

  test("exports grouped object builds with stable target ordering", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "keymorph-pptx-test-"));
    const out = path.join(dir, "group-build.pptx");

    await exportIrToPptx(createGroupedBuildDeck(), out);
    const slideXml = await readPptxXml(out, "ppt/slides/slide1.xml");
    const shapeIds = Array.from(slideXml.matchAll(/<p:cNvPr id="([^"]+)" name="(Group child [AB])"/g)).map((match) => ({
      id: match[1],
      name: match[2]
    }));
    const buildIds = Array.from(slideXml.matchAll(/<p:bldP spid="([^"]+)"/g)).map((match) => match[1]);

    assert.deepEqual(shapeIds.map((shape) => shape.name), ["Group child A", "Group child B"]);
    assert.deepEqual(buildIds, shapeIds.map((shape) => shape.id));

    const parsed = await parsePptxToIr(out);
    const events = parsed.deck.slides[0].timeline?.events ?? [];
    const edges = parsed.deck.slides[0].timeline?.dependencyGraph?.edges ?? [];
    assert.equal(events.length, 2);
    assert.ok(edges.some((edge) => edge.relation === "after" && edge.from === events[0]?.id && edge.to === events[1]?.id));
    assert.ok(edges.every((edge) => edge.from !== edge.to));
    assert.ok(parsed.conversion?.degradedFeatures.some((feature) => feature.code === "presentationml-build-list"));
  });

  test("reports unsupported timing condition semantics while importing supported effects", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "keymorph-pptx-test-"));
    const out = path.join(dir, "unsupported-condition.pptx");

    await exportIrToPptx(createAnimatedDeck(), out);
    await patchPptxXml(out, "ppt/slides/slide1.xml", (source) =>
      source.replace(
        "</p:cTn>\n          </p:seq>",
        "<p:nextCondLst><p:cond evt=\"onNext\" delay=\"0\"/></p:nextCondLst></p:cTn>\n          </p:seq>"
      )
    );

    const parsed = await parsePptxToIr(out);
    assert.equal(parsed.deck.slides[0].timeline?.events.length, 5);
    assert.ok(parsed.conversion?.degradedFeatures.some((feature) => feature.code === "presentationml-next-condition"));
    assert.ok(parsed.conversion?.degradedFeatures.some((feature) => feature.description.includes('evt="onNext"')));
  });

  test("imports numeric p:anim opacity and preserves repeat/autoreverse/fill metadata", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "keymorph-pptx-test-"));
    const out = path.join(dir, "property-anim.pptx");

    await exportIrToPptx(createAnimatedDeck(), out);
    await patchPptxXml(out, "ppt/slides/slide1.xml", (source) =>
      source.replace(
        /<p:animEffect transition="in" filter="fade">[\s\S]*?<\/p:animEffect>/,
        `<p:anim calcmode="lin" valueType="num">
  <p:cBhvr>
    <p:cTn id="4" dur="800" fill="remove" repeatCount="2" autoRev="1">
      <p:stCondLst><p:cond delay="300"/></p:stCondLst>
    </p:cTn>
    <p:tgtEl><p:spTgt spid="2"/></p:tgtEl>
    <p:attrNameLst><p:attrName>style.opacity</p:attrName></p:attrNameLst>
  </p:cBhvr>
  <p:tavLst>
    <p:tav tm="0"><p:val><p:fltVal val="0"/></p:val></p:tav>
    <p:tav tm="50000"><p:val><p:fltVal val="50000"/></p:val></p:tav>
    <p:tav tm="100000"><p:val><p:fltVal val="100000"/></p:val></p:tav>
  </p:tavLst>
</p:anim>`
      )
    );

    const parsed = await parsePptxToIr(out);
    const opacity = (parsed.deck.slides[0].timeline?.events ?? []).find(
      (event) => event.kind === "keyframes" && event.metadata?.pptxTag === "anim"
    );

    assert.equal(opacity?.kind, "keyframes");
    if (opacity?.kind !== "keyframes") throw new Error("Expected imported p:anim opacity keyframes.");
    assert.equal(opacity.durationMs, 800);
    assert.equal(opacity.fill, "none");
    assert.equal(opacity.metadata?.pptxRepeatCount, "2");
    assert.equal(opacity.metadata?.pptxAutoReverse, true);
    assert.deepEqual(opacity.tracks.find((track) => track.property === "opacity")?.keyframes, [
      { offset: 0, value: 0 },
      { offset: 0.5, value: 0.5 },
      { offset: 1, value: 1 }
    ]);
    assert.ok(parsed.conversion?.degradedFeatures.some((feature) => feature.code === "presentationml-repeat-behavior"));
    assert.ok(parsed.conversion?.degradedFeatures.some((feature) => feature.code === "presentationml-autoreverse-behavior"));
  });

  test("reports unsupported command and effect metadata without dropping supported effects", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "keymorph-pptx-test-"));
    const out = path.join(dir, "unsupported-metadata.pptx");

    await exportIrToPptx(createAnimatedDeck(), out);
    await patchPptxXml(out, "ppt/slides/slide1.xml", (source) =>
      source.replace(
        "</p:cTn>\n          </p:seq>",
        `<p:par>
  <p:cTn id="99" fill="hold">
    <p:stCondLst><p:cond delay="0"/></p:stCondLst>
    <p:childTnLst>
      <p:animEffect transition="in" filter="checkerboard(across)" presetClass="entr" presetID="10" presetSubtype="5">
        <p:cBhvr>
          <p:cTn id="100" dur="400" fill="hold"/>
          <p:tgtEl><p:spTgt spid="2"/></p:tgtEl>
        </p:cBhvr>
      </p:animEffect>
      <p:cmd type="call" cmd="playFrom(0.0)">
        <p:cBhvr>
          <p:cTn id="101" dur="1"/>
          <p:tgtEl><p:spTgt spid="2"/></p:tgtEl>
        </p:cBhvr>
      </p:cmd>
    </p:childTnLst>
  </p:cTn>
</p:par></p:cTn>\n          </p:seq>`
      )
    );

    const parsed = await parsePptxToIr(out);
    const unsupported = parsed.conversion?.unsupportedFeatures ?? [];
    assert.equal(parsed.deck.slides[0].timeline?.events.length, 5);
    assert.ok(
      unsupported.some(
        (feature) =>
          feature.code === "presentationml-animation-effect" &&
          feature.description.includes('filter="checkerboard(across)"') &&
          feature.description.includes('presetID="10"')
      )
    );
    assert.ok(
      unsupported.some(
        (feature) =>
          feature.code === "presentationml-cmd" &&
          feature.description.includes('type "call"') &&
          feature.description.includes("playFrom(0.0)")
      )
    );
  });
});

function createAnimatedDeck(): DeckIR {
  return {
    irVersion: IR_VERSION,
    metadata: { title: "Animated PPTX fixture", sourceApplication: "KeyMorph tests" },
    deck: {
      id: "animated-fixture",
      title: "Animated PPTX fixture",
      size: { width: 1280, height: 720, unit: "px" },
      slides: [
        {
          id: "slide-1",
          index: 0,
          name: "Animations",
          background: { type: "solid", color: "#ffffff" },
          objects: [
            {
              id: "headline",
              type: "text",
              name: "Headline",
              bounds: { x: 96, y: 96, width: 500, height: 96 },
              opacity: 1,
              text: {
                plainText: "Animated headline",
                runs: [{ text: "Animated headline", style: { fontFamily: "Arial", fontSize: 36, color: "#111827" } }]
              }
            },
            {
              id: "badge",
              type: "shape",
              name: "Badge",
              shape: "roundRect",
              bounds: { x: 120, y: 260, width: 180, height: 72 },
              opacity: 1,
              style: {
                fill: { type: "solid", color: "#0f766e" },
                stroke: { color: "#0f766e", width: 0 }
              }
            }
          ],
          timeline: {
            durationMs: 1800,
            events: [
              {
                id: "headline-fade",
                kind: "keyframes",
                targetId: "headline",
                start: { type: "absolute", atMs: 250 },
                durationMs: 700,
                fill: "both",
                tracks: [
                  {
                    property: "opacity",
                    keyframes: [
                      { offset: 0, value: 0 },
                      { offset: 1, value: 1 }
                    ]
                  }
                ]
              },
              {
                id: "headline-hide",
                kind: "visibility",
                targetId: "headline",
                start: { type: "absolute", atMs: 1100 },
                durationMs: 0,
                fill: "both",
                visible: false
              },
              {
                id: "badge-slide",
                kind: "keyframes",
                targetId: "badge",
                start: { type: "absolute", atMs: 500 },
                durationMs: 600,
                fill: "both",
                tracks: [
                  {
                    property: "transform.translateX",
                    keyframes: [
                      { offset: 0, value: 0 },
                      { offset: 1, value: 120 }
                    ]
                  },
                  {
                    property: "transform.translateY",
                    keyframes: [
                      { offset: 0, value: 0 },
                      { offset: 1, value: 45 }
                    ]
                  }
                ]
              },
              {
                id: "badge-scale",
                kind: "keyframes",
                targetId: "badge",
                start: { type: "absolute", atMs: 950 },
                durationMs: 400,
                fill: "both",
                tracks: [
                  {
                    property: "transform.scaleX",
                    keyframes: [
                      { offset: 0, value: 1 },
                      { offset: 1, value: 1.25 }
                    ]
                  },
                  {
                    property: "transform.scaleY",
                    keyframes: [
                      { offset: 0, value: 1 },
                      { offset: 1, value: 0.75 }
                    ]
                  }
                ]
              },
              {
                id: "badge-rotate",
                kind: "keyframes",
                targetId: "badge",
                start: { type: "absolute", atMs: 1250 },
                durationMs: 300,
                fill: "both",
                tracks: [
                  {
                    property: "transform.rotateDeg",
                    keyframes: [
                      { offset: 0, value: 0 },
                      { offset: 1, value: 45 }
                    ]
                  }
                ]
              }
            ],
            dependencyGraph: { edges: [] }
          }
        }
      ]
    },
    conversion: { status: "success", messages: [] }
  };
}

function createImageDeck(): DeckIR {
  const imageDataUri = `data:image/png;base64,${Buffer.from(pngBytes(4, 3)).toString("base64")}`;
  return {
    irVersion: IR_VERSION,
    metadata: { title: "Image PPTX fixture", sourceApplication: "KeyMorph tests" },
    deck: {
      id: "image-fixture",
      title: "Image PPTX fixture",
      size: { width: 1280, height: 720, unit: "px" },
      assets: [
        {
          id: "asset-image",
          kind: "image",
          uri: imageDataUri,
          mimeType: "image/png",
          name: "fixture.png",
          width: 4,
          height: 3
        }
      ],
      slides: [
        {
          id: "slide-1",
          index: 0,
          name: "Image",
          background: { type: "solid", color: "#ffffff" },
          objects: [
            {
              id: "picture",
              type: "image",
              name: "Picture",
              bounds: { x: 100, y: 80, width: 640, height: 360 },
              opacity: 1,
              source: { assetId: "asset-image" }
            }
          ],
          timeline: { durationMs: 1000, events: [], dependencyGraph: { edges: [] } }
        }
      ]
    },
    conversion: { status: "success", messages: [] }
  };
}

function createSequencedDeck(): DeckIR {
  const deck = createAnimatedDeck();
  const slide = deck.deck.slides[0];
  slide.timeline = {
    durationMs: 2200,
    triggers: [{ id: "start-click", type: "onClick", targetId: "headline", clickIndex: 1 }],
    events: [
      {
        id: "headline-click-fade",
        kind: "keyframes",
        targetId: "headline",
        start: { type: "trigger", triggerId: "start-click" },
        durationMs: 500,
        fill: "both",
        tracks: [
          {
            property: "opacity",
            keyframes: [
              { offset: 0, value: 0 },
              { offset: 1, value: 1 }
            ]
          }
        ]
      },
      {
        id: "badge-with-previous",
        kind: "keyframes",
        targetId: "badge",
        start: { type: "withPrevious", offsetMs: 125 },
        durationMs: 450,
        fill: "both",
        tracks: [
          {
            property: "transform.translateX",
            keyframes: [
              { offset: 0, value: 0 },
              { offset: 1, value: 80 }
            ]
          }
        ]
      },
      {
        id: "badge-after-previous",
        kind: "keyframes",
        targetId: "badge",
        start: { type: "afterPrevious", offsetMs: 250 },
        durationMs: 300,
        fill: "both",
        tracks: [
          {
            property: "transform.rotateDeg",
            keyframes: [
              { offset: 0, value: 0 },
              { offset: 1, value: 30 }
            ]
          }
        ]
      }
    ],
    dependencyGraph: {
      edges: [
        { from: "start-click", to: "headline-click-fade", relation: "triggers" },
        { from: "headline-click-fade", to: "badge-with-previous", relation: "with", offsetMs: 125 },
        { from: "badge-with-previous", to: "badge-after-previous", relation: "after", offsetMs: 250 }
      ]
    }
  };
  return deck;
}

function createGroupedBuildDeck(): DeckIR {
  return {
    irVersion: IR_VERSION,
    metadata: { title: "Grouped build PPTX fixture", sourceApplication: "KeyMorph tests" },
    deck: {
      id: "grouped-fixture",
      title: "Grouped build PPTX fixture",
      size: { width: 1280, height: 720, unit: "px" },
      slides: [
        {
          id: "slide-1",
          index: 0,
          name: "Grouped builds",
          background: { type: "solid", color: "#ffffff" },
          objects: [
            {
              id: "group",
              type: "group",
              name: "Group",
              bounds: { x: 90, y: 90, width: 520, height: 220 },
              opacity: 1,
              children: [
                {
                  id: "child-a",
                  type: "shape",
                  name: "Group child A",
                  shape: "rect",
                  bounds: { x: 120, y: 140, width: 140, height: 80 },
                  opacity: 1,
                  style: { fill: { type: "solid", color: "#2563eb" }, stroke: { color: "#2563eb", width: 0 } }
                },
                {
                  id: "child-b",
                  type: "shape",
                  name: "Group child B",
                  shape: "ellipse",
                  bounds: { x: 320, y: 140, width: 140, height: 80 },
                  opacity: 1,
                  style: { fill: { type: "solid", color: "#dc2626" }, stroke: { color: "#dc2626", width: 0 } }
                }
              ]
            }
          ],
          timeline: {
            durationMs: 1800,
            triggers: [{ id: "build-click", type: "onClick", clickIndex: 1 }],
            events: [
              {
                id: "child-a-fade",
                kind: "keyframes",
                targetId: "child-a",
                start: { type: "trigger", triggerId: "build-click" },
                durationMs: 350,
                fill: "both",
                tracks: [
                  {
                    property: "opacity",
                    keyframes: [
                      { offset: 0, value: 0 },
                      { offset: 1, value: 1 }
                    ]
                  }
                ]
              },
              {
                id: "child-b-fade",
                kind: "keyframes",
                targetId: "child-b",
                start: { type: "afterPrevious", offsetMs: 100 },
                durationMs: 350,
                fill: "both",
                tracks: [
                  {
                    property: "opacity",
                    keyframes: [
                      { offset: 0, value: 0 },
                      { offset: 1, value: 1 }
                    ]
                  }
                ]
              }
            ],
            dependencyGraph: {
              edges: [
                { from: "build-click", to: "child-a-fade", relation: "triggers" },
                { from: "child-a-fade", to: "child-b-fade", relation: "after", offsetMs: 100 }
              ]
            }
          }
        }
      ]
    },
    conversion: { status: "success", messages: [] }
  };
}

async function readPptxXml(filePath: string, entryPath: string): Promise<string> {
  const entries = readZipEntries(await readFile(filePath));
  const entry = entries.get(entryPath);
  if (!entry) throw new Error(`Missing PPTX entry ${entryPath}`);
  return new TextDecoder().decode(entry);
}

async function patchPptxXml(filePath: string, entryPath: string, patch: (source: string) => string): Promise<void> {
  const entries = readZipEntries(await readFile(filePath));
  const entry = entries.get(entryPath);
  if (!entry) throw new Error(`Missing PPTX entry ${entryPath}`);
  entries.set(entryPath, new TextEncoder().encode(patch(new TextDecoder().decode(entry))));
  await writeFile(filePath, createZip(entries));
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

function readZipEntries(data: Uint8Array): Map<string, Uint8Array> {
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  const entries = new Map<string, Uint8Array>();
  let offset = 0;

  while (offset + 30 <= data.length && view.getUint32(offset, true) === 0x04034b50) {
    const flags = view.getUint16(offset + 6, true);
    const compression = view.getUint16(offset + 8, true);
    const compressedSize = view.getUint32(offset + 18, true);
    const nameLength = view.getUint16(offset + 26, true);
    const extraLength = view.getUint16(offset + 28, true);
    const nameStart = offset + 30;
    const dataStart = nameStart + nameLength + extraLength;
    const dataEnd = dataStart + compressedSize;
    const name = new TextDecoder().decode(data.slice(nameStart, nameStart + nameLength));

    if ((flags & 0x08) !== 0) throw new Error("ZIP data descriptors are not supported in test helper.");
    const compressed = data.slice(dataStart, dataEnd);
    entries.set(name, compression === 8 ? inflateRawSync(compressed) : compressed);
    offset = dataEnd;
  }

  return entries;
}

function createZip(files: Map<string, Uint8Array>): Uint8Array {
  const chunks: Uint8Array[] = [];
  const centralDirectory: Uint8Array[] = [];
  let offset = 0;

  for (const [name, data] of files) {
    const nameBytes = new TextEncoder().encode(name);
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
  return concat([
    ...chunks,
    central,
    concat([u32(0x06054b50), u16(0), u16(0), u16(files.size), u16(files.size), u32(central.length), u32(centralStart), u16(0)])
  ]);
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
