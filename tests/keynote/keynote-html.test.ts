import { describe, test } from "node:test";
import assert from "node:assert/strict";
import { mkdir, mkdtemp, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { validateIR } from "../../src/ir/index.ts";
import { parseKeynoteHtmlExportToIr } from "../../src/keynote/index.ts";

describe("Keynote HTML export parser", () => {
  test("maps synthetic Keynote HTML layer JSON into DeckIR objects, assets, timeline, and report", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "keymorph-keynote-html-"));
    await mkdir(path.join(dir, "assets", "slide-1"), { recursive: true });
    await mkdir(path.join(dir, "assets", "slide-2"), { recursive: true });
    await writeJson(path.join(dir, "assets", "header.json"), {
      slideWidth: 1024,
      slideHeight: 768,
      title: "HTML Fixture",
      fonts: ["Helvetica Neue"],
      slideList: [
        { id: "slide-1", title: "Hero" },
        { id: "slide-2", title: "Video" }
      ]
    });
    await writeJson(path.join(dir, "assets", "slide-1", "slide-1.json"), {
      assets: {
        img1: {
          type: "texture",
          path: "assets/slide-1/hero.png",
          width: 320,
          height: 180,
          mimeType: "image/png"
        }
      },
      baseLayer: {
        layers: [
          {
            id: "group",
            name: "Hero group",
            initialState: {
              position: { x: 512, y: 384 },
              width: 640,
              height: 360,
              anchorPoint: { x: 0.5, y: 0.5 },
              opacity: 1,
              scale: { x: 1, y: 1 }
            },
            layers: [
              {
                id: "image",
                name: "Hero image",
                initialState: {
                  position: { x: 220, y: 140 },
                  width: 320,
                  height: 180,
                  anchorPoint: { x: 0.5, y: 0.5 },
                  opacity: 0.75,
                  rotation: Math.PI / 2,
                  scale: { x: 1.2, y: 0.8 },
                  contents: "img1",
                  contentsRect: { x: 0.1, y: 0.2, width: 0.8, height: 0.7 }
                },
                animations: [
                  {
                    property: "opacity",
                    from: 0,
                    to: 1,
                    beginTime: 0.25,
                    duration: 0.5,
                    timingFunction: "ease-in-out"
                  },
                  {
                    property: "position",
                    keyTimes: [0, 0.5, 1],
                    values: [
                      { x: 100, y: 120 },
                      { x: 180, y: 150 },
                      { x: 220, y: 140 }
                    ],
                    beginTime: 0,
                    duration: 1,
                    timingFunction: "linear"
                  },
                  {
                    property: "rotation",
                    from: 0,
                    to: Math.PI,
                    beginTime: 1,
                    duration: 0.25
                  }
                ]
              }
            ]
          }
        ]
      }
    });
    await writeJson(path.join(dir, "assets", "slide-2", "slide-2.json"), {
      assets: {
        clip: {
          type: "video",
          path: "assets/slide-2/clip.mp4",
          width: 640,
          height: 360,
          duration: 4
        }
      },
      layers: [
        {
          id: "video",
          name: "Demo clip",
          initialState: {
            position: [512, 384],
            width: 640,
            height: 360,
            anchorPoint: [0.5, 0.5],
            hidden: false,
            contents: "clip"
          },
          animations: [
            {
              property: "bounds",
              values: [
                { x: 192, y: 204, width: 640, height: 360 },
                { x: 128, y: 168, width: 768, height: 432 }
              ],
              beginTime: 0,
              duration: 0.75
            },
            {
              property: "filters.gaussianBlur",
              from: 0,
              to: 12,
              beginTime: 0.1,
              duration: 0.2
            }
          ]
        }
      ]
    });

    const deck = await parseKeynoteHtmlExportToIr(dir);

    assert.equal(validateIR(deck).valid, true);
    assert.equal(deck.deck.title, "HTML Fixture");
    assert.deepEqual(deck.deck.size, { width: 1024, height: 768, unit: "px" });
    assert.equal(deck.deck.slides.length, 2);
    assert.equal(deck.deck.assets?.length, 2);
    assert.equal(deck.deck.assets?.find((asset) => asset.name === "hero.png")?.kind, "image");
    assert.equal(deck.deck.assets?.find((asset) => asset.name === "clip.mp4")?.kind, "video");
    assert.equal(deck.deck.assets?.find((asset) => asset.name === "clip.mp4")?.durationMs, 4000);

    const group = deck.deck.slides[0]?.objects[0];
    assert.equal(group?.type, "group");
    assert.deepEqual(group?.bounds, { x: 192, y: 204, width: 640, height: 360 });
    assert.equal(group?.type === "group" ? group.children.length : 0, 1);
    const image = group?.type === "group" ? group.children[0] : undefined;
    assert.equal(image?.type, "image");
    assert.deepEqual(image?.bounds, { x: 60, y: 50, width: 320, height: 180 });
    assert.equal(image?.opacity, 0.75);
    assert.equal(image?.transform?.scaleX, 1.2);
    assert.equal(image?.transform?.scaleY, 0.8);
    assert.equal(image?.transform?.rotateDeg, 90);
    assert.deepEqual(image?.type === "image" ? image.crop : undefined, { x: 0.1, y: 0.2, width: 0.8, height: 0.7, unit: "ratio" });

    const slideOneEvents = deck.deck.slides[0]?.timeline?.events ?? [];
    assert.equal(slideOneEvents.length, 3);
    const opacity = slideOneEvents.find((event) => event.kind === "keyframes" && event.tracks.some((track) => track.property === "opacity"));
    assert.equal(opacity?.kind, "keyframes");
    assert.equal(opacity?.durationMs, 500);
    assert.deepEqual(opacity?.kind === "keyframes" ? opacity.tracks[0]?.keyframes : undefined, [
      { offset: 0, value: 0 },
      { offset: 1, value: 1 }
    ]);
    const position = slideOneEvents.find((event) => event.kind === "keyframes" && event.tracks.some((track) => track.property === "bounds.x"));
    assert.equal(position?.kind, "keyframes");
    assert.deepEqual(position?.kind === "keyframes" ? position.tracks.find((track) => track.property === "bounds.x")?.keyframes : undefined, [
      { offset: 0, value: 100 },
      { offset: 0.5, value: 180 },
      { offset: 1, value: 220 }
    ]);
    const rotation = slideOneEvents.find((event) => event.kind === "keyframes" && event.tracks.some((track) => track.property === "transform.rotateDeg"));
    assert.deepEqual(rotation?.kind === "keyframes" ? rotation.tracks[0]?.keyframes : undefined, [
      { offset: 0, value: 0 },
      { offset: 1, value: 180 }
    ]);

    const video = deck.deck.slides[1]?.objects[0];
    assert.equal(video?.type, "media");
    assert.equal(video?.type === "media" ? video.mediaType : undefined, "video");
    assert.deepEqual(video?.bounds, { x: 192, y: 204, width: 640, height: 360 });
    const boundsEvent = deck.deck.slides[1]?.timeline?.events.find(
      (event) => event.kind === "keyframes" && event.tracks.some((track) => track.property === "bounds.width")
    );
    assert.equal(boundsEvent?.kind, "keyframes");
    assert.deepEqual(boundsEvent?.kind === "keyframes" ? boundsEvent.tracks.find((track) => track.property === "bounds.width")?.keyframes : undefined, [
      { offset: 0, value: 640 },
      { offset: 1, value: 768 }
    ]);
    const customEvent = deck.deck.slides[1]?.timeline?.events.find(
      (event) => event.kind === "keyframes" && event.tracks.some((track) => track.property === "custom.keynote.filters.gaussianBlur")
    );
    assert.equal(customEvent?.kind, "keyframes");

    assert.equal(deck.conversion?.status, "partial");
    assert.equal(deck.conversion?.tool, "keymorph-keynote-html-parser");
    assert.equal(deck.conversion?.statistics?.slideCount, 2);
    assert.equal(deck.conversion?.statistics?.assetCount, 2);
    assert.equal(deck.conversion?.statistics?.animationCount, 5);
    assert.equal(deck.conversion?.uncertainMappings?.some((mapping) => mapping.code === "keynote-html-animation-custom-property"), true);
    assert.equal(deck.conversion?.messages.some((message) => message.code === "keynote-html-static-import"), true);
  });
});

async function writeJson(filePath: string, value: unknown): Promise<void> {
  await writeFile(filePath, JSON.stringify(value, null, 2));
}
