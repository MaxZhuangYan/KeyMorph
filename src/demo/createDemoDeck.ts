import { IR_VERSION, type DeckIR } from "../ir/index.ts";

export function createDemoDeck(): DeckIR {
  return {
    irVersion: IR_VERSION,
    metadata: {
      title: "KeyMorph demo",
      sourceApplication: "KeyMorph"
    },
    deck: {
      id: "deck-demo",
      title: "KeyMorph demo",
      size: {
        width: 1280,
        height: 720,
        unit: "px"
      },
      slides: [
        {
          id: "slide-1",
          index: 0,
          name: "Intro",
          background: { type: "solid", color: "#f7f3ea" },
          objects: [
            {
              id: "title-slide-1",
              type: "text",
              name: "Title",
              morphKey: "deck-title",
              bounds: { x: 96, y: 112, width: 760, height: 120 },
              opacity: 1,
              transform: { scaleX: 1, scaleY: 1, rotateDeg: 0 },
              text: {
                runs: [
                  {
                    text: "KeyMorph",
                    style: {
                      fontFamily: "Inter, Arial, sans-serif",
                      fontSize: 74,
                      fontWeight: 800,
                      color: "#1e293b"
                    }
                  }
                ]
              }
            },
            {
              id: "tagline",
              type: "text",
              name: "Tagline",
              bounds: { x: 102, y: 244, width: 740, height: 82 },
              opacity: 0,
              transform: { scaleX: 0.98, scaleY: 0.98, rotateDeg: 0 },
              text: {
                runs: [
                  {
                    text: "Any slide deck becomes a programmable animation runtime.",
                    style: {
                      fontFamily: "Inter, Arial, sans-serif",
                      fontSize: 28,
                      color: "#334155"
                    }
                  }
                ]
              }
            },
            {
              id: "accent",
              type: "shape",
              name: "Accent bar",
              shape: "roundRect",
              bounds: { x: 102, y: 360, width: 260, height: 18 },
              opacity: 1,
              transform: { scaleX: 0.15, scaleY: 1, rotateDeg: 0 },
              style: {
                fill: { type: "solid", color: "#0f766e" },
                stroke: { color: "#0f766e", width: 0 }
              }
            }
          ],
          timeline: {
            durationMs: 2600,
            events: [
              {
                id: "tagline-fade",
                kind: "keyframes",
                targetId: "tagline",
                start: { type: "absolute", atMs: 350 },
                durationMs: 700,
                easing: "easeOutCubic",
                fill: "both",
                tracks: [
                  {
                    property: "opacity",
                    keyframes: [
                      { offset: 0, value: 0 },
                      { offset: 1, value: 1 }
                    ]
                  },
                  {
                    property: "bounds.y",
                    keyframes: [
                      { offset: 0, value: 264 },
                      { offset: 1, value: 244 }
                    ]
                  },
                  {
                    property: "transform.scaleX",
                    keyframes: [
                      { offset: 0, value: 0.98 },
                      { offset: 1, value: 1 }
                    ]
                  },
                  {
                    property: "transform.scaleY",
                    keyframes: [
                      { offset: 0, value: 0.98 },
                      { offset: 1, value: 1 }
                    ]
                  }
                ]
              },
              {
                id: "accent-grow",
                kind: "keyframes",
                targetId: "accent",
                start: { type: "absolute", atMs: 900 },
                durationMs: 900,
                easing: "easeInOutCubic",
                fill: "both",
                dependencies: [{ eventId: "tagline-fade", relation: "after", offsetMs: 550 }],
                tracks: [
                  {
                    property: "transform.scaleX",
                    keyframes: [
                      { offset: 0, value: 0.15 },
                      { offset: 1, value: 1 }
                    ]
                  }
                ]
              }
            ],
            dependencyGraph: {
              edges: [
                {
                  from: "tagline-fade",
                  to: "accent-grow",
                  relation: "after",
                  offsetMs: 550
                }
              ]
            }
          }
        },
        {
          id: "slide-2",
          index: 1,
          name: "Morph",
          background: { type: "solid", color: "#eef6ff" },
          objects: [
            {
              id: "title-slide-2",
              type: "text",
              name: "Title",
              morphKey: "deck-title",
              bounds: { x: 396, y: 96, width: 760, height: 110 },
              opacity: 1,
              transform: { scaleX: 0.82, scaleY: 0.82, rotateDeg: 0 },
              text: {
                runs: [
                  {
                    text: "KeyMorph",
                    style: {
                      fontFamily: "Inter, Arial, sans-serif",
                      fontSize: 74,
                      fontWeight: 800,
                      color: "#1e293b"
                    }
                  }
                ]
              }
            },
            {
              id: "runtime-card",
              type: "shape",
              name: "Runtime panel",
              shape: "roundRect",
              bounds: { x: 126, y: 244, width: 1028, height: 332 },
              opacity: 0,
              transform: { scaleX: 0.96, scaleY: 0.96, rotateDeg: 0 },
              style: {
                fill: { type: "solid", color: "#ffffff" },
                stroke: { color: "#94a3b8", width: 2 }
              }
            }
          ],
          transition: {
            type: "morph",
            durationMs: 900,
            easing: "easeInOutCubic",
            morph: {
              strategy: "morph",
              pairs: [
                {
                  fromObjectId: "title-slide-1",
                  toObjectId: "title-slide-2",
                  morphKey: "deck-title"
                }
              ],
              properties: ["bounds", "transform"]
            }
          },
          timeline: {
            durationMs: 1800,
            events: [
              {
                id: "runtime-card-in",
                kind: "keyframes",
                targetId: "runtime-card",
                start: { type: "absolute", atMs: 200 },
                durationMs: 620,
                easing: "easeOutCubic",
                fill: "both",
                tracks: [
                  {
                    property: "opacity",
                    keyframes: [
                      { offset: 0, value: 0 },
                      { offset: 1, value: 1 }
                    ]
                  },
                  {
                    property: "bounds.y",
                    keyframes: [
                      { offset: 0, value: 276 },
                      { offset: 1, value: 244 }
                    ]
                  },
                  {
                    property: "transform.scaleX",
                    keyframes: [
                      { offset: 0, value: 0.96 },
                      { offset: 1, value: 1 }
                    ]
                  },
                  {
                    property: "transform.scaleY",
                    keyframes: [
                      { offset: 0, value: 0.96 },
                      { offset: 1, value: 1 }
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
    conversion: {
      source: {
        kind: "json",
        name: "generated-demo",
        application: "KeyMorph"
      },
      status: "success",
      generatedAt: new Date(0).toISOString(),
      tool: "keymorph-demo",
      messages: [],
      unsupportedFeatures: [],
      degradedFeatures: [],
      uncertainMappings: [],
      statistics: {
        slideCount: 2,
        objectCount: 5,
        animationCount: 3,
        unsupportedFeatureCount: 0,
        degradedFeatureCount: 0,
        uncertainMappingCount: 0
      }
    }
  };
}
