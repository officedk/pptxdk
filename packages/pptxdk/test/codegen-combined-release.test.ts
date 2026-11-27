import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { readFileSync } from "node:fs";

import {
  DeckContainer,
  build,
  codegen,
  defineTheme,
  importPptx,
  type CodegenResult,
} from "@officedk/pptxdk";
import {
  Box,
  Cell,
  Chart,
  Deck,
  Group,
  Image,
  MotionScope,
  Notes,
  P,
  Placeholder,
  Row,
  Section,
  Shape,
  Slide,
  SlideLayout,
  SlideMaster,
  SlideNumber,
  Span,
  Step,
  Table,
  Text,
  after,
  createMorphScope,
  defineSlideLayout,
  expandDocument,
  fade,
  font as reactFont,
  linearGradient as reactLinearGradient,
  path as reactPath,
  transitions,
  useMotionScope,
  wipe,
} from "@officedk/pptxdk/react";
import {
  box,
  cell,
  chart,
  createMotionScopeHandle,
  deck,
  font as vanillaFont,
  group,
  image,
  linearGradient as vanillaLinearGradient,
  motionScope,
  notes,
  paragraph,
  path as vanillaPath,
  placeholder,
  row,
  section,
  shape,
  slide,
  slideLayout,
  slideMaster,
  slideNumber,
  span,
  step,
  table,
  text,
} from "@officedk/pptxdk/vanilla";
import { type ValidationResult } from "@pptxdk/validation";
import { strFromU8, unzipSync } from "fflate";
import { PNG } from "pngjs";
import { createElement } from "react";
import { describe, it } from "vitest";

import {
  previewGeneratedSourceRelocated,
  rebuildGeneratedSourceRelocated,
  verifyGeneratedSourcesTypecheck,
} from "./helpers/codegen-acceptance.ts";

const BRAND_FIXTURE_SHA256 =
  "de3376969353a2322ccce48182393f63eea9a44f5a231565985bcb5d3d336536";

const jostBookBytes = new Uint8Array(
  readFileSync(new URL("./assets/fonts/Jost-400-Book.ttf", import.meta.url)),
);
const jostBoldBytes = new Uint8Array(
  readFileSync(new URL("./assets/fonts/Jost-700-Bold.ttf", import.meta.url)),
);

function createSharedBrandPng(): Uint8Array {
  const png = new PNG({ width: 32, height: 32 });
  for (let y = 0; y < png.height; y += 1) {
    for (let x = 0; x < png.width; x += 1) {
      const color =
        x < 16 && y < 16
          ? [29, 78, 216, 255]
          : x >= 16 && y < 16
            ? [5, 150, 105, 255]
            : x < 16
              ? [217, 119, 6, 255]
              : [15, 23, 42, 255];
      png.data.set(color, (y * png.width + x) * 4);
    }
  }
  return new Uint8Array(PNG.sync.write(png));
}

function assertByteIdenticalCodegen(
  first: CodegenResult,
  second: CodegenResult,
  label: string,
): void {
  assert.equal(second.target, first.target, `${label} target`);
  assert.equal(second.entry, first.entry, `${label} entry`);
  const firstPaths = Object.keys(first.files).toSorted();
  const secondPaths = Object.keys(second.files).toSorted();
  assert.deepEqual(secondPaths, firstPaths, `${label} relative paths`);
  for (const relPath of firstPaths) {
    const a = first.files[relPath];
    const b = second.files[relPath];
    if (typeof a === "string") {
      assert.equal(typeof b, "string", `${label} ${relPath} type`);
      assert.equal(b, a, `${label} ${relPath} source text`);
    } else {
      assert.ok(a instanceof Uint8Array && b instanceof Uint8Array);
      assert.equal(Buffer.compare(a, b), 0, `${label} ${relPath} binary bytes`);
    }
  }
}

function countNonWhitePixelsInBox(
  pngBytes: Uint8Array,
  region: {
    readonly x: number;
    readonly y: number;
    readonly width: number;
    readonly height: number;
  },
): number {
  const png = PNG.sync.read(Buffer.from(pngBytes));
  const minX = Math.max(0, Math.floor(region.x));
  const minY = Math.max(0, Math.floor(region.y));
  const maxX = Math.min(png.width - 1, Math.ceil(region.x + region.width));
  const maxY = Math.min(png.height - 1, Math.ceil(region.y + region.height));
  let count = 0;
  for (let y = minY; y <= maxY; y += 1) {
    for (let x = minX; x <= maxX; x += 1) {
      const idx = (y * png.width + x) * 4;
      const r = png.data[idx]!;
      const g = png.data[idx + 1]!;
      const b = png.data[idx + 2]!;
      const a = png.data[idx + 3]!;
      if (a > 0 && (r < 245 || g < 245 || b < 245)) {
        count += 1;
      }
    }
  }
  return count;
}

function buildCombinedReleaseDecks() {
  const sharedPng = createSharedBrandPng();
  const vFont = vanillaFont("Jost*", [
    { weight: 400, style: "normal", data: jostBookBytes },
    { weight: 700, style: "normal", data: jostBoldBytes },
  ]);
  const rFont = reactFont("Jost*", [
    { weight: 400, style: "normal", data: jostBookBytes },
    { weight: 700, style: "normal", data: jostBoldBytes },
  ]);

  const makeTheme = (fontValue: typeof vFont) =>
    defineTheme({
      name: "Combined Brand",
      colors: {
        accent1: "#1D4ED8",
        accent2: "#059669",
      },
      fonts: { major: fontValue, minor: fontValue },
    });

  const vTheme = makeTheme(vFont);
  const vThemeB = defineTheme({
    name: "Combined Brand Secondary",
    colors: {
      accent1: "#1D4ED8",
      accent2: "#059669",
    },
    fonts: { major: vFont, minor: vFont },
  });
  const rTheme = defineTheme({
    name: "Combined Brand",
    colors: {
      accent1: "#1D4ED8",
      accent2: "#059669",
    },
    fonts: { major: rFont, minor: rFont },
  });
  const rThemeB = defineTheme({
    name: "Combined Brand Secondary",
    colors: {
      accent1: "#1D4ED8",
      accent2: "#059669",
    },
    fonts: { major: rFont, minor: rFont },
  });

  // Duplicate display name "Shared Content" across active layout, idle layout, and second master layout
  const primaryContentLayout = defineSlideLayout({
    id: "combo-content-a",
    name: "Shared Content",
    type: "obj",
    preserve: true,
    showMasterPlaceholderAnimations: true,
    headerFooter: { dateTime: false, footer: true, slideNumber: true },
    placeholders: {
      title: {
        type: "title",
        style: {
          position: "absolute",
          left: 40,
          top: 52,
          width: 560,
          height: 54,
          fontSize: "20pt",
          color: "#0F172A",
        },
      },
      picture: {
        type: "pic",
        explicitType: true,
        idx: 1,
        style: {
          position: "absolute",
          left: 40,
          top: 116,
          width: 150,
          height: 110,
        },
      },
      table: {
        type: "tbl",
        explicitType: true,
        idx: 2,
        style: {
          position: "absolute",
          left: 210,
          top: 116,
          width: 320,
          height: 150,
        },
      },
      chart: {
        type: "chart",
        explicitType: true,
        idx: 3,
        style: {
          position: "absolute",
          left: 550,
          top: 116,
          width: 370,
          height: 230,
        },
      },
      emptySlot: {
        type: "obj",
        explicitType: false,
        idx: 4,
        style: {
          position: "absolute",
          left: 40,
          top: 242,
          width: 150,
          height: 84,
        },
      },
    },
  });

  const idleLayoutA = defineSlideLayout({
    id: "combo-idle-a",
    name: "Shared Content",
    type: "obj",
    preserve: true,
    placeholders: {
      title: {
        type: "title",
        style: {
          position: "absolute",
          left: 40,
          top: 52,
          width: 880,
          height: 48,
        },
      },
      body: {
        type: "body",
        explicitType: true,
        idx: 1,
        style: {
          position: "absolute",
          left: 40,
          top: 116,
          width: 880,
          height: 300,
        },
      },
    },
  });

  const secondaryFlowLayout = defineSlideLayout({
    id: "combo-flow-b",
    name: "Shared Content",
    type: "title",
    preserve: true,
    placeholders: {
      title: {
        type: "ctrTitle",
        style: {
          position: "absolute",
          left: 40,
          top: 52,
          width: 880,
          height: 52,
          fontSize: "24pt",
          color: "#0F172A",
        },
      },
    },
  });

  const morphScopeA = createMorphScope();
  const morphScopeB = createMorphScope();
  const heroCardKey = morphScopeA.handle("hero-card");
  const isolatedScopeKey = morphScopeB.handle("hero-card");

  const leadScopeV = createMotionScopeHandle();
  const followerScopeV = createMotionScopeHandle();
  const readyCueV = leadScopeV.cue("ready");

  const comboChartProps = {
    id: "combo-chart",
    name: "Combined Combo Chart",
    placeholder: "chart" as const,
    title: {
      text: "Revenue & Target",
      style: { fontFamily: vFont, fontSize: 13, color: "#0F172A" as const },
    },
    displayBlanksAs: "gap" as const,
    legend: "bottom" as const,
    style: {
      position: "absolute" as const,
      left: 550,
      top: 116,
      width: 370,
      height: 230,
      fontFamily: vFont,
      fontSize: 10,
    },
    enter: fade({ duration: 240, by: "series" }),
    axes: {
      period: { kind: "category" as const, title: "Period" },
      amount: {
        kind: "value" as const,
        position: "left" as const,
        min: 0,
        max: 60,
      },
      rate: {
        kind: "value" as const,
        position: "right" as const,
        min: 0,
        max: 100,
      },
    },
    groups: [
      {
        kind: "column" as const,
        axes: { category: "period", value: "amount" },
        data: {
          kind: "shared-categories" as const,
          categories: ["Q1", "Q2", "Q3"],
        },
        series: [
          {
            name: "Revenue",
            values: [0, null, 42],
            trendlines: [{ type: "linear" as const, displayEquation: true }],
            errorBars: {
              type: "custom" as const,
              include: "both" as const,
              plus: [4, null, 6],
              minus: [2, 0, 3],
            },
          },
        ],
      },
      {
        kind: "line" as const,
        axes: { category: "period", value: "rate" },
        data: {
          kind: "shared-categories" as const,
          categories: ["Q1", "Q2", "Q3"],
        },
        series: [{ name: "Target", values: [25, 55, 85] }],
      },
    ],
  };

  const emptySeriesRingProps = {
    id: "ring-empty-series",
    name: "Ring With Empty Series",
    style: {
      position: "absolute" as const,
      left: 620,
      top: 300,
      width: 280,
      height: 180,
      fontFamily: vFont,
      fontSize: 10,
    },
    groups: [
      {
        kind: "doughnut" as const,
        data: { kind: "independent-categories" as const },
        series: [
          { name: "Backlog", categories: [], values: [] },
          { name: "Active", categories: ["A", "B"], values: [0, 40] },
        ],
      },
    ],
  };

  const masterMotion = {
    steps: [
      {
        trigger: "click" as const,
        effects: [
          {
            target: { placeholder: "ftr" as const },
            effect: wipe({ from: "left" as const, duration: 220 }),
          },
        ],
      },
    ],
  };

  const masterFooterPlaceholders = {
    footer: { text: "Combined Release Deck", style: { fontSize: "10pt" } },
    slideNumber: { idx: 5, style: { fontSize: "10pt" } },
  } as const;

  const vanillaDoc = deck({ width: 960, height: 540, theme: vTheme }).append(
    slideMaster({
      id: "brand-master-a",
      name: "Brand Master",
      preserve: true,
      theme: vTheme,
      background: "#F8FAFC",
      headerFooter: { dateTime: false, footer: true, slideNumber: true },
      footerPlaceholders: masterFooterPlaceholders,
      placeholders: {
        ftr: {
          type: "ftr",
          idx: 6,
          style: {
            position: "absolute",
            left: 40,
            top: 500,
            width: 320,
            height: 24,
          },
        },
      },
      motion: masterMotion,
    }).append(
      image({
        id: "master-logo",
        name: "Master Logo",
        src: sharedPng,
        style: {
          position: "absolute",
          left: 24,
          top: 12,
          width: 40,
          height: 28,
        },
      }),
      shape("RELEASE MASTER", {
        id: "master-badge",
        name: "Master Badge",
        style: {
          position: "absolute",
          left: 72,
          top: 14,
          width: 150,
          height: 24,
          background: "#1E293B",
          fontFamily: vFont,
          fontSize: 10,
          color: "#F8FAFC",
        },
      }),
      slideLayout(primaryContentLayout),
      slideLayout(idleLayoutA),
    ),
    slideMaster({
      id: "brand-master-b",
      name: "Brand Master",
      preserve: true,
      theme: vThemeB,
      background: "#FFFFFF",
    }).append(slideLayout(secondaryFlowLayout)),
    section({ name: "Executive Overview" }).append(
      slide({
        id: "overview-slide",
        layout: primaryContentLayout,
        footer: "Combined Release Deck",
        slideNumber: true,
        showMasterPlaceholderAnimations: true,
        transition: transitions.fade({ duration: 400 }),
        advanceAfter: 6000,
      }).append(
        step({ n: 1 }).append(
          text(
            [
              paragraph().append(
                span("Release ", {
                  style: { fontWeight: "bold", color: "#1D4ED8" },
                }),
                slideNumber({ style: { color: "#059669" } }),
              ),
              paragraph("Executive Summary"),
            ],
            {
              id: "s1-title",
              placeholder: "title",
              link: { url: "https://example.com/release" },
              enter: fade({ duration: 220, by: "paragraph" }),
              style: { fontFamily: vFont },
            },
          ),
          chart(comboChartProps),
        ),
        step({ n: 2 }).append(
          image({
            id: "ph-cropped-img",
            name: "Cropped Placeholder Image",
            placeholder: "picture",
            src: sharedPng,
            crop: { left: 0.1, top: 0.1, right: 0.1, bottom: 0.1 },
          }),
        ),
        table({
          id: "ph-merged-table",
          name: "Combined Merged Table",
          placeholder: "table",
          columns: [110, 105, 105],
          style: { fontFamily: vFont, fontSize: 11 },
        }).append(
          row({ style: { height: 34 } }).append(
            cell("Quarterly Plan", { colSpan: 3 }),
          ),
          row({ style: { height: 38 } }).append(
            cell("North", { rowSpan: 2 }),
            cell("120"),
            cell("140"),
          ),
          row({ style: { height: 38 } }).append(cell("95"), cell("110")),
        ),
        placeholder({ id: "ph-empty-slot", placeholder: "emptySlot" }),
        shape("Origin", {
          id: "morph-origin",
          morphKey: heroCardKey,
          style: {
            position: "absolute",
            left: 210,
            top: 356,
            width: 140,
            height: 58,
            background: "#1D4ED8",
            fontFamily: vFont,
            fontSize: 12,
            color: "#FFFFFF",
          },
        }),
        shape("Isolated", {
          id: "isolated-morph",
          morphKey: isolatedScopeKey,
          style: {
            position: "absolute",
            left: 370,
            top: 356,
            width: 140,
            height: 58,
            background: "#7C3AED",
            fontFamily: vFont,
            fontSize: 12,
            color: "#FFFFFF",
          },
        }),
        notes("Overview notes for release acceptance"),
      ),
    ),
    section({ name: "Visuals & Choreography" }).append(
      slide({
        id: "motion-slide",
        layout: secondaryFlowLayout,
        footer: "Combined Release Deck",
        slideNumber: true,
        transition: transitions.morph({
          duration: 900,
          fallback: transitions.fade({ duration: 400 }),
        }),
        advanceAfter: 4500,
      }).append(
        text("Visuals & Choreography", {
          id: "s2-title",
          placeholder: "title",
          style: { fontFamily: vFont },
        }),
        shape({
          id: "img-fill-shape",
          name: "Image Filled Shape",
          geometry: vanillaPath("M 0 0 L 140 0 L 140 96 L 0 96 Z", {
            viewBox: [140, 96],
          }),
          style: {
            position: "absolute",
            left: 40,
            top: 124,
            width: 140,
            height: 96,
            background: { image: sharedPng },
            border: { color: "#0F172A", width: 2 },
          },
        }),
        shape({
          id: "grad-shape",
          name: "Gradient Shape",
          style: {
            position: "absolute",
            left: 200,
            top: 124,
            width: 120,
            height: 96,
            background: vanillaLinearGradient(135, "#1D4ED8", "#0EA5E9"),
          },
        }),
        box({
          id: "flex-panel",
          style: {
            position: "absolute",
            left: 340,
            top: 124,
            width: 240,
            height: 96,
            display: "flex",
            flexDirection: "row",
            gap: 12,
            padding: 10,
            background: "#E2E8F0",
          },
        }).append(
          text("Flex A", {
            id: "flex-a",
            style: { fontFamily: vFont, fontSize: 12, color: "#0F172A" },
          }),
          text("Flex B", {
            id: "flex-b",
            style: { fontFamily: vFont, fontSize: 12, color: "#1D4ED8" },
          }),
        ),
        group({
          id: "outer-group",
          name: "Outer Group",
        }).append(
          shape("Outer", {
            id: "outer-chip",
            style: {
              position: "absolute",
              left: 600,
              top: 124,
              width: 120,
              height: 50,
              background: "#0284C7",
              fontFamily: vFont,
              fontSize: 11,
              color: "#FFFFFF",
            },
          }),
          group({
            id: "inner-group",
            name: "Inner Group",
          }).append(
            shape("Inner", {
              id: "inner-chip",
              style: {
                position: "absolute",
                left: 730,
                top: 144,
                width: 140,
                height: 60,
                background: "#059669",
                fontFamily: vFont,
                fontSize: 11,
                color: "#FFFFFF",
              },
            }),
          ),
        ),
        shape("Target", {
          id: "morph-target",
          morphKey: heroCardKey,
          style: {
            position: "absolute",
            left: 40,
            top: 332,
            width: 220,
            height: 82,
            background: "#059669",
            fontFamily: vFont,
            fontSize: 14,
            color: "#FFFFFF",
          },
        }),
        step().append(
          motionScope(leadScopeV, {
            mode: "sequence",
            publishEnd: readyCueV,
          }).append(
            shape({
              id: "cue-lead",
              name: "Cue Lead Shape",
              style: {
                position: "absolute",
                left: 300,
                top: 332,
                width: 140,
                height: 64,
                background: "#D97706",
              },
              enter: fade({ duration: 240 }),
            }),
          ),
          motionScope(followerScopeV, {
            startAt: after(readyCueV, 120),
          }).append(
            shape({
              id: "cue-follower",
              name: "Cue Follower Shape",
              style: {
                position: "absolute",
                left: 460,
                top: 332,
                width: 140,
                height: 64,
                background: "#0F172A",
              },
              enter: wipe({ from: "left", duration: 220 }),
            }),
          ),
        ),
        chart(emptySeriesRingProps),
        notes("Visual and choreography notes"),
      ),
    ),
  );

  function MotionSlideReact() {
    const leadScope = useMotionScope();
    const followerScope = useMotionScope();
    const readyCue = leadScope.cue("ready");
    return createElement(
      Slide,
      {
        id: "motion-slide",
        layout: secondaryFlowLayout,
        footer: "Combined Release Deck",
        slideNumber: true,
        transition: transitions.morph({
          duration: 900,
          fallback: transitions.fade({ duration: 400 }),
        }),
        advanceAfter: 4500,
      },
      createElement(
        Text,
        { id: "s2-title", placeholder: "title", style: { fontFamily: rFont } },
        "Visuals & Choreography",
      ),
      createElement(Shape, {
        id: "img-fill-shape",
        name: "Image Filled Shape",
        geometry: reactPath("M 0 0 L 140 0 L 140 96 L 0 96 Z", {
          viewBox: [140, 96],
        }),
        style: {
          position: "absolute",
          left: 40,
          top: 124,
          width: 140,
          height: 96,
          background: { image: sharedPng },
          border: { color: "#0F172A", width: 2 },
        },
      }),
      createElement(Shape, {
        id: "grad-shape",
        name: "Gradient Shape",
        style: {
          position: "absolute",
          left: 200,
          top: 124,
          width: 120,
          height: 96,
          background: reactLinearGradient(135, "#1D4ED8", "#0EA5E9"),
        },
      }),
      createElement(
        Box,
        {
          id: "flex-panel",
          style: {
            position: "absolute",
            left: 340,
            top: 124,
            width: 240,
            height: 96,
            display: "flex",
            flexDirection: "row",
            gap: 12,
            padding: 10,
            background: "#E2E8F0",
          },
        },
        createElement(
          Text,
          {
            id: "flex-a",
            style: { fontFamily: rFont, fontSize: 12, color: "#0F172A" },
          },
          "Flex A",
        ),
        createElement(
          Text,
          {
            id: "flex-b",
            style: { fontFamily: rFont, fontSize: 12, color: "#1D4ED8" },
          },
          "Flex B",
        ),
      ),
      createElement(
        Group,
        {
          id: "outer-group",
          name: "Outer Group",
        },
        createElement(
          Shape,
          {
            id: "outer-chip",
            style: {
              position: "absolute",
              left: 600,
              top: 124,
              width: 120,
              height: 50,
              background: "#0284C7",
              fontFamily: rFont,
              fontSize: 11,
              color: "#FFFFFF",
            },
          },
          "Outer",
        ),
        createElement(
          Group,
          {
            id: "inner-group",
            name: "Inner Group",
          },
          createElement(
            Shape,
            {
              id: "inner-chip",
              style: {
                position: "absolute",
                left: 730,
                top: 144,
                width: 140,
                height: 60,
                background: "#059669",
                fontFamily: rFont,
                fontSize: 11,
                color: "#FFFFFF",
              },
            },
            "Inner",
          ),
        ),
      ),
      createElement(
        Shape,
        {
          id: "morph-target",
          morphKey: heroCardKey,
          style: {
            position: "absolute",
            left: 40,
            top: 332,
            width: 220,
            height: 82,
            background: "#059669",
            fontFamily: rFont,
            fontSize: 14,
            color: "#FFFFFF",
          },
        },
        "Target",
      ),
      createElement(
        Step,
        null,
        createElement(
          MotionScope,
          { scope: leadScope, mode: "sequence", publishEnd: readyCue },
          createElement(Shape, {
            id: "cue-lead",
            name: "Cue Lead Shape",
            style: {
              position: "absolute",
              left: 300,
              top: 332,
              width: 140,
              height: 64,
              background: "#D97706",
            },
            enter: fade({ duration: 240 }),
          }),
        ),
        createElement(
          MotionScope,
          { scope: followerScope, startAt: after(readyCue, 120) },
          createElement(Shape, {
            id: "cue-follower",
            name: "Cue Follower Shape",
            style: {
              position: "absolute",
              left: 460,
              top: 332,
              width: 140,
              height: 64,
              background: "#0F172A",
            },
            enter: wipe({ from: "left", duration: 220 }),
          }),
        ),
      ),
      createElement(Chart, {
        ...emptySeriesRingProps,
        style: {
          ...emptySeriesRingProps.style,
          fontFamily: rFont,
          fontSize: 10,
        },
      }),
      createElement(Notes, null, "Visual and choreography notes"),
    );
  }

  const reactElement = createElement(
    Deck,
    { width: 960, height: 540, theme: rTheme },
    createElement(
      SlideMaster,
      {
        id: "brand-master-a",
        name: "Brand Master",
        preserve: true,
        theme: rTheme,
        background: "#F8FAFC",
        headerFooter: { dateTime: false, footer: true, slideNumber: true },
        footerPlaceholders: masterFooterPlaceholders,
        placeholders: {
          ftr: {
            type: "ftr",
            idx: 6,
            style: {
              position: "absolute",
              left: 40,
              top: 500,
              width: 320,
              height: 24,
            },
          },
        },
        motion: masterMotion,
      },
      createElement(Image, {
        id: "master-logo",
        name: "Master Logo",
        src: sharedPng,
        style: {
          position: "absolute",
          left: 24,
          top: 12,
          width: 40,
          height: 28,
        },
      }),
      createElement(
        Shape,
        {
          id: "master-badge",
          name: "Master Badge",
          style: {
            position: "absolute",
            left: 72,
            top: 14,
            width: 150,
            height: 24,
            background: "#1E293B",
            fontFamily: rFont,
            fontSize: 10,
            color: "#F8FAFC",
          },
        },
        "RELEASE MASTER",
      ),
      createElement(SlideLayout, { layout: primaryContentLayout }),
      createElement(SlideLayout, { layout: idleLayoutA }),
    ),
    createElement(
      SlideMaster,
      {
        id: "brand-master-b",
        name: "Brand Master",
        preserve: true,
        theme: rThemeB,
        background: "#FFFFFF",
      },
      createElement(SlideLayout, { layout: secondaryFlowLayout }),
    ),
    createElement(
      Section,
      { name: "Executive Overview" },
      createElement(
        Slide,
        {
          id: "overview-slide",
          layout: primaryContentLayout,
          footer: "Combined Release Deck",
          slideNumber: true,
          showMasterPlaceholderAnimations: true,
          transition: transitions.fade({ duration: 400 }),
          advanceAfter: 6000,
        },
        createElement(
          Step,
          { n: 1 },
          createElement(
            Text,
            {
              id: "s1-title",
              placeholder: "title",
              link: { url: "https://example.com/release" },
              enter: fade({ duration: 220, by: "paragraph" }),
              style: { fontFamily: rFont },
            },
            createElement(
              P,
              null,
              createElement(
                Span,
                {
                  style: { fontWeight: "bold", color: "#1D4ED8" },
                },
                "Release ",
              ),
              createElement(SlideNumber, { style: { color: "#059669" } }),
            ),
            createElement(P, null, "Executive Summary"),
          ),
          createElement(Chart, {
            ...comboChartProps,
            title: {
              text: "Revenue & Target",
              style: { fontFamily: rFont, fontSize: 13, color: "#0F172A" },
            },
            style: {
              ...comboChartProps.style,
              fontFamily: rFont,
              fontSize: 10,
            },
          }),
        ),
        createElement(
          Step,
          { n: 2 },
          createElement(Image, {
            id: "ph-cropped-img",
            name: "Cropped Placeholder Image",
            placeholder: "picture",
            src: sharedPng,
            crop: { left: 0.1, top: 0.1, right: 0.1, bottom: 0.1 },
          }),
        ),
        createElement(
          Table,
          {
            id: "ph-merged-table",
            name: "Combined Merged Table",
            placeholder: "table",
            columns: [110, 105, 105],
            style: { fontFamily: rFont, fontSize: 11 },
          },
          createElement(
            Row,
            { style: { height: 34 } },
            createElement(Cell, { colSpan: 3 }, "Quarterly Plan"),
          ),
          createElement(
            Row,
            { style: { height: 38 } },
            createElement(Cell, { rowSpan: 2 }, "North"),
            createElement(Cell, null, "120"),
            createElement(Cell, null, "140"),
          ),
          createElement(
            Row,
            { style: { height: 38 } },
            createElement(Cell, null, "95"),
            createElement(Cell, null, "110"),
          ),
        ),
        createElement(Placeholder, {
          id: "ph-empty-slot",
          placeholder: "emptySlot",
        }),
        createElement(
          Shape,
          {
            id: "morph-origin",
            morphKey: heroCardKey,
            style: {
              position: "absolute",
              left: 210,
              top: 356,
              width: 140,
              height: 58,
              background: "#1D4ED8",
              fontFamily: rFont,
              fontSize: 12,
              color: "#FFFFFF",
            },
          },
          "Origin",
        ),
        createElement(
          Shape,
          {
            id: "isolated-morph",
            morphKey: isolatedScopeKey,
            style: {
              position: "absolute",
              left: 370,
              top: 356,
              width: 140,
              height: 58,
              background: "#7C3AED",
              fontFamily: rFont,
              fontSize: 12,
              color: "#FFFFFF",
            },
          },
          "Isolated",
        ),
        createElement(Notes, null, "Overview notes for release acceptance"),
      ),
    ),
    createElement(
      Section,
      { name: "Visuals & Choreography" },
      createElement(MotionSlideReact),
    ),
  );

  return { vanillaDoc, reactElement, sharedPng };
}

function assertCombinedRebuiltPptx(
  pptx: Uint8Array,
  file: ValidationResult,
  label: string,
  originKind: "direct" | "imported",
): void {
  assert.deepEqual(
    file.errors,
    [],
    `${label} validation errors: ${JSON.stringify(file.errors)}`,
  );
  assert.deepEqual(
    file.presentation.sections?.map((sec) => sec.name),
    ["Executive Overview", "Visuals & Choreography"],
    `${label} sections`,
  );
  assert.equal(file.masters?.length, 2, `${label} master count`);
  assert.equal(
    file.masters?.[0]?.layouts.length,
    2,
    `${label} master 1 layout count (includes idle layout)`,
  );
  assert.equal(
    file.masters?.[1]?.layouts.length,
    1,
    `${label} master 2 layout count`,
  );
  assert.equal(file.slides.length, 2, `${label} slide count`);
  assert.equal(
    file.masters?.[0]?.animations?.[0]?.effects[0]?.preset,
    "wipe",
    `${label} master OwnerMotion`,
  );

  const s1 = file.slides[0]!;
  const s2 = file.slides[1]!;
  assert.equal(s1.section, "Executive Overview", `${label} s1 section`);
  assert.equal(s2.section, "Visuals & Choreography", `${label} s2 section`);
  assert.equal(s1.transition?.kind, "fade", `${label} s1 transition`);
  assert.equal(s1.advanceAfterMs, 6000, `${label} s1 advanceAfter`);
  assert.equal(s2.transition?.kind, "morph", `${label} s2 morph transition`);
  assert.equal(
    s2.transition?.fallback?.kind,
    "fade",
    `${label} s2 morph fallback`,
  );
  assert.equal(s2.advanceAfterMs, 4500, `${label} s2 advanceAfter`);
  assert.equal(
    s1.notes?.text.text,
    "Overview notes for release acceptance",
    `${label} s1 notes`,
  );
  assert.equal(
    s2.notes?.text.text,
    "Visual and choreography notes",
    `${label} s2 notes`,
  );

  // Content placeholders on Slide 1
  const s1Title = s1.objects.find((o) => o.placeholder?.type === "title");
  assert.ok(s1Title, `${label} s1 title placeholder`);
  assert.match(s1Title.text?.text ?? "", /Release/u);
  assert.deepEqual(
    s1Title.link,
    { url: "https://example.com/release" },
    `${label} s1 title link`,
  );
  const s1Image = s1.objects.find((o) => o.kind === "image");
  assert.ok(s1Image?.crop, `${label} s1 cropped image placeholder`);
  const s1Table = s1.objects.find((o) => o.kind === "table");
  assert.ok(s1Table?.table, `${label} s1 merged table placeholder`);
  assert.ok(
    (s1Table.table.merges?.length ?? 0) >= 2,
    `${label} s1 table merges preserved`,
  );
  const s1Chart = s1.objects.find((o) => o.kind === "chart");
  assert.ok(s1Chart?.chart, `${label} s1 combo chart placeholder`);
  assert.equal(
    s1Chart.chart.series.length,
    2,
    `${label} combo chart retains revenue + target`,
  );
  const revenueSeries = s1Chart.chart.series.find((s) => s.name === "Revenue");
  assert.deepEqual(revenueSeries?.values, [0, null, 42]);
  assert.ok(
    (revenueSeries?.trendlines?.length ?? 0) >= 1 &&
      revenueSeries?.errorBars !== undefined,
    `${label} chart trendlines and errorBars preserved`,
  );
  const s2Chart = s2.objects.find((o) => o.kind === "chart");
  assert.ok(s2Chart?.chart, `${label} s2 doughnut chart`);
  const backlogSeries = s2Chart.chart.series.find((s) => s.name === "Backlog");
  assert.deepEqual(backlogSeries?.values, []);
  const s1Unpopulated = s1.objects.find(
    (o) =>
      o.kind === "shape" && o.placeholder?.idx === 4 && o.text === undefined,
  );
  assert.ok(s1Unpopulated, `${label} unpopulated placeholder preserved`);

  // Cross-slide Morph pairing & isolated MorphScope
  const morphOrigin = s1.objects.find((o) => o.text?.text === "Origin");
  const isolatedMorph = s1.objects.find((o) => o.text?.text === "Isolated");
  const morphTarget = s2.objects.find((o) => o.text?.text === "Target");
  assert.ok(morphOrigin?.name?.startsWith("!!morph-"));
  assert.equal(morphTarget?.name, morphOrigin?.name);
  assert.ok(isolatedMorph?.name?.startsWith("!!morph-"));
  assert.notEqual(isolatedMorph?.name, morphOrigin?.name);

  // Slide 2 nested Group & image-filled Shape
  const outerGroup = s2.objects.find(
    (o) => o.kind === "group" && o.name === "Outer Group",
  );
  assert.ok(outerGroup, `${label} outer group`);
  assert.ok(
    outerGroup.members?.some(
      (m) => m.kind === "group" && m.name === "Inner Group",
    ),
    `${label} nested inner group`,
  );

  const zip = unzipSync(pptx);
  const slide2Xml = strFromU8(zip["ppt/slides/slide2.xml"]!);
  assert.match(
    slide2Xml,
    /<a:blipFill/u,
    `${label} shape backgroundImage blipFill`,
  );
  assert.match(slide2Xml, /<a:gradFill/u, `${label} shape gradient fill`);
  assert.ok(
    Object.keys(zip).some((entry) => entry.startsWith("ppt/fonts/")),
    `${label} embedded Jost* font part`,
  );

  // Animations: paragraph + series segmented animations, click steps, cue timing
  const s1Effects = (s1.animations ?? []).flatMap((st) => st.effects);
  assert.ok(
    s1Effects.some((e) => e.targetRange?.type === "paragraph"),
    `${label} paragraph segmented animation`,
  );
  assert.ok(
    s1Effects.some((e) => e.targetRange?.type === "chart"),
    `${label} chart segmented animation`,
  );
  assert.ok(
    (s2.animations?.[0]?.effects.length ?? 0) >= 2,
    `${label} slide 2 cue-chained effects`,
  );

  if (originKind === "direct") {
    assert.ok(
      s2.objects.some((o) => o.text?.text === "Flex A") &&
        s2.objects.some((o) => o.text?.text === "Flex B"),
      `${label} box flex children rendered`,
    );
  }
}

describe("pptxdk combined release acceptance", () => {
  it("deterministically generates, typechecks, relocates, rebuilds, validates, and previews the full multi-capability deck across direct & imported sources and vanilla/react targets", async () => {
    const { vanillaDoc, reactElement } = buildCombinedReleaseDecks();
    const reactDoc = await expandDocument(reactElement);

    const builtDirect = await build(vanillaDoc, { embedFonts: "full" });

    const importBytes = new Uint8Array(builtDirect.pptx);
    const imported = importPptx(importBytes);
    importBytes.fill(0);
    assert.deepEqual(imported.report.diagnostics, []);

    const directVanilla = codegen(vanillaDoc, { target: "vanilla" });
    const directVanillaRepeat = codegen(vanillaDoc, { target: "vanilla" });
    assertByteIdenticalCodegen(
      directVanilla,
      directVanillaRepeat,
      "direct-vanilla",
    );

    const directReact = codegen(reactDoc, { target: "react" });
    const directReactRepeat = codegen(reactDoc, { target: "react" });
    assertByteIdenticalCodegen(directReact, directReactRepeat, "direct-react");

    const importedVanilla = codegen(imported.document, { target: "vanilla" });
    const importedVanillaRepeat = codegen(imported.document, {
      target: "vanilla",
    });
    assertByteIdenticalCodegen(
      importedVanilla,
      importedVanillaRepeat,
      "imported-vanilla",
    );

    const importedReact = codegen(imported.document, { target: "react" });
    const importedReactRepeat = codegen(imported.document, { target: "react" });
    assertByteIdenticalCodegen(
      importedReact,
      importedReactRepeat,
      "imported-react",
    );

    const allCases = [
      {
        label: "direct:vanilla",
        originKind: "direct" as const,
        result: directVanilla,
      },
      {
        label: "direct:react",
        originKind: "direct" as const,
        result: directReact,
      },
      {
        label: "imported:vanilla",
        originKind: "imported" as const,
        result: importedVanilla,
      },
      {
        label: "imported:react",
        originKind: "imported" as const,
        result: importedReact,
      },
    ];

    // Shared PNG deduplicated into a single image asset; direct Box retains layout rules
    for (const item of allCases) {
      const imageAssets = Object.keys(item.result.files).filter((path) =>
        /^assets\/image-.*\.png$/u.test(path),
      );
      assert.equal(
        imageAssets.length,
        1,
        `${item.label} shared PNG must deduplicate to 1 asset`,
      );
      const source = item.result.files[item.result.entry] as string;
      const layoutDecls = source.match(/\bdefineSlideLayout\(/gu) ?? [];
      assert.equal(
        layoutDecls.length,
        3,
        `${item.label} all 3 layouts (including duplicate-named & idle) emitted`,
      );
      if (item.originKind === "direct") {
        assert.match(
          source,
          /display:\s*"flex"/u,
          `${item.label} Box layout preserved`,
        );
      }
    }

    await verifyGeneratedSourcesTypecheck(allCases);

    for (const item of allCases) {
      const { built, file } = await rebuildGeneratedSourceRelocated(
        item.result,
        { buildOptions: { embedFonts: "full" } },
      );
      assertCombinedRebuiltPptx(built.pptx, file, item.label, item.originKind);

      const preview = await previewGeneratedSourceRelocated(item.result);
      assert.equal(preview.pages.length, 2, `${item.label} preview page count`);
      for (const page of preview.pages) {
        assert.deepEqual(
          page.errors,
          [],
          `${item.label} preview page ${page.page} errors`,
        );
        assert.ok(page.png instanceof Uint8Array);
      }
      // Verify visible pixels in Slide 1 combo chart and Slide 2 image-filled shape
      assert.ok(
        countNonWhitePixelsInBox(preview.pages[0]!.png!, {
          x: 550,
          y: 116,
          width: 370,
          height: 230,
        }) > 100,
        `${item.label} chart must paint visible pixels`,
      );
      assert.ok(
        countNonWhitePixelsInBox(preview.pages[1]!.png!, {
          x: 40,
          y: 124,
          width: 140,
          height: 96,
        }) > 100,
        `${item.label} image-filled shape must paint visible pixels`,
      );
    }
  }, 120_000);

  it("matches brand-master-inheritance-override.pptx correspondence across codegen and relocated rebuild", async () => {
    const rawFixture = new Uint8Array(
      readFileSync(
        new URL(
          "./fixtures/brand-master-inheritance-override.pptx",
          import.meta.url,
        ),
      ),
    );
    assert.equal(
      createHash("sha256").update(rawFixture).digest("hex"),
      BRAND_FIXTURE_SHA256,
    );
    const imported = importPptx(rawFixture);
    rawFixture.fill(0);
    assert.ok(imported.document instanceof DeckContainer);
    assert.equal(imported.document.masters.length, 1);

    const vanilla = codegen(imported.document, { target: "vanilla" });
    const react = codegen(imported.document, { target: "react" });
    await verifyGeneratedSourcesTypecheck([
      { label: "brand-master-vanilla", result: vanilla },
      { label: "brand-master-react", result: react },
    ]);

    for (const [label, result] of [
      ["brand-master-vanilla", vanilla],
      ["brand-master-react", react],
    ] as const) {
      const { built, file } = await rebuildGeneratedSourceRelocated(result);
      assert.deepEqual(file.errors, [], `${label} validation errors`);
      assert.equal(file.masters?.length, 1);
      assert.equal(file.masters?.[0]?.layouts.length, 2);
      assert.equal(file.masters?.[0]?.layouts[0]?.name, "Brand Hero");
      assert.equal(file.masters?.[0]?.layouts[1]?.name, "Brand Idle Content");
      assert.equal(file.masters?.[0]?.theme?.fonts.major, "Arial");
      assert.deepEqual(
        (file.masters?.[0]?.objects ?? []).map((o) => o.name),
        ["BrandBadge"],
      );
      assert.equal(file.slides.length, 2);
      assert.equal(
        file.slides[0]?.objects[0]?.text?.paragraphs[0]?.runs[0]?.fontSizePt,
        undefined,
        `${label} slide 1 inherits Brand Hero 32pt`,
      );
      assert.equal(
        file.slides[1]?.objects[0]?.text?.paragraphs[0]?.runs[0]?.fontSizePt,
        32,
        `${label} slide 2 retains explicit equal 32pt override`,
      );
      const zip = unzipSync(built.pptx);
      assert.match(
        strFromU8(zip["ppt/slideMasters/slideMaster1.xml"]!),
        /F8FAFC/u,
      );
    }
  }, 60_000);
});
