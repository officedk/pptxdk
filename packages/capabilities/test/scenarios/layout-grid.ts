import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import { Deck, Slide, type CommonNode } from "@react-pptx/core";
import { createElement } from "react";

import { Box } from "../../src/box-elements.ts";
import { box, deck, font, shape, slide, text } from "../../src/document.ts";
import { Shape } from "../../src/shape-elements.ts";
import { Text } from "../../src/text-elements.ts";
import type { CapabilityScenario } from "../acceptance-runner.ts";
import type { LayoutRejectionCase } from "./layout-block-flex.ts";

const jostFont = font("Jost*", [
  {
    weight: 400,
    data: readFileSync(
      new URL(
        "../../../react-pptx/test/assets/fonts/Jost-400-Book.ttf",
        import.meta.url,
      ),
    ),
  },
  {
    weight: 700,
    data: readFileSync(
      new URL(
        "../../../react-pptx/test/assets/fonts/Jost-700-Bold.ttf",
        import.meta.url,
      ),
    ),
  },
]);

const notoFont = font("Noto Sans CJK SC", [
  {
    weight: 400,
    data: readFileSync(
      new URL(
        "../../../react-pptx/test/assets/fonts/NotoSansSC-Regular.woff2",
        import.meta.url,
      ),
    ),
  },
]);

export const layoutGridPositiveScenario: CapabilityScenario = {
  id: "layout-grid-semantics",
  capability: "layout:grid",
  description:
    "Dual-author Grid layout scenario covering px/auto/fr columns & rows, explicit placement & spans, percentage sizing inside grid areas, width-constrained Text wrapping driving auto row height, nested Grid <-> Flex Box combinations, scoped ThemeScope, and grid-overflow / grid-item-overflow warnings.",
  visualCheckpoints: [
    "Page 1 (grid-tracks-span-wrap): light #F8FAFC canvas; top full-span blue banner ('Grid Layout Dashboard' + pale blue badge) spans all 3 columns; row 2 has a 120px white left card ('Fixed 120px'), an auto-width blue card ('AutoColTrack'), and a wider emerald 1fr card whose wrapped 2-line text expands row 2 height so all three row-2 cards share the same height; row 3 has a 2-column spanning bottom-left white card and a centered purple 60%x50% Shape in column 3.",
    "Page 2 (grid-nested-flex-and-theme): white canvas; top Flex header ('Nested Flex <-> Grid & Scoped Theme'); left half holds an amber-themed nested Grid ('Scoped Theme Grid', orange Shape, and white Flex card '网格嵌套 Flex') overridden via local theme; right half holds a slate-themed Flex column containing an auto-rows Grid whose top row is fixed 36px and bottom row grows with wrapped text while retaining outer theme colors.",
    "Page 3 (grid-overflow-warnings): pale rose #FEF2F2 canvas; top red-bordered 280x110 Grid lets its 180px+150px fixed columns visibly extend past the right container border while emitting grid-overflow warnings for both axes; bottom dark-red-bordered Grid contains a 200x110 coral Shape intentionally overflowing its 160x80 cell (overlapping the left edge of the green Shape in column 2 and emitting grid-item-overflow).",
  ],
  nativeCheckpoints: [
    "Open exported PPTX in PowerPoint/WPS and verify identical EMU geometry across JSX and /document builds, including retained coordinates for grid-overflow and grid-item-overflow items.",
  ],
  jsx() {
    return createElement(
      Deck,
      {
        width: 640,
        height: 360,
        theme: {
          colors: {
            surface: "#FFFFFF",
            accent: "#2563EB",
            ink: "#0F172A",
            cardBg: "#F1F5F9",
            badge: "#0EA5E9",
            textInk: "#0F172A",
          },
        },
      },
      createElement(
        Slide,
        {
          id: "grid-tracks-span-wrap",
          background: "#F8FAFC",
          style: {
            display: "grid",
            padding: 20,
            gap: 12,
            gridTemplateColumns: "120px auto 1fr",
            gridTemplateRows: "auto auto 1fr",
          },
        },
        createElement(
          Box,
          {
            id: "grid-header",
            style: {
              gridColumn: "1 / span 3",
              gridRow: 1,
              display: "flex",
              flexDirection: "row",
              justifyContent: "space-between",
              alignItems: "center",
              padding: [8, 12],
              backgroundColor: "accent",
            },
          },
          createElement(
            Text,
            {
              id: "grid-header-title",
              style: {
                fontFamily: jostFont,
                fontSize: 16,
                fontWeight: "bold",
                color: "#FFFFFF",
                width: "auto",
                height: "auto",
              },
            },
            "Grid Layout Dashboard",
          ),
          createElement(Shape, {
            id: "grid-header-badge",
            width: 72,
            height: 20,
            fill: "#DBEAFE",
          }),
        ),
        createElement(
          Box,
          {
            id: "sidebar-fixed",
            style: {
              gridColumn: 1,
              gridRow: 2,
              padding: 8,
              backgroundColor: "surface",
              borderColor: "#CBD5E1",
              borderWidth: 1,
            },
          },
          createElement(
            Text,
            {
              id: "sidebar-text",
              style: {
                fontFamily: jostFont,
                fontSize: 13,
                color: "ink",
              },
            },
            "Fixed 120px",
          ),
        ),
        createElement(
          Box,
          {
            id: "auto-col-card",
            style: {
              gridColumn: 2,
              gridRow: 2,
              padding: 8,
              backgroundColor: "#EFF6FF",
              borderColor: "#3B82F6",
              borderWidth: 1,
            },
          },
          createElement(
            Text,
            {
              id: "auto-col-text",
              style: {
                fontFamily: jostFont,
                fontSize: 13,
                color: "#1E3A8A",
              },
            },
            "AutoColTrack",
          ),
        ),
        createElement(
          Box,
          {
            id: "wrap-fr-card",
            style: {
              gridColumn: 3,
              gridRow: 2,
              padding: 8,
              backgroundColor: "#ECFDF5",
              borderColor: "#10B981",
              borderWidth: 1,
            },
          },
          createElement(
            Text,
            {
              id: "wrap-fr-text",
              style: {
                fontFamily: jostFont,
                fontSize: 13,
                lineHeight: 18,
                color: "#065F46",
              },
            },
            "Wrapped text constrained by the 1fr column width expands row 2 automatically across multiple lines",
          ),
        ),
        createElement(
          Box,
          {
            id: "span-panel",
            style: {
              gridColumn: "1 / span 2",
              gridRow: 3,
              padding: 10,
              backgroundColor: "surface",
              borderColor: "#94A3B8",
              borderWidth: 1,
            },
          },
          createElement(
            Text,
            {
              id: "span-panel-text",
              style: {
                fontFamily: jostFont,
                fontSize: 13,
                color: "ink",
              },
            },
            "Spans columns 1-2 (120px + gap + auto)",
          ),
        ),
        createElement(Shape, {
          id: "percent-cell-shape",
          fill: "#7C3AED",
          style: {
            gridColumn: 3,
            gridRow: 3,
            width: "60%",
            height: "50%",
            justifySelf: "center",
            alignSelf: "center",
          },
        }),
      ),
      createElement(
        Slide,
        {
          id: "grid-nested-flex-and-theme",
          background: "#FFFFFF",
          style: {
            display: "flex",
            flexDirection: "column",
            padding: 20,
            gap: 12,
          },
        },
        createElement(
          Box,
          {
            id: "flex-top",
            style: {
              display: "flex",
              flexDirection: "row",
              justifyContent: "space-between",
              alignItems: "center",
              padding: [8, 12],
              backgroundColor: "cardBg",
              borderColor: "#CBD5E1",
              borderWidth: 1,
            },
          },
          createElement(
            Text,
            {
              id: "flex-top-title",
              style: {
                fontFamily: jostFont,
                fontSize: 15,
                fontWeight: "bold",
                color: "textInk",
                width: "auto",
                height: "auto",
              },
            },
            "Nested Flex <-> Grid & Scoped Theme",
          ),
          createElement(Shape, {
            id: "outer-theme-dot",
            width: 22,
            height: 22,
            fill: "badge",
          }),
        ),
        createElement(
          Box,
          {
            id: "outer-grid",
            style: {
              flex: 1,
              display: "grid",
              gridTemplateColumns: "repeat(2, minmax(0, 1fr))",
              gridTemplateRows: "1fr",
              gap: 14,
            },
          },
          createElement(
            Box,
            {
              id: "nested-grid",
              theme: {
                colors: {
                  cardBg: "#FEF3C7",
                  badge: "#D97706",
                  textInk: "#78350F",
                },
              },
              style: {
                display: "grid",
                gridTemplateColumns: "1fr 1fr",
                gridTemplateRows: "auto 1fr",
                gap: 8,
                padding: 10,
                backgroundColor: "cardBg",
                borderColor: "badge",
                borderWidth: 2,
              },
            },
            createElement(
              Text,
              {
                id: "scoped-grid-title",
                style: {
                  gridColumn: "1 / span 2",
                  fontFamily: jostFont,
                  fontSize: 14,
                  fontWeight: "bold",
                  color: "textInk",
                },
              },
              "Scoped Theme Grid",
            ),
            createElement(Shape, {
              id: "nested-shape-a",
              fill: "badge",
              style: { height: "100%" },
            }),
            createElement(
              Box,
              {
                id: "nested-flex-b",
                style: {
                  display: "flex",
                  flexDirection: "column",
                  justifyContent: "center",
                  alignItems: "center",
                  backgroundColor: "#FFFFFF",
                  borderColor: "badge",
                  borderWidth: 1,
                  padding: 6,
                },
              },
              createElement(
                Text,
                {
                  id: "nested-cjk-label",
                  style: {
                    fontFamily: notoFont,
                    fontSize: 12,
                    color: "textInk",
                    align: "center",
                  },
                },
                "网格嵌套 Flex",
              ),
            ),
          ),
          createElement(
            Box,
            {
              id: "right-flex",
              style: {
                display: "flex",
                flexDirection: "column",
                gap: 8,
                padding: 10,
                backgroundColor: "cardBg",
                borderColor: "#94A3B8",
                borderWidth: 1,
              },
            },
            createElement(
              Text,
              {
                id: "right-flex-title",
                style: {
                  fontFamily: jostFont,
                  fontSize: 14,
                  fontWeight: "bold",
                  color: "textInk",
                },
              },
              "Outer Theme Preserved",
            ),
            createElement(
              Box,
              {
                id: "auto-rows-grid",
                style: {
                  display: "grid",
                  gridTemplateColumns: "1fr 1fr",
                  gridAutoRows: "36px auto",
                  gap: 8,
                },
              },
              createElement(Shape, {
                id: "auto-r1-c1",
                fill: "badge",
                style: { height: "100%" },
              }),
              createElement(Shape, {
                id: "auto-r1-c2",
                fill: "#38BDF8",
                style: { height: "100%" },
              }),
              createElement(
                Box,
                {
                  id: "auto-r2-c1",
                  style: {
                    padding: 6,
                    backgroundColor: "#FFFFFF",
                    borderColor: "#CBD5E1",
                    borderWidth: 1,
                  },
                },
                createElement(
                  Text,
                  {
                    id: "auto-r2-c1-text",
                    style: {
                      fontFamily: jostFont,
                      fontSize: 12,
                      color: "textInk",
                    },
                  },
                  "Auto row wraps nicely here",
                ),
              ),
              createElement(Shape, {
                id: "auto-r2-c2",
                fill: "#0284C7",
                style: { height: "100%" },
              }),
            ),
          ),
        ),
      ),
      createElement(
        Slide,
        {
          id: "grid-overflow-warnings",
          background: "#FEF2F2",
          style: {
            display: "flex",
            flexDirection: "column",
            padding: 20,
            gap: 14,
          },
        },
        createElement(
          Box,
          {
            id: "container-overflow-grid",
            style: {
              display: "grid",
              width: 280,
              height: 110,
              gridTemplateColumns: "180px 150px",
              gridTemplateRows: "70px 60px",
              gap: 10,
              padding: 10,
              backgroundColor: "#FFFFFF",
              borderColor: "#EF4444",
              borderWidth: 2,
            },
          },
          createElement(Shape, {
            id: "ov-c1",
            fill: "#F87171",
            style: { height: "100%" },
          }),
          createElement(Shape, {
            id: "ov-c2",
            fill: "#FB923C",
            style: { height: "100%" },
          }),
        ),
        createElement(
          Box,
          {
            id: "item-overflow-grid",
            style: {
              display: "grid",
              width: 420,
              height: 150,
              gridTemplateColumns: "160px 1fr",
              gridTemplateRows: "80px",
              padding: 12,
              gap: 12,
              backgroundColor: "#FFFFFF",
              borderColor: "#DC2626",
              borderWidth: 2,
            },
          },
          createElement(Shape, {
            id: "tall-item",
            fill: "#FCA5A5",
            width: 200,
            height: 110,
          }),
          createElement(Shape, {
            id: "normal-item",
            fill: "#34D399",
            height: 60,
          }),
        ),
      ),
    );
  },
  document(): CommonNode {
    return deck({
      width: 640,
      height: 360,
      theme: {
        colors: {
          surface: "#FFFFFF",
          accent: "#2563EB",
          ink: "#0F172A",
          cardBg: "#F1F5F9",
          badge: "#0EA5E9",
          textInk: "#0F172A",
        },
      },
    }).append(
      slide({
        id: "grid-tracks-span-wrap",
        background: "#F8FAFC",
        style: {
          display: "grid",
          padding: 20,
          gap: 12,
          gridTemplateColumns: "120px auto 1fr",
          gridTemplateRows: "auto auto 1fr",
        },
      }).append(
        box({
          id: "grid-header",
          style: {
            gridColumn: "1 / span 3",
            gridRow: 1,
            display: "flex",
            flexDirection: "row",
            justifyContent: "space-between",
            alignItems: "center",
            padding: [8, 12],
            backgroundColor: "accent",
          },
        }).append(
          text("Grid Layout Dashboard", {
            id: "grid-header-title",
            style: {
              fontFamily: jostFont,
              fontSize: 16,
              fontWeight: "bold",
              color: "#FFFFFF",
              width: "auto",
              height: "auto",
            },
          }),
          shape({
            id: "grid-header-badge",
            width: 72,
            height: 20,
            fill: "#DBEAFE",
          }),
        ),
        box({
          id: "sidebar-fixed",
          style: {
            gridColumn: 1,
            gridRow: 2,
            padding: 8,
            backgroundColor: "surface",
            borderColor: "#CBD5E1",
            borderWidth: 1,
          },
        }).append(
          text("Fixed 120px", {
            id: "sidebar-text",
            style: {
              fontFamily: jostFont,
              fontSize: 13,
              color: "ink",
            },
          }),
        ),
        box({
          id: "auto-col-card",
          style: {
            gridColumn: 2,
            gridRow: 2,
            padding: 8,
            backgroundColor: "#EFF6FF",
            borderColor: "#3B82F6",
            borderWidth: 1,
          },
        }).append(
          text("AutoColTrack", {
            id: "auto-col-text",
            style: {
              fontFamily: jostFont,
              fontSize: 13,
              color: "#1E3A8A",
            },
          }),
        ),
        box({
          id: "wrap-fr-card",
          style: {
            gridColumn: 3,
            gridRow: 2,
            padding: 8,
            backgroundColor: "#ECFDF5",
            borderColor: "#10B981",
            borderWidth: 1,
          },
        }).append(
          text(
            "Wrapped text constrained by the 1fr column width expands row 2 automatically across multiple lines",
            {
              id: "wrap-fr-text",
              style: {
                fontFamily: jostFont,
                fontSize: 13,
                lineHeight: 18,
                color: "#065F46",
              },
            },
          ),
        ),
        box({
          id: "span-panel",
          style: {
            gridColumn: "1 / span 2",
            gridRow: 3,
            padding: 10,
            backgroundColor: "surface",
            borderColor: "#94A3B8",
            borderWidth: 1,
          },
        }).append(
          text("Spans columns 1-2 (120px + gap + auto)", {
            id: "span-panel-text",
            style: {
              fontFamily: jostFont,
              fontSize: 13,
              color: "ink",
            },
          }),
        ),
        shape({
          id: "percent-cell-shape",
          fill: "#7C3AED",
          style: {
            gridColumn: 3,
            gridRow: 3,
            width: "60%",
            height: "50%",
            justifySelf: "center",
            alignSelf: "center",
          },
        }),
      ),
      slide({
        id: "grid-nested-flex-and-theme",
        background: "#FFFFFF",
        style: {
          display: "flex",
          flexDirection: "column",
          padding: 20,
          gap: 12,
        },
      }).append(
        box({
          id: "flex-top",
          style: {
            display: "flex",
            flexDirection: "row",
            justifyContent: "space-between",
            alignItems: "center",
            padding: [8, 12],
            backgroundColor: "cardBg",
            borderColor: "#CBD5E1",
            borderWidth: 1,
          },
        }).append(
          text("Nested Flex <-> Grid & Scoped Theme", {
            id: "flex-top-title",
            style: {
              fontFamily: jostFont,
              fontSize: 15,
              fontWeight: "bold",
              color: "textInk",
              width: "auto",
              height: "auto",
            },
          }),
          shape({
            id: "outer-theme-dot",
            width: 22,
            height: 22,
            fill: "badge",
          }),
        ),
        box({
          id: "outer-grid",
          style: {
            flex: 1,
            display: "grid",
            gridTemplateColumns: "repeat(2, minmax(0, 1fr))",
            gridTemplateRows: "1fr",
            gap: 14,
          },
        }).append(
          box({
            id: "nested-grid",
            theme: {
              colors: {
                cardBg: "#FEF3C7",
                badge: "#D97706",
                textInk: "#78350F",
              },
            },
            style: {
              display: "grid",
              gridTemplateColumns: "1fr 1fr",
              gridTemplateRows: "auto 1fr",
              gap: 8,
              padding: 10,
              backgroundColor: "cardBg",
              borderColor: "badge",
              borderWidth: 2,
            },
          }).append(
            text("Scoped Theme Grid", {
              id: "scoped-grid-title",
              style: {
                gridColumn: "1 / span 2",
                fontFamily: jostFont,
                fontSize: 14,
                fontWeight: "bold",
                color: "textInk",
              },
            }),
            shape({
              id: "nested-shape-a",
              fill: "badge",
              style: { height: "100%" },
            }),
            box({
              id: "nested-flex-b",
              style: {
                display: "flex",
                flexDirection: "column",
                justifyContent: "center",
                alignItems: "center",
                backgroundColor: "#FFFFFF",
                borderColor: "badge",
                borderWidth: 1,
                padding: 6,
              },
            }).append(
              text("网格嵌套 Flex", {
                id: "nested-cjk-label",
                style: {
                  fontFamily: notoFont,
                  fontSize: 12,
                  color: "textInk",
                  align: "center",
                },
              }),
            ),
          ),
          box({
            id: "right-flex",
            style: {
              display: "flex",
              flexDirection: "column",
              gap: 8,
              padding: 10,
              backgroundColor: "cardBg",
              borderColor: "#94A3B8",
              borderWidth: 1,
            },
          }).append(
            text("Outer Theme Preserved", {
              id: "right-flex-title",
              style: {
                fontFamily: jostFont,
                fontSize: 14,
                fontWeight: "bold",
                color: "textInk",
              },
            }),
            box({
              id: "auto-rows-grid",
              style: {
                display: "grid",
                gridTemplateColumns: "1fr 1fr",
                gridAutoRows: "36px auto",
                gap: 8,
              },
            }).append(
              shape({
                id: "auto-r1-c1",
                fill: "badge",
                style: { height: "100%" },
              }),
              shape({
                id: "auto-r1-c2",
                fill: "#38BDF8",
                style: { height: "100%" },
              }),
              box({
                id: "auto-r2-c1",
                style: {
                  padding: 6,
                  backgroundColor: "#FFFFFF",
                  borderColor: "#CBD5E1",
                  borderWidth: 1,
                },
              }).append(
                text("Auto row wraps nicely here", {
                  id: "auto-r2-c1-text",
                  style: {
                    fontFamily: jostFont,
                    fontSize: 12,
                    color: "textInk",
                  },
                }),
              ),
              shape({
                id: "auto-r2-c2",
                fill: "#0284C7",
                style: { height: "100%" },
              }),
            ),
          ),
        ),
      ),
      slide({
        id: "grid-overflow-warnings",
        background: "#FEF2F2",
        style: {
          display: "flex",
          flexDirection: "column",
          padding: 20,
          gap: 14,
        },
      }).append(
        box({
          id: "container-overflow-grid",
          style: {
            display: "grid",
            width: 280,
            height: 110,
            gridTemplateColumns: "180px 150px",
            gridTemplateRows: "70px 60px",
            gap: 10,
            padding: 10,
            backgroundColor: "#FFFFFF",
            borderColor: "#EF4444",
            borderWidth: 2,
          },
        }).append(
          shape({
            id: "ov-c1",
            fill: "#F87171",
            style: { height: "100%" },
          }),
          shape({
            id: "ov-c2",
            fill: "#FB923C",
            style: { height: "100%" },
          }),
        ),
        box({
          id: "item-overflow-grid",
          style: {
            display: "grid",
            width: 420,
            height: 150,
            gridTemplateColumns: "160px 1fr",
            gridTemplateRows: "80px",
            padding: 12,
            gap: 12,
            backgroundColor: "#FFFFFF",
            borderColor: "#DC2626",
            borderWidth: 2,
          },
        }).append(
          shape({
            id: "tall-item",
            fill: "#FCA5A5",
            width: 200,
            height: 110,
          }),
          shape({
            id: "normal-item",
            fill: "#34D399",
            height: 60,
          }),
        ),
      ),
    );
  },
  assert({ author, preview, validation }) {
    assert.equal(preview.pages.length, 3, author);
    assert.deepEqual(validation.errors, [], author);

    // Page 1: px/auto/fr columns & rows, span placement, text wrapping driving auto row height, and percentage sizing.
    const page1 = preview.pages[0]!;
    assert.deepEqual(page1.diagnostics, [], author);
    assert.ok(page1.png !== undefined, author);

    const s1 = preview.build.query.getSlide("grid-tracks-span-wrap");
    const headerBox = s1.get("grid-header").observe().box!;
    const sidebarBox = s1.get("sidebar-fixed").observe().box!;
    const autoColBox = s1.get("auto-col-card").observe().box!;
    const wrapCardBox = s1.get("wrap-fr-card").observe().box!;
    const wrapTextBox = s1
      .get("wrap-fr-card")
      .get("wrap-fr-text")
      .observe().box!;
    const spanPanelBox = s1.get("span-panel").observe().box!;
    const percentShapeBox = s1.get("percent-cell-shape").observe().box!;

    // Total inner width = 640 - 40 = 600; spanning all 3 columns fills 600px.
    assert.equal(headerBox.width, 600, author);
    assert.equal(sidebarBox.width, 120, author);
    assert.ok(autoColBox.width > 60 && autoColBox.width < 150, author);
    // 1fr column takes remaining width: 600 - 120 - autoColBox.width - 2*12.
    assert.ok(
      Math.abs(wrapCardBox.width - (600 - 120 - autoColBox.width - 24)) < 0.5,
      author,
    );
    // Wrapped multi-line text height drives row 2 height; all 3 cards stretch to identical height.
    assert.ok(wrapTextBox.height > 30, author);
    assert.equal(sidebarBox.height, wrapCardBox.height, author);
    assert.equal(autoColBox.height, wrapCardBox.height, author);
    // Spanning 2 columns matches col1 + gap + col2.
    assert.ok(
      Math.abs(
        spanPanelBox.width - (sidebarBox.width + 12 + autoColBox.width),
      ) < 0.5,
      author,
    );
    // Percentage child in col 3 row 3: 60% width of col 3 and 50% height of row 3, centered.
    assert.ok(
      Math.abs(percentShapeBox.width - wrapCardBox.width * 0.6) < 0.5,
      author,
    );
    assert.ok(
      Math.abs(percentShapeBox.height - spanPanelBox.height * 0.5) < 0.5,
      author,
    );
    assert.ok(percentShapeBox.x > wrapCardBox.x, author);
    assert.ok(percentShapeBox.y > spanPanelBox.y, author);

    // Page 2: Nested Grid <-> Flex Box, gridAutoRows cycling, and scoped ThemeScope isolation.
    const page2 = preview.pages[1]!;
    assert.deepEqual(page2.diagnostics, [], author);
    assert.ok(page2.png !== undefined, author);

    const s2 = preview.build.query.getSlide("grid-nested-flex-and-theme");
    const nestedGridBox = s2
      .get("outer-grid")
      .get("nested-grid")
      .observe().box!;
    const rightFlexBox = s2.get("outer-grid").get("right-flex").observe().box!;
    assert.ok(Math.abs(nestedGridBox.width - rightFlexBox.width) < 0.5, author);
    assert.equal(nestedGridBox.height, rightFlexBox.height, author);

    const autoR1C1 = s2
      .get("outer-grid")
      .get("right-flex")
      .get("auto-rows-grid")
      .get("auto-r1-c1")
      .observe().box!;
    const autoR2C1 = s2
      .get("outer-grid")
      .get("right-flex")
      .get("auto-rows-grid")
      .get("auto-r2-c1")
      .observe().box!;
    assert.equal(autoR1C1.height, 36, author);
    assert.ok(autoR2C1.height > 20, author);
    assert.ok(autoR2C1.y >= autoR1C1.y + 36 + 8 - 0.5, author);

    // Scoped theme overrides inside nested-grid must not leak to sibling right-flex.
    assert.ok(page2.svg?.includes('fill="#D97706"'), author);
    assert.ok(page2.svg?.includes('fill="#0EA5E9"'), author);

    // Page 3: grid-overflow and grid-item-overflow warnings with retained PNG & PPTX.
    const page3 = preview.pages[2]!;
    assert.deepEqual(page3.diagnostics, [], author);
    assert.ok(page3.png !== undefined, author);

    const s3 = preview.build.query.getSlide("grid-overflow-warnings");
    const ovC1 = s3.get("container-overflow-grid").get("ov-c1").observe().box!;
    const ovC2 = s3.get("container-overflow-grid").get("ov-c2").observe().box!;
    assert.equal(ovC1.width, 180, author);
    assert.equal(ovC1.height, 70, author);
    assert.equal(ovC2.width, 150, author);

    const tallItem = s3
      .get("item-overflow-grid")
      .get("tall-item")
      .observe().box!;
    assert.deepEqual(
      { width: tallItem.width, height: tallItem.height },
      { width: 200, height: 110 },
      author,
    );

    assert.ok(
      preview.build.report.diagnostics.some(
        (d) =>
          d.severity === "warning" &&
          d.code === "grid-overflow" &&
          d.path.includes("box[container-overflow-grid]"),
      ),
      author,
    );
    assert.ok(
      preview.build.report.diagnostics.some(
        (d) =>
          d.severity === "warning" &&
          d.code === "grid-item-overflow" &&
          d.path.includes("shape[tall-item]"),
      ),
      author,
    );
    const nativeTallItem = validation.slides[2]?.objects.find(
      (o) => o.boxEmu.width === 200 * 9525 && o.boxEmu.height === 110 * 9525,
    );
    assert.ok(nativeTallItem !== undefined, author);
  },
};

export const layoutGridRejectionCases: readonly LayoutRejectionCase[] = [
  {
    id: "grid-template-columns-on-block-container",
    expectedCode: "invalid-layout-context",
    expectedPathSubstring: "box[bad-grid-prop]",
    jsx() {
      return createElement(
        Deck,
        { width: 320, height: 180 },
        createElement(
          Slide,
          { id: "s1" },
          createElement(Box, {
            id: "bad-grid-prop",
            style: { gridTemplateColumns: "1fr 1fr" },
          }),
        ),
      );
    },
    document() {
      return deck({ width: 320, height: 180 }).append(
        slide({ id: "s1" }).append(
          box({
            id: "bad-grid-prop",
            style: { gridTemplateColumns: "1fr 1fr" },
          }),
        ),
      );
    },
  },
  {
    id: "grid-auto-rows-unsupported-track",
    expectedCode: "invalid-grid-tracks",
    expectedPathSubstring: "box[bad-auto-rows]",
    jsx() {
      return createElement(
        Deck,
        { width: 320, height: 180 },
        createElement(
          Slide,
          { id: "s1" },
          createElement(Box, {
            id: "bad-auto-rows",
            style: { display: "grid", gridAutoRows: "fit-content(100px)" },
          }),
        ),
      );
    },
    document() {
      return deck({ width: 320, height: 180 }).append(
        slide({ id: "s1" }).append(
          box({
            id: "bad-auto-rows",
            style: { display: "grid", gridAutoRows: "fit-content(100px)" },
          }),
        ),
      );
    },
  },
  {
    id: "grid-invalid-placement-zero",
    expectedCode: "invalid-grid-placement",
    expectedPathSubstring: "shape[bad-col]",
    jsx() {
      return createElement(
        Deck,
        { width: 320, height: 180 },
        createElement(
          Slide,
          { id: "s1" },
          createElement(
            Box,
            {
              id: "grid",
              style: { display: "grid", gridTemplateColumns: "1fr" },
            },
            createElement(Shape, {
              id: "bad-col",
              width: 40,
              height: 40,
              style: { gridColumn: 0 },
            }),
          ),
        ),
      );
    },
    document() {
      return deck({ width: 320, height: 180 }).append(
        slide({ id: "s1" }).append(
          box({
            id: "grid",
            style: { display: "grid", gridTemplateColumns: "1fr" },
          }).append(
            shape({
              id: "bad-col",
              width: 40,
              height: 40,
              style: { gridColumn: 0 },
            }),
          ),
        ),
      );
    },
  },
  {
    id: "grid-container-baseline-unsupported",
    expectedCode: "unsupported-grid-alignment",
    expectedPathSubstring: "box[grid-baseline]",
    jsx() {
      return createElement(
        Deck,
        { width: 320, height: 180 },
        createElement(
          Slide,
          { id: "s1" },
          createElement(Box, {
            id: "grid-baseline",
            style: { display: "grid", alignItems: "baseline" },
          }),
        ),
      );
    },
    document() {
      return deck({ width: 320, height: 180 }).append(
        slide({ id: "s1" }).append(
          box({
            id: "grid-baseline",
            style: { display: "grid", alignItems: "baseline" },
          }),
        ),
      );
    },
  },
  {
    id: "grid-item-baseline-unsupported",
    expectedCode: "unsupported-grid-alignment",
    expectedPathSubstring: "shape[item-baseline]",
    jsx() {
      return createElement(
        Deck,
        { width: 320, height: 180 },
        createElement(
          Slide,
          { id: "s1" },
          createElement(
            Box,
            {
              id: "grid",
              style: { display: "grid", gridTemplateColumns: "1fr" },
            },
            createElement(Shape, {
              id: "item-baseline",
              width: 40,
              height: 20,
              style: { alignSelf: "baseline" },
            }),
          ),
        ),
      );
    },
    document() {
      return deck({ width: 320, height: 180 }).append(
        slide({ id: "s1" }).append(
          box({
            id: "grid",
            style: { display: "grid", gridTemplateColumns: "1fr" },
          }).append(
            shape({
              id: "item-baseline",
              width: 40,
              height: 20,
              style: { alignSelf: "baseline" },
            }),
          ),
        ),
      );
    },
  },
  {
    id: "grid-absolute-item-placement-unsupported",
    expectedCode: "unsupported-grid-positioning",
    expectedPathSubstring: "shape[abs-placed]",
    jsx() {
      return createElement(
        Deck,
        { width: 320, height: 180 },
        createElement(
          Slide,
          { id: "s1" },
          createElement(
            Box,
            {
              id: "grid",
              style: { display: "grid", width: 200, height: 100 },
            },
            createElement(Shape, {
              id: "abs-placed",
              width: 40,
              height: 20,
              style: { position: "absolute", gridColumn: 1 },
            }),
          ),
        ),
      );
    },
    document() {
      return deck({ width: 320, height: 180 }).append(
        slide({ id: "s1" }).append(
          box({
            id: "grid",
            style: { display: "grid", width: 200, height: 100 },
          }).append(
            shape({
              id: "abs-placed",
              width: 40,
              height: 20,
              style: { position: "absolute", gridColumn: 1 },
            }),
          ),
        ),
      );
    },
  },
  {
    id: "grid-invalid-justify-items",
    expectedCode: "unsupported-layout-value",
    expectedPathSubstring: "box[bad-justify]",
    jsx() {
      return createElement(
        Deck,
        { width: 320, height: 180 },
        createElement(
          Slide,
          { id: "s1" },
          createElement(Box, {
            id: "bad-justify",
            style: {
              display: "grid",
              justifyItems: "space-between" as unknown as "start",
            },
          }),
        ),
      );
    },
    document() {
      return deck({ width: 320, height: 180 }).append(
        slide({ id: "s1" }).append(
          box({
            id: "bad-justify",
            style: {
              display: "grid",
              justifyItems: "space-between" as unknown as "start",
            },
          }),
        ),
      );
    },
  },
  {
    id: "grid-mismatched-track-parentheses",
    expectedCode: "invalid-grid-tracks",
    expectedPathSubstring: "box[bad-parens]",
    jsx() {
      return createElement(
        Deck,
        { width: 320, height: 180 },
        createElement(
          Slide,
          { id: "s1" },
          createElement(Box, {
            id: "bad-parens",
            style: {
              display: "grid",
              gridTemplateColumns: "repeat(2, 1fr",
            },
          }),
        ),
      );
    },
    document() {
      return deck({ width: 320, height: 180 }).append(
        slide({ id: "s1" }).append(
          box({
            id: "bad-parens",
            style: {
              display: "grid",
              gridTemplateColumns: "repeat(2, 1fr",
            },
          }),
        ),
      );
    },
  },
];
