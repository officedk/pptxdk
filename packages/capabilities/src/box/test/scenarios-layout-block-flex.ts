import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import { Box } from "@pptxdk/capabilities/box-elements";
import {
  box,
  deck,
  font,
  shape,
  slide,
  text,
} from "@pptxdk/capabilities/document";
import { Shape } from "@pptxdk/capabilities/shape-elements";
import { Text } from "@pptxdk/capabilities/text-elements";
import type { CommonNode } from "@pptxdk/core";
import { Deck, Slide } from "@pptxdk/react";
import { createElement, type ReactNode } from "react";

import type { CapabilityScenario } from "../../../test/acceptance-runner.ts";

const jostFont = font("Jost*", [
  {
    weight: 400,
    data: readFileSync(
      new URL("../../../test/assets/fonts/Jost-400-Book.ttf", import.meta.url),
    ),
  },
  {
    weight: 700,
    data: readFileSync(
      new URL("../../../test/assets/fonts/Jost-700-Bold.ttf", import.meta.url),
    ),
  },
]);

const notoFont = font("Noto Sans CJK SC", [
  {
    weight: 400,
    data: readFileSync(
      new URL(
        "../../../test/assets/fonts/NotoSansSC-Regular.woff2",
        import.meta.url,
      ),
    ),
  },
]);

export const layoutBlockFlexPositiveScenario: CapabilityScenario = {
  id: "layout-block-flex-semantics",
  capability: "layout:block-flex",
  description:
    "Dual-author Block and Flex layout scenario covering row/column flow, auto sizing, min-content automatic minimums, percentages, flexWrap, text/shape baseline alignment, absolute containing-block escaping, and canvas overflow warnings.",
  visualCheckpoints: [
    "Page 1 (block-flex-flow-and-auto-min): light #F8FAFC canvas; top white header row aligns auto-sized 'Block & Flex Layout' text on the left and a 30%-width blue badge ('30% Width') on the right; middle row displays 3 cards—left blue card protects its unbreakable 'MinimumContentGuard' width via automatic minimum, middle emerald card (minWidth: 0) shrinks and wraps text cleanly, right amber card takes 25% width; bottom gray block section stacks full-width column items.",
    "Page 2 (flex-wrap-and-baseline): white canvas; top-left bordered panel wraps four 110x42 colored chips across two rows (2 per row); top-right amber-bordered Flex row aligns baselines across 28px Text ('28px'), a nested Box wrapping 16px Text ('16px Boxed'), and 12px CJK/Latin Text ('12px 基线') alongside a centered teal Shape badge; bottom indigo-bordered relative box anchors an escaped absolute indigo Shape at its bottom-right corner and a dark slate Shape at its top-left.",
    "Page 3 (canvas-overflow-and-escape): pale rose #FEF2F2 canvas; dark slate Shape escapes an enclosing static Box to position at (32, 32) on the Slide; bottom-right coral Box ('Overflow retained') intentionally extends past the right/bottom canvas edges (clipped at PNG border while retaining full coordinates in PPTX and reporting an overflow warning).",
  ],
  nativeCheckpoints: [
    "Open exported PPTX in PowerPoint/WPS and verify identical geometry across JSX and /document builds, including clipped-at-slide-edge overflowing Box retained at its declared EMU bounds.",
  ],
  jsx() {
    return createElement(
      Deck,
      { width: 640, height: 360 },
      createElement(
        Slide,
        {
          id: "block-flex-flow-and-auto-min",
          background: "#F8FAFC",
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
            id: "header-row",
            style: {
              display: "flex",
              flexDirection: "row",
              justifyContent: "space-between",
              alignItems: "center",
              padding: [10, 14],
              background: "#FFFFFF",
              border: { color: "#CBD5E1", width: 1 },
            },
          },
          createElement(
            Text,
            {
              id: "header-title",
              style: {
                fontFamily: jostFont,
                fontSize: 18,
                fontWeight: "bold",
                color: "#0F172A",
                width: "auto",
                height: "auto",
              },
            },
            "Block & Flex Layout",
          ),
          createElement(
            Box,
            {
              id: "header-badge",
              style: {
                width: "30%",
                height: 30,
                display: "flex",
                justifyContent: "center",
                alignItems: "center",
                background: "#2563EB",
              },
            },
            createElement(
              Text,
              {
                id: "header-badge-label",
                style: {
                  fontFamily: jostFont,
                  fontSize: 13,
                  color: "#FFFFFF",
                  align: "center",
                  width: "auto",
                  height: "auto",
                },
              },
              "30% Width",
            ),
          ),
        ),
        createElement(
          Box,
          {
            id: "auto-min-row",
            style: {
              display: "flex",
              flexDirection: "row",
              gap: 12,
              width: "100%",
              alignItems: "stretch",
            },
          },
          createElement(
            Box,
            {
              id: "card-automin",
              style: {
                width: 240,
                flexShrink: 1,
                padding: 8,
                background: "#EFF6FF",
                border: { color: "#3B82F6", width: 2 },
              },
            },
            createElement(
              Text,
              {
                id: "automin-text",
                style: {
                  fontFamily: jostFont,
                  fontSize: 16,
                  color: "#1E3A8A",
                },
              },
              "MinimumContentGuard",
            ),
          ),
          createElement(
            Box,
            {
              id: "card-released",
              style: {
                flex: 1,
                minWidth: 0,
                padding: 8,
                background: "#ECFDF5",
                border: { color: "#10B981", width: 2 },
              },
            },
            createElement(
              Text,
              {
                id: "released-text",
                style: {
                  fontFamily: jostFont,
                  fontSize: 13,
                  lineHeight: 18,
                  color: "#065F46",
                },
              },
              "Released minWidth: 0 wraps inside remaining space",
            ),
          ),
          createElement(
            Box,
            {
              id: "card-percent",
              style: {
                width: "25%",
                padding: 8,
                background: "#FEF3C7",
                border: { color: "#D97706", width: 2 },
              },
            },
            createElement(
              Text,
              {
                id: "percent-text",
                style: {
                  fontFamily: jostFont,
                  fontSize: 14,
                  color: "#92400E",
                },
              },
              "25% card",
            ),
          ),
        ),
        createElement(
          Box,
          {
            id: "block-column-section",
            style: {
              padding: 10,
              gap: 6,
              background: "#E2E8F0",
              display: "flex",
              flexDirection: "column",
            },
          },
          createElement(
            Box,
            {
              id: "block-child-default",
              style: {
                padding: [6, 10],
                background: "#FFFFFF",
                border: { color: "#94A3B8", width: 1 },
              },
            },
            createElement(
              Text,
              {
                id: "block-child-text",
                style: {
                  fontFamily: notoFont,
                  fontSize: 14,
                  color: "#334155",
                },
              },
              "Default block child stretches to full container width",
            ),
          ),
          createElement(Shape, {
            id: "block-accent-bar",
            style: { height: 18, background: "#0F766E" },
          }),
        ),
      ),
      createElement(
        Slide,
        { id: "flex-wrap-and-baseline", background: "#FFFFFF" },
        createElement(
          Box,
          {
            id: "wrap-panel",
            style: {
              position: "absolute",
              left: 24,
              top: 24,
              width: 276,
              height: 196,
              display: "flex",
              flexDirection: "row",
              flexWrap: "wrap",
              alignContent: "flex-start",
              gap: 12,
              padding: 14,
              background: "#F8FAFC",
              border: { color: "#94A3B8", width: 1 },
            },
          },
          createElement(Shape, {
            id: "chip-1",
            style: { width: 112, height: 44, background: "#2563EB" },
          }),
          createElement(Shape, {
            id: "chip-2",
            style: { width: 112, height: 44, background: "#0F766E" },
          }),
          createElement(Shape, {
            id: "chip-3",
            style: { width: 112, height: 44, background: "#D97706" },
          }),
          createElement(Shape, {
            id: "chip-4",
            style: { width: 112, height: 44, background: "#7C3AED" },
          }),
        ),
        createElement(
          Box,
          {
            id: "baseline-row",
            style: {
              position: "absolute",
              left: 320,
              top: 24,
              width: 296,
              display: "flex",
              flexDirection: "row",
              alignItems: "baseline",
              gap: 10,
              padding: 12,
              background: "#FFFBEB",
              border: { color: "#F59E0B", width: 2 },
            },
          },
          createElement(
            Text,
            {
              id: "base-large",
              style: {
                fontFamily: jostFont,
                fontSize: 28,
                color: "#78350F",
                padding: 4,
                background: "#FFFFFF",
              },
            },
            "28px",
          ),
          createElement(
            Box,
            {
              id: "base-passthrough",
              style: {
                display: "flex",
                flexDirection: "row",
                alignItems: "flex-start",
              },
            },
            createElement(
              Text,
              {
                id: "base-boxed",
                style: {
                  fontFamily: jostFont,
                  fontSize: 16,
                  color: "#FFFFFF",
                  padding: 4,
                  background: "#0F766E",
                },
              },
              "16px Boxed",
            ),
          ),
          createElement(
            Text,
            {
              id: "base-small",
              style: {
                fontFamily: notoFont,
                fontSize: 12,
                color: "#0C4A6E",
                padding: 4,
                background: "#E0F2FE",
              },
            },
            "12px 基线",
          ),
          createElement(Shape, {
            id: "base-non-participant",
            style: {
              width: 18,
              height: 18,
              background: "#0F766E",
              alignSelf: "center",
            },
          }),
        ),
        createElement(
          Box,
          {
            id: "relative-cb",
            style: {
              position: "relative",
              left: 24,
              top: 244,
              width: 592,
              height: 92,
              background: "#EEF2FF",
              border: { color: "#6366F1", width: 1 },
              padding: 10,
            },
          },
          createElement(
            Box,
            {
              id: "static-nested",
              style: { padding: 4 },
            },
            createElement(Shape, {
              id: "escaped-to-relative",
              style: {
                position: "absolute",
                right: 14,
                bottom: 12,
                width: 150,
                height: 36,
                background: "#4F46E5",
              },
            }),
          ),
          createElement(Shape, {
            id: "anchored-top-left",
            style: {
              position: "absolute",
              left: 14,
              top: 12,
              width: 150,
              height: 36,
              background: "#1E293B",
            },
          }),
        ),
      ),
      createElement(
        Slide,
        { id: "canvas-overflow-and-escape", background: "#FEF2F2", style: {} },
        createElement(
          Box,
          {
            id: "static-outer-box",
            style: {
              marginLeft: 240,
              marginTop: 140,
              width: 140,
              height: 60,
              background: "#FEE2E2",
              border: { color: "#EF4444", width: 1 },
            },
          },
          createElement(Shape, {
            id: "escaped-to-slide",
            style: {
              position: "absolute",
              left: 32,
              top: 32,
              width: 180,
              height: 54,
              background: "#0F172A",
            },
          }),
        ),
        createElement(
          Box,
          {
            id: "overflowing-box",
            style: {
              position: "absolute",
              left: 520,
              top: 260,
              width: 180,
              height: 120,
              padding: 12,
              background: "#FCA5A5",
              border: { color: "#DC2626", width: 2 },
            },
          },
          createElement(
            Text,
            {
              id: "overflow-label",
              style: {
                fontFamily: jostFont,
                fontSize: 16,
                color: "#7F1D1D",
              },
            },
            "Overflow retained",
          ),
        ),
      ),
    );
  },
  document() {
    return deck({ width: 640, height: 360 }).append(
      slide({
        id: "block-flex-flow-and-auto-min",
        background: "#F8FAFC",
        style: {
          display: "flex",
          flexDirection: "column",
          padding: 20,
          gap: 14,
        },
      }).append(
        box({
          id: "header-row",
          style: {
            display: "flex",
            flexDirection: "row",
            justifyContent: "space-between",
            alignItems: "center",
            padding: [10, 14],
            background: "#FFFFFF",
            border: { color: "#CBD5E1", width: 1 },
          },
        }).append(
          text("Block & Flex Layout", {
            id: "header-title",
            style: {
              fontFamily: jostFont,
              fontSize: 18,
              fontWeight: "bold",
              color: "#0F172A",
              width: "auto",
              height: "auto",
            },
          }),
          box({
            id: "header-badge",
            style: {
              width: "30%",
              height: 30,
              display: "flex",
              justifyContent: "center",
              alignItems: "center",
              background: "#2563EB",
            },
          }).append(
            text("30% Width", {
              id: "header-badge-label",
              style: {
                fontFamily: jostFont,
                fontSize: 13,
                color: "#FFFFFF",
                align: "center",
                width: "auto",
                height: "auto",
              },
            }),
          ),
        ),
        box({
          id: "auto-min-row",
          style: {
            display: "flex",
            flexDirection: "row",
            gap: 12,
            width: "100%",
            alignItems: "stretch",
          },
        }).append(
          box({
            id: "card-automin",
            style: {
              width: 240,
              flexShrink: 1,
              padding: 8,
              background: "#EFF6FF",
              border: { color: "#3B82F6", width: 2 },
            },
          }).append(
            text("MinimumContentGuard", {
              id: "automin-text",
              style: {
                fontFamily: jostFont,
                fontSize: 16,
                color: "#1E3A8A",
              },
            }),
          ),
          box({
            id: "card-released",
            style: {
              flex: 1,
              minWidth: 0,
              padding: 8,
              background: "#ECFDF5",
              border: { color: "#10B981", width: 2 },
            },
          }).append(
            text("Released minWidth: 0 wraps inside remaining space", {
              id: "released-text",
              style: {
                fontFamily: jostFont,
                fontSize: 13,
                lineHeight: 18,
                color: "#065F46",
              },
            }),
          ),
          box({
            id: "card-percent",
            style: {
              width: "25%",
              padding: 8,
              background: "#FEF3C7",
              border: { color: "#D97706", width: 2 },
            },
          }).append(
            text("25% card", {
              id: "percent-text",
              style: {
                fontFamily: jostFont,
                fontSize: 14,
                color: "#92400E",
              },
            }),
          ),
        ),
        box({
          id: "block-column-section",
          style: {
            padding: 10,
            gap: 6,
            background: "#E2E8F0",
            display: "flex",
            flexDirection: "column",
          },
        }).append(
          box({
            id: "block-child-default",
            style: {
              padding: [6, 10],
              background: "#FFFFFF",
              border: { color: "#94A3B8", width: 1 },
            },
          }).append(
            text("Default block child stretches to full container width", {
              id: "block-child-text",
              style: {
                fontFamily: notoFont,
                fontSize: 14,
                color: "#334155",
              },
            }),
          ),
          shape({
            id: "block-accent-bar",
            style: { height: 18, background: "#0F766E" },
          }),
        ),
      ),
      slide({ id: "flex-wrap-and-baseline", background: "#FFFFFF" }).append(
        box({
          id: "wrap-panel",
          style: {
            position: "absolute",
            left: 24,
            top: 24,
            width: 276,
            height: 196,
            display: "flex",
            flexDirection: "row",
            flexWrap: "wrap",
            alignContent: "flex-start",
            gap: 12,
            padding: 14,
            background: "#F8FAFC",
            border: { color: "#94A3B8", width: 1 },
          },
        }).append(
          shape({
            id: "chip-1",
            style: { width: 112, height: 44, background: "#2563EB" },
          }),
          shape({
            id: "chip-2",
            style: { width: 112, height: 44, background: "#0F766E" },
          }),
          shape({
            id: "chip-3",
            style: { width: 112, height: 44, background: "#D97706" },
          }),
          shape({
            id: "chip-4",
            style: { width: 112, height: 44, background: "#7C3AED" },
          }),
        ),
        box({
          id: "baseline-row",
          style: {
            position: "absolute",
            left: 320,
            top: 24,
            width: 296,
            display: "flex",
            flexDirection: "row",
            alignItems: "baseline",
            gap: 10,
            padding: 12,
            background: "#FFFBEB",
            border: { color: "#F59E0B", width: 2 },
          },
        }).append(
          text("28px", {
            id: "base-large",
            style: {
              fontFamily: jostFont,
              fontSize: 28,
              color: "#78350F",
              padding: 4,
              background: "#FFFFFF",
            },
          }),
          box({
            id: "base-passthrough",
            style: {
              display: "flex",
              flexDirection: "row",
              alignItems: "flex-start",
            },
          }).append(
            text("16px Boxed", {
              id: "base-boxed",
              style: {
                fontFamily: jostFont,
                fontSize: 16,
                color: "#FFFFFF",
                padding: 4,
                background: "#0F766E",
              },
            }),
          ),
          text("12px 基线", {
            id: "base-small",
            style: {
              fontFamily: notoFont,
              fontSize: 12,
              color: "#0C4A6E",
              padding: 4,
              background: "#E0F2FE",
            },
          }),
          shape({
            id: "base-non-participant",
            style: {
              width: 18,
              height: 18,
              background: "#0F766E",
              alignSelf: "center",
            },
          }),
        ),
        box({
          id: "relative-cb",
          style: {
            position: "relative",
            left: 24,
            top: 244,
            width: 592,
            height: 92,
            background: "#EEF2FF",
            border: { color: "#6366F1", width: 1 },
            padding: 10,
          },
        }).append(
          box({
            id: "static-nested",
            style: { padding: 4 },
          }).append(
            shape({
              id: "escaped-to-relative",
              style: {
                position: "absolute",
                right: 14,
                bottom: 12,
                width: 150,
                height: 36,
                background: "#4F46E5",
              },
            }),
          ),
          shape({
            id: "anchored-top-left",
            style: {
              position: "absolute",
              left: 14,
              top: 12,
              width: 150,
              height: 36,
              background: "#1E293B",
            },
          }),
        ),
      ),
      slide({
        id: "canvas-overflow-and-escape",
        background: "#FEF2F2",
        style: {},
      }).append(
        box({
          id: "static-outer-box",
          style: {
            marginLeft: 240,
            marginTop: 140,
            width: 140,
            height: 60,
            background: "#FEE2E2",
            border: { color: "#EF4444", width: 1 },
          },
        }).append(
          shape({
            id: "escaped-to-slide",
            style: {
              position: "absolute",
              left: 32,
              top: 32,
              width: 180,
              height: 54,
              background: "#0F172A",
            },
          }),
        ),
        box({
          id: "overflowing-box",
          style: {
            position: "absolute",
            left: 520,
            top: 260,
            width: 180,
            height: 120,
            padding: 12,
            background: "#FCA5A5",
            border: { color: "#DC2626", width: 2 },
          },
        }).append(
          text("Overflow retained", {
            id: "overflow-label",
            style: {
              fontFamily: jostFont,
              fontSize: 16,
              color: "#7F1D1D",
            },
          }),
        ),
      ),
    );
  },
  assert({ author, resolution, preview, validation }) {
    assert.equal(preview.pages.length, 3, author);
    assert.deepEqual(validation.errors, [], author);

    // Page 1: Flex row/column, percentage badge, automatic minimum, and minWidth: 0 wrapping.
    const page1 = preview.pages[0]!;
    assert.deepEqual(page1.errors, [], author);
    assert.ok(page1.png !== undefined, author);
    const headerRow = resolution
      .slide("block-flex-flow-and-auto-min")
      .get("header-row");
    const badge = resolution
      .slide("block-flex-flow-and-auto-min")
      .get("header-row")
      .get("header-badge");
    assert.ok(
      Math.abs(badge.box!.width - (headerRow.box!.width - 28) * 0.3) < 0.5,
      author,
    );
    const autominCard = resolution
      .slide("block-flex-flow-and-auto-min")
      .get("auto-min-row")
      .get("card-automin");
    const releasedCard = resolution
      .slide("block-flex-flow-and-auto-min")
      .get("auto-min-row")
      .get("card-released");
    const percentCard = resolution
      .slide("block-flex-flow-and-auto-min")
      .get("auto-min-row")
      .get("card-percent");
    assert.ok(autominCard.box!.width > 160, author);
    assert.ok(releasedCard.box!.width > 80, author);
    assert.equal(percentCard.box!.width, 150, author);

    // Page 2: Flex wrap across rows, baseline alignment, and relative containing-block escape.
    const page2 = preview.pages[1]!;
    assert.deepEqual(page2.errors, [], author);
    assert.ok(page2.png !== undefined, author);
    const chip1 = resolution
      .slide("flex-wrap-and-baseline")
      .get("wrap-panel")
      .get("chip-1").box!;
    const chip2 = resolution
      .slide("flex-wrap-and-baseline")
      .get("wrap-panel")
      .get("chip-2").box!;
    const chip3 = resolution
      .slide("flex-wrap-and-baseline")
      .get("wrap-panel")
      .get("chip-3").box!;
    assert.equal(chip1.y, chip2.y, author);
    assert.ok(chip3.y > chip1.y, author);
    assert.equal(chip3.x, chip1.x, author);

    const baseLarge = resolution
      .slide("flex-wrap-and-baseline")
      .get("baseline-row")
      .get("base-large").box!;
    const baseShape = resolution
      .slide("flex-wrap-and-baseline")
      .get("baseline-row")
      .get("base-passthrough")
      .get("base-boxed").box!;
    const baseSmall = resolution
      .slide("flex-wrap-and-baseline")
      .get("baseline-row")
      .get("base-small").box!;
    assert.ok(baseShape.y > baseLarge.y, author);
    assert.ok(baseSmall.y > baseShape.y, author);

    const escapedRelative = resolution
      .slide("flex-wrap-and-baseline")
      .get("relative-cb")
      .get("static-nested")
      .get("escaped-to-relative").box!;
    // Anchored relative to relative-cb (x=24, y=244, width=592, height=92), right: 14, bottom: 12.
    assert.deepEqual(
      escapedRelative,
      {
        x: 24 + 592 - 14 - 150,
        y: 244 + 92 - 12 - 36,
        width: 150,
        height: 36,
      },
      author,
    );

    // Page 3: Static ancestor escape to Slide canvas + overflowing Box retained at declared coordinates.
    const page3 = preview.pages[2]!;
    assert.deepEqual(page3.errors, [], author);
    assert.ok(page3.png !== undefined, author);
    const escapedSlide = resolution
      .slide("canvas-overflow-and-escape")
      .get("static-outer-box")
      .get("escaped-to-slide").box!;
    assert.deepEqual(
      escapedSlide,
      { x: 32, y: 32, width: 180, height: 54 },
      author,
    );
    const overflowLocator = resolution
      .slide("canvas-overflow-and-escape")
      .get("overflowing-box");
    assert.deepEqual(
      overflowLocator.box,
      { x: 520, y: 260, width: 180, height: 120 },
      author,
    );
    assert.deepEqual(
      overflowLocator.bounds,
      { x: 520, y: 260, width: 180, height: 120 },
      author,
    );
    assert.ok(
      overflowLocator.bounds!.x + overflowLocator.bounds!.width >
        resolution.report.dimensions.width &&
        overflowLocator.bounds!.y + overflowLocator.bounds!.height >
          resolution.report.dimensions.height,
      author,
    );
    const nativeOverflow = validation.slides[2]?.objects.find(
      (o) => o.boxEmu.x === 520 * 9525 && o.boxEmu.y === 260 * 9525,
    );
    assert.ok(nativeOverflow !== undefined, author);
  },
};

export type LayoutRejectionCase = {
  readonly id: string;
  readonly expectedCode: string;
  readonly expectedPathSubstring: string;
  jsx(): ReactNode;
  document(): CommonNode;
};

export const layoutBlockFlexRejectionCases: readonly LayoutRejectionCase[] = [
  {
    id: "flex-direction-on-block-container",
    expectedCode: "invalid-structure",
    expectedPathSubstring: "box[bad-block]",
    jsx() {
      return createElement(
        Deck,
        { width: 320, height: 180 },
        createElement(
          Slide,
          { id: "s1" },
          createElement(Box, {
            id: "bad-block",
            style: { flexDirection: "row" },
          }),
        ),
      );
    },
    document() {
      return deck({ width: 320, height: 180 }).append(
        slide({ id: "s1" }).append(
          box({
            id: "bad-block",
            style: { flexDirection: "row" },
          }),
        ),
      );
    },
  },
  {
    id: "position-offset-on-static-element",
    expectedCode: "invalid-structure",
    expectedPathSubstring: "shape[static-offset]",
    jsx() {
      return createElement(
        Deck,
        { width: 320, height: 180 },
        createElement(
          Slide,
          { id: "s1" },
          createElement(Shape, {
            id: "static-offset",
            style: { width: 80, height: 40, left: 10 },
          }),
        ),
      );
    },
    document() {
      return deck({ width: 320, height: 180 }).append(
        slide({ id: "s1" }).append(
          shape({
            id: "static-offset",
            style: { width: 80, height: 40, left: 10 },
          }),
        ),
      );
    },
  },
  {
    id: "baseline-on-wrapping-flex-row",
    expectedCode: "invalid-structure",
    expectedPathSubstring: "box[wrap-baseline]",
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
              id: "wrap-baseline",
              style: {
                width: 200,
                display: "flex",
                flexDirection: "row",
                flexWrap: "wrap",
                alignItems: "baseline",
              },
            },
            createElement(
              Text,
              { style: { fontFamily: jostFont, fontSize: 16 } },
              "Wrap",
            ),
          ),
        ),
      );
    },
    document() {
      return deck({ width: 320, height: 180 }).append(
        slide({ id: "s1" }).append(
          box({
            id: "wrap-baseline",
            style: {
              width: 200,
              display: "flex",
              flexDirection: "row",
              flexWrap: "wrap",
              alignItems: "baseline",
            },
          }).append(
            text("Wrap", {
              style: { fontFamily: jostFont, fontSize: 16 },
            }),
          ),
        ),
      );
    },
  },
  {
    id: "multi-child-box-in-baseline-chain",
    expectedCode: "invalid-structure",
    expectedPathSubstring: "box[multi-child-box]",
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
              id: "base-row",
              style: {
                display: "flex",
                flexDirection: "row",
                alignItems: "baseline",
              },
            },
            createElement(
              Box,
              {
                id: "multi-child-box",
                style: { display: "flex", flexDirection: "row" },
              },
              createElement(
                Text,
                { style: { fontFamily: jostFont, fontSize: 14 } },
                "A",
              ),
              createElement(
                Text,
                { style: { fontFamily: jostFont, fontSize: 14 } },
                "B",
              ),
            ),
          ),
        ),
      );
    },
    document() {
      return deck({ width: 320, height: 180 }).append(
        slide({ id: "s1" }).append(
          box({
            id: "base-row",
            style: {
              display: "flex",
              flexDirection: "row",
              alignItems: "baseline",
            },
          }).append(
            box({
              id: "multi-child-box",
              style: { display: "flex", flexDirection: "row" },
            }).append(
              text("A", { style: { fontFamily: jostFont, fontSize: 14 } }),
              text("B", { style: { fontFamily: jostFont, fontSize: 14 } }),
            ),
          ),
        ),
      );
    },
  },
  {
    id: "min-width-greater-than-max-width",
    expectedCode: "invalid-style",
    expectedPathSubstring: "box[bad-bounds]",
    jsx() {
      return createElement(
        Deck,
        { width: 320, height: 180 },
        createElement(
          Slide,
          { id: "s1" },
          createElement(Box, {
            id: "bad-bounds",
            style: { minWidth: 120, maxWidth: 80 },
          }),
        ),
      );
    },
    document() {
      return deck({ width: 320, height: 180 }).append(
        slide({ id: "s1" }).append(
          box({
            id: "bad-bounds",
            style: { minWidth: 120, maxWidth: 80 },
          }),
        ),
      );
    },
  },
  {
    id: "non-text-shape-in-baseline-row",
    expectedCode: "invalid-structure",
    expectedPathSubstring: "shape[no-text-shape]",
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
              id: "base-row",
              style: {
                display: "flex",
                flexDirection: "row",
                alignItems: "baseline",
              },
            },
            createElement(Shape, {
              id: "no-text-shape",
              style: { width: 60, height: 30, background: "#2563EB" },
            }),
          ),
        ),
      );
    },
    document() {
      return deck({ width: 320, height: 180 }).append(
        slide({ id: "s1" }).append(
          box({
            id: "base-row",
            style: {
              display: "flex",
              flexDirection: "row",
              alignItems: "baseline",
            },
          }).append(
            shape({
              id: "no-text-shape",
              style: { width: 60, height: 30, background: "#2563EB" },
            }),
          ),
        ),
      );
    },
  },
];
