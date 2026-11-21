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

import type { CapabilityScenario } from "../acceptance-runner.ts";

const jostFont = font("Jost*", [
  {
    weight: 400,
    data: readFileSync(
      new URL("../assets/fonts/Jost-400-Book.ttf", import.meta.url),
    ),
  },
  {
    weight: 700,
    data: readFileSync(
      new URL("../assets/fonts/Jost-700-Bold.ttf", import.meta.url),
    ),
  },
]);

const notoFont = font("Noto Sans CJK SC", [
  {
    weight: 400,
    data: readFileSync(
      new URL("../assets/fonts/NotoSansSC-Regular.woff2", import.meta.url),
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
              backgroundColor: "#FFFFFF",
              borderColor: "#CBD5E1",
              borderWidth: 1,
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
                backgroundColor: "#2563EB",
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
                backgroundColor: "#EFF6FF",
                borderColor: "#3B82F6",
                borderWidth: 2,
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
                backgroundColor: "#ECFDF5",
                borderColor: "#10B981",
                borderWidth: 2,
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
                backgroundColor: "#FEF3C7",
                borderColor: "#D97706",
                borderWidth: 2,
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
              backgroundColor: "#E2E8F0",
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
                backgroundColor: "#FFFFFF",
                borderColor: "#94A3B8",
                borderWidth: 1,
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
            height: 18,
            fill: "#0F766E",
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
              backgroundColor: "#F8FAFC",
              borderColor: "#94A3B8",
              borderWidth: 1,
            },
          },
          createElement(Shape, {
            id: "chip-1",
            width: 112,
            height: 44,
            fill: "#2563EB",
          }),
          createElement(Shape, {
            id: "chip-2",
            width: 112,
            height: 44,
            fill: "#0F766E",
          }),
          createElement(Shape, {
            id: "chip-3",
            width: 112,
            height: 44,
            fill: "#D97706",
          }),
          createElement(Shape, {
            id: "chip-4",
            width: 112,
            height: 44,
            fill: "#7C3AED",
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
              backgroundColor: "#FFFBEB",
              borderColor: "#F59E0B",
              borderWidth: 2,
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
                backgroundColor: "#FFFFFF",
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
                  backgroundColor: "#0F766E",
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
                backgroundColor: "#E0F2FE",
              },
            },
            "12px 基线",
          ),
          createElement(Shape, {
            id: "base-non-participant",
            width: 18,
            height: 18,
            fill: "#0F766E",
            style: { alignSelf: "center" },
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
              backgroundColor: "#EEF2FF",
              borderColor: "#6366F1",
              borderWidth: 1,
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
              fill: "#4F46E5",
              style: {
                position: "absolute",
                right: 14,
                bottom: 12,
                width: 150,
                height: 36,
              },
            }),
          ),
          createElement(Shape, {
            id: "anchored-top-left",
            fill: "#1E293B",
            style: {
              position: "absolute",
              left: 14,
              top: 12,
              width: 150,
              height: 36,
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
              backgroundColor: "#FEE2E2",
              borderColor: "#EF4444",
              borderWidth: 1,
            },
          },
          createElement(Shape, {
            id: "escaped-to-slide",
            fill: "#0F172A",
            style: {
              position: "absolute",
              left: 32,
              top: 32,
              width: 180,
              height: 54,
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
              backgroundColor: "#FCA5A5",
              borderColor: "#DC2626",
              borderWidth: 2,
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
            backgroundColor: "#FFFFFF",
            borderColor: "#CBD5E1",
            borderWidth: 1,
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
              backgroundColor: "#2563EB",
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
              backgroundColor: "#EFF6FF",
              borderColor: "#3B82F6",
              borderWidth: 2,
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
              backgroundColor: "#ECFDF5",
              borderColor: "#10B981",
              borderWidth: 2,
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
              backgroundColor: "#FEF3C7",
              borderColor: "#D97706",
              borderWidth: 2,
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
            backgroundColor: "#E2E8F0",
            display: "flex",
            flexDirection: "column",
          },
        }).append(
          box({
            id: "block-child-default",
            style: {
              padding: [6, 10],
              backgroundColor: "#FFFFFF",
              borderColor: "#94A3B8",
              borderWidth: 1,
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
            height: 18,
            fill: "#0F766E",
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
            backgroundColor: "#F8FAFC",
            borderColor: "#94A3B8",
            borderWidth: 1,
          },
        }).append(
          shape({ id: "chip-1", width: 112, height: 44, fill: "#2563EB" }),
          shape({ id: "chip-2", width: 112, height: 44, fill: "#0F766E" }),
          shape({ id: "chip-3", width: 112, height: 44, fill: "#D97706" }),
          shape({ id: "chip-4", width: 112, height: 44, fill: "#7C3AED" }),
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
            backgroundColor: "#FFFBEB",
            borderColor: "#F59E0B",
            borderWidth: 2,
          },
        }).append(
          text("28px", {
            id: "base-large",
            style: {
              fontFamily: jostFont,
              fontSize: 28,
              color: "#78350F",
              padding: 4,
              backgroundColor: "#FFFFFF",
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
                backgroundColor: "#0F766E",
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
              backgroundColor: "#E0F2FE",
            },
          }),
          shape({
            id: "base-non-participant",
            width: 18,
            height: 18,
            fill: "#0F766E",
            style: { alignSelf: "center" },
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
            backgroundColor: "#EEF2FF",
            borderColor: "#6366F1",
            borderWidth: 1,
            padding: 10,
          },
        }).append(
          box({
            id: "static-nested",
            style: { padding: 4 },
          }).append(
            shape({
              id: "escaped-to-relative",
              fill: "#4F46E5",
              style: {
                position: "absolute",
                right: 14,
                bottom: 12,
                width: 150,
                height: 36,
              },
            }),
          ),
          shape({
            id: "anchored-top-left",
            fill: "#1E293B",
            style: {
              position: "absolute",
              left: 14,
              top: 12,
              width: 150,
              height: 36,
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
            backgroundColor: "#FEE2E2",
            borderColor: "#EF4444",
            borderWidth: 1,
          },
        }).append(
          shape({
            id: "escaped-to-slide",
            fill: "#0F172A",
            style: {
              position: "absolute",
              left: 32,
              top: 32,
              width: 180,
              height: 54,
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
            backgroundColor: "#FCA5A5",
            borderColor: "#DC2626",
            borderWidth: 2,
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
    assert.deepEqual(page1.diagnostics, [], author);
    assert.ok(page1.png !== undefined, author);
    const headerRow = resolution.query
      .getSlide("block-flex-flow-and-auto-min")
      .get("header-row")
      .observe();
    const badge = resolution.query
      .getSlide("block-flex-flow-and-auto-min")
      .get("header-row")
      .get("header-badge")
      .observe();
    assert.ok(
      Math.abs(badge.box!.width - (headerRow.box!.width - 28) * 0.3) < 0.5,
      author,
    );
    const autominCard = resolution.query
      .getSlide("block-flex-flow-and-auto-min")
      .get("auto-min-row")
      .get("card-automin")
      .observe();
    const releasedCard = resolution.query
      .getSlide("block-flex-flow-and-auto-min")
      .get("auto-min-row")
      .get("card-released")
      .observe();
    const percentCard = resolution.query
      .getSlide("block-flex-flow-and-auto-min")
      .get("auto-min-row")
      .get("card-percent")
      .observe();
    assert.ok(autominCard.box!.width > 160, author);
    assert.ok(releasedCard.box!.width > 80, author);
    assert.equal(percentCard.box!.width, 150, author);

    // Page 2: Flex wrap across rows, baseline alignment, and relative containing-block escape.
    const page2 = preview.pages[1]!;
    assert.deepEqual(page2.diagnostics, [], author);
    assert.ok(page2.png !== undefined, author);
    const chip1 = resolution.query
      .getSlide("flex-wrap-and-baseline")
      .get("wrap-panel")
      .get("chip-1")
      .observe().box!;
    const chip2 = resolution.query
      .getSlide("flex-wrap-and-baseline")
      .get("wrap-panel")
      .get("chip-2")
      .observe().box!;
    const chip3 = resolution.query
      .getSlide("flex-wrap-and-baseline")
      .get("wrap-panel")
      .get("chip-3")
      .observe().box!;
    assert.equal(chip1.y, chip2.y, author);
    assert.ok(chip3.y > chip1.y, author);
    assert.equal(chip3.x, chip1.x, author);

    const baseLarge = resolution.query
      .getSlide("flex-wrap-and-baseline")
      .get("baseline-row")
      .get("base-large")
      .observe().box!;
    const baseShape = resolution.query
      .getSlide("flex-wrap-and-baseline")
      .get("baseline-row")
      .get("base-passthrough")
      .get("base-boxed")
      .observe().box!;
    const baseSmall = resolution.query
      .getSlide("flex-wrap-and-baseline")
      .get("baseline-row")
      .get("base-small")
      .observe().box!;
    assert.ok(baseShape.y > baseLarge.y, author);
    assert.ok(baseSmall.y > baseShape.y, author);

    const escapedRelative = resolution.query
      .getSlide("flex-wrap-and-baseline")
      .get("relative-cb")
      .get("static-nested")
      .get("escaped-to-relative")
      .observe().box!;
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

    // Page 3: Static ancestor escape to Slide canvas + canvas overflow warning with retained PNG & PPTX.
    const page3 = preview.pages[2]!;
    assert.deepEqual(page3.diagnostics, [], author);
    assert.ok(page3.png !== undefined, author);
    const escapedSlide = resolution.query
      .getSlide("canvas-overflow-and-escape")
      .get("static-outer-box")
      .get("escaped-to-slide")
      .observe().box!;
    assert.deepEqual(
      escapedSlide,
      { x: 32, y: 32, width: 180, height: 54 },
      author,
    );
    const overflowBox = resolution.query
      .getSlide("canvas-overflow-and-escape")
      .get("overflowing-box")
      .observe().box!;
    assert.deepEqual(
      overflowBox,
      { x: 520, y: 260, width: 180, height: 120 },
      author,
    );
    assert.ok(
      resolution.report.diagnostics.some(
        (d) =>
          d.severity === "warning" &&
          d.code === "overflow" &&
          d.path.includes("box[overflowing-box]"),
      ),
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
    expectedCode: "invalid-layout-context",
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
    expectedCode: "invalid-layout-context",
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
            width: 80,
            height: 40,
            style: { left: 10 },
          }),
        ),
      );
    },
    document() {
      return deck({ width: 320, height: 180 }).append(
        slide({ id: "s1" }).append(
          shape({
            id: "static-offset",
            width: 80,
            height: 40,
            style: { left: 10 },
          }),
        ),
      );
    },
  },
  {
    id: "baseline-on-wrapping-flex-row",
    expectedCode: "unsupported-baseline-context",
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
    expectedCode: "unsupported-baseline-context",
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
    expectedCode: "invalid-layout-style",
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
    expectedCode: "unsupported-baseline-context",
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
              width: 60,
              height: 30,
              fill: "#2563EB",
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
              width: 60,
              height: 30,
              fill: "#2563EB",
            }),
          ),
        ),
      );
    },
  },
];
