import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import { createMorphScope } from "@react-pptx/capabilities";
import { boxCapability } from "@react-pptx/capabilities/box";
import { Box } from "@react-pptx/capabilities/box-elements";
import {
  box,
  deck,
  font,
  group,
  shape,
  slide,
  text,
} from "@react-pptx/capabilities/document";
import { groupCapability } from "@react-pptx/capabilities/group";
import { Group } from "@react-pptx/capabilities/group-elements";
import { fade, type MotionObservation } from "@react-pptx/capabilities/motion";
import { path } from "@react-pptx/capabilities/shape";
import { shapeCapability } from "@react-pptx/capabilities/shape";
import { Shape } from "@react-pptx/capabilities/shape-elements";
import { Text } from "@react-pptx/capabilities/text-elements";
import { transitions } from "@react-pptx/capabilities/transition";
import { BuildError } from "@react-pptx/core";
import { createResolver } from "@react-pptx/core/build";
import { emit } from "@react-pptx/emit";
import { Deck, Slide, expandDocument } from "@react-pptx/react";
import { strFromU8, unzipSync } from "fflate";
import { createElement } from "react";
import { expect } from "vitest";

import type { CapabilityScenario } from "../acceptance-runner.ts";

const jostFont = font("Jost*", [
  {
    weight: 400,
    data: readFileSync(
      new URL("../../assets/fonts/Jost-400-Book.ttf", import.meta.url),
    ),
  },
]);

const notoFont = font("Noto Sans CJK SC", [
  {
    weight: 400,
    data: readFileSync(
      new URL("../../assets/fonts/NotoSansSC-Regular.woff2", import.meta.url),
    ),
  },
]);

const trianglePath = path("M0 0 L100 0 L0 100 Z", { viewBox: [100, 100] });

const morphScope = createMorphScope();
const groupMorphKey = morphScope.handle("group-morph-card");

export const groupPositiveScenario: CapabilityScenario = {
  id: "group-native-composition-and-transforms",
  capability: "rp:group",
  description:
    "Dual-author Group scenario covering loose and layout-participating Groups containing Text, Shape, Box, nested rotated Groups, flipped members, scoped ThemeScope, Motion members, Morph continuity, and slide links.",
  visualCheckpoints: [
    "Page 1 (group-hierarchy-and-transforms): light slate #F8FAFC canvas; top-left blue-bordered card (#EFF6FF fill, 2px #2563EB border) holds a #1D4ED8 badge shape and text '原生组合 Native Group' overlapped at top-right by an amber #F59E0B tag shape; on the right half, a 20deg clockwise rotated cluster contains an emerald green triangle (#10B981) and a nested -35deg rotated inner group holding a horizontally flipped pink triangle (#DB2777) above a dark slate #1E293B 'Nested' caption box.",
    "Page 2 (group-layout-morph-and-link): white canvas; light gray #F1F5F9 Flex container frames a -6deg tilted Group on the left (dark teal #0F766E rectangle beside mint #CCFBF1 'Morph & Link' text card) followed on the right by an amber #FEF3C7 sibling Box with text 'Sibling card follows layout-sized Group cleanly.'",
    "Page 3 (group-morph-target): white canvas; centered 8deg rotated Group shows a sky-blue #0284C7 rectangle beside a deep-blue #0369A1 'Morphed Group' card.",
  ],
  nativeCheckpoints: [
    "Open exported PPTX in PowerPoint/WPS and verify Groups remain native <p:grpSp> containers preserving member <p:sp> hierarchy, nested <p:grpSp> rotation, Morph matching name, and <a:hlinkClick> on Page 2 Group.",
  ],
  jsx() {
    return createElement(
      Deck,
      {
        width: 640,
        height: 360,
        theme: { colors: { badgeFill: "#991B1B" } },
      },
      createElement(
        Slide,
        { id: "group-hierarchy-and-transforms", background: "#F8FAFC" },
        createElement(
          Group,
          {
            id: "badge-group",
            theme: { colors: { badgeFill: "#1D4ED8" } },
            enter: fade({ duration: 400 }),
          },
          createElement(
            Box,
            {
              id: "badge-card",
              style: {
                position: "absolute",
                left: 28,
                top: 28,
                width: 240,
                padding: 12,
                display: "flex",
                flexDirection: "column",
                gap: 8,
                backgroundColor: "#EFF6FF",
                borderColor: "#2563EB",
                borderWidth: 2,
                borderStyle: "solid",
              },
            },
            createElement(Shape, {
              id: "badge-dot",
              width: 48,
              height: 22,
              fill: "badgeFill",
            }),
            createElement(
              Text,
              {
                id: "badge-title",
                style: { fontFamily: notoFont, fontSize: 16, color: "#0F172A" },
              },
              "原生组合 Native Group",
            ),
          ),
          createElement(Shape, {
            id: "badge-tag",
            x: 204,
            y: 40,
            width: 52,
            height: 24,
            fill: "#F59E0B",
          }),
        ),
        createElement(
          Group,
          {
            id: "rotated-outer",
            name: "Rotated Cluster",
            style: { rotate: 20 },
          },
          createElement(Shape, {
            id: "triangle-base",
            x: 350,
            y: 70,
            width: 110,
            height: 110,
            geometry: trianglePath,
            fill: "#10B981",
          }),
          createElement(
            Group,
            {
              id: "nested-inner",
              style: { rotate: -35 },
            },
            createElement(Shape, {
              id: "flipped-triangle",
              x: 470,
              y: 120,
              width: 85,
              height: 75,
              geometry: trianglePath,
              fill: "#DB2777",
              flipH: true,
            }),
            createElement(
              Text,
              {
                id: "inner-caption",
                x: 455,
                y: 208,
                width: 110,
                style: {
                  fontFamily: jostFont,
                  fontSize: 14,
                  color: "#FFFFFF",
                  backgroundColor: "#1E293B",
                  padding: 6,
                  align: "center",
                },
              },
              "Nested",
            ),
          ),
        ),
      ),
      createElement(
        Slide,
        { id: "group-layout-morph-and-link", background: "#FFFFFF" },
        createElement(
          Box,
          {
            id: "layout-row",
            style: {
              position: "absolute",
              left: 32,
              top: 40,
              width: 576,
              display: "flex",
              flexDirection: "row",
              gap: 24,
              alignItems: "flex-start",
              padding: 16,
              backgroundColor: "#F1F5F9",
              borderColor: "#CBD5E1",
              borderWidth: 1,
              borderStyle: "solid",
            },
          },
          createElement(
            Group,
            {
              id: "morph-group",
              morphKey: groupMorphKey,
              link: { id: "group-hierarchy-and-transforms" },
              style: {
                display: "flex",
                flexDirection: "row",
                gap: 10,
                width: 220,
                height: 110,
                padding: 12,
                rotate: -6,
              },
            },
            createElement(Shape, {
              id: "left-pill",
              width: 80,
              height: 86,
              fill: "#0F766E",
            }),
            createElement(
              Text,
              {
                id: "right-note",
                width: 106,
                style: {
                  fontFamily: jostFont,
                  fontSize: 15,
                  color: "#0F172A",
                  backgroundColor: "#CCFBF1",
                  padding: 8,
                },
              },
              "Morph & Link",
            ),
          ),
          createElement(
            Box,
            {
              id: "sibling-card",
              style: {
                width: 290,
                padding: 14,
                backgroundColor: "#FEF3C7",
                borderColor: "#D97706",
                borderWidth: 2,
                borderStyle: "solid",
              },
            },
            createElement(
              Text,
              {
                id: "sibling-copy",
                style: { fontFamily: jostFont, fontSize: 15, color: "#78350F" },
              },
              "Sibling card follows layout-sized Group cleanly.",
            ),
          ),
        ),
      ),
      createElement(
        Slide,
        {
          id: "group-morph-target",
          background: "#FFFFFF",
          transition: transitions.morph({ fallback: transitions.fade() }),
        },
        createElement(
          Group,
          {
            id: "morph-group-end",
            morphKey: groupMorphKey,
            style: {
              position: "absolute",
              left: 150,
              top: 100,
              width: 340,
              height: 140,
              display: "flex",
              flexDirection: "row",
              gap: 14,
              padding: 16,
              rotate: 8,
            },
          },
          createElement(Shape, {
            id: "end-pill",
            width: 110,
            height: 108,
            fill: "#0284C7",
          }),
          createElement(
            Text,
            {
              id: "end-note",
              width: 184,
              style: {
                fontFamily: jostFont,
                fontSize: 18,
                color: "#FFFFFF",
                backgroundColor: "#0369A1",
                padding: 10,
              },
            },
            "Morphed Group",
          ),
        ),
      ),
    );
  },
  document() {
    return deck({
      width: 640,
      height: 360,
      theme: { colors: { badgeFill: "#991B1B" } },
    }).append(
      slide({
        id: "group-hierarchy-and-transforms",
        background: "#F8FAFC",
      }).append(
        group({
          id: "badge-group",
          theme: { colors: { badgeFill: "#1D4ED8" } },
          enter: fade({ duration: 400 }),
        }).append(
          box({
            id: "badge-card",
            style: {
              position: "absolute",
              left: 28,
              top: 28,
              width: 240,
              padding: 12,
              display: "flex",
              flexDirection: "column",
              gap: 8,
              backgroundColor: "#EFF6FF",
              borderColor: "#2563EB",
              borderWidth: 2,
              borderStyle: "solid",
            },
          }).append(
            shape({
              id: "badge-dot",
              width: 48,
              height: 22,
              fill: "badgeFill",
            }),
            text("原生组合 Native Group", {
              id: "badge-title",
              style: { fontFamily: notoFont, fontSize: 16, color: "#0F172A" },
            }),
          ),
          shape({
            id: "badge-tag",
            x: 204,
            y: 40,
            width: 52,
            height: 24,
            fill: "#F59E0B",
          }),
        ),
        group({
          id: "rotated-outer",
          name: "Rotated Cluster",
          style: { rotate: 20 },
        }).append(
          shape({
            id: "triangle-base",
            x: 350,
            y: 70,
            width: 110,
            height: 110,
            geometry: trianglePath,
            fill: "#10B981",
          }),
          group({
            id: "nested-inner",
            style: { rotate: -35 },
          }).append(
            shape({
              id: "flipped-triangle",
              x: 470,
              y: 120,
              width: 85,
              height: 75,
              geometry: trianglePath,
              fill: "#DB2777",
              flipH: true,
            }),
            text("Nested", {
              id: "inner-caption",
              x: 455,
              y: 208,
              width: 110,
              style: {
                fontFamily: jostFont,
                fontSize: 14,
                color: "#FFFFFF",
                backgroundColor: "#1E293B",
                padding: 6,
                align: "center",
              },
            }),
          ),
        ),
      ),
      slide({
        id: "group-layout-morph-and-link",
        background: "#FFFFFF",
      }).append(
        box({
          id: "layout-row",
          style: {
            position: "absolute",
            left: 32,
            top: 40,
            width: 576,
            display: "flex",
            flexDirection: "row",
            gap: 24,
            alignItems: "flex-start",
            padding: 16,
            backgroundColor: "#F1F5F9",
            borderColor: "#CBD5E1",
            borderWidth: 1,
            borderStyle: "solid",
          },
        }).append(
          group({
            id: "morph-group",
            morphKey: groupMorphKey,
            link: { id: "group-hierarchy-and-transforms" },
            style: {
              display: "flex",
              flexDirection: "row",
              gap: 10,
              width: 220,
              height: 110,
              padding: 12,
              rotate: -6,
            },
          }).append(
            shape({
              id: "left-pill",
              width: 80,
              height: 86,
              fill: "#0F766E",
            }),
            text("Morph & Link", {
              id: "right-note",
              width: 106,
              style: {
                fontFamily: jostFont,
                fontSize: 15,
                color: "#0F172A",
                backgroundColor: "#CCFBF1",
                padding: 8,
              },
            }),
          ),
          box({
            id: "sibling-card",
            style: {
              width: 290,
              padding: 14,
              backgroundColor: "#FEF3C7",
              borderColor: "#D97706",
              borderWidth: 2,
              borderStyle: "solid",
            },
          }).append(
            text("Sibling card follows layout-sized Group cleanly.", {
              id: "sibling-copy",
              style: { fontFamily: jostFont, fontSize: 15, color: "#78350F" },
            }),
          ),
        ),
      ),
      slide({
        id: "group-morph-target",
        background: "#FFFFFF",
        transition: transitions.morph({ fallback: transitions.fade() }),
      }).append(
        group({
          id: "morph-group-end",
          morphKey: groupMorphKey,
          style: {
            position: "absolute",
            left: 150,
            top: 100,
            width: 340,
            height: 140,
            display: "flex",
            flexDirection: "row",
            gap: 14,
            padding: 16,
            rotate: 8,
          },
        }).append(
          shape({
            id: "end-pill",
            width: 110,
            height: 108,
            fill: "#0284C7",
          }),
          text("Morphed Group", {
            id: "end-note",
            width: 184,
            style: {
              fontFamily: jostFont,
              fontSize: 18,
              color: "#FFFFFF",
              backgroundColor: "#0369A1",
              padding: 10,
            },
          }),
        ),
      ),
    );
  },
  assert({ resolution, pptx, preview, validation }) {
    assert.deepEqual(resolution.report.diagnostics, []);
    assert.equal(preview.pages.length, 3);
    for (const page of preview.pages) {
      assert.deepEqual(page.diagnostics, []);
      assert.ok(page.png !== undefined);
    }

    // Page 1: verify Query hierarchy, transformed screenGeometry, and Motion group members.
    const page1 = preview.pages[0]!;
    const badgeGroupQuery = resolution.query
      .getSlide("group-hierarchy-and-transforms")
      .get("badge-group");
    const badgeObjects = badgeGroupQuery.observe().objects;
    assert.equal(badgeObjects[0]?.kind, "group");
    assert.deepEqual(
      badgeObjects.map((o) => o.authorId),
      ["badge-group", "badge-card", "badge-dot", "badge-title", "badge-tag"],
    );

    const outerNode = page1.nodes.find((n) => n.authorId === "rotated-outer")!;
    const innerNode = page1.nodes.find((n) => n.authorId === "nested-inner")!;
    const flippedNode = page1.nodes.find(
      (n) => n.authorId === "flipped-triangle",
    )!;
    assert.ok(outerNode.screenBox !== undefined);
    assert.ok(outerNode.screenBox.width > outerNode.box!.width);
    assert.ok(innerNode.corners?.length === 4);
    assert.ok(flippedNode.screenBox !== undefined);

    // Motion capability observation retains group target and native animation target id.
    const slide1Motion = resolution.report.slides[0]?.capabilities?.[
      "motion"
    ] as MotionObservation;
    const badgeEffect = slide1Motion.steps
      .flatMap((step) => step.effects)
      .find((e) => e.target.authorId === "badge-group");
    assert.equal(badgeEffect?.kind, "entrance");
    assert.equal(badgeEffect?.target.id, badgeObjects[0]?.id);
    assert.equal(badgeEffect?.preset, "fade");
    assert.equal(
      validation.slides[0]?.animations?.[0]?.effects.some(
        (e) => e.targetId === badgeObjects[0]?.id,
      ),
      true,
    );

    // Verify independent PPTX structure: <p:grpSp>, nested <p:grpSp>, morph name, and hyperlink.
    const fileSlide1 = validation.slides[0];
    const fileBadgeGroup = fileSlide1.objects.find(
      (obj) => obj.kind === "group" && obj.members?.length === 4,
    );
    assert.ok(fileBadgeGroup !== undefined);
    const fileRotatedOuter = fileSlide1.objects.find(
      (obj) => obj.name === "Rotated Cluster",
    );
    assert.equal(fileRotatedOuter?.kind, "group");
    assert.equal(fileRotatedOuter?.members?.[1]?.kind, "group");
    assert.equal(fileRotatedOuter?.members?.[1]?.members?.length, 2);

    const page2 = preview.pages[1]!;
    const morphGroupNode = page2.nodes.find(
      (n) => n.authorId === "morph-group",
    )!;
    const siblingCardNode = page2.nodes.find(
      (n) => n.authorId === "sibling-card",
    )!;
    assert.deepEqual(morphGroupNode.box, {
      x: 48,
      y: 56,
      width: 220,
      height: 110,
    });
    assert.equal(siblingCardNode.box?.x, 292);

    const unzipped = unzipSync(pptx);
    const slide1Xml = strFromU8(unzipped["ppt/slides/slide1.xml"]!);
    const slide2Xml = strFromU8(unzipped["ppt/slides/slide2.xml"]!);
    const slide3Xml = strFromU8(unzipped["ppt/slides/slide3.xml"]!);
    expect(slide1Xml).toContain("<p:grpSp>");
    expect(slide1Xml).toContain("1D4ED8");
    expect(slide2Xml).toContain("!!react-pptx-morph-");
    expect(slide2Xml).toContain("<a:hlinkClick");
    expect(slide3Xml).toContain("!!react-pptx-morph-");
  },
};

export const groupRejectionScenario: CapabilityScenario = {
  id: "group-preview-and-continuity-rejections",
  capability: "rp:group",
  description:
    "Dual-author Group rejection scenario verifying member preview failure propagation, morph-inside-group restriction, empty-group rejection, and missing member capability rejection.",
  visualCheckpoints: [
    "Page 1 (group-member-preview-failure): fails static preview (no PNG delivered) with located diagnostic pointing to group[bad-preview-group]/shape[unsupported-member] while retaining formal PPTX.",
    "Page 2 (group-valid-sibling): white canvas with valid teal #0D9488 rectangle inside Group.",
  ],
  nativeCheckpoints: [
    "Formal PPTX retains both slides as valid native <p:grpSp> objects despite static preview rejecting boxShadow on Page 1.",
  ],
  jsx() {
    return createElement(
      Deck,
      { width: 360, height: 200 },
      createElement(
        Slide,
        { id: "group-member-preview-failure" },
        createElement(
          Group,
          { id: "bad-preview-group", style: { rotate: 12 } },
          createElement(Shape, {
            id: "ok-member",
            x: 20,
            y: 20,
            width: 60,
            height: 40,
            fill: "#10B981",
          }),
          createElement(Shape, {
            id: "unsupported-member",
            x: 100,
            y: 20,
            width: 60,
            height: 40,
            fill: "#EF4444",
            style: { boxShadow: { blur: 4, distance: 2 } },
          }),
        ),
      ),
      createElement(
        Slide,
        { id: "group-valid-sibling", background: "#FFFFFF" },
        createElement(
          Group,
          { id: "valid-group" },
          createElement(Shape, {
            id: "valid-shape",
            x: 24,
            y: 24,
            width: 100,
            height: 60,
            fill: "#0D9488",
          }),
        ),
      ),
    );
  },
  document() {
    return deck({ width: 360, height: 200 }).append(
      slide({ id: "group-member-preview-failure" }).append(
        group({ id: "bad-preview-group", style: { rotate: 12 } }).append(
          shape({
            id: "ok-member",
            x: 20,
            y: 20,
            width: 60,
            height: 40,
            fill: "#10B981",
          }),
          shape({
            id: "unsupported-member",
            x: 100,
            y: 20,
            width: 60,
            height: 40,
            fill: "#EF4444",
            style: { boxShadow: { blur: 4, distance: 2 } },
          }),
        ),
      ),
      slide({ id: "group-valid-sibling", background: "#FFFFFF" }).append(
        group({ id: "valid-group" }).append(
          shape({
            id: "valid-shape",
            x: 24,
            y: 24,
            width: 100,
            height: 60,
            fill: "#0D9488",
          }),
        ),
      ),
    );
  },
  async assert({ preview }) {
    const failedPage = preview.pages[0]!;
    assert.equal(failedPage.png, undefined);
    assert.equal(failedPage.svg, undefined);
    assert.ok(
      failedPage.diagnostics.some(
        (d) =>
          d.severity === "error" &&
          /group\[bad-preview-group\]\/shape\[unsupported-member\]/u.test(
            d.path,
          ) &&
          /boxShadow/i.test(d.message),
      ),
    );
    assert.ok(
      failedPage.nodes.find((n) => n.authorId === "unsupported-member")
        ?.screenBox !== undefined,
    );

    const validPage = preview.pages[1]!;
    assert.deepEqual(validPage.diagnostics, []);
    assert.ok(validPage.png !== undefined);

    // Build-time continuity, empty-group, and missing-capability rejections via formal builder.
    const trimmedResolve = createResolver([
      groupCapability,
      boxCapability,
      shapeCapability,
    ]);
    const trimmedBuilder = async (
      element: Parameters<typeof expandDocument>[0],
    ) => emit(trimmedResolve(await expandDocument(element)));
    await assert.rejects(
      trimmedBuilder(
        createElement(
          Deck,
          { width: 360, height: 200 },
          createElement(
            Slide,
            { id: "morph-member" },
            createElement(
              Group,
              { id: "g" },
              createElement(
                Box,
                { style: { width: 100, height: 60 } },
                createElement(Shape, {
                  id: "m",
                  morphKey: groupMorphKey,
                  width: 40,
                  height: 30,
                }),
              ),
            ),
          ),
        ),
      ),
      (error: unknown) => {
        assert(error instanceof BuildError);
        assert.ok(
          error.diagnostics.some((d) => d.code === "morph-inside-group"),
        );
        return true;
      },
    );
    await assert.rejects(
      trimmedBuilder(
        createElement(
          Deck,
          { width: 360, height: 200 },
          createElement(
            Slide,
            { id: "empty" },
            createElement(Group, { id: "empty-group" }),
          ),
        ),
      ),
      (error: unknown) => {
        assert(error instanceof BuildError);
        assert.ok(error.diagnostics.some((d) => d.code === "empty-group"));
        return true;
      },
    );
    await assert.rejects(
      trimmedBuilder(
        createElement(
          Deck,
          { width: 360, height: 200 },
          createElement(
            Slide,
            { id: "missing-member-cap" },
            createElement(
              Group,
              { id: "g" },
              createElement(Text, { x: 10, y: 10, width: 80 }, "Missing Text"),
            ),
          ),
        ),
      ),
      (error: unknown) => {
        assert(error instanceof BuildError);
        assert.ok(
          error.diagnostics.some((d) => d.code === "missing-capability"),
        );
        return true;
      },
    );
  },
};
