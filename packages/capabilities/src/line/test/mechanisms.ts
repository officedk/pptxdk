import assert from "node:assert/strict";

import { Box } from "@pptxdk/capabilities/box-elements";
import {
  box,
  deck,
  group,
  line,
  shape,
  slide,
} from "@pptxdk/capabilities/document";
import { Group } from "@pptxdk/capabilities/group-elements";
import { Line } from "@pptxdk/capabilities/line-elements";
import { Shape } from "@pptxdk/capabilities/shape-elements";
import { CommonNode } from "@pptxdk/core";
import { Deck, Slide } from "@pptxdk/react";
import { strFromU8, unzipSync } from "fflate";
import { createElement } from "react";

import type { CapabilityScenario } from "../../../test/acceptance-runner.ts";
import type { ExhaustiveMechanismVerificationMap } from "../../../test/mechanism-verification.ts";

function findLineByName(
  root: CommonNode,
  name: string,
): CommonNode | undefined {
  let matched: CommonNode | undefined;
  const visit = (node: CommonNode): void => {
    if (matched !== undefined) return;
    if (node.type === "line" && node.props["name"] === name) {
      matched = node;
      return;
    }
    for (const child of node.children) {
      if (child instanceof CommonNode) {
        visit(child);
      }
    }
  };
  visit(root);
  return matched;
}

function extractConnectorXmlByName(slideXml: string, name: string): string {
  const marker = `name="${name}"`;
  const markerIdx = slideXml.indexOf(marker);
  assert.ok(markerIdx >= 0, `Expected <p:cxnSp> with ${marker} in slide XML`);
  const start = slideXml.lastIndexOf("<p:cxnSp>", markerIdx);
  const end = slideXml.indexOf("</p:cxnSp>", markerIdx);
  assert.ok(
    start >= 0 && end > start,
    `Expected enclosing <p:cxnSp>...</p:cxnSp> for ${marker}`,
  );
  return slideXml.slice(start, end + "</p:cxnSp>".length);
}

export const connectorLineRoundtripScenario: CapabilityScenario = {
  id: "mechanism-connector-line-roundtrip",
  capability: "line",
  description:
    "Minimal dual-author Transitional scenario covering 1D box-mode horizontal/vertical dividers, arrow styles, elbow/curved routing, anchored cross-object connectors, and group-scoped connectors.",
  visualCheckpoints: [
    "Page 1: horizontal 0-height divider, vertical 0-width line, all arrowhead styles, elbow & curved routed connectors anchored across shapes, and a group containing connected shapes and a connector line.",
  ],
  nativeCheckpoints: [
    "Initial and roundtrip PPTX retain <p:cxnSp> single-axis zero extents, <a:headEnd>/<a:tailEnd>, routing preset geometries, <a:stCxn>/<a:endCxn> anchors, and group-scoped connectors.",
  ],
  document() {
    return deck({ width: 960, height: 540 }).append(
      slide({ id: "s1", name: "Connectors" }).append(
        box({
          style: {
            position: "absolute",
            left: 40,
            top: 30,
            width: 880,
            display: "flex",
            flexDirection: "column",
            gap: 10,
          },
        }).append(
          line({
            id: "divider-h",
            name: "Horizontal Divider",
            stroke: { color: "#CBD5E1", width: 2, style: "dashed" },
            style: { width: "100%", height: 0 },
          }),
        ),
        line({
          id: "divider-v",
          name: "Vertical Divider",
          stroke: {
            color: "#0F172A",
            width: 2,
            style: "dotted",
            startArrow: "oval",
            endArrow: "diamond",
          },
          style: {
            position: "absolute",
            left: 60,
            top: 80,
            width: 0,
            height: 220,
          },
        }),
        shape({
          id: "box-a",
          name: "Start Shape",
          geometry: "roundRect",
          style: {
            position: "absolute",
            left: 120,
            top: 100,
            width: 140,
            height: 60,
            background: "#DBEAFE",
            border: { color: "#1D4ED8", width: 1.5 },
          },
        }),
        line({
          id: "conn-elbow",
          name: "Elbow Connector",
          routing: "elbow",
          adjust: { adj1: 50000 },
          stroke: {
            color: "#1D4ED8",
            width: 2,
            startArrow: "none",
            endArrow: "triangle",
          },
          from: { id: "box-a", port: "right" },
          to: { id: "box-b", port: "left" },
        }),
        shape({
          id: "box-b",
          name: "End Shape",
          geometry: "roundRect",
          style: {
            position: "absolute",
            left: 420,
            top: 180,
            width: 140,
            height: 60,
            background: "#DCFCE7",
            border: { color: "#15803D", width: 1.5 },
          },
        }),
        line({
          id: "conn-curved",
          name: "Curved Connector",
          routing: "curved",
          adjust: { adj1: 35000, adj2: 65000 },
          stroke: {
            color: "#9333EA",
            width: 2,
            style: "dashed",
            startArrow: "arrow",
            endArrow: "stealth",
          },
          from: { id: "box-b", port: 0 },
          to: { x: 640, y: 90 },
        }),
        group({ id: "flow-group", name: "Flow Group" }).append(
          shape({
            id: "g-src",
            name: "Group Source",
            style: {
              position: "absolute",
              left: 120,
              top: 320,
              width: 120,
              height: 50,
              background: "#FEF3C7",
            },
          }),
          shape({
            id: "g-dst",
            name: "Group Target",
            style: {
              position: "absolute",
              left: 340,
              top: 320,
              width: 120,
              height: 50,
              background: "#FEE2E2",
            },
          }),
          line({
            id: "g-conn",
            name: "Group Connector",
            stroke: { color: "#B91C1C", width: 2, endArrow: "triangle" },
            from: { id: "g-src", port: "right" },
            to: { id: "g-dst", port: "left" },
          }),
        ),
      ),
    );
  },
  jsx() {
    return createElement(
      Deck,
      { width: 960, height: 540 },
      createElement(
        Slide,
        { id: "s1", name: "Connectors" },
        createElement(
          Box,
          {
            style: {
              position: "absolute",
              left: 40,
              top: 30,
              width: 880,
              display: "flex",
              flexDirection: "column",
              gap: 10,
            },
          },
          createElement(Line, {
            id: "divider-h",
            name: "Horizontal Divider",
            stroke: { color: "#CBD5E1", width: 2, style: "dashed" },
            style: { width: "100%", height: 0 },
          }),
        ),
        createElement(Line, {
          id: "divider-v",
          name: "Vertical Divider",
          stroke: {
            color: "#0F172A",
            width: 2,
            style: "dotted",
            startArrow: "oval",
            endArrow: "diamond",
          },
          style: {
            position: "absolute",
            left: 60,
            top: 80,
            width: 0,
            height: 220,
          },
        }),
        createElement(Shape, {
          id: "box-a",
          name: "Start Shape",
          geometry: "roundRect",
          style: {
            position: "absolute",
            left: 120,
            top: 100,
            width: 140,
            height: 60,
            background: "#DBEAFE",
            border: { color: "#1D4ED8", width: 1.5 },
          },
        }),
        createElement(Line, {
          id: "conn-elbow",
          name: "Elbow Connector",
          routing: "elbow",
          adjust: { adj1: 50000 },
          stroke: {
            color: "#1D4ED8",
            width: 2,
            startArrow: "none",
            endArrow: "triangle",
          },
          from: { id: "box-a", port: "right" },
          to: { id: "box-b", port: "left" },
        }),
        createElement(Shape, {
          id: "box-b",
          name: "End Shape",
          geometry: "roundRect",
          style: {
            position: "absolute",
            left: 420,
            top: 180,
            width: 140,
            height: 60,
            background: "#DCFCE7",
            border: { color: "#15803D", width: 1.5 },
          },
        }),
        createElement(Line, {
          id: "conn-curved",
          name: "Curved Connector",
          routing: "curved",
          adjust: { adj1: 35000, adj2: 65000 },
          stroke: {
            color: "#9333EA",
            width: 2,
            style: "dashed",
            startArrow: "arrow",
            endArrow: "stealth",
          },
          from: { id: "box-b", port: 0 },
          to: { x: 640, y: 90 },
        }),
        createElement(
          Group,
          { id: "flow-group", name: "Flow Group" },
          createElement(Shape, {
            id: "g-src",
            name: "Group Source",
            style: {
              position: "absolute",
              left: 120,
              top: 320,
              width: 120,
              height: 50,
              background: "#FEF3C7",
            },
          }),
          createElement(Shape, {
            id: "g-dst",
            name: "Group Target",
            style: {
              position: "absolute",
              left: 340,
              top: 320,
              width: 120,
              height: 50,
              background: "#FEE2E2",
            },
          }),
          createElement(Line, {
            id: "g-conn",
            name: "Group Connector",
            stroke: { color: "#B91C1C", width: 2, endArrow: "triangle" },
            from: { id: "g-src", port: "right" },
            to: { id: "g-dst", port: "left" },
          }),
        ),
      ),
    );
  },
  assert(context) {
    assert.deepEqual(context.validation.errors, []);
    const slide1 = context.validation.slides[0];
    assert.ok(slide1 !== undefined);
    const lines = slide1.objects.filter((o) => o.kind === "line");
    assert.ok(lines.length >= 4);
    assert.ok(context.preview.pages[0]?.svg?.includes("<path"));
  },
};

function assertConnectorRoundtripNative(
  ctx: Parameters<
    ExhaustiveMechanismVerificationMap["connector-line"]["assertNative"]
  >[0],
): void {
  const { stage, pptx, validation, importedDocument } = ctx;
  assert.deepEqual(validation.errors, []);

  const entries = unzipSync(pptx);
  const slide1Xml = strFromU8(entries["ppt/slides/slide1.xml"]!);

  const horizXml = extractConnectorXmlByName(slide1Xml, "Horizontal Divider");
  assert.ok(horizXml.includes('cy="0"'));

  const vertXml = extractConnectorXmlByName(slide1Xml, "Vertical Divider");
  assert.ok(vertXml.includes('cx="0"'));
  assert.ok(vertXml.includes('<a:headEnd type="oval"/>'));
  assert.ok(vertXml.includes('<a:tailEnd type="diamond"/>'));

  const elbowXml = extractConnectorXmlByName(slide1Xml, "Elbow Connector");
  assert.ok(elbowXml.includes('prst="bentConnector3"'));
  assert.ok(elbowXml.includes("<a:stCxn "));
  assert.ok(elbowXml.includes("<a:endCxn "));

  const curvedXml = extractConnectorXmlByName(slide1Xml, "Curved Connector");
  assert.ok(curvedXml.includes('prst="curvedConnector4"'));
  assert.ok(curvedXml.includes("<a:stCxn "));
  assert.ok(!curvedXml.includes("<a:endCxn "));

  if (stage === "roundtrip") {
    assert.ok(importedDocument !== undefined);
    const importedElbow = findLineByName(importedDocument, "Elbow Connector");
    assert.ok(importedElbow !== undefined);
    assert.equal(importedElbow.props["routing"], "elbow");
    assert.ok(importedElbow.props["from"] !== undefined);
    assert.ok(importedElbow.props["to"] !== undefined);

    const importedGroupConn = findLineByName(
      importedDocument,
      "Group Connector",
    );
    assert.ok(importedGroupConn !== undefined);
    assert.ok(importedGroupConn.props["from"] !== undefined);
    assert.ok(importedGroupConn.props["to"] !== undefined);
  }
}

export const lineMechanismVerifications: Pick<
  ExhaustiveMechanismVerificationMap,
  "connector-line" | "line-normalized-shape"
> = {
  "connector-line": {
    scenario: connectorLineRoundtripScenario,
    assertNative: assertConnectorRoundtripNative,
    assertPaint: ({ initial, roundtrip }) => {
      assert.equal(initial.pages.length, 1);
      assert.equal(roundtrip.pages.length, 1);
      assert.ok((initial.pages[0]?.svg ?? "").includes("<path"));
      assert.ok((roundtrip.pages[0]?.svg ?? "").includes("<path"));
    },
  },
  "line-normalized-shape": {
    scenario: connectorLineRoundtripScenario,
    assertNative: assertConnectorRoundtripNative,
    assertPaint: ({ initial, roundtrip }) => {
      assert.equal(initial.pages.length, 1);
      assert.equal(roundtrip.pages.length, 1);
      assert.ok((initial.pages[0]?.svg ?? "").includes("<path"));
      assert.ok((roundtrip.pages[0]?.svg ?? "").includes("<path"));
    },
  },
};
