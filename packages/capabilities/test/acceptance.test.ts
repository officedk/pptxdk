import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

import {
  BuildError,
  CommonNode,
  Deck,
  Slide,
  createPreview,
  expandDocument,
  pxToEmu,
  xmlEscape,
  type BoundObjectCapabilityModule,
  type CapabilityInput,
  type CapabilityParseContext,
} from "@react-pptx/core";
import { parseCapabilityLayoutContract } from "@react-pptx/core/capability";
import { RichTextContainer } from "@react-pptx/core/document";
import { createDocumentPreview } from "@react-pptx/core/preview";
import { validatePptx } from "@react-pptx/validation";
import { PNG } from "pngjs";
import { createElement } from "react";
import { describe, expect, it } from "vitest";

import {
  deck,
  documentCapabilities,
  font,
  shape,
  slide,
} from "../src/document.ts";
import { shapeCapability } from "../src/shape.ts";
import { Text } from "../src/text-elements.ts";
import { runCapabilityScenario } from "./acceptance-runner.ts";
import {
  shapePositiveScenario,
  shapeRejectionScenario,
} from "./scenarios/shape.ts";
import {
  slideBackgroundPositiveScenario,
  slideBackgroundRejectionScenario,
} from "./scenarios/slide-background.ts";
import {
  textPositiveScenario,
  textRejectionScenario,
} from "./scenarios/text.ts";

type CustomMarkerProps = {
  readonly id?: string;
  readonly x: number;
  readonly y: number;
  readonly size: number;
  readonly color: string;
  readonly warnMessage?: string;
};

function createMarkerNode(
  props: CustomMarkerProps,
): CommonNode<"marker", CustomMarkerProps> {
  return new CommonNode("marker", "rp:custom-marker", props);
}

describe("Capability-owned acceptance scenarios & custom bound preview extensions", () => {
  it("runs Shape, slide-background, and Text dual-author scenarios through shared runner and collects artifacts", async () => {
    const outDir = await mkdtemp(join(tmpdir(), "rp-cap-accept-"));
    try {
      for (const scenario of [
        shapePositiveScenario,
        shapeRejectionScenario,
        slideBackgroundPositiveScenario,
        slideBackgroundRejectionScenario,
        textPositiveScenario,
        textRejectionScenario,
      ]) {
        const runs = await runCapabilityScenario(scenario, {
          outputDir: outDir,
        });
        assert.equal(runs.length, 2);
        assert.deepEqual(
          runs.map((r) => r.author),
          ["jsx", "document"],
        );
      }
    } finally {
      await rm(outDir, { recursive: true, force: true });
    }
  });

  it("supports JSX expansion followed by imperative Text edit, rebuild, and preview without mutating prior outputs", async () => {
    const jostFont = font("Jost*", [
      {
        weight: 400,
        data: readFileSync(
          new URL(
            "../../react-pptx/test/assets/fonts/Jost-400-Book.ttf",
            import.meta.url,
          ),
        ),
      },
    ]);
    const expanded = await expandDocument(
      createElement(
        Deck,
        { width: 480, height: 270 },
        createElement(
          Slide,
          { id: "handoff" },
          createElement(
            Text,
            {
              id: "headline",
              x: 24,
              y: 24,
              width: 200,
              style: {
                fontFamily: jostFont,
                fontSize: 20,
                color: "#1E293B",
                padding: 8,
                backgroundColor: "#F1F5F9",
              },
            },
            "Initial title",
          ),
        ),
      ),
    );

    const previewDocument = createDocumentPreview(documentCapabilities);
    const first = await previewDocument(expanded);
    assert.deepEqual(first.pages[0]?.diagnostics, []);
    assert.deepEqual((await validatePptx(first.build.pptx)).errors, []);
    const firstPng = first.pages[0]!.png!.slice();
    const firstSvg = first.pages[0]!.svg!;
    const firstPptx = first.build.pptx.slice();
    const firstHeight = first.build.query
      .getSlide("handoff")
      .get("headline")
      .observe().box!.height;

    const headline = expanded.query
      .get("handoff")
      .getAs(
        "headline",
        (node): node is RichTextContainer => node instanceof RichTextContainer,
      );
    headline.content("Updated longer headline wrapping across multiple lines");
    headline.style({ color: "#059669" });

    const second = await previewDocument(expanded);
    assert.deepEqual(second.pages[0]?.diagnostics, []);
    assert.deepEqual((await validatePptx(second.build.pptx)).errors, []);
    assert.equal(
      second.build.query.getSlide("handoff").get("headline").observe().text,
      "Updated longer headline wrapping across multiple lines",
    );
    const secondHeight = second.build.query
      .getSlide("handoff")
      .get("headline")
      .observe().box!.height;
    assert.ok(secondHeight > firstHeight);
    expect(second.pages[0]?.svg).toContain('fill="#059669"');
    assert.notDeepEqual(second.pages[0]?.png, firstPng);

    // Prior BuildResult and PreviewResult remain unchanged.
    assert.equal(first.pages[0]?.svg, firstSvg);
    assert.deepEqual(first.pages[0]?.png, firstPng);
    assert.deepEqual(first.build.pptx, firstPptx);
    assert.equal(
      first.build.query.getSlide("handoff").get("headline").observe().text,
      "Initial title",
    );
  });

  it("supports registering custom bound extension modules, warning retention, and missing preview rejection", async () => {
    const calls: string[] = [];

    const customMarkerWithPreview: BoundObjectCapabilityModule = {
      declaration: "rp:custom-marker",
      bind(input: CapabilityInput, parseContext: CapabilityParseContext) {
        const x = Number(input.props["x"]);
        const y = Number(input.props["y"]);
        const size = Number(input.props["size"]);
        const color = String(input.props["color"]);
        const warnMessage =
          typeof input.props["warnMessage"] === "string"
            ? input.props["warnMessage"]
            : undefined;
        const contract = parseCapabilityLayoutContract({
          nativeKind: "sp",
          props: { x, y, width: size, height: size },
          allowedStyleKeys: new Set<string>(),
          allowPadding: false,
          path: parseContext.path,
        });
        calls.push(`bind:${parseContext.path}`);

        return {
          contract,
          compile(compileContext) {
            calls.push(`compile:${compileContext.path}`);
            const box = compileContext.layoutBox ?? {
              x,
              y,
              width: size,
              height: size,
            };
            const name = xmlEscape(`Marker ${compileContext.id}`);
            const xml = `<p:sp><p:nvSpPr><p:cNvPr id="${compileContext.id}" name="${name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${pxToEmu(box.x)}" y="${pxToEmu(box.y)}"/><a:ext cx="${pxToEmu(box.width)}" cy="${pxToEmu(box.height)}"/></a:xfrm><a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom><a:noFill/></p:spPr></p:sp>`;
            return {
              xml,
              observation: {
                id: compileContext.id,
                kind: "shape",
                name,
                geometry: "ellipse",
                box,
              },
              extensions: [],
            };
          },
          preview(previewContext) {
            calls.push(`preview:${previewContext.path}`);
            const box = previewContext.box!;
            const resolved = previewContext.theme.resolveColor(color);
            return {
              svg: `<circle cx="${box.x + box.width / 2}" cy="${box.y + box.height / 2}" r="${box.width / 2}" fill="#${resolved.hex}"/>`,
              ...(warnMessage === undefined
                ? {}
                : { warnings: [{ message: warnMessage }] }),
            };
          },
        };
      },
    };

    const customMarkerWithoutPreview: BoundObjectCapabilityModule = {
      declaration: "rp:custom-no-preview",
      bind(input: CapabilityInput, parseContext: CapabilityParseContext) {
        const contract = parseCapabilityLayoutContract({
          nativeKind: "sp",
          props: input.props,
          allowedStyleKeys: new Set<string>(),
          allowPadding: false,
          path: parseContext.path,
        });
        return {
          contract,
          compile(compileContext) {
            const box = compileContext.layoutBox ?? {
              x: 0,
              y: 0,
              width: 20,
              height: 20,
            };
            return {
              xml: `<p:sp><p:nvSpPr><p:cNvPr id="${compileContext.id}" name="NoPreview"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="190500" cy="190500"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr></p:sp>`,
              observation: {
                id: compileContext.id,
                kind: "shape",
                name: "NoPreview",
                geometry: "rect",
                box,
              },
              extensions: [],
            };
          },
        };
      },
    };

    const previewCustom = createDocumentPreview([
      shapeCapability,
      customMarkerWithPreview,
      customMarkerWithoutPreview,
    ]);

    const doc = deck({ width: 200, height: 120 }).append(
      slide({ id: "warn-keeps-png" }).append(
        createMarkerNode({
          id: "dot",
          x: 60,
          y: 20,
          size: 80,
          color: "#DC2626",
          warnMessage: "Custom approximation notice",
        }),
      ),
      slide({ id: "warn-cannot-clear-error" }).append(
        new CommonNode("noprev", "rp:custom-no-preview", {
          id: "opaque",
          x: 10,
          y: 10,
          width: 20,
          height: 20,
        }),
        createMarkerNode({
          id: "later-warn",
          x: 100,
          y: 20,
          size: 40,
          color: "#2563EB",
          warnMessage: "Later warning must not clear earlier error",
        }),
      ),
    );

    const result = await previewCustom(doc);
    // Page 1: warning preserves PNG and draws the red circle.
    assert.equal(result.pages[0]?.diagnostics.length, 1);
    assert.equal(result.pages[0]?.diagnostics[0]?.severity, "warning");
    assert.ok(result.pages[0]?.png !== undefined);
    const pixels = PNG.sync.read(Buffer.from(result.pages[0].png));
    const center = (60 * pixels.width + 100) * 4;
    assert.deepEqual(
      [...pixels.data.subarray(center, center + 4)],
      [0xdc, 0x26, 0x26, 255],
    );
    assert.ok(
      result.build.query.getSlide("warn-keeps-png").get("dot").observe().box
        ?.width === 80,
    );

    // Page 2: missing preview handler causes error; subsequent warning does NOT clear error or fabricate PNG.
    assert.equal(result.pages[1]?.png, undefined);
    expect(result.pages[1]?.diagnostics).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          severity: "error",
          path: "deck/slide[warn-cannot-clear-error]/custom-no-preview[opaque]",
        }),
        expect.objectContaining({
          severity: "warning",
          path: "deck/slide[warn-cannot-clear-error]/custom-marker[later-warn]",
        }),
      ]),
    );

    // Same custom bound extension works via JSX createPreview and shares trimmed constraints.
    const jsxCustomPreview = createPreview([shapeCapability]);
    const trimmedOk = await jsxCustomPreview(
      createElement(
        Deck,
        { width: 200, height: 120 },
        createElement(Slide, { id: "trimmed" }),
      ),
    );
    assert.ok(trimmedOk.pages[0]?.png !== undefined);

    await assert.rejects(
      () =>
        jsxCustomPreview(
          createElement(
            Deck,
            { width: 200, height: 120 },
            createElement(Slide, { id: "missing-bg", background: "#123456" }),
          ),
        ),
      BuildError,
    );
    void shape;
    assert.ok(calls.some((entry) => entry.startsWith("bind:")));
    assert.ok(calls.some((entry) => entry.startsWith("compile:")));
    assert.ok(calls.some((entry) => entry.startsWith("preview:")));
  });
});
