import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

import {
  BuildError,
  CommonNode,
  Deck,
  Slide,
  createPreview,
  pxToEmu,
  xmlEscape,
  type BoundObjectCapabilityModule,
  type CapabilityInput,
  type CapabilityParseContext,
} from "@react-pptx/core";
import { parseCapabilityLayoutContract } from "@react-pptx/core/capability";
import { createDocumentPreview } from "@react-pptx/core/preview";
import { PNG } from "pngjs";
import { createElement } from "react";
import { describe, expect, it } from "vitest";

import { deck, shape, slide } from "../src/document.ts";
import { shapeCapability } from "../src/shape.ts";
import { runCapabilityScenario } from "./acceptance-runner.ts";
import {
  imagePositiveScenario,
  imageRejectionScenario,
} from "./scenarios/image.ts";
import {
  shapePositiveScenario,
  shapeRejectionScenario,
} from "./scenarios/shape.ts";
import {
  slideBackgroundPositiveScenario,
  slideBackgroundRejectionScenario,
} from "./scenarios/slide-background.ts";
import {
  runTransitionRejectionScenarios,
  transitionPositiveScenario,
} from "./scenarios/transition.ts";

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
  it("runs Shape, Image, and slide-background dual-author scenarios through shared runner and collects artifacts", async () => {
    const outDir = await mkdtemp(join(tmpdir(), "rp-cap-accept-"));
    try {
      for (const scenario of [
        shapePositiveScenario,
        shapeRejectionScenario,
        imagePositiveScenario,
        imageRejectionScenario,
        slideBackgroundPositiveScenario,
        slideBackgroundRejectionScenario,
        transitionPositiveScenario,
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
      await runTransitionRejectionScenarios();
    } finally {
      await rm(outDir, { recursive: true, force: true });
    }
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
