import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

import { createResolver } from "@react-pptx/core/build";
import {
  parseCapabilityLayoutContract,
  textStyleKeys,
  type BoundObjectCapabilityModule,
  type CapabilityInput,
  type CapabilityParseContext,
} from "@react-pptx/core/capability";
import {
  createDeckNode,
  createSlideNode,
  RichTextContainer,
} from "@react-pptx/core/document";
import { emit } from "@react-pptx/emit";
import { describe, expect, it } from "vitest";

import { font } from "../src/font.ts";
import { slideBackgroundCapability } from "../src/slide-background.ts";
import {
  emitTextBody,
  parseTextContent,
  renderTextOutlinesSvg,
  resolveTextBody,
  textInsets,
} from "../src/text-service.ts";

describe("shared text-service and module isolation", () => {
  it("shares text-service layout and glyph outlines without registering textCapability or loading text.ts", () => {
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

    let outlinedSvg = "";
    const customLabelCapability: BoundObjectCapabilityModule = {
      declaration: "rp:custom-label",
      bind(input: CapabilityInput, parseContext: CapabilityParseContext) {
        const contract = parseCapabilityLayoutContract({
          nativeKind: "sp",
          props: input.props,
          allowedStyleKeys: new Set(["width", "height", ...textStyleKeys]),
          allowPadding: false,
          path: parseContext.path,
        });
        const insets = textInsets(input.props["style"]);
        const { paragraphs, style } = parseTextContent(
          input.children,
          input.props["style"],
          parseContext.theme,
          { slideNumber: parseContext.slideNumber },
        );
        return {
          contract,
          resolve(resolveContext) {
            const box = resolveContext.layoutBox ?? {
              x: Number(input.props["x"] ?? 0),
              y: Number(input.props["y"] ?? 0),
              width: Number(input.props["width"] ?? 120),
              height: Number(input.props["height"] ?? 40),
            };
            const body = resolveTextBody(
              paragraphs,
              style,
              resolveContext.theme,
              "middle",
              { width: box.width, height: box.height, insets },
              "text-box",
            );
            outlinedSvg = renderTextOutlinesSvg({
              paragraphs,
              box,
              insets,
              verticalAlign: "middle",
              theme: resolveContext.theme,
              measuredLayout: body.measuredLayout,
            });
            return {
              observation: {
                id: resolveContext.id,
                kind: "shape",
                name: "CustomLabel",
                geometry: "rect",
                box,
                text: body.text,
              },
              extensions: [],
              resourceClaims: body.resourceClaims,
              emit(emitContext) {
                const emitted = emitTextBody(body, emitContext.theme, {
                  createFieldId: emitContext.createTextFieldId,
                  ...(emitContext.options.embedFonts === undefined
                    ? {}
                    : { embedFonts: emitContext.options.embedFonts }),
                });
                return {
                  xml: `<p:sp><p:nvSpPr><p:cNvPr id="${emitContext.id}" name="CustomLabel"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="381000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>${emitted.xml}</p:sp>`,
                  ...(emitted.presentationContributions.length === 0
                    ? {}
                    : {
                        presentationContributions:
                          emitted.presentationContributions,
                      }),
                };
              },
            };
          },
        };
      },
    };

    const resolveWithoutTextCapability = createResolver([
      slideBackgroundCapability,
      customLabelCapability,
    ]);
    const doc = createDeckNode({ width: 320, height: 180 }).append(
      createSlideNode({ id: "shared-only" }).append(
        new RichTextContainer("custom-label", "rp:custom-label", {
          id: "lbl",
          x: 20,
          y: 20,
          width: 180,
          height: 50,
          style: { fontFamily: jostFont, fontSize: 18, color: "#1D4ED8" },
        }).content("Shared glyphs"),
      ),
    );
    const resolution = resolveWithoutTextCapability(doc);
    const emitted = emit(resolution);
    assert.ok(emitted.pptx.byteLength > 0);
    expect(outlinedSvg).toContain("<path");

    // Verify text-service never pulls in text.ts at runtime.
    const packageRoot = fileURLToPath(new URL("../", import.meta.url));
    const isolated = spawnSync(
      process.execPath,
      [
        "--conditions=source",
        "--input-type=module",
        "--eval",
        `
import { registerHooks } from "node:module";
registerHooks({
  resolve(specifier, context, nextResolve) {
    if (specifier.endsWith("/text.ts") || specifier.endsWith("/text.js") || specifier.endsWith("/text")) {
      throw new Error("text-service loaded text capability module: " + specifier);
    }
    return nextResolve(specifier, context);
  },
});
await import("./src/text-service.ts");
process.stdout.write("ok");
`,
      ],
      { cwd: packageRoot, encoding: "utf8" },
    );
    assert.equal(isolated.status, 0, isolated.stderr);
    assert.equal(isolated.stdout, "ok");
  });

  it("never loads text.ts from chart.ts at runtime", () => {
    const packageRoot = fileURLToPath(new URL("../", import.meta.url));
    const isolated = spawnSync(
      process.execPath,
      [
        "--conditions=source",
        "--input-type=module",
        "--eval",
        `
import { registerHooks } from "node:module";
registerHooks({
  resolve(specifier, context, nextResolve) {
    if (specifier.endsWith("/text.ts") || specifier.endsWith("/text.js") || specifier.endsWith("/text")) {
      throw new Error("chart capability loaded text.ts: " + specifier);
    }
    return nextResolve(specifier, context);
  },
});
await import("./src/chart.ts");
process.stdout.write("ok");
`,
      ],
      { cwd: packageRoot, encoding: "utf8" },
    );
    assert.equal(isolated.status, 0, isolated.stderr);
    assert.equal(isolated.stdout, "ok");
  });
});
