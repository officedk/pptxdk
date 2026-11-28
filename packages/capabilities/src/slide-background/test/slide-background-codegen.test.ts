import assert from "node:assert/strict";
import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { pathToFileURL } from "node:url";

import { documentCapabilities } from "@pptxdk/capabilities/all";
import {
  deck,
  defineSlideLayout,
  linearGradient,
  radialGradient,
  slide,
  slideLayout,
  slideMaster,
  text,
} from "@pptxdk/capabilities/document";
import { CodegenError, createCodegen } from "@pptxdk/codegen";
import { darker, defineTheme, lighter } from "@pptxdk/core";
import { PNG } from "pngjs";
import { describe, it } from "vitest";

function createSolidPng(r: number, g: number, b: number): Uint8Array {
  const png = new PNG({ width: 4, height: 4 });
  for (let i = 0; i < png.data.length; i += 4) {
    png.data[i] = r;
    png.data[i + 1] = g;
    png.data[i + 2] = b;
    png.data[i + 3] = 255;
  }
  return new Uint8Array(PNG.sync.write(png));
}

describe("slide-background capability codegen", () => {
  it("emits solid, theme shade (lighter/darker), bgRef, linear/radial gradients, and image backgrounds on master/layout/slide plus style.background gradients", () => {
    const bgPng = createSolidPng(0x1e, 0x3a, 0x8a);
    const tempDir = mkdtempSync(join(tmpdir(), "pptxdk-bg-codegen-"));
    try {
      const bgFilePath = join(tempDir, "bg.png");
      writeFileSync(bgFilePath, bgPng);

      const brandTheme = defineTheme({
        name: "Brand",
        colors: { accent1: "#2563EB", accent2: "#7C3AED" },
      });

      const gradLayout = defineSlideLayout({
        id: "grad-layout",
        name: "Gradient Layout",
        type: "blank",
        background: radialGradient(
          { color: "#FEF3C7", position: 0 },
          { color: darker("accent2", 0.25), position: 1 },
        ),
      });

      const doc = deck({ width: 960, height: 540, theme: brandTheme }).append(
        slideMaster({
          id: "master-1",
          theme: brandTheme,
          background: lighter("accent1", 0.4),
        }).append(slideLayout(gradLayout)),
        slide({
          id: "s-linear",
          layout: gradLayout,
          background: linearGradient(
            90,
            { color: lighter("accent1", 0.3), position: 0 },
            { color: { color: "#38BDF8", alpha: 0.85 }, position: 1 },
          ),
        }).append(
          text("Gradient card", {
            id: "grad-text",
            style: {
              position: "absolute",
              left: 40,
              top: 40,
              width: 320,
              height: 80,
              background: linearGradient(45, "#0F172A", "#1D4ED8"),
            },
          }),
        ),
        slide({
          id: "s-ref",
          layout: gradLayout,
          background: { ref: 1001, color: "bg1" },
        }),
        slide({
          id: "s-img-bytes",
          layout: gradLayout,
          background: { image: bgPng },
        }),
        slide({
          id: "s-img-url",
          layout: gradLayout,
          background: { image: pathToFileURL(bgFilePath) },
        }),
      );

      const generate = createCodegen(documentCapabilities);
      for (const target of ["vanilla", "react"] as const) {
        const res = generate(doc, { target });
        const source = res.files[res.entry] as string;
        const assetKeys = Object.keys(res.files).filter((k) =>
          k.startsWith("assets/"),
        );
        // bgPng and bgFilePath share identical bytes -> deduplicated into 1 asset.
        assert.equal(assetKeys.length, 1, `${target} image background asset`);
        const assetKey = assetKeys[0]!;
        assert.deepEqual(res.files[assetKey], bgPng);
        assert.ok(
          source.includes(`new URL("./${assetKey}", import.meta.url)`),
          `${target} asset URL`,
        );

        assert.match(source, /lighter\("accent1", 0\.4\)/u);
        assert.match(source, /darker\("accent2", 0\.25\)/u);
        assert.match(source, /lighter\("accent1", 0\.3\)/u);
        assert.match(source, /linearGradient\(/u);
        assert.match(source, /radialGradient\(/u);
        assert.match(source, /ref:\s*1001/u);
        assert.match(source, /image:\s*imageAsset/u);
      }
    } finally {
      rmSync(tempDir, { recursive: true, force: true });
    }
  });

  it("reports CodegenError with slide/layout background path on invalid or unreadable background declarations", () => {
    const generate = createCodegen(documentCapabilities);

    const missingBgDoc = deck({ width: 960, height: 540 }).append(
      slide({
        id: "s-missing-bg",
        background: { image: "/nonexistent/pptxdk-missing-bg.png" },
      }),
    );
    assert.throws(
      () => generate(missingBgDoc, { target: "vanilla" }),
      (error: unknown) =>
        error instanceof CodegenError &&
        error.diagnostics[0]?.code === "unexpressible-value" &&
        error.diagnostics[0]?.path === "deck/slide[s-missing-bg]@background",
    );

    const corruptBgDoc = deck({ width: 960, height: 540 }).append(
      slide({
        id: "s-corrupt-bg",
        background: { image: new Uint8Array([9, 8, 7]) },
      }),
    );
    assert.throws(
      () => generate(corruptBgDoc, { target: "react" }),
      (error: unknown) =>
        error instanceof CodegenError &&
        error.diagnostics[0]?.code === "unexpressible-value" &&
        error.diagnostics[0]?.path === "deck/slide[s-corrupt-bg]@background",
    );

    const badShapeBgDoc = deck({ width: 960, height: 540 }).append(
      slide({
        id: "s-invalid-bg",
        background: { unknownField: true } as never,
      }),
    );
    assert.throws(
      () => generate(badShapeBgDoc, { target: "vanilla" }),
      (error: unknown) =>
        error instanceof CodegenError &&
        error.diagnostics[0]?.code === "unexpressible-value" &&
        error.diagnostics[0]?.path === "deck/slide[s-invalid-bg]@background",
    );
  });
});
