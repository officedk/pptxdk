import assert from "node:assert/strict";
import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { pathToFileURL } from "node:url";

import { CodegenError, createCodegen } from "@pptxdk/codegen";
import { PNG } from "pngjs";
import { describe, it } from "vitest";

import { documentCapabilities } from "../../all.ts";
import { deck, image, slide, text } from "../../document.ts";
import { preset } from "../../shared/geometry.ts";

function createSamplePngBytes(): Uint8Array {
  const png = new PNG({ width: 4, height: 4 });
  for (let i = 0; i < png.data.length; i += 4) {
    png.data[i] = 0x10;
    png.data[i + 1] = 0x70;
    png.data[i + 2] = 0xc0;
    png.data[i + 3] = 255;
  }
  return new Uint8Array(PNG.sync.write(png));
}

const sampleGifBytes = new Uint8Array(
  Buffer.from(
    "47494638396101000100800000ffffff0000002c00000000010001000002024401003b",
    "hex",
  ),
);

describe("image capability codegen", () => {
  it("extracts embedded images from bytes, file path, and file URL, deduplicates shared assets, and preserves crop/transforms/style", () => {
    const pngBytes = createSamplePngBytes();
    const tempDir = mkdtempSync(join(tmpdir(), "pptxdk-img-codegen-"));
    try {
      const pngPath = join(tempDir, "local-photo.png");
      const gifPath = join(tempDir, "badge.gif");
      writeFileSync(pngPath, pngBytes);
      writeFileSync(gifPath, sampleGifBytes);

      const doc = deck({ width: 960, height: 540 }).append(
        slide({ id: "s1" }).append(
          image({
            id: "hero-bytes",
            name: "Hero Image",
            alt: "Blue square",
            src: pngBytes,
            fit: "cover",
            crop: { left: 0.1, top: 0.2 },
            geometry: preset("roundRect", { adj: 20000 }),
            link: { url: "https://example.com" },
            style: {
              position: "absolute",
              left: 40,
              top: 50,
              width: 240,
              height: 160,
              flipH: true,
              rotate: 15,
              opacity: 0.85,
              border: {
                color: "#0F172A",
                width: 2,
                style: "dashed",
              },
            },
          }),
          image({
            id: "hero-path-dup",
            src: pngPath,
            style: {
              position: "absolute",
              left: 320,
              top: 50,
              width: 120,
              height: 120,
            },
          }),
          image({
            id: "badge-url",
            src: pathToFileURL(gifPath),
            style: {
              position: "absolute",
              left: 480,
              top: 50,
              width: 40,
              height: 40,
              flipV: true,
            },
          }),
        ),
      );

      const generate = createCodegen(documentCapabilities);

      for (const target of ["vanilla", "react"] as const) {
        const result = generate(doc, { target });
        const source = result.files[result.entry] as string;
        const assetPaths = Object.keys(result.files).filter((key) =>
          key.startsWith("assets/"),
        );
        // pngBytes and pngPath share the exact same SHA-256 -> deduplicated into 1 png + 1 gif.
        assert.equal(assetPaths.length, 2, `${target} asset count`);
        const pngAssetPath = assetPaths.find((p) => p.endsWith(".png"));
        const gifAssetPath = assetPaths.find((p) => p.endsWith(".gif"));
        assert.ok(pngAssetPath !== undefined, `${target} png asset`);
        assert.ok(gifAssetPath !== undefined, `${target} gif asset`);
        assert.deepEqual(result.files[pngAssetPath], pngBytes);
        assert.deepEqual(result.files[gifAssetPath], sampleGifBytes);

        assert.ok(
          source.includes(`new URL("./${pngAssetPath}", import.meta.url)`),
          `${target} relative URL for PNG`,
        );
        assert.ok(
          source.includes(`new URL("./${gifAssetPath}", import.meta.url)`),
          `${target} relative URL for GIF`,
        );
        assert.match(source, /rotate:\s*15/u);
        assert.match(source, /opacity:\s*0\.85/u);
        assert.match(source, /style:\s*"dashed"/u);
        assert.match(source, /flipH:\s*true/u);
        assert.match(source, /flipV:\s*true/u);
        assert.match(source, /url:\s*"https:\/\/example\.com"/u);
        if (target === "vanilla") {
          assert.match(
            source,
            /crop:\s*\{\s*left:\s*0\.1,\s*top:\s*0\.2\s*\}/u,
          );
          assert.match(source, /\bimage\(/u);
        } else {
          assert.match(
            source,
            /crop=\{\{\s*left:\s*0\.1,\s*top:\s*0\.2\s*\}\}/u,
          );
          assert.match(source, /<Image\b/u);
        }
      }
    } finally {
      rmSync(tempDir, { recursive: true, force: true });
    }
  });

  it("preserves externalUrl images with zero asset files and reports CodegenError with node path on missing or corrupt sources", () => {
    const generate = createCodegen(documentCapabilities);

    const externalDoc = deck({ width: 960, height: 540 }).append(
      slide({ id: "ext-slide" }).append(
        image({
          id: "remote-banner",
          src: { externalUrl: "https://assets.example.com/banner.png" },
          style: {
            position: "absolute",
            left: 40,
            top: 40,
            width: 400,
            height: 200,
          },
        }),
      ),
    );
    for (const target of ["vanilla", "react"] as const) {
      const res = generate(externalDoc, { target });
      assert.deepEqual(Object.keys(res.files), [res.entry]);
      assert.match(
        res.files[res.entry] as string,
        /externalUrl:\s*"https:\/\/assets\.example\.com\/banner\.png"/u,
      );
    }

    const missingPathDoc = deck({ width: 960, height: 540 }).append(
      slide({ id: "s-err" }).append(
        image({
          id: "missing-pic",
          src: "/nonexistent/pptxdk-missing-image-file.png",
          style: {
            position: "absolute",
            left: 0,
            top: 0,
            width: 100,
            height: 100,
          },
        }),
      ),
    );
    assert.throws(
      () => generate(missingPathDoc, { target: "vanilla" }),
      (error: unknown) =>
        error instanceof CodegenError &&
        error.diagnostics[0]?.code === "unexpressible-value" &&
        error.diagnostics[0]?.path ===
          "deck/slide[s-err]/image[missing-pic]@src",
    );

    const corruptBytesDoc = deck({ width: 960, height: 540 }).append(
      slide({ id: "s-err2" }).append(
        image({
          id: "corrupt-pic",
          src: new Uint8Array([1, 2, 3, 4]),
          style: {
            position: "absolute",
            left: 0,
            top: 0,
            width: 100,
            height: 100,
          },
        }),
      ),
    );
    assert.throws(
      () => generate(corruptBytesDoc, { target: "react" }),
      (error: unknown) =>
        error instanceof CodegenError &&
        error.diagnostics[0]?.code === "unexpressible-value" &&
        error.diagnostics[0]?.path ===
          "deck/slide[s-err2]/image[corrupt-pic]@src",
    );

    const nonFileUrlDoc = deck({ width: 960, height: 540 }).append(
      slide({ id: "s-err3" }).append(
        image({
          id: "http-url-pic",
          src: new URL("https://example.com/not-local.png"),
          style: {
            position: "absolute",
            left: 0,
            top: 0,
            width: 100,
            height: 100,
          },
        }),
      ),
    );
    assert.throws(
      () => generate(nonFileUrlDoc, { target: "vanilla" }),
      (error: unknown) =>
        error instanceof CodegenError &&
        error.diagnostics[0]?.code === "unexpressible-value" &&
        error.diagnostics[0]?.path ===
          "deck/slide[s-err3]/image[http-url-pic]@src",
    );

    const emptyExternalDoc = deck({ width: 960, height: 540 }).append(
      slide({ id: "s-err4" }).append(
        image({
          id: "empty-ext",
          src: { externalUrl: "" },
          style: {
            position: "absolute",
            left: 0,
            top: 0,
            width: 100,
            height: 100,
          },
        }),
      ),
    );
    assert.throws(
      () => generate(emptyExternalDoc, { target: "vanilla" }),
      (error: unknown) =>
        error instanceof CodegenError &&
        error.diagnostics[0]?.code === "unexpressible-value" &&
        error.diagnostics[0]?.path ===
          "deck/slide[s-err4]/image[empty-ext]@src",
    );
  });

  it("extracts style.background image assets via valueSerializers and fails with @style.background on bad sources", () => {
    const pngBytes = createSamplePngBytes();
    const generate = createCodegen(documentCapabilities);

    const doc = deck({ width: 960, height: 540 }).append(
      slide({ id: "s1" }).append(
        text("Card", {
          id: "card-with-bg-img",
          style: {
            position: "absolute",
            left: 40,
            top: 40,
            width: 300,
            height: 120,
            background: { image: pngBytes },
          },
        }),
        image({
          id: "shared-with-bg",
          src: pngBytes,
          style: {
            position: "absolute",
            left: 400,
            top: 40,
            width: 120,
            height: 120,
          },
        }),
      ),
    );

    for (const target of ["vanilla", "react"] as const) {
      const res = generate(doc, { target });
      const assetKeys = Object.keys(res.files).filter((k) =>
        k.startsWith("assets/"),
      );
      assert.equal(assetKeys.length, 1, `${target} shared asset dedup`);
      const source = res.files[res.entry] as string;
      assert.match(source, /background:\s*\{\s*image:\s*imageAsset\s*\}/u);
    }

    const badBgImgDoc = deck({ width: 960, height: 540 }).append(
      slide({ id: "s-bad-bgimg" }).append(
        text("Bad", {
          id: "bad-card",
          style: {
            position: "absolute",
            left: 40,
            top: 40,
            width: 200,
            height: 80,
            background: { image: "/missing/style-bg-img.png" },
          },
        }),
      ),
    );
    assert.throws(
      () => generate(badBgImgDoc, { target: "vanilla" }),
      (error: unknown) =>
        error instanceof CodegenError &&
        error.diagnostics[0]?.code === "unexpressible-value" &&
        error.diagnostics[0]?.path ===
          "deck/slide[s-bad-bgimg]/text[bad-card]@style.background",
    );
  });
});
