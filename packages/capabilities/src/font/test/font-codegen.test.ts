import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import { documentCapabilities } from "@pptxdk/capabilities/all";
import {
  deck,
  paragraph,
  slide,
  span,
  text,
} from "@pptxdk/capabilities/document";
import { font } from "@pptxdk/capabilities/font";
import { createCodegen } from "@pptxdk/codegen";
import { defineTheme } from "@pptxdk/core";
import { describe, it } from "vitest";

const jostRegular = new Uint8Array(
  readFileSync(
    new URL("../../../test/assets/fonts/Jost-400-Book.ttf", import.meta.url),
  ),
);
const jostBold = new Uint8Array(
  readFileSync(
    new URL("../../../test/assets/fonts/Jost-700-Bold.ttf", import.meta.url),
  ),
);
const jostItalic = new Uint8Array(
  readFileSync(
    new URL("../../../test/assets/fonts/Jost-400-Italic.ttf", import.meta.url),
  ),
);

describe("font codegen", () => {
  it("emits relative URL + new Uint8Array(readFileSync(...)) for TTF/OTF faces, preserves substitute/embeddable options, and deduplicates shared fonts", () => {
    const jostPrimary = font("Jost", [
      { weight: 400, style: "normal", data: jostRegular },
      { weight: 700, style: "normal", data: jostBold },
      { weight: 400, style: "italic", data: jostItalic },
    ]);
    // Distinct Font object with identical family, faces, and bytes should deduplicate.
    const jostDuplicate = font("Jost", [
      { weight: 400, style: "normal", data: new Uint8Array(jostRegular) },
      { weight: 700, style: "normal", data: new Uint8Array(jostBold) },
      { weight: 400, style: "italic", data: new Uint8Array(jostItalic) },
    ]);

    const syntheticOtfBytes = new Uint8Array([
      0x4f, 0x54, 0x54, 0x4f, 0x00, 0x0c, 0x00, 0x80, 0x00, 0x03, 0x00, 0x40,
    ]);
    const otfSubFont = font(
      "Cff Display",
      [{ weight: 400, style: "normal", data: syntheticOtfBytes }],
      { substitute: true },
    );
    const nonEmbeddableFont = font(
      "Locked Sans",
      [{ weight: 400, style: "normal", data: jostRegular }],
      { embeddable: false },
    );
    const metriclessNonEmbeddable = font("System Only", [], {
      embeddable: false,
    });

    const sharedTheme = defineTheme({
      name: "Font Theme",
      fonts: { major: jostPrimary, minor: jostDuplicate },
    });

    const doc = deck({ width: 960, height: 540, theme: sharedTheme }).append(
      slide({ id: "fonts-1" }).append(
        text(
          [
            paragraph([
              span("Primary ", { style: { fontFamily: jostPrimary } }),
              span("Duplicate ", { style: { fontFamily: jostDuplicate } }),
              span("OTF Sub ", { style: { fontFamily: otfSubFont } }),
              span("Locked ", { style: { fontFamily: nonEmbeddableFont } }),
              span("Metricless ", {
                style: { fontFamily: metriclessNonEmbeddable },
              }),
              span("Unembedded", {
                style: { fontFamily: "MissingPreviewFamily" },
              }),
            ]),
          ],
          {
            id: "font-sample",
            style: {
              position: "absolute",
              left: 40,
              top: 40,
              width: 800,
              height: 200,
              fontFamily: jostPrimary,
              fontSize: 18,
            },
          },
        ),
      ),
    );

    const generate = createCodegen(documentCapabilities);

    for (const target of ["vanilla", "react"] as const) {
      const result = generate(doc, { target });
      const source = result.files[result.entry] as string;
      const fontAssets = Object.keys(result.files).filter((key) =>
        key.startsWith("assets/font-"),
      );

      // Deduplicated across jostPrimary, jostDuplicate, and nonEmbeddableFont (which shares jostRegular bytes):
      // 3 Jost TTF faces + 1 OTF face = 4 font asset files.
      assert.equal(fontAssets.length, 4, `${target} font asset count`);
      const otfAsset = fontAssets.find(
        (path) => path.includes("cff-display") && path.endsWith(".otf"),
      );
      assert.ok(otfAsset !== undefined, `${target} OTF extension detected`);
      assert.deepEqual(result.files[otfAsset], syntheticOtfBytes);

      const ttfAssets = fontAssets.filter((path) => path.endsWith(".ttf"));
      assert.equal(ttfAssets.length, 3, `${target} TTF count`);

      assert.match(source, /import \{ readFileSync \} from "node:fs";/u);
      assert.match(
        source,
        /new Uint8Array\(\s*readFileSync\(new URL\("\.\/assets\/font-[^"]+\.(?:ttf|otf)", import\.meta\.url\)\),\s*\)/u,
      );
      // Only one font("Jost", ...) declaration despite jostPrimary + jostDuplicate.
      const jostDeclMatches = source.match(/font\("Jost",/gu) ?? [];
      assert.equal(
        jostDeclMatches.length,
        1,
        `${target} deduplicated Jost font() declaration`,
      );
      assert.match(source, /substitute:\s*true/u);
      assert.match(source, /embeddable:\s*false/u);
      assert.match(
        source,
        /font\("System Only", \[\], \{ embeddable: false \}\)/u,
      );
      assert.match(source, /fontFamily:\s*"MissingPreviewFamily"/u);
    }
  });

  it("omits node:fs import and font assets when only string font families are declared", () => {
    const doc = deck({ width: 960, height: 540 }).append(
      slide({ id: "plain-font" }).append(
        text("System family", {
          id: "t1",
          style: {
            position: "absolute",
            left: 40,
            top: 40,
            width: 300,
            height: 60,
            fontFamily: "Calibri",
            fontSize: 16,
          },
        }),
      ),
    );
    const generate = createCodegen(documentCapabilities);
    const result = generate(doc, { target: "vanilla" });
    assert.deepEqual(Object.keys(result.files), ["deck.ts"]);
    assert.doesNotMatch(result.files["deck.ts"] as string, /node:fs/u);
    assert.match(result.files["deck.ts"] as string, /fontFamily:\s*"Calibri"/u);
  });
});
