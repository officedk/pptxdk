import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { documentCapabilities } from "@pptxdk/capabilities/all";
import { deck, image, slide } from "@pptxdk/capabilities/document";
import { createResolver } from "@pptxdk/core/build";
import { describe, it } from "vitest";

import { runCapabilityScenario } from "../../../test/acceptance-runner.ts";
import {
  imageMasksAndEffectsScenario,
  imagePositiveScenario,
  imageRejectionScenario,
  imageSizingSoftBorderAndLinkScenario,
} from "./scenarios.ts";

describe("Image capability acceptance", () => {
  it("runs Image dual-author positive and rejection scenarios through shared runner", async () => {
    const outDir = await mkdtemp(join(tmpdir(), "rp-image-accept-"));
    try {
      for (const scenario of [
        imagePositiveScenario,
        imageMasksAndEffectsScenario,
        imageSizingSoftBorderAndLinkScenario,
        imageRejectionScenario,
      ]) {
        const { runs } = await runCapabilityScenario(scenario, {
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
  }, 15_000);

  it("resolves large Uint8Array PNG buffers without enumerating byte indices in externalUrl schema", () => {
    const minimalPng = Buffer.from(
      "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==",
      "base64",
    );
    // Append a valid trailing PNG tEXt-free padding chunk or extra bytes after IEND? No, keep valid PNG with a large ancillary tEXt chunk before IEND.
    const payloadLen = 2_000_000;
    const chunkLen = payloadLen;
    const out = new Uint8Array(minimalPng.byteLength + 12 + chunkLen);
    const idatEnd = minimalPng.byteLength - 12; // before IEND
    out.set(minimalPng.subarray(0, idatEnd), 0);
    const view = new DataView(out.buffer);
    view.setUint32(idatEnd, chunkLen);
    out.set([0x74, 0x45, 0x58, 0x74], idatEnd + 4); // tEXt
    out.set(minimalPng.subarray(idatEnd), idatEnd + 12 + chunkLen);

    const resolveDeck = createResolver(documentCapabilities);
    const start = performance.now();
    const res = resolveDeck(
      deck({ width: 320, height: 180 }).append(
        slide({ id: "s1" }).append(
          image({
            src: out,
            style: {
              position: "absolute",
              left: 0,
              top: 0,
              width: 100,
              height: 100,
            },
          }),
        ),
      ),
    );
    const elapsedMs = performance.now() - start;
    assert.equal(res.report.slides.length, 1);
    assert.ok(
      elapsedMs < 1000,
      `Expected resolve < 1000ms, got ${elapsedMs}ms`,
    );
  });
});
