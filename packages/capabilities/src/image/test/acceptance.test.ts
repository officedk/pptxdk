import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

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
});
