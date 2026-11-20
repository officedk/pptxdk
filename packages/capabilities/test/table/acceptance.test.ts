import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { describe, it } from "vitest";

import { runCapabilityScenario } from "../acceptance-runner.ts";
import {
  tablePositiveScenario,
  tableRejectionScenario,
  tableRichTextBordersAndOversetScenario,
} from "./scenarios.ts";

describe("Table capability acceptance", () => {
  it("runs Table dual-author positive and rejection scenarios through shared runner", async () => {
    const outDir = await mkdtemp(join(tmpdir(), "rp-table-accept-"));
    try {
      for (const scenario of [
        tablePositiveScenario,
        tableRichTextBordersAndOversetScenario,
        tableRejectionScenario,
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
  }, 15_000);
});
