import assert from "node:assert/strict";

import { documentCapabilities } from "@pptxdk/capabilities/all";
import { deck, line, shape, slide } from "@pptxdk/capabilities/document";
import { BuildError } from "@pptxdk/core";
import { createResolver } from "@pptxdk/core/build";
import { describe, it } from "vitest";

import {
  runCapabilityScenario,
  type StageCounters,
} from "../../../test/acceptance-runner.ts";
import { connectorLineRoundtripScenario } from "./mechanisms.ts";

const resolveDeck = createResolver(documentCapabilities);

function expectBuildErrorCode(fn: () => unknown, code: string): void {
  assert.throws(
    fn,
    (error) =>
      error instanceof BuildError &&
      error.diagnostics.some((entry) => entry.code === code),
  );
}

describe("Line capability acceptance", () => {
  it("runs connector-line dual-author roundtrip scenario across vanilla, react, validation, import, and preview", async () => {
    const counters: StageCounters = {
      expandJsx: 0,
      buildDocument: 0,
      resolve: 0,
      emit: 0,
      paint: 0,
      import: 0,
      validate: 0,
    };
    const execution = await runCapabilityScenario(
      connectorLineRoundtripScenario,
      { counters },
    );
    assert.deepEqual(execution.primary.validation.errors, []);
    assert.deepEqual(execution.roundtrip.validation.errors, []);
    assert.deepEqual(execution.roundtrip.importReport.diagnostics, []);
    assert.ok(counters.expandJsx >= 1);
    assert.ok(counters.buildDocument >= 1);
    assert.ok(counters.validate >= 2);
  });

  it("rejects invalid endpoint combinations, broken references, and zero-area 0x0 lines", () => {
    expectBuildErrorCode(
      () =>
        resolveDeck(
          deck({ width: 960, height: 540 }).append(
            slide({ id: "s1" }).append(
              line({
                from: { id: "missing", port: "right" },
                to: { x: 200, y: 100 },
              }),
            ),
          ),
        ),
      "broken-reference",
    );

    expectBuildErrorCode(
      () =>
        resolveDeck(
          deck({ width: 960, height: 540 }).append(
            slide({ id: "s1" }).append(
              shape({
                id: "a",
                style: {
                  position: "absolute",
                  left: 20,
                  top: 20,
                  width: 40,
                  height: 40,
                },
              }),
              line({ from: { id: "a", port: "right" }, to: { x: 60, y: 40 } }),
            ),
          ),
        ),
      "invalid-style",
    );
  });
});
