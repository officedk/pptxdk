import assert from "node:assert/strict";

import { documentCapabilities } from "@pptxdk/capabilities/all";
import { deck, line, shape, slide } from "@pptxdk/capabilities/document";
import { createCodegen } from "@pptxdk/codegen";
import { BuildError, CommonNode } from "@pptxdk/core";
import { createResolver } from "@pptxdk/core/build";
import { emit } from "@pptxdk/emit";
import { createImporter } from "@pptxdk/import";
import { paint } from "@pptxdk/preview";
import { validatePptx } from "@pptxdk/validation";
import { strFromU8, unzipSync } from "fflate";
import { describe, it } from "vitest";

import {
  runCapabilityScenario,
  type StageCounters,
} from "../../../test/acceptance-runner.ts";
import { verifyBatchGeneratedSourcesAndRelocatedRebuilds } from "../../../test/codegen-batch-acceptance.ts";
import { connectorLineRoundtripScenario } from "./mechanisms.ts";

const resolveDeck = createResolver(documentCapabilities);
const importDeck = createImporter(documentCapabilities);
const codegenDeck = createCodegen(documentCapabilities);

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

  it("roundtrips octagon port 4..7 connectors through resolve, emit, validatePptx, paint, importPptx, and dual-target codegen", async () => {
    const doc = deck({ width: 640, height: 360 }).append(
      slide({ id: "s1", background: "#FFFFFF" }).append(
        shape({
          id: "oct-a",
          name: "Octagon A",
          geometry: "octagon",
          style: {
            position: "absolute",
            left: 80,
            top: 80,
            width: 120,
            height: 120,
            background: "#DBEAFE",
            border: { color: "#1D4ED8", width: 2 },
          },
        }),
        shape({
          id: "oct-b",
          name: "Octagon B",
          geometry: "octagon",
          style: {
            position: "absolute",
            left: 360,
            top: 120,
            width: 120,
            height: 120,
            background: "#DCFCE7",
            border: { color: "#15803D", width: 2 },
          },
        }),
        line({
          id: "oct-conn-4-6",
          name: "Octagon Link 4-6",
          routing: "elbow",
          stroke: { color: "#0F172A", width: 2, endArrow: "triangle" },
          from: { id: "oct-a", port: 5 },
          to: { id: "oct-b", port: 4 },
        }),
        line({
          id: "oct-conn-7",
          name: "Octagon Link 7",
          stroke: { color: "#DC2626", width: 2 },
          from: { id: "oct-a", port: 7 },
          to: { id: "oct-b", port: 6 },
        }),
      ),
    );

    const resolution = resolveDeck(doc);
    const emitted = emit(resolution);
    const validated = await validatePptx(emitted.pptx);
    assert.deepEqual(validated.errors, []);

    const slideXml = strFromU8(
      unzipSync(emitted.pptx)["ppt/slides/slide1.xml"]!,
    );
    assert.match(slideXml, /<a:stCxn id="\d+" idx="5"\/>/u);
    assert.match(slideXml, /<a:endCxn id="\d+" idx="4"\/>/u);
    assert.match(slideXml, /<a:stCxn id="\d+" idx="7"\/>/u);
    assert.match(slideXml, /<a:endCxn id="\d+" idx="6"\/>/u);

    const painted = paint(resolution);
    assert.deepEqual(painted.pages[0]?.errors, []);
    assert.ok((painted.pages[0]?.png?.byteLength ?? 0) > 0);

    const imported = importDeck(emitted.pptx);
    assert.deepEqual(imported.report.diagnostics, []);

    const collect = (root: CommonNode): CommonNode[] => {
      const out: CommonNode[] = [root];
      for (const ch of root.children) {
        if (ch instanceof CommonNode) out.push(...collect(ch));
      }
      return out;
    };
    const importedNodes = collect(imported.document);
    const conn1 = importedNodes.find(
      (n) => n.props["name"] === "Octagon Link 4-6",
    );
    const conn2 = importedNodes.find(
      (n) => n.props["name"] === "Octagon Link 7",
    );
    assert.equal((conn1?.props["from"] as { port: unknown })?.port, 5);
    assert.equal((conn1?.props["to"] as { port: unknown })?.port, 4);
    assert.equal((conn2?.props["from"] as { port: unknown })?.port, 7);
    assert.equal((conn2?.props["to"] as { port: unknown })?.port, 6);

    const vanilla = codegenDeck(imported.document, { target: "vanilla" });
    const react = codegenDeck(imported.document, { target: "react" });
    const vanillaEntry = vanilla.files[vanilla.entry];
    const reactEntry = react.files[react.entry];
    assert.equal(typeof vanillaEntry, "string");
    assert.equal(typeof reactEntry, "string");
    assert.match(vanillaEntry as string, /port:\s*5/u);
    assert.match(reactEntry as string, /port:\s*7/u);
    const batch = await verifyBatchGeneratedSourcesAndRelocatedRebuilds(
      [
        {
          scenarioId: "octagon-port-ge-4",
          originLabel: "roundtrip-imported",
          result: vanilla,
        },
        {
          scenarioId: "octagon-port-ge-4",
          originLabel: "roundtrip-imported",
          result: react,
        },
      ],
      [],
    );
    assert.equal(batch.typecheckedCaseCount, 2);
  });
});
