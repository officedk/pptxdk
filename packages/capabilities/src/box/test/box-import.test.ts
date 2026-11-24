import assert from "node:assert/strict";

import { documentCapabilities } from "@pptxdk/capabilities/all";
import { box, deck, shape, slide } from "@pptxdk/capabilities/document";
import { CommonNode } from "@pptxdk/core";
import { createResolver } from "@pptxdk/core/build";
import { emit } from "@pptxdk/emit";
import { createImporter } from "@pptxdk/import";
import { paint } from "@pptxdk/preview";
import { validatePptx } from "@pptxdk/validation";
import { describe, it } from "vitest";

const resolveDocument = createResolver(documentCapabilities);
const importDocument = createImporter(documentCapabilities);

function findNodesByType(
  root: CommonNode,
  type: string,
): readonly CommonNode[] {
  const result: CommonNode[] = [];
  const visit = (node: CommonNode): void => {
    if (node.type === type) result.push(node);
    for (const child of node.children) {
      if (child instanceof CommonNode) visit(child);
    }
  };
  visit(root);
  return result;
}

describe("Box expanded layout OOXML import", () => {
  it("reconstructs expanded Flex and Grid containers preserving visible Boxes as Shapes and dropping transparent Boxes", async () => {
    const doc = deck({ width: 640, height: 360 }).append(
      slide({ id: "layout-reconstruction" }).append(
        box({
          id: "transparent-flex",
          style: {
            position: "absolute",
            left: 24,
            top: 24,
            width: 260,
            display: "flex",
            flexDirection: "row",
            gap: 16,
          },
        }).append(
          shape({
            name: "Flex A",
            width: 80,
            height: 44,
            fill: "#2563EB",
          }),
          shape({
            name: "Flex B",
            width: 90,
            height: 44,
            fill: "#16A34A",
          }),
        ),
        box({
          id: "transparent-grid",
          style: {
            position: "absolute",
            left: 24,
            top: 100,
            width: 300,
            display: "grid",
            gridTemplateColumns: "1fr 1fr",
            columnGap: 12,
            rowGap: 12,
          },
        }).append(
          box({
            id: "visible-grid-card",
            style: {
              padding: 10,
              display: "flex",
              flexDirection: "column",
              backgroundColor: "#FEF3C7",
              borderColor: "#D97706",
              borderWidth: 2,
              borderRadius: 8,
            },
          }).append(
            shape({
              name: "Grid Card Leaf",
              width: 60,
              height: 30,
              fill: "#B45309",
            }),
          ),
          shape({
            name: "Grid Sibling",
            width: 80,
            height: 50,
            fill: "#7C3AED",
          }),
        ),
      ),
    );

    const initialResolution = resolveDocument(doc);
    const initialPptx = emit(initialResolution).pptx;
    const initialValidation = await validatePptx(initialPptx);
    assert.deepEqual(initialValidation.errors, []);

    const imported = importDocument(initialPptx);
    assert.deepEqual(imported.report.diagnostics, []);

    const roundtripResolution = resolveDocument(imported.document);
    const roundtripPptx = emit(roundtripResolution).pptx;
    const roundtripPaint = paint(roundtripResolution);
    const roundtripValidation = await validatePptx(roundtripPptx);

    assert.deepEqual(roundtripValidation.errors, []);
    assert.deepEqual(roundtripPaint.pages[0]?.diagnostics, []);
    assert.ok(roundtripPaint.pages[0]?.png instanceof Uint8Array);

    const initialObjects = initialValidation.slides[0]!.objects;
    const roundtripObjects = roundtripValidation.slides[0]!.objects;
    assert.equal(initialObjects.length, 5);
    assert.equal(roundtripObjects.length, 5);
    assert.deepEqual(
      roundtripObjects.map((o) => ({
        name: o.name,
        geometry: o.geometry,
        boxEmu: o.boxEmu,
      })),
      initialObjects.map((o) => ({
        name: o.name,
        geometry: o.geometry,
        boxEmu: o.boxEmu,
      })),
    );

    const importedShapes = findNodesByType(imported.document, "shape");
    assert.equal(importedShapes.length, 5);
  });
});
