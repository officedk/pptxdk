import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import {
  build,
  defineSlideLayout,
  font,
  placeholderSizeSchema,
  resolve,
  type PlaceholderSize,
} from "@officedk/pptxdk";
import {
  Box,
  Cell,
  Deck,
  Group,
  Image,
  Placeholder,
  Row,
  Shape,
  Slide,
  SlideLayout,
  SlideMaster,
  Table,
  Text,
  expandDocument,
} from "@officedk/pptxdk/react";
import {
  box,
  cell,
  deck,
  group,
  image,
  placeholder,
  row,
  shape,
  slide,
  slideLayout,
  slideMaster,
  table,
  text,
} from "@officedk/pptxdk/vanilla";
import { validatePptx } from "@pptxdk/validation";
import { createElement } from "react";
import { describe, it } from "vitest";

const jostBookBytes = new Uint8Array(
  readFileSync(new URL("./assets/fonts/Jost-400-Book.ttf", import.meta.url)),
);

const jostFont = font("Jost*", [{ weight: 400, data: jostBookBytes }]);

const pixelPng = new Uint8Array(
  Buffer.from(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/n1sAAAAASUVORK5CYII=",
    "base64",
  ),
);

describe("Ticket 03: single-track layout and containers", () => {
  it("keeps Slide, SlideMaster, and SlideLayout as unconditional block layout roots and initial containing blocks regardless of inherited text styles", async () => {
    const layoutDef = defineSlideLayout({
      id: "styled-layout",
      style: {
        padding: 10,
        fontFamily: jostFont,
        fontSize: 16,
      },
    });

    const jsxDoc = createElement(
      Deck,
      { width: 640, height: 360 },
      createElement(
        SlideMaster,
        {
          id: "main-master",
          style: {
            padding: 16,
            fontFamily: jostFont,
            fontSize: 14,
            color: "#223344",
          },
        },
        createElement(Shape, {
          id: "master-flow",
          style: { height: 24, marginBottom: 8 },
        }),
        createElement(
          Text,
          {
            id: "master-text",
            style: { width: 200, height: 30 },
          },
          "Master inherited",
        ),
        createElement(Shape, {
          id: "master-abs",
          style: {
            position: "absolute",
            right: 12,
            bottom: 12,
            width: 60,
            height: 20,
          },
        }),
        createElement(
          SlideLayout,
          { layout: layoutDef },
          createElement(Shape, {
            id: "layout-flow",
            style: { height: 20 },
          }),
          createElement(
            Text,
            {
              id: "layout-text",
              style: { width: 180, height: 28 },
            },
            "Layout inherited",
          ),
          createElement(Shape, {
            id: "layout-abs",
            style: {
              position: "absolute",
              left: "50%",
              top: 40,
              width: 80,
              height: 24,
            },
          }),
        ),
      ),
      createElement(
        Slide,
        { id: "plain-slide", layout: layoutDef },
        createElement(Shape, {
          id: "flow-1",
          style: { height: 40, marginBottom: 10 },
        }),
        createElement(Shape, {
          id: "flow-2",
          style: { height: 30 },
        }),
        createElement(Shape, {
          id: "abs-pin",
          style: {
            position: "absolute",
            right: 20,
            bottom: 15,
            width: 100,
            height: 50,
          },
        }),
        createElement(Image, {
          id: "auto-img",
          src: pixelPng,
          style: { height: 40 },
        }),
      ),
      createElement(
        Slide,
        {
          id: "text-styled-slide",
          layout: layoutDef,
          style: { fontFamily: jostFont, fontSize: 18, color: "#112233" },
        },
        createElement(Shape, {
          id: "flow-1",
          style: { height: 40, marginBottom: 10 },
        }),
        createElement(Shape, {
          id: "flow-2",
          style: { height: 30 },
        }),
        createElement(Shape, {
          id: "abs-pin",
          style: {
            position: "absolute",
            right: 20,
            bottom: 15,
            width: 100,
            height: 50,
          },
        }),
        createElement(Image, {
          id: "auto-img",
          src: pixelPng,
          style: { height: 40 },
        }),
      ),
      createElement(
        Slide,
        {
          id: "flex-slide",
          layout: layoutDef,
          style: {
            display: "flex",
            flexDirection: "row",
            padding: 20,
            gap: 10,
          },
        },
        createElement(Box, {
          id: "flex-a",
          style: { width: 100, height: 50 },
        }),
        createElement(Box, {
          id: "flex-b",
          style: { flexGrow: 1, height: 50 },
        }),
      ),
    );

    const vanillaDoc = deck({ width: 640, height: 360 }).append(
      slideMaster({
        id: "main-master",
        style: {
          padding: 16,
          fontFamily: jostFont,
          fontSize: 14,
          color: "#223344",
        },
      }).append(
        shape({ id: "master-flow", style: { height: 24, marginBottom: 8 } }),
        text("Master inherited", {
          id: "master-text",
          style: { width: 200, height: 30 },
        }),
        shape({
          id: "master-abs",
          style: {
            position: "absolute",
            right: 12,
            bottom: 12,
            width: 60,
            height: 20,
          },
        }),
        slideLayout(layoutDef).append(
          shape({ id: "layout-flow", style: { height: 20 } }),
          text("Layout inherited", {
            id: "layout-text",
            style: { width: 180, height: 28 },
          }),
          shape({
            id: "layout-abs",
            style: {
              position: "absolute",
              left: "50%",
              top: 40,
              width: 80,
              height: 24,
            },
          }),
        ),
      ),
      slide({ id: "plain-slide", layout: layoutDef }).append(
        shape({ id: "flow-1", style: { height: 40, marginBottom: 10 } }),
        shape({ id: "flow-2", style: { height: 30 } }),
        shape({
          id: "abs-pin",
          style: {
            position: "absolute",
            right: 20,
            bottom: 15,
            width: 100,
            height: 50,
          },
        }),
        image({
          id: "auto-img",
          src: pixelPng,
          style: { height: 40 },
        }),
      ),
      slide({
        id: "text-styled-slide",
        layout: layoutDef,
        style: { fontFamily: jostFont, fontSize: 18, color: "#112233" },
      }).append(
        shape({ id: "flow-1", style: { height: 40, marginBottom: 10 } }),
        shape({ id: "flow-2", style: { height: 30 } }),
        shape({
          id: "abs-pin",
          style: {
            position: "absolute",
            right: 20,
            bottom: 15,
            width: 100,
            height: 50,
          },
        }),
        image({
          id: "auto-img",
          src: pixelPng,
          style: { height: 40 },
        }),
      ),
      slide({
        id: "flex-slide",
        layout: layoutDef,
        style: { display: "flex", flexDirection: "row", padding: 20, gap: 10 },
      }).append(
        box({ id: "flex-a", style: { width: 100, height: 50 } }),
        box({ id: "flex-b", style: { flexGrow: 1, height: 50 } }),
      ),
    );

    const jsxBuilt = await build(await expandDocument(jsxDoc));
    const vanillaRes = resolve(vanillaDoc);

    // Block root defaults: vertical stacking, width stretched to 640
    const plain = jsxBuilt.slide("plain-slide");
    const styled = jsxBuilt.slide("text-styled-slide");
    for (const s of [
      plain,
      styled,
      vanillaRes.slide("plain-slide"),
      vanillaRes.slide("text-styled-slide"),
    ]) {
      assert.deepEqual(s.get("flow-1").box, {
        x: 0,
        y: 0,
        width: 640,
        height: 40,
      });
      assert.deepEqual(s.get("flow-2").box, {
        x: 0,
        y: 50,
        width: 640,
        height: 30,
      });
      assert.deepEqual(s.get("abs-pin").box, {
        x: 520,
        y: 295,
        width: 100,
        height: 50,
      });
    }
    // Embedded auto-width Image with height=40 preserves aspect ratio (1:1 -> 40x40) instead of stretching to 640
    for (const s of [
      plain,
      styled,
      vanillaRes.slide("plain-slide"),
      vanillaRes.slide("text-styled-slide"),
    ]) {
      assert.deepEqual(s.get("auto-img").box, {
        x: 0,
        y: 80,
        width: 40,
        height: 40,
      });
    }

    // Slide.style.display = "flex" overrides default block root
    for (const flexSlide of [
      jsxBuilt.slide("flex-slide"),
      vanillaRes.slide("flex-slide"),
    ]) {
      assert.deepEqual(flexSlide.get("flex-a").box, {
        x: 20,
        y: 20,
        width: 100,
        height: 50,
      });
      assert.deepEqual(flexSlide.get("flex-b").box, {
        x: 130,
        y: 20,
        width: 490,
        height: 50,
      });
    }

    // SlideMaster and SlideLayout block layout roots + text style inheritance
    for (const report of [jsxBuilt.report, vanillaRes.report]) {
      const master = report.masters![0]!;
      const mFlow = master.objects.find((o) => o.authorId === "master-flow")!;
      const mText = master.objects.find((o) => o.authorId === "master-text")!;
      const mAbs = master.objects.find((o) => o.authorId === "master-abs")!;
      assert.deepEqual(mFlow.box, { x: 16, y: 16, width: 608, height: 24 });
      assert.deepEqual(mText.box, { x: 16, y: 48, width: 200, height: 30 });
      assert.deepEqual(mAbs.box, { x: 568, y: 328, width: 60, height: 20 });
      assert.equal(mText.textMetrics?.status, "measured");

      const layoutObs = master.layouts[0]!;
      const lFlow = layoutObs.objects.find(
        (o) => o.authorId === "layout-flow",
      )!;
      const lText = layoutObs.objects.find(
        (o) => o.authorId === "layout-text",
      )!;
      const lAbs = layoutObs.objects.find((o) => o.authorId === "layout-abs")!;
      assert.deepEqual(lFlow.box, { x: 10, y: 10, width: 620, height: 20 });
      assert.deepEqual(lText.box, { x: 10, y: 30, width: 180, height: 28 });
      assert.deepEqual(lAbs.box, { x: 320, y: 40, width: 80, height: 24 });
      assert.equal(lText.textMetrics?.status, "measured");
    }

    const validated = await validatePptx(jsxBuilt.pptx);
    assert.deepEqual(validated.errors, []);
  });

  it("applies CSS containing-block and frame rules to unpositioned and positioned Group across JSX and vanilla", async () => {
    const jsxDoc = createElement(
      Deck,
      { width: 800, height: 450 },
      createElement(
        Slide,
        { id: "group-semantics" },
        // 1. Unpositioned Group on Slide root -> absolute children escape to Slide; Group frame = union
        createElement(
          Group,
          { id: "root-unpositioned-group" },
          createElement(Shape, {
            id: "u1",
            style: {
              position: "absolute",
              left: 20,
              top: 30,
              width: 60,
              height: 40,
            },
          }),
          createElement(Shape, {
            id: "u2",
            style: {
              position: "absolute",
              left: 110,
              top: 50,
              width: 50,
              height: 60,
            },
          }),
        ),
        // 2. Unpositioned Group inside positioned Box -> absolute children reference Box; Group frame = union
        createElement(
          Box,
          {
            id: "pos-box",
            style: {
              position: "absolute",
              left: 200,
              top: 40,
              width: 240,
              height: 180,
            },
          },
          createElement(
            Group,
            { id: "nested-unpositioned-group", style: { position: "static" } },
            createElement(Shape, {
              id: "bu1",
              style: {
                position: "absolute",
                left: 15,
                top: 20,
                width: 50,
                height: 30,
              },
            }),
            createElement(Shape, {
              id: "bu2",
              style: {
                position: "absolute",
                right: 10,
                bottom: 15,
                width: 40,
                height: 25,
              },
            }),
          ),
        ),
        // 3. Positioned Group (absolute) with explicit width/height -> acts as CB and preserves explicit size
        createElement(
          Group,
          {
            id: "pos-abs-group",
            style: {
              position: "absolute",
              left: 480,
              top: 60,
              width: 260,
              height: 150,
            },
          },
          createElement(Shape, {
            id: "pa1",
            style: {
              position: "absolute",
              left: 20,
              top: 15,
              width: 70,
              height: 30,
            },
          }),
          createElement(Shape, {
            id: "pa2",
            style: {
              position: "absolute",
              right: 10,
              bottom: 10,
              width: 60,
              height: 40,
            },
          }),
        ),
        // 4. Positioned Group (relative) in block flow with explicit size -> acts as CB
        createElement(
          Group,
          {
            id: "pos-rel-group",
            style: {
              position: "relative",
              width: 200,
              height: 90,
              marginTop: 250,
              marginLeft: 30,
            },
          },
          createElement(Shape, {
            id: "pr1",
            style: {
              position: "absolute",
              left: "50%",
              top: 10,
              width: 80,
              height: 30,
            },
          }),
        ),
      ),
    );

    const vanillaDoc = deck({ width: 800, height: 450 }).append(
      slide({ id: "group-semantics" }).append(
        group({ id: "root-unpositioned-group" }).append(
          shape({
            id: "u1",
            style: {
              position: "absolute",
              left: 20,
              top: 30,
              width: 60,
              height: 40,
            },
          }),
          shape({
            id: "u2",
            style: {
              position: "absolute",
              left: 110,
              top: 50,
              width: 50,
              height: 60,
            },
          }),
        ),
        box({
          id: "pos-box",
          style: {
            position: "absolute",
            left: 200,
            top: 40,
            width: 240,
            height: 180,
          },
        }).append(
          group({
            id: "nested-unpositioned-group",
            style: { position: "static" },
          }).append(
            shape({
              id: "bu1",
              style: {
                position: "absolute",
                left: 15,
                top: 20,
                width: 50,
                height: 30,
              },
            }),
            shape({
              id: "bu2",
              style: {
                position: "absolute",
                right: 10,
                bottom: 15,
                width: 40,
                height: 25,
              },
            }),
          ),
        ),
        group({
          id: "pos-abs-group",
          style: {
            position: "absolute",
            left: 480,
            top: 60,
            width: 260,
            height: 150,
          },
        }).append(
          shape({
            id: "pa1",
            style: {
              position: "absolute",
              left: 20,
              top: 15,
              width: 70,
              height: 30,
            },
          }),
          shape({
            id: "pa2",
            style: {
              position: "absolute",
              right: 10,
              bottom: 10,
              width: 60,
              height: 40,
            },
          }),
        ),
        group({
          id: "pos-rel-group",
          style: {
            position: "relative",
            width: 200,
            height: 90,
            marginTop: 250,
            marginLeft: 30,
          },
        }).append(
          shape({
            id: "pr1",
            style: {
              position: "absolute",
              left: "50%",
              top: 10,
              width: 80,
              height: 30,
            },
          }),
        ),
      ),
    );

    const jsxBuilt = await build(await expandDocument(jsxDoc));
    const vanillaBuilt = await build(vanillaDoc);

    for (const built of [jsxBuilt, vanillaBuilt]) {
      const s = built.slide("group-semantics");
      const rootGroup = s.get("root-unpositioned-group");
      assert.deepEqual(rootGroup.get("u1").box, {
        x: 20,
        y: 30,
        width: 60,
        height: 40,
      });
      assert.deepEqual(rootGroup.get("u2").box, {
        x: 110,
        y: 50,
        width: 50,
        height: 60,
      });
      assert.deepEqual(rootGroup.box, {
        x: 20,
        y: 30,
        width: 140,
        height: 80,
      });

      const nestedGroup = s.get("pos-box").get("nested-unpositioned-group");
      assert.deepEqual(nestedGroup.get("bu1").box, {
        x: 215,
        y: 60,
        width: 50,
        height: 30,
      });
      assert.deepEqual(nestedGroup.get("bu2").box, {
        x: 390,
        y: 180,
        width: 40,
        height: 25,
      });
      assert.deepEqual(nestedGroup.box, {
        x: 215,
        y: 60,
        width: 215,
        height: 145,
      });

      const posAbsGroup = s.get("pos-abs-group");
      assert.deepEqual(posAbsGroup.box, {
        x: 480,
        y: 60,
        width: 260,
        height: 150,
      });
      assert.deepEqual(posAbsGroup.get("pa1").box, {
        x: 500,
        y: 75,
        width: 70,
        height: 30,
      });
      assert.deepEqual(posAbsGroup.get("pa2").box, {
        x: 670,
        y: 160,
        width: 60,
        height: 40,
      });

      const posRelGroup = s.get("pos-rel-group");
      assert.deepEqual(posRelGroup.box, {
        x: 30,
        y: 250,
        width: 200,
        height: 90,
      });
      assert.deepEqual(posRelGroup.get("pr1").box, {
        x: 130,
        y: 260,
        width: 80,
        height: 30,
      });
    }

    const validated = await validatePptx(jsxBuilt.pptx);
    assert.deepEqual(validated.errors, []);
    const slideValidation = validated.slides[0]!;
    const posAbsGroupObj = slideValidation.objects.find(
      (obj) => obj.kind === "group" && obj.boxEmu.x === 480 * 9525,
    );
    assert.ok(posAbsGroupObj !== undefined);
    assert.deepEqual(posAbsGroupObj.boxEmu, {
      x: 480 * 9525,
      y: 60 * 9525,
      width: 260 * 9525,
      height: 150 * 9525,
    });
  });

  it("enforces Table target dimensions, auto-row excess distribution, explicit-row preservation, and min-content feedback to prevent sibling overlap", async () => {
    const jsxDoc = createElement(
      Deck,
      { width: 800, height: 500 },
      createElement(
        Slide,
        { id: "table-grid-sizing", style: { padding: 20 } },
        // 1. Table with target height > base rows and auto rows -> excess distributed equally to auto rows
        createElement(
          Box,
          {
            id: "auto-grow-wrap",
            style: {
              display: "flex",
              flexDirection: "row",
              gap: 16,
              alignItems: "flex-start",
            },
          },
          createElement(
            Table,
            {
              id: "tbl-auto-grow",
              columns: [100, 100],
              style: {
                width: 200,
                height: 150,
                fontFamily: jostFont,
                fontSize: 12,
              },
            },
            createElement(
              Row,
              { style: { height: 30 } },
              createElement(Cell, null, "Fixed"),
              createElement(Cell, null, "30"),
            ),
            createElement(
              Row,
              null,
              createElement(Cell, null, "Auto A"),
              createElement(Cell, null, "A"),
            ),
            createElement(
              Row,
              null,
              createElement(Cell, null, "Auto B"),
              createElement(Cell, null, "B"),
            ),
          ),
          // 2. Table with ALL explicit rows and larger target height -> keeps explicit row sum (80)
          createElement(
            Box,
            {
              id: "explicit-col",
              style: { display: "flex", flexDirection: "column" },
            },
            createElement(
              Table,
              {
                id: "tbl-all-explicit",
                columns: [90, 90],
                style: {
                  width: 180,
                  height: 180,
                  fontFamily: jostFont,
                  fontSize: 12,
                },
              },
              createElement(
                Row,
                { style: { height: 35 } },
                createElement(Cell, null, "R1"),
                createElement(Cell, null, "35"),
              ),
              createElement(
                Row,
                { style: { height: 45 } },
                createElement(Cell, null, "R2"),
                createElement(Cell, null, "45"),
              ),
            ),
            createElement(Shape, {
              id: "after-explicit-tbl",
              style: { width: 180, height: 24 },
            }),
          ),
          // 3. Table whose column sum (240) and base row sum (100) exceed target style (140 x 50) -> feeds actual size to flex sibling
          createElement(
            Table,
            {
              id: "tbl-overflow-target",
              columns: [120, 120],
              style: {
                width: 140,
                height: 50,
                fontFamily: jostFont,
                fontSize: 12,
              },
            },
            createElement(
              Row,
              { style: { height: 40 } },
              createElement(Cell, null, "C1"),
              createElement(Cell, null, "C2"),
            ),
            createElement(
              Row,
              { style: { height: 60 } },
              createElement(Cell, null, "C3"),
              createElement(Cell, null, "C4"),
            ),
          ),
          createElement(Shape, {
            id: "flex-sibling-after-overflow",
            style: { width: 60, height: 40 },
          }),
        ),
        // 4. Block flow sibling below an overflowing table never overlaps
        createElement(
          Box,
          { id: "block-flow-check", style: { marginTop: 12 } },
          createElement(
            Table,
            {
              id: "tbl-block-overflow",
              columns: [150, 150],
              style: {
                width: 200,
                height: 40,
                fontFamily: jostFont,
                fontSize: 12,
              },
            },
            createElement(
              Row,
              { style: { height: 55 } },
              createElement(Cell, null, "Row 1"),
              createElement(Cell, null, "55"),
            ),
            createElement(
              Row,
              { style: { height: 55 } },
              createElement(Cell, null, "Row 2"),
              createElement(Cell, null, "55"),
            ),
          ),
          createElement(Shape, {
            id: "block-sibling-after",
            style: { width: 200, height: 30 },
          }),
        ),
        // 5. Grid container with full LayoutStyle hosting Table & Shape
        createElement(
          Box,
          {
            id: "grid-host",
            style: {
              position: "absolute",
              left: 20,
              top: 330,
              width: 500,
              height: 140,
              display: "grid",
              gridTemplateColumns: "auto 1fr",
              gridTemplateRows: "auto",
              columnGap: 16,
            },
          },
          createElement(
            Table,
            {
              id: "tbl-in-grid",
              columns: [110, 110],
              style: {
                width: 150,
                height: 40,
                fontFamily: jostFont,
                fontSize: 12,
              },
            },
            createElement(
              Row,
              { style: { height: 45 } },
              createElement(Cell, null, "G1"),
              createElement(Cell, null, "G2"),
            ),
            createElement(
              Row,
              { style: { height: 45 } },
              createElement(Cell, null, "G3"),
              createElement(Cell, null, "G4"),
            ),
          ),
          createElement(Shape, {
            id: "grid-sibling",
            style: { height: 50 },
          }),
        ),
      ),
    );

    const vanillaDoc = deck({ width: 800, height: 500 }).append(
      slide({ id: "table-grid-sizing", style: { padding: 20 } }).append(
        box({
          id: "auto-grow-wrap",
          style: {
            display: "flex",
            flexDirection: "row",
            gap: 16,
            alignItems: "flex-start",
          },
        }).append(
          table({
            id: "tbl-auto-grow",
            columns: [100, 100],
            style: {
              width: 200,
              height: 150,
              fontFamily: jostFont,
              fontSize: 12,
            },
          }).append(
            row({ style: { height: 30 } }).append(cell("Fixed"), cell("30")),
            row().append(cell("Auto A"), cell("A")),
            row().append(cell("Auto B"), cell("B")),
          ),
          box({
            id: "explicit-col",
            style: { display: "flex", flexDirection: "column" },
          }).append(
            table({
              id: "tbl-all-explicit",
              columns: [90, 90],
              style: {
                width: 180,
                height: 180,
                fontFamily: jostFont,
                fontSize: 12,
              },
            }).append(
              row({ style: { height: 35 } }).append(cell("R1"), cell("35")),
              row({ style: { height: 45 } }).append(cell("R2"), cell("45")),
            ),
            shape({
              id: "after-explicit-tbl",
              style: { width: 180, height: 24 },
            }),
          ),
          table({
            id: "tbl-overflow-target",
            columns: [120, 120],
            style: {
              width: 140,
              height: 50,
              fontFamily: jostFont,
              fontSize: 12,
            },
          }).append(
            row({ style: { height: 40 } }).append(cell("C1"), cell("C2")),
            row({ style: { height: 60 } }).append(cell("C3"), cell("C4")),
          ),
          shape({
            id: "flex-sibling-after-overflow",
            style: { width: 60, height: 40 },
          }),
        ),
        box({ id: "block-flow-check", style: { marginTop: 12 } }).append(
          table({
            id: "tbl-block-overflow",
            columns: [150, 150],
            style: {
              width: 200,
              height: 40,
              fontFamily: jostFont,
              fontSize: 12,
            },
          }).append(
            row({ style: { height: 55 } }).append(cell("Row 1"), cell("55")),
            row({ style: { height: 55 } }).append(cell("Row 2"), cell("55")),
          ),
          shape({
            id: "block-sibling-after",
            style: { width: 200, height: 30 },
          }),
        ),
        box({
          id: "grid-host",
          style: {
            position: "absolute",
            left: 20,
            top: 330,
            width: 500,
            height: 140,
            display: "grid",
            gridTemplateColumns: "auto 1fr",
            gridTemplateRows: "auto",
            columnGap: 16,
          },
        }).append(
          table({
            id: "tbl-in-grid",
            columns: [110, 110],
            style: {
              width: 150,
              height: 40,
              fontFamily: jostFont,
              fontSize: 12,
            },
          }).append(
            row({ style: { height: 45 } }).append(cell("G1"), cell("G2")),
            row({ style: { height: 45 } }).append(cell("G3"), cell("G4")),
          ),
          shape({ id: "grid-sibling", style: { height: 50 } }),
        ),
      ),
    );

    const jsxBuilt = await build(await expandDocument(jsxDoc));
    const vanillaBuilt = await build(vanillaDoc);

    for (const built of [jsxBuilt, vanillaBuilt]) {
      const s = built.slide("table-grid-sizing");
      const wrap = s.get("auto-grow-wrap");
      const autoGrow = wrap.get("tbl-auto-grow");
      assert.deepEqual(autoGrow.box, { x: 20, y: 20, width: 200, height: 150 });
      assert.equal(autoGrow.objects[0]?.table?.rows[0]?.height, 30);
      assert.equal(autoGrow.objects[0]?.table?.rows[1]?.height, 60);
      assert.equal(autoGrow.objects[0]?.table?.rows[2]?.height, 60);

      const explicitCol = wrap.get("explicit-col");
      const explicitTbl = explicitCol.get("tbl-all-explicit");
      assert.deepEqual(explicitTbl.box, {
        x: 236,
        y: 20,
        width: 180,
        height: 80,
      });
      assert.deepEqual(explicitTbl.objects[0]?.table?.rowHeights, [35, 45]);
      assert.deepEqual(explicitCol.get("after-explicit-tbl").box, {
        x: 236,
        y: 100,
        width: 180,
        height: 24,
      });

      const overflowTbl = wrap.get("tbl-overflow-target");
      assert.deepEqual(overflowTbl.box, {
        x: 432,
        y: 20,
        width: 240,
        height: 100,
      });
      const flexSibling = wrap.get("flex-sibling-after-overflow");
      assert.deepEqual(flexSibling.box, {
        x: 688,
        y: 20,
        width: 60,
        height: 40,
      });

      const blockFlow = s.get("block-flow-check");
      const blockTbl = blockFlow.get("tbl-block-overflow");
      assert.deepEqual(blockTbl.box, {
        x: 20,
        y: 182,
        width: 300,
        height: 110,
      });
      assert.deepEqual(blockFlow.get("block-sibling-after").box, {
        x: 20,
        y: 292,
        width: 200,
        height: 30,
      });

      const gridHost = s.get("grid-host");
      const gridTbl = gridHost.get("tbl-in-grid");
      assert.deepEqual(gridTbl.box, { x: 20, y: 330, width: 220, height: 90 });
      assert.deepEqual(gridHost.get("grid-sibling").box, {
        x: 256,
        y: 330,
        width: 264,
        height: 50,
      });
    }

    const validated = await validatePptx(jsxBuilt.pptx);
    assert.deepEqual(validated.errors, []);
  });

  it("supports style (LayoutStyle) and PlaceholderSize (size) across master, layout, footer, and slide Placeholder specs", async () => {
    const phSize: PlaceholderSize = placeholderSizeSchema.parse("half");
    assert.equal(phSize, "half");

    const phLayout = defineSlideLayout({
      id: "ph-layout",
      placeholders: {
        title: {
          type: "title",
          size: "full",
          style: {
            position: "absolute",
            left: 30,
            top: 24,
            width: 580,
            height: 48,
            fontFamily: jostFont,
            fontSize: 22,
          },
        },
        content: {
          type: "obj",
          idx: 1,
          size: phSize,
          style: {
            position: "absolute",
            left: 30,
            top: 90,
            right: 30,
            height: 180,
          },
        },
      },
      footerPlaceholders: {
        footer: {
          text: "Layout Footer",
          size: "quarter",
          style: {
            position: "absolute",
            left: 24,
            bottom: 12,
            width: 200,
            height: 24,
            fontFamily: jostFont,
            fontSize: 10,
          },
        },
      },
    });

    const jsxDoc = createElement(
      Deck,
      { width: 640, height: 360 },
      createElement(
        SlideMaster,
        {
          id: "ph-master",
          placeholders: {
            title: {
              type: "title",
              size: "full",
              style: { left: 40, top: 20, width: 560, height: 50 },
            },
          },
          headerFooter: { footer: true },
        },
        createElement(SlideLayout, { layout: phLayout }),
      ),
      createElement(
        Slide,
        {
          id: "inherited-ph",
          layout: phLayout,
          placeholders: {
            title: "Inherited Title",
          },
        },
        createElement(Placeholder, {
          id: "slot-override",
          placeholder: "content",
          style: {
            position: "absolute",
            left: 50,
            top: 110,
            width: 300,
            height: 140,
          },
        }),
      ),
      createElement(
        Slide,
        {
          id: "pct-ph",
          layout: phLayout,
        },
        createElement(
          Text,
          {
            id: "body-override",
            placeholder: "content",
            style: {
              position: "absolute",
              right: 40,
              bottom: 50,
              width: "50%",
              height: 120,
            },
          },
          "Percent positioned placeholder",
        ),
      ),
    );

    const vanillaDoc = deck({ width: 640, height: 360 }).append(
      slideMaster({
        id: "ph-master",
        placeholders: {
          title: {
            type: "title",
            size: "full",
            style: { left: 40, top: 20, width: 560, height: 50 },
          },
        },
        headerFooter: { footer: true },
      }).append(slideLayout(phLayout)),
      slide({
        id: "inherited-ph",
        layout: phLayout,
        placeholders: {
          title: "Inherited Title",
        },
      }).append(
        placeholder({
          id: "slot-override",
          placeholder: "content",
          style: {
            position: "absolute",
            left: 50,
            top: 110,
            width: 300,
            height: 140,
          },
        }),
      ),
      slide({
        id: "pct-ph",
        layout: phLayout,
      }).append(
        text("Percent positioned placeholder", {
          id: "body-override",
          placeholder: "content",
          style: {
            position: "absolute",
            right: 40,
            bottom: 50,
            width: "50%",
            height: 120,
          },
        }),
      ),
    );

    const jsxBuilt = await build(await expandDocument(jsxDoc));
    const vanillaBuilt = await build(vanillaDoc);

    for (const built of [jsxBuilt, vanillaBuilt]) {
      const master = built.report.masters![0]!;
      const masterTitlePh = master.placeholders.find((p) => p.key === "title");
      assert.equal(masterTitlePh?.size, "full");
      assert.deepEqual(masterTitlePh?.box, {
        x: 40,
        y: 20,
        width: 560,
        height: 50,
      });

      const layoutObs = master.layouts[0]!;
      const layoutTitlePh = layoutObs.placeholders.find(
        (p) => p.key === "title",
      );
      const layoutContentPh = layoutObs.placeholders.find(
        (p) => p.key === "content",
      );
      assert.equal(layoutTitlePh?.size, "full");
      assert.deepEqual(layoutTitlePh?.box, {
        x: 30,
        y: 24,
        width: 580,
        height: 48,
      });
      assert.equal(layoutContentPh?.size, "half");
      assert.deepEqual(layoutContentPh?.box, {
        x: 30,
        y: 90,
        width: 580,
        height: 180,
      });

      const s1 = built.slide("inherited-ph");
      const titleObj = s1.objects.find((o) => o.kind === "text");
      assert.deepEqual(titleObj?.box, { x: 30, y: 24, width: 580, height: 48 });
      assert.deepEqual(s1.get("slot-override").box, {
        x: 50,
        y: 110,
        width: 300,
        height: 140,
      });
      assert.deepEqual(s1.footer?.footer?.box, {
        x: 24,
        y: 324,
        width: 200,
        height: 24,
      });

      const s2 = built.slide("pct-ph");
      assert.deepEqual(s2.get("body-override").box, {
        x: 280,
        y: 190,
        width: 320,
        height: 120,
      });
    }

    const validated = await validatePptx(jsxBuilt.pptx);
    assert.deepEqual(validated.errors, []);
  });
});
