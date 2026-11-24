import assert from "node:assert/strict";
import {
  mkdir,
  mkdtemp,
  readFile,
  rm,
  symlink,
  writeFile,
} from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";

import { validatePptx } from "@pptxdk/validation";
import { strFromU8, unzipSync } from "fflate";
import { PNG } from "pngjs";
import { build, codegen, importPptx } from "pptxdk";
import { buildFromSource, previewFromSource } from "pptxdk/node";
import {
  Box,
  Deck,
  Group,
  P,
  Shape,
  Slide,
  Span,
  Text,
  blockArc,
  createMorphScope,
  expandDocument,
  path,
  preset,
} from "pptxdk/react";
import { box, deck, group, shape, slide, text } from "pptxdk/vanilla";
import { createElement } from "react";
import { describe, it } from "vitest";

import {
  rebuildGeneratedSourceRelocated,
  verifyGeneratedSourcesTypecheck,
  writeCodegenFiles,
} from "./helpers/codegen-acceptance.ts";

const packageRoot = fileURLToPath(new URL("../", import.meta.url));
const emuPerPx = 9525;

async function linkConsumerDeps(dir: string): Promise<void> {
  const nodeModulesDir = join(dir, "node_modules");
  await mkdir(nodeModulesDir, { recursive: true });
  await symlink(packageRoot, join(nodeModulesDir, "pptxdk"), "dir");
  await symlink(
    join(packageRoot, "node_modules/react"),
    join(nodeModulesDir, "react"),
    "dir",
  );
}

const compositeCustomGeometry = {
  type: "custom" as const,
  width: 100,
  height: 80,
  paths: [
    {
      commands: [
        { op: "M" as const, x: 0, y: 0 },
        { op: "L" as const, x: 100, y: 0 },
        { op: "L" as const, x: 100, y: 80 },
        { op: "Z" as const },
      ],
      fill: true,
      stroke: false,
    },
    {
      commands: [
        { op: "M" as const, x: 10, y: 70 },
        { op: "C" as const, x1: 20, y1: 10, x2: 80, y2: 10, x: 90, y: 70 },
        { op: "Q" as const, x1: 50, y1: 80, x: 10, y: 70 },
      ],
      fill: false,
      stroke: true,
    },
  ],
};

function createDirectVanillaDeck() {
  const morphScope = createMorphScope();
  return deck({ width: 960, height: 540 }).append(
    slide({ id: "geometry-slide" }).append(
      shape({
        id: "preset-round",
        name: "Preset Round",
        link: { url: "https://example.com" },
        geometry: preset("roundRect", { adj: 20000 }),
        x: 24,
        y: 24,
        width: 150,
        height: 80,
        fill: "#2563EB",
        style: {
          borderColor: "#1E293B",
          borderWidth: 2,
          borderStyle: "dashed",
        },
      }),
      shape({
        id: "morph-badge",
        name: undefined,
        morphKey: morphScope.handle("card"),
        geometry: "ellipse",
        x: 250,
        y: 160,
        width: 80,
        height: 60,
        fill: "#0284C7",
      }),
      shape({
        id: "preset-arc",
        name: "Preset BlockArc",
        geometry: blockArc({ startAngle: 0, endAngle: 180, thickness: 0.3 }),
        x: 200,
        y: 24,
        width: 100,
        height: 100,
        fill: "#10B981",
      }),
      shape({
        id: "single-bezier",
        name: "Single Bezier",
        geometry: path("M 10 80 C 10 15 90 15 90 80 Q 50 95 10 80 Z", {
          viewBox: [100, 100],
          fill: true,
          stroke: true,
        }),
        x: 330,
        y: 24,
        width: 120,
        height: 100,
        fill: "#F59E0B",
        flipH: true,
        flipV: true,
        style: {
          borderColor: "#78350F",
          borderWidth: 2,
          rotate: 15,
        },
      }),
      shape({
        id: "composite-custom",
        name: "Composite Custom",
        geometry: compositeCustomGeometry,
        x: 480,
        y: 24,
        width: 120,
        height: 96,
        fill: "#0EA5E9",
        style: {
          borderColor: "#0369A1",
          borderWidth: 2,
        },
      }),
      shape({
        id: "arrow-connector",
        name: "Arrow Connector",
        geometry: "line",
        x: 640,
        y: 64,
        width: 180,
        height: 2,
        style: {
          borderColor: "#DC2626",
          borderWidth: 3,
          startArrow: "oval",
          endArrow: "triangle",
        },
      }),
    ),
    slide({ id: "groups-and-layout" }).append(
      group({
        id: "outer-cluster",
        name: "Outer Cluster",
        style: { rotate: 12 },
      }).append(
        shape({
          id: "outer-hex",
          name: "Outer Hex",
          geometry: "hexagon",
          x: 40,
          y: 50,
          width: 90,
          height: 60,
          fill: "#2563EB",
          flipH: true,
        }),
        group({
          id: "inner-cluster",
          name: "Inner Cluster",
          style: { rotate: -18 },
        }).append(
          shape({
            id: "inner-diamond",
            name: "Inner Diamond",
            geometry: "diamond",
            x: 160,
            y: 60,
            width: 80,
            height: 60,
            fill: "#10B981",
            flipV: true,
            style: { rotate: 30 },
          }),
        ),
      ),
      box({
        id: "flex-wrap-box",
        style: {
          position: "absolute",
          left: 360,
          top: 40,
          width: 300,
          display: "flex",
          flexDirection: "row",
          flexWrap: "wrap",
          gap: 12,
          padding: 12,
          backgroundColor: "#F1F5F9",
          borderColor: "#94A3B8",
          borderWidth: 1,
          borderRadius: 6,
        },
      }).append(
        shape({
          id: "flex-a",
          name: "Flex A",
          width: 120,
          height: 44,
          fill: "#3B82F6",
        }),
        shape({
          id: "flex-b",
          name: "Flex B",
          width: 120,
          height: 44,
          fill: "#22C55E",
        }),
      ),
      box({
        id: "grid-box",
        style: {
          position: "absolute",
          left: 360,
          top: 220,
          width: 280,
          display: "grid",
          gridTemplateColumns: "1fr 1fr",
          columnGap: 12,
          rowGap: 12,
        },
      }).append(
        shape({
          id: "grid-a",
          name: "Grid A",
          height: 44,
          fill: "#F97316",
        }),
        shape({
          id: "grid-b",
          name: "Grid B",
          height: 44,
          fill: "#A855F7",
        }),
      ),
    ),
    slide({ id: "shape-text-slide" }).append(
      shape({
        id: "shorthand-shape",
        name: "Shorthand Shape",
        geometry: "roundRect",
        x: 24,
        y: 24,
        width: 180,
        height: 80,
        fill: "#1E293B",
        text: "Rounded",
        textStyle: { fontSize: 16, color: "#FFFFFF" },
      }),
      shape({
        id: "rich-shape",
        name: "Rich Shape",
        geometry: "diamond",
        x: 240,
        y: 24,
        width: 200,
        height: 90,
        fill: "#7C3AED",
      }).append("Rich Text"),
      group({
        id: "text-cluster",
        name: "Text Cluster",
      }).append(
        text("Group Label", {
          id: "inner-label",
          name: "Inner Label",
          x: 40,
          y: 160,
          width: 120,
          height: 32,
        }),
      ),
    ),
  );
}

async function createDirectReactDeck() {
  const morphScope = createMorphScope();
  return expandDocument(
    createElement(
      Deck,
      { width: 960, height: 540 },
      createElement(
        Slide,
        { id: "geometry-slide" },
        createElement(Shape, {
          id: "preset-round",
          name: "Preset Round",
          link: { url: "https://example.com" },
          geometry: preset("roundRect", { adj: 20000 }),
          x: 24,
          y: 24,
          width: 150,
          height: 80,
          fill: "#2563EB",
          style: {
            borderColor: "#1E293B",
            borderWidth: 2,
            borderStyle: "dashed",
          },
        }),
        createElement(Shape, {
          id: "morph-badge",
          morphKey: morphScope.handle("card"),
          geometry: "ellipse",
          x: 250,
          y: 160,
          width: 80,
          height: 60,
          fill: "#0284C7",
        }),
        createElement(Shape, {
          id: "preset-arc",
          name: "Preset BlockArc",
          geometry: blockArc({ startAngle: 0, endAngle: 180, thickness: 0.3 }),
          x: 200,
          y: 24,
          width: 100,
          height: 100,
          fill: "#10B981",
        }),
        createElement(Shape, {
          id: "single-bezier",
          name: "Single Bezier",
          geometry: path("M 10 80 C 10 15 90 15 90 80 Q 50 95 10 80 Z", {
            viewBox: [100, 100],
            fill: true,
            stroke: true,
          }),
          x: 330,
          y: 24,
          width: 120,
          height: 100,
          fill: "#F59E0B",
          flipH: true,
          flipV: true,
          style: {
            borderColor: "#78350F",
            borderWidth: 2,
            rotate: 15,
          },
        }),
        createElement(Shape, {
          id: "composite-custom",
          name: "Composite Custom",
          geometry: compositeCustomGeometry,
          x: 480,
          y: 24,
          width: 120,
          height: 96,
          fill: "#0EA5E9",
          style: {
            borderColor: "#0369A1",
            borderWidth: 2,
          },
        }),
        createElement(Shape, {
          id: "arrow-connector",
          name: "Arrow Connector",
          geometry: "line",
          x: 640,
          y: 64,
          width: 180,
          height: 2,
          style: {
            borderColor: "#DC2626",
            borderWidth: 3,
            startArrow: "oval",
            endArrow: "triangle",
          },
        }),
      ),
      createElement(
        Slide,
        { id: "groups-and-layout" },
        createElement(
          Group,
          {
            id: "outer-cluster",
            name: "Outer Cluster",
            style: { rotate: 12 },
          },
          createElement(Shape, {
            id: "outer-hex",
            name: "Outer Hex",
            geometry: "hexagon",
            x: 40,
            y: 50,
            width: 90,
            height: 60,
            fill: "#2563EB",
            flipH: true,
          }),
          createElement(
            Group,
            {
              id: "inner-cluster",
              name: "Inner Cluster",
              style: { rotate: -18 },
            },
            createElement(Shape, {
              id: "inner-diamond",
              name: "Inner Diamond",
              geometry: "diamond",
              x: 160,
              y: 60,
              width: 80,
              height: 60,
              fill: "#10B981",
              flipV: true,
              style: { rotate: 30 },
            }),
          ),
        ),
        createElement(
          Box,
          {
            id: "flex-wrap-box",
            style: {
              position: "absolute",
              left: 360,
              top: 40,
              width: 300,
              display: "flex",
              flexDirection: "row",
              flexWrap: "wrap",
              gap: 12,
              padding: 12,
              backgroundColor: "#F1F5F9",
              borderColor: "#94A3B8",
              borderWidth: 1,
              borderRadius: 6,
            },
          },
          createElement(Shape, {
            id: "flex-a",
            name: "Flex A",
            width: 120,
            height: 44,
            fill: "#3B82F6",
          }),
          createElement(Shape, {
            id: "flex-b",
            name: "Flex B",
            width: 120,
            height: 44,
            fill: "#22C55E",
          }),
        ),
        createElement(
          Box,
          {
            id: "grid-box",
            style: {
              position: "absolute",
              left: 360,
              top: 220,
              width: 280,
              display: "grid",
              gridTemplateColumns: "1fr 1fr",
              columnGap: 12,
              rowGap: 12,
            },
          },
          createElement(Shape, {
            id: "grid-a",
            name: "Grid A",
            height: 44,
            fill: "#F97316",
          }),
          createElement(Shape, {
            id: "grid-b",
            name: "Grid B",
            height: 44,
            fill: "#A855F7",
          }),
        ),
      ),
      createElement(
        Slide,
        { id: "shape-text-slide" },
        createElement(Shape, {
          id: "shorthand-shape",
          name: "Shorthand Shape",
          geometry: "roundRect",
          x: 24,
          y: 24,
          width: 180,
          height: 80,
          fill: "#1E293B",
          text: "Rounded",
          textStyle: { fontSize: 16, color: "#FFFFFF" },
        }),
        createElement(
          Shape,
          {
            id: "rich-shape",
            name: "Rich Shape",
            geometry: "diamond",
            x: 240,
            y: 24,
            width: 200,
            height: 90,
            fill: "#7C3AED",
          },
          createElement(P, null, createElement(Span, null, "Rich Text")),
        ),
        createElement(
          Group,
          {
            id: "text-cluster",
            name: "Text Cluster",
          },
          createElement(
            Text,
            {
              id: "inner-label",
              name: "Inner Label",
              x: 40,
              y: 160,
              width: 120,
              height: 32,
            },
            "Group Label",
          ),
        ),
      ),
    ),
  );
}

function countPaintedPixelsInBox(
  pngBytes: Uint8Array,
  region: {
    readonly x: number;
    readonly y: number;
    readonly width: number;
    readonly height: number;
  },
): number {
  const png = PNG.sync.read(Buffer.from(pngBytes));
  const minX = Math.max(0, Math.floor(region.x));
  const minY = Math.max(0, Math.floor(region.y));
  const maxX = Math.min(png.width - 1, Math.ceil(region.x + region.width));
  const maxY = Math.min(png.height - 1, Math.ceil(region.y + region.height));
  let painted = 0;
  for (let y = minY; y <= maxY; y += 1) {
    for (let x = minX; x <= maxX; x += 1) {
      const idx = (y * png.width + x) * 4;
      const r = png.data[idx]!;
      const g = png.data[idx + 1]!;
      const b = png.data[idx + 2]!;
      const a = png.data[idx + 3]!;
      if (a > 0 && (r < 245 || g < 245 || b < 245)) {
        painted += 1;
      }
    }
  }
  return painted;
}

function extractShapeXmlByName(slideXml: string, name: string): string {
  const marker = `name="${name}"`;
  const markerIdx = slideXml.indexOf(marker);
  assert.ok(markerIdx >= 0, `Expected <p:sp> with ${marker}`);
  const spStart = slideXml.lastIndexOf("<p:sp>", markerIdx);
  const spEnd = slideXml.indexOf("</p:sp>", markerIdx);
  assert.ok(
    spStart >= 0 && spEnd > spStart,
    `Expected closing </p:sp> for ${marker}`,
  );
  return slideXml.slice(spStart, spEnd + "</p:sp>".length);
}

describe("pptxdk geometry, layout container, and group codegen acceptance", () => {
  it("generates typechecked, relocated-rebuildable vanilla and react sources preserving preset adjusts, custom bezier control points, arrows, groups, and per-object preview", async () => {
    const directVanilla = createDirectVanillaDeck();
    const directReact = await createDirectReactDeck();

    const originalBuilt = await build(directVanilla);
    assert.deepEqual(
      originalBuilt.report.diagnostics.filter((d) => d.severity === "error"),
      [],
    );
    const importedResult = importPptx(originalBuilt.pptx);
    assert.deepEqual(importedResult.report.diagnostics, []);

    const sources = [
      { label: "direct-vanilla", doc: directVanilla, isImported: false },
      { label: "direct-react", doc: directReact, isImported: false },
      {
        label: "imported-pptx",
        doc: importedResult.document,
        isImported: true,
      },
    ] as const;

    const generatedCases = sources.flatMap(({ label, doc, isImported }) => {
      const vanillaGen = codegen(doc, { target: "vanilla" });
      const reactGen = codegen(doc, { target: "react" });
      return [
        { label: `${label}:vanilla`, result: vanillaGen, isImported },
        { label: `${label}:react`, result: reactGen, isImported },
      ];
    });

    await verifyGeneratedSourcesTypecheck(generatedCases);

    for (const item of generatedCases) {
      const { built, file } = await rebuildGeneratedSourceRelocated(
        item.result,
      );
      assert.deepEqual(
        file.errors,
        [],
        `${item.label} validation errors: ${JSON.stringify(file.errors)}`,
      );
      assert.equal(file.slides.length, 3, `${item.label} slide count`);

      // Independent XML reading of rebuilt PPTX for Slide 1 geometry & control points
      const parts = unzipSync(built.pptx);
      const slide1Xml = strFromU8(parts["ppt/slides/slide1.xml"]!);

      const roundXml = extractShapeXmlByName(slide1Xml, "Preset Round");
      assert.match(roundXml, /<a:prstGeom prst="roundRect">/u);
      assert.match(roundXml, /<a:gd name="adj" fmla="val 20000"\/>/u);

      const arcXml = extractShapeXmlByName(slide1Xml, "Preset BlockArc");
      assert.match(arcXml, /<a:prstGeom prst="blockArc">/u);
      assert.match(arcXml, /<a:gd name="adj1" fmla="val 0"\/>/u);
      assert.match(arcXml, /<a:gd name="adj2" fmla="val 10800000"\/>/u);
      assert.match(arcXml, /<a:gd name="adj3" fmla="val 30000"\/>/u);

      const bezierXml = extractShapeXmlByName(slide1Xml, "Single Bezier");
      assert.match(bezierXml, /flipH="1"/u);
      assert.match(bezierXml, /flipV="1"/u);
      assert.match(bezierXml, /rot="900000"/u);
      assert.match(bezierXml, /<a:custGeom>/u);
      assert.match(
        bezierXml,
        /<a:moveTo><a:pt x="10000" y="80000"\/><\/a:moveTo>/u,
      );
      assert.match(
        bezierXml,
        /<a:cubicBezTo><a:pt x="10000" y="15000"\/><a:pt x="90000" y="15000"\/><a:pt x="90000" y="80000"\/><\/a:cubicBezTo>/u,
      );
      assert.match(
        bezierXml,
        /<a:quadBezTo><a:pt x="50000" y="95000"\/><a:pt x="10000" y="80000"\/><\/a:quadBezTo>/u,
      );

      const compositeXml = extractShapeXmlByName(slide1Xml, "Composite Custom");
      assert.match(compositeXml, /<a:path w="100000" h="80000" stroke="0">/u);
      assert.match(compositeXml, /<a:path w="100000" h="80000" fill="none">/u);

      const arrowXml = extractShapeXmlByName(slide1Xml, "Arrow Connector");
      assert.match(arrowXml, /<a:headEnd type="oval"\/>/u);
      assert.match(arrowXml, /<a:tailEnd type="triangle"\/>/u);

      // Slide 2: Nested Group hierarchy & member order
      const slide2Objects = file.slides[1]!.objects;
      const outer = slide2Objects.find((o) => o.name === "Outer Cluster");
      assert.ok(outer !== undefined, `${item.label} Outer Cluster`);
      assert.equal(outer.kind, "group");
      assert.equal(outer.rotationDegrees, 12);
      assert.deepEqual(
        outer.members?.map((m) => m.name),
        ["Outer Hex", "Inner Cluster"],
      );
      const inner = outer.members?.[1];
      assert.equal(inner?.kind, "group");
      assert.equal(inner?.rotationDegrees, 342);
      assert.deepEqual(
        inner?.members?.map((m) => m.name),
        ["Inner Diamond"],
      );

      // Slide 3: Shape shorthand text, rich-text children, and Group text member
      const slide3Objects = file.slides[2]!.objects;
      assert.equal(
        slide3Objects.find((o) => o.name === "Shorthand Shape")?.text?.text,
        "Rounded",
      );
      assert.equal(
        slide3Objects.find((o) => o.name === "Rich Shape")?.text?.text,
        "Rich Text",
      );
      const textCluster = slide3Objects.find((o) => o.name === "Text Cluster");
      assert.equal(textCluster?.kind, "group");
      assert.equal(textCluster?.members?.[0]?.name, "Inner Label");
      assert.equal(textCluster?.members?.[0]?.text?.text, "Group Label");

      // Independent import + per-object preview check on pages 1 and 2
      const reimported = importPptx(built.pptx);
      assert.deepEqual(reimported.report.diagnostics, []);
      const tempDir = await mkdtemp(join(tmpdir(), "pptxdk-prev-geom-"));
      try {
        await linkConsumerDeps(tempDir);
        await writeCodegenFiles(tempDir, item.result.files);
        const previewed = await previewFromSource(
          join(tempDir, item.result.entry),
          { pages: [1, 2] },
        );
        const page1 = previewed.pages[0]!;
        assert.deepEqual(page1.diagnostics, []);
        assert.ok(page1.png instanceof Uint8Array);

        for (const name of [
          "Preset Round",
          "Preset BlockArc",
          "Single Bezier",
          "Composite Custom",
          "Arrow Connector",
        ]) {
          const node = page1.nodes.find((n) => n.name === name);
          assert.ok(
            node?.screenBox !== undefined,
            `${item.label} ${name} screenBox`,
          );
          const paintedPixels = countPaintedPixelsInBox(
            page1.png,
            node.screenBox,
          );
          assert.ok(
            paintedPixels > 10,
            `${item.label} ${name} should paint visible pixels in preview, got ${paintedPixels}`,
          );
        }

        const page2 = previewed.pages[1]!;
        assert.deepEqual(page2.diagnostics, []);
        assert.ok(page2.png instanceof Uint8Array);
        for (const name of ["Outer Hex", "Inner Diamond", "Flex A", "Grid A"]) {
          const node = page2.nodes.find((n) => n.name === name);
          assert.ok(
            node?.screenBox !== undefined,
            `${item.label} ${name} screenBox`,
          );
          const paintedPixels = countPaintedPixelsInBox(
            page2.png,
            node.screenBox,
          );
          assert.ok(
            paintedPixels > 10,
            `${item.label} ${name} should paint visible pixels in preview, got ${paintedPixels}`,
          );
        }
      } finally {
        await rm(tempDir, { recursive: true, force: true });
      }
    }
  }, 120_000);

  it("detects target shape omission in per-object preview while sibling shapes continue rendering", async () => {
    const completeDoc = deck({ width: 400, height: 200 }).append(
      slide({ id: "s1" }).append(
        shape({
          id: "target-shape",
          name: "Target Shape",
          geometry: preset("roundRect", { adj: 20000 }),
          x: 24,
          y: 40,
          width: 100,
          height: 80,
          fill: "#2563EB",
        }),
        shape({
          id: "sibling-shape",
          name: "Sibling Shape",
          geometry: "ellipse",
          x: 200,
          y: 40,
          width: 100,
          height: 80,
          fill: "#16A34A",
        }),
      ),
    );

    const omittedTargetDoc = deck({ width: 400, height: 200 }).append(
      slide({ id: "s1" }).append(
        shape({
          id: "sibling-shape",
          name: "Sibling Shape",
          geometry: "ellipse",
          x: 200,
          y: 40,
          width: 100,
          height: 80,
          fill: "#16A34A",
        }),
      ),
    );

    const dir = await mkdtemp(join(tmpdir(), "pptxdk-omit-check-"));
    try {
      await linkConsumerDeps(dir);
      const completeGen = codegen(completeDoc, { target: "vanilla" });
      await writeCodegenFiles(join(dir, "complete"), completeGen.files);
      const completePreview = await previewFromSource(
        join(dir, "complete", completeGen.entry),
      );
      const completePage = completePreview.pages[0]!;
      const targetNode = completePage.nodes.find(
        (n) => n.authorId === "target-shape",
      )!;
      const siblingNode = completePage.nodes.find(
        (n) => n.authorId === "sibling-shape",
      )!;
      assert.ok(
        countPaintedPixelsInBox(completePage.png!, targetNode.screenBox!) > 100,
      );
      assert.ok(
        countPaintedPixelsInBox(completePage.png!, siblingNode.screenBox!) >
          100,
      );

      const omittedGen = codegen(omittedTargetDoc, { target: "vanilla" });
      await writeCodegenFiles(join(dir, "omitted"), omittedGen.files);
      const omittedPreview = await previewFromSource(
        join(dir, "omitted", omittedGen.entry),
      );
      const omittedPage = omittedPreview.pages[0]!;
      assert.equal(
        omittedPage.nodes.find((n) => n.authorId === "target-shape"),
        undefined,
      );
      assert.equal(
        countPaintedPixelsInBox(omittedPage.png!, targetNode.screenBox!),
        0,
      );
      assert.ok(
        countPaintedPixelsInBox(omittedPage.png!, siblingNode.screenBox!) > 100,
      );
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  }, 30_000);

  it("preserves Box Flex and Grid layout rules so resizing containers in generated TS and TSX triggers child reflow", async () => {
    const directDoc = createDirectVanillaDeck();

    for (const target of ["vanilla", "react"] as const) {
      const generated = codegen(directDoc, { target });
      const dir = await mkdtemp(join(tmpdir(), `pptxdk-reflow-${target}-`));
      try {
        await linkConsumerDeps(dir);
        const initialDir = join(dir, "initial");
        const editedDir = join(dir, "edited");
        await writeCodegenFiles(initialDir, generated.files);
        const initialEntryPath = join(initialDir, generated.entry);

        const initialBuilt = await buildFromSource(initialEntryPath);
        const initialValidation = await validatePptx(initialBuilt.pptx);
        assert.deepEqual(initialValidation.errors, []);
        const initialObjs = initialValidation.slides[1]!.objects;
        const initFlexA = initialObjs.find((o) => o.name === "Flex A")!;
        const initFlexB = initialObjs.find((o) => o.name === "Flex B")!;
        const initGridA = initialObjs.find((o) => o.name === "Grid A")!;
        const initGridB = initialObjs.find((o) => o.name === "Grid B")!;

        assert.equal(initFlexB.boxEmu?.y, initFlexA.boxEmu?.y);
        assert.ok((initFlexB.boxEmu?.x ?? 0) > (initFlexA.boxEmu?.x ?? 0));

        // Edit container widths in a separate directory so ESM module cache resolves the edited file cleanly
        await writeCodegenFiles(editedDir, generated.files);
        const editedEntryPath = join(editedDir, generated.entry);
        let editedSource = await readFile(editedEntryPath, "utf8");
        assert.ok(editedSource.includes("width: 300"), `${target} flex width`);
        assert.ok(editedSource.includes("width: 280"), `${target} grid width`);
        editedSource = editedSource
          .replace("width: 300", "width: 180")
          .replace("width: 280", "width: 400");
        await writeFile(editedEntryPath, editedSource, "utf8");

        const reflowBuilt = await buildFromSource(editedEntryPath);
        const reflowValidation = await validatePptx(reflowBuilt.pptx);
        assert.deepEqual(reflowValidation.errors, []);
        const reflowObjs = reflowValidation.slides[1]!.objects;
        const nextFlexA = reflowObjs.find((o) => o.name === "Flex A")!;
        const nextFlexB = reflowObjs.find((o) => o.name === "Flex B")!;
        const nextGridA = reflowObjs.find((o) => o.name === "Grid A")!;
        const nextGridB = reflowObjs.find((o) => o.name === "Grid B")!;

        assert.ok(
          (nextFlexB.boxEmu?.y ?? 0) > (nextFlexA.boxEmu?.y ?? 0),
          `${target}: Flex B should wrap onto second row after narrowing Box`,
        );
        assert.equal(
          nextFlexB.boxEmu?.x,
          nextFlexA.boxEmu?.x,
          `${target}: wrapped Flex B should align with Flex A x`,
        );
        assert.ok(
          (nextGridA.boxEmu?.width ?? 0) > (initGridA.boxEmu?.width ?? 0),
          `${target}: Grid A should widen when Grid Box widens`,
        );
        assert.ok(
          (nextGridB.boxEmu?.x ?? 0) > (initGridB.boxEmu?.x ?? 0),
          `${target}: Grid B should shift right when Grid Box widens`,
        );
      } finally {
        await rm(dir, { recursive: true, force: true });
      }
    }
  }, 60_000);

  it("transforms native Group members together when modifying an imported Group in generated TS and TSX", async () => {
    const directBuilt = await build(createDirectVanillaDeck());
    const imported = importPptx(directBuilt.pptx);
    assert.deepEqual(imported.report.diagnostics, []);

    for (const target of ["vanilla", "react"] as const) {
      const generated = codegen(imported.document, { target });
      const dir = await mkdtemp(join(tmpdir(), `pptxdk-grp-edit-${target}-`));
      try {
        await linkConsumerDeps(dir);
        const initialDir = join(dir, "initial");
        const editedDir = join(dir, "edited");
        await writeCodegenFiles(initialDir, generated.files);
        const initialEntryPath = join(initialDir, generated.entry);

        const initialPreview = await previewFromSource(initialEntryPath, {
          pages: [2],
        });
        const initialSlide2 = initialPreview.pages[0]!;
        const initHex = initialSlide2.nodes.find(
          (n) => n.name === "Outer Hex",
        )!;
        const initDiamond = initialSlide2.nodes.find(
          (n) => n.name === "Inner Diamond",
        )!;
        const initHexScreen = initHex.screenBox!;
        const initDiamondScreen = initDiamond.screenBox!;

        await writeCodegenFiles(editedDir, generated.files);
        const editedEntryPath = join(editedDir, generated.entry);
        let editedSource = await readFile(editedEntryPath, "utf8");
        assert.ok(
          editedSource.includes("rotate: 12"),
          `${target}: expected imported Outer Cluster rotate: 12`,
        );
        editedSource = editedSource.replace("rotate: 12", "rotate: 90");
        await writeFile(editedEntryPath, editedSource, "utf8");

        const rebuilt = await buildFromSource(editedEntryPath);
        const validation = await validatePptx(rebuilt.pptx);
        assert.deepEqual(validation.errors, []);
        const outer = validation.slides[1]!.objects.find(
          (o) => o.name === "Outer Cluster",
        );
        assert.equal(outer?.rotationDegrees, 90);
        assert.equal(
          Math.round((outer?.boxEmu?.width ?? 0) / emuPerPx) > 0,
          true,
        );

        const editedPreview = await previewFromSource(editedEntryPath, {
          pages: [2],
        });
        const editedSlide2 = editedPreview.pages[0]!;
        const nextHex = editedSlide2.nodes.find((n) => n.name === "Outer Hex")!;
        const nextDiamond = editedSlide2.nodes.find(
          (n) => n.name === "Inner Diamond",
        )!;

        assert.notDeepEqual(
          nextHex.screenBox,
          initHexScreen,
          `${target}: Outer Hex screenBox should transform with Outer Cluster`,
        );
        assert.notDeepEqual(
          nextDiamond.screenBox,
          initDiamondScreen,
          `${target}: Inner Diamond screenBox should transform with Outer Cluster`,
        );
      } finally {
        await rm(dir, { recursive: true, force: true });
      }
    }
  }, 60_000);
});
