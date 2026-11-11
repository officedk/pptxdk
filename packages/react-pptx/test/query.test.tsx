import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import {
  Deck,
  Cell,
  Box,
  Chart,
  Group,
  Image,
  font,
  MotionScope,
  Shape,
  Slide,
  Span,
  Text,
  Row,
  Table,
  build,
  spin,
  useMotionScope,
  type MotionObservation,
} from "@react-pptx/react-pptx";
import { validatePptx } from "@react-pptx/validation";
import { strFromU8, unzipSync } from "fflate";
import { createElement } from "react";
import { describe, it } from "vitest";

const png = new Uint8Array(
  Buffer.from(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/n1sAAAAASUVORK5CYII=",
    "base64",
  ),
);

const jost = font("Jost*", [
  {
    weight: 400,
    data: new Uint8Array(
      readFileSync(
        new URL("./assets/fonts/Jost-400-Book.ttf", import.meta.url),
      ),
    ),
  },
  {
    weight: 700,
    data: new Uint8Array(
      readFileSync(
        new URL("./assets/fonts/Jost-700-Bold.ttf", import.meta.url),
      ),
    ),
  },
]);

function FrozenMotionDeck() {
  const scope = useMotionScope();
  return createElement(
    Deck,
    { width: 160, height: 100 },
    createElement(
      Slide,
      { id: "frozen-motion" },
      createElement(
        MotionScope,
        { scope },
        createElement(Shape, {
          id: "moving-shape",
          x: 0,
          y: 0,
          width: 20,
          height: 20,
          emphasis: spin({ duration: 100, easing: { accelerate: 0.2 } }),
        }),
      ),
    ),
  );
}

function ContentQueryDeck() {
  return createElement(
    Deck,
    { width: 320, height: 180 },
    createElement(
      Slide,
      { id: "content-query" },
      createElement(
        Box,
        {
          id: "sales-card",
          style: { width: 200, height: 100, backgroundColor: "#FFFFFF" },
        },
        createElement(
          Text,
          { id: "sales-title", x: 12, y: 16, width: 160, height: 32 },
          "Quarterly sales",
        ),
      ),
    ),
  );
}

function PartQueryDeck() {
  return createElement(
    Deck,
    { width: 320, height: 180 },
    createElement(
      Slide,
      { id: "part-query" },
      createElement(
        Box,
        {
          id: "sales-first",
          style: { width: 140, height: 100 },
        },
        createElement(
          Text,
          { id: "title", x: 0, y: 0, width: 100, height: 20 },
          "First sales",
        ),
        createElement(
          Text,
          {
            id: "featured-title",
            x: 0,
            y: 24,
            width: 100,
            height: 20,
          },
          "Nested feature",
        ),
        createElement(
          Text,
          { id: "first-match", x: 0, y: 48, width: 100, height: 20 },
          "First match",
        ),
      ),
      createElement(
        Box,
        {
          id: "sales-second",
          style: { width: 140, height: 100, left: 150, position: "absolute" },
        },
        createElement(
          Text,
          { id: "title", x: 0, y: 0, width: 100, height: 20 },
          "Second sales",
        ),
        createElement(
          Text,
          { id: "second-match", x: 0, y: 24, width: 100, height: 20 },
          "Second match",
        ),
      ),
    ),
  );
}

function ResourceQueryDeck() {
  return createElement(
    Deck,
    { width: 320, height: 180 },
    createElement(
      Slide,
      { id: "resource-query" },
      createElement(Image, {
        id: "cover-image",
        src: png,
        x: 0,
        y: 0,
        width: 120,
        height: 80,
      }),
      createElement(Chart, {
        id: "sales-chart",
        kind: "column",
        x: 130,
        y: 0,
        width: 160,
        height: 120,
        data: {
          categories: ["Q1"],
          series: [{ name: "Bookings", values: [17] }],
        },
      }),
    ),
  );
}

function FontResourceQueryDeck() {
  return createElement(
    Deck,
    { width: 320, height: 180 },
    createElement(
      Slide,
      { id: "font-resource-query" },
      createElement(
        Text,
        {
          id: "fonted-text",
          x: 8,
          y: 8,
          width: 240,
          height: 36,
          style: { fontFamily: jost },
        },
        "Regular ",
        createElement(Span, { style: { fontWeight: "bold" } }, "Bold"),
      ),
    ),
  );
}

function MissingFontResourceQueryDeck() {
  return createElement(
    Deck,
    { width: 320, height: 180 },
    createElement(
      Slide,
      { id: "missing-font-resource-query" },
      createElement(
        Text,
        {
          id: "font-name-only-text",
          x: 8,
          y: 8,
          width: 240,
          height: 36,
          style: { fontFamily: "Aptos" },
        },
        "No embedded face",
      ),
    ),
  );
}

function TablePartQueryDeck() {
  return createElement(
    Deck,
    { width: 320, height: 180 },
    createElement(
      Slide,
      { id: "table-part-query" },
      createElement(
        Table,
        {
          id: "forecast-table",
          x: 16,
          y: 16,
          columns: [80, 80],
          style: { fontFamily: "Aptos", fontSize: 12 },
        },
        createElement(
          Row,
          { id: "first-row", style: { height: 24 } },
          createElement(Cell, { id: "metric" }, "Bookings"),
          createElement(Cell, { id: "value" }, "17"),
        ),
      ),
    ),
  );
}

function GroupQueryDeck() {
  return createElement(
    Deck,
    { width: 320, height: 180 },
    createElement(
      Slide,
      { id: "group-query" },
      createElement(
        Box,
        { id: "group-host", style: { width: 160, height: 100 } },
        createElement(
          Group,
          { id: "native-group" },
          createElement(
            Box,
            { id: "group-box", style: { width: 80, height: 50 } },
            createElement(Shape, {
              id: "group-member",
              x: 0,
              y: 0,
              width: 20,
              height: 20,
              fill: "#336699",
            }),
          ),
        ),
      ),
    ),
  );
}

describe("read-only build query", () => {
  it("freezes already shallowly frozen motion scopes and their members", async () => {
    const result = await build(createElement(FrozenMotionDeck));
    const motion = result.report.slides[0]?.capabilities?.[
      "motion"
    ] as MotionObservation;
    const scope = motion.scopes?.[0];
    assert.ok(scope);
    assert.equal(scope.mode, "parallel");
    assert.throws(() => {
      (scope as { mode: "parallel" | "sequence" }).mode = "sequence";
    }, TypeError);
    assert.throws(() => {
      (
        scope.members as { kind: "scope" | "object"; id: string | number }[]
      )[0] = {
        kind: "object",
        id: 999,
      };
    }, TypeError);
    assert.equal(scope.mode, "parallel");
    assert.deepEqual(scope.members, [{ kind: "object", id: 2 }]);
  });

  it("locates content in the current build by author id and observes only its objects and text body", async () => {
    const result = await build(createElement(ContentQueryDeck));
    const slide = result.query.getSlide("content-query");
    const card = slide.get("sales-card");
    const title = card.get("sales-title");

    assert.equal(result.query.evidence, "generation");
    assert.equal(slide.id, "content-query");
    assert.deepEqual(
      card.observe().objects.map(({ kind }) => kind),
      ["shape", "text"],
    );
    assert.equal(title.observe().text, "Quarterly sales");
    assert.equal(title.observe().objects[0]?.authorId, "sales-title");
  });

  it("associates image and chart resources with located objects and verifies file bytes from the same build", async () => {
    const result = await build(createElement(ResourceQueryDeck));
    const slide = result.query.getSlide("resource-query");
    const image = slide.get("cover-image").observe();
    const chart = slide.get("sales-chart").observe();
    const imageResource = image.resources?.[0];

    assert.equal(imageResource?.contentType, "image/png");
    assert.equal(imageResource?.objectId, image.objects[0]?.id);
    assert.ok(imageResource);
    const packageParts = unzipSync(result.pptx);
    assert.equal(
      packageParts[imageResource.part]?.byteLength,
      imageResource.byteLength,
    );
    assert.deepEqual(packageParts[imageResource.part], png);

    assert.equal(chart.resources?.length, 2);
    assert.ok(
      chart.resources?.some(
        ({ part, contentType }) =>
          part.startsWith("ppt/charts/") &&
          contentType ===
            "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
      ),
    );
    assert.ok(
      chart.resources?.some(
        ({ part, contentType }) =>
          part.startsWith("ppt/embeddings/") &&
          contentType ===
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      ),
    );
    assert.ok(
      chart.resources?.every(
        ({ part, byteLength }) => packageParts[part]?.byteLength === byteLength,
      ),
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.equal(file.slides[0]?.objects[0]?.image?.part, imageResource.part);
  });

  it("queries only actually written embedded fonts and their relationships in the same PPTX", async () => {
    const defaultResult = await build(createElement(FontResourceQueryDeck));
    const defaultParts = unzipSync(defaultResult.pptx);
    assert.equal(
      defaultResult.query
        .getSlide("font-resource-query")
        .get("fonted-text")
        .observe().resources?.length ?? 0,
      0,
    );
    assert.deepEqual(
      Object.keys(defaultParts).filter((part) => part.startsWith("ppt/fonts/")),
      [],
    );

    const missingResult = await build(
      createElement(MissingFontResourceQueryDeck),
    );
    assert.equal(
      missingResult.query
        .getSlide("missing-font-resource-query")
        .get("font-name-only-text")
        .observe().resources?.length ?? 0,
      0,
    );

    for (const embedFonts of ["full", "subset"] as const) {
      const result = await build(createElement(FontResourceQueryDeck), {
        embedFonts,
      });
      const slide = result.query.getSlide("font-resource-query");
      const text = slide.get("fonted-text").observe();
      const resources = text.resources ?? [];
      const parts = unzipSync(result.pptx);
      const presentation = strFromU8(parts["ppt/presentation.xml"]!);
      const relationships = strFromU8(
        parts["ppt/_rels/presentation.xml.rels"]!,
      );
      const embeddedFontRecords = [
        ...presentation.matchAll(/<p:embeddedFont>(.*?)<\/p:embeddedFont>/gu),
      ];
      const fontRelationships = [
        ...relationships.matchAll(
          /<Relationship Id="([^"]+)" Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/font" Target="([^"]+)"\/>/gu,
        ),
      ];
      const writtenFontParts = embeddedFontRecords.flatMap(([, record]) => {
        const family = record?.match(/<p:font typeface="([^"]+)"/u)?.[1];
        assert.equal(family, "Jost*");
        return [
          ...(record?.matchAll(
            /<p:(regular|bold|italic|boldItalic) r:id="(rId\d+)"\/>/gu,
          ) ?? []),
        ].map(([, style, id]) => {
          const relationship = fontRelationships.find(
            ([, relationId]) => relationId === id,
          );
          assert.ok(relationship?.[2]);
          return { style, part: `ppt/${relationship[2]}` };
        });
      });

      assert.deepEqual(
        writtenFontParts.map(({ style }) => style),
        ["regular", "bold"],
      );
      assert.equal(resources.length, 2);
      assert.deepEqual(
        resources.map(({ part }) => part).sort(),
        writtenFontParts.map(({ part }) => part).sort(),
      );
      assert.ok(
        resources.every(
          ({ objectId, part, byteLength }) =>
            objectId === text.objects[0]?.id &&
            parts[part]?.byteLength === byteLength,
        ),
      );
      if (embedFonts === "subset") {
        assert.ok(
          resources.some(({ part }) => part === "ppt/fonts/font1.fntdata"),
        );
      }
      assert.deepEqual((await validatePptx(result.pptx)).errors, []);
    }
  });

  it("keeps anonymous layers transparent and stops at each named author scope", async () => {
    const result = await build(createElement(PartQueryDeck));
    const slide = result.query.getSlide("part-query");
    const first = slide.get("sales-first");
    const second = slide.get("sales-second");

    assert.equal(first.get("title").observe().text, "First sales");
    assert.equal(second.get("title").observe().text, "Second sales");
    assert.equal(
      first.observe().text,
      "First sales\nNested feature\nFirst match",
    );
    assert.deepEqual(slide.getAll("title"), []);
    assert.throws(() => slide.get("title"), { code: "missing-match" });
    assert.throws(() => first.get("missing"), { code: "missing-match" });
  });

  it("locates cell text body and final geometry from table owner without forging native objects", async () => {
    const result = await build(createElement(TablePartQueryDeck));
    const table = result.query
      .getSlide("table-part-query")
      .get("forecast-table");
    const row = table.get("first-row");
    const metric = row.get("metric");
    const value = row.get("value");

    assert.equal(row.observe().text, "Bookings\t17");
    assert.equal(row.children().length, 2);
    assert.equal(metric.observe().text, "Bookings");
    assert.deepEqual(metric.observe().box, {
      x: 16,
      y: 16,
      width: 80,
      height: 24,
    });
    assert.equal(value.observe().text, "17");
    assert.deepEqual(row.observe().objects, []);
    assert.equal(table.observe().objects[0]?.kind, "table");
  });

  it("traverses along Box layout and Group member hierarchies and deduplicates native objects", async () => {
    const result = await build(createElement(GroupQueryDeck));
    const host = result.query.getSlide("group-query").get("group-host");
    const group = host.children()[0];
    const groupBox = group?.children()[0];

    assert.ok(group);
    assert.ok(groupBox);
    assert.equal(group.observe().objects[0]?.kind, "group");
    assert.equal(
      groupBox.get("group-member").observe().objects[0]?.kind,
      "shape",
    );
    assert.equal(groupBox.observe().box?.width, 80);
    assert.deepEqual(
      group.observe().objects.map(({ kind }) => kind),
      ["group", "shape"],
    );
    assert.equal(group.descendants().length, 2);
  });
});
