import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import {
  build,
  Box,
  BuildError,
  Cell,
  Chart,
  createMorphScope,
  Deck,
  Group,
  Image,
  MotionScope,
  Row,
  Shape,
  Slide,
  Table,
  Text,
  targets,
  transitions,
  useMotionScope,
} from "@react-pptx/react-pptx";
import { validatePptx } from "@react-pptx/validation";
import { createElement } from "react";
import { describe, it } from "vitest";

describe("跨页对象连续性", () => {
  it("只让可见 Box 的自有背景参与原生 Morph，子项仍保持独立", async () => {
    const scope = createMorphScope();
    const background = scope.handle("card-background");
    const declarations = createElement(
      Deck,
      {},
      createElement(
        Slide,
        { slideId: "card-before" },
        createElement(
          Box,
          {
            testId: "card-before",
            morphKey: background,
            style: { width: 160, height: 100, backgroundColor: "accent1" },
          },
          createElement(Shape, {
            testId: "card-child-before",
            x: 12,
            y: 12,
            width: 40,
            height: 30,
          }),
        ),
      ),
      createElement(
        Slide,
        {
          slideId: "card-after",
          transition: transitions.morph({ fallback: transitions.fade() }),
        },
        createElement(
          Box,
          {
            testId: "card-after",
            morphKey: background,
            style: { width: 180, height: 110, backgroundColor: "accent2" },
          },
          createElement(Shape, {
            testId: "card-child-after",
            x: 16,
            y: 16,
            width: 44,
            height: 32,
          }),
        ),
      ),
    );
    const result = await build(declarations);
    const file = await validatePptx(result.pptx, {
      allowedExtensions: ["p14", "p159"],
    });

    assert.deepEqual(file.errors, []);
    assert.deepEqual(
      result.report.slides.map(({ objects }) =>
        objects.map(({ testId, kind }) => ({ testId, kind })),
      ),
      [
        [
          { testId: "card-before", kind: "shape" },
          { testId: "card-child-before", kind: "shape" },
        ],
        [
          { testId: "card-after", kind: "shape" },
          { testId: "card-child-after", kind: "shape" },
        ],
      ],
    );
    assert.equal(
      result.report.slides[0]?.objects[0]?.name,
      result.report.slides[1]?.objects[0]?.name,
    );
    assert.ok(result.report.slides[0]?.objects[0]?.name.startsWith("!!"));
    assert.deepEqual(
      file.slides.map(({ objects }) => objects.map(({ kind }) => kind)),
      [
        ["shape", "shape"],
        ["shape", "shape"],
      ],
    );
  });

  it("通过同一个显式句柄写出相邻页的原生 Morph 对应", async () => {
    const scope = createMorphScope();
    const hero = scope.handle("hero");
    const declarations = createElement(
      Deck,
      { width: 960, height: 540 },
      createElement(
        Slide,
        { slideId: "before" },
        createElement(Shape, {
          testId: "hero-before",
          morphKey: hero,
          x: 24,
          y: 24,
          width: 160,
          height: 100,
          fill: "#336699",
        }),
      ),
      createElement(
        Slide,
        {
          slideId: "after",
          transition: transitions.morph({
            duration: 1200,
            fallback: transitions.fade(),
          }),
        },
        createElement(Shape, {
          testId: "hero-after",
          morphKey: hero,
          x: 280,
          y: 80,
          width: 180,
          height: 120,
          fill: "#CC6633",
        }),
      ),
    );
    const result = await build(declarations);

    assert.deepEqual(result.report.extensions, ["p14", "p159"]);
    const [before, after] = result.report.slides;
    assert.equal(before?.objects[0]?.name, after?.objects[0]?.name);
    assert.match(before?.objects[0]?.name ?? "", /^!!/u);
    assert.equal(before?.objects[0]?.testId, "hero-before");
    assert.equal(after?.objects[0]?.testId, "hero-after");

    const file = await validatePptx(result.pptx, {
      allowedExtensions: ["p14", "p159"],
    });
    assert.deepEqual(file.errors, []);
    assert.deepEqual(file.extensions, ["p14", "p159"]);
    assert.equal(
      file.slides[0]?.objects[0]?.name,
      file.slides[1]?.objects[0]?.name,
    );
    assert.deepEqual(file.slides[1]?.transition, {
      kind: "morph",
      speed: "slow",
      durationMs: 1200,
      by: "object",
      fallback: { kind: "fade", speed: "slow" },
    });

    for (const target of [
      targets.baselineOnly,
      {
        ...targets.transitional,
        name: "p159-only",
        extensions: ["p159"] as const,
      },
    ]) {
      await assert.rejects(
        build(declarations, { target }),
        (error: unknown) => {
          assert.ok(error instanceof BuildError);
          assert.equal(error.diagnostics[0]?.code, "forbidden-mechanism");
          return true;
        },
      );
    }
  });

  it("透明 Box 和 MotionScope 在页面顶层保留 Morph 对应", async () => {
    const key = createMorphScope().handle("hero");
    function MotionWrappedHero() {
      const scope = useMotionScope();
      return createElement(
        MotionScope,
        { scope },
        createElement(Shape, {
          morphKey: key,
          x: 240,
          y: 20,
          width: 60,
          height: 40,
        }),
      );
    }
    const result = await build(
      createElement(
        Deck,
        {},
        createElement(
          Slide,
          { slideId: "before" },
          createElement(
            Box,
            { style: { width: 100, height: 80 } },
            createElement(Shape, {
              morphKey: key,
              x: 20,
              y: 20,
              width: 60,
              height: 40,
            }),
          ),
        ),
        createElement(
          Slide,
          {
            slideId: "after",
            transition: transitions.morph({ fallback: transitions.fade() }),
          },
          createElement(MotionWrappedHero),
        ),
      ),
    );
    const file = await validatePptx(result.pptx, {
      allowedExtensions: ["p14", "p159"],
    });

    assert.deepEqual(file.errors, []);
    assert.equal(
      file.slides[0]?.objects[0]?.name,
      file.slides[1]?.objects[0]?.name,
    );
  });

  it("隔离重复组件实例中的同名句柄并按显式 scope 匹配", async () => {
    const leftScope = createMorphScope();
    const rightScope = createMorphScope();
    const Marker = (props: {
      readonly scope: ReturnType<typeof createMorphScope>;
      readonly testId: string;
      readonly x: number;
    }) =>
      createElement(Shape, {
        morphKey: props.scope.handle("marker"),
        testId: props.testId,
        x: props.x,
        y: 20,
        width: 80,
        height: 40,
        fill: "#336699",
      });
    const result = await build(
      createElement(
        Deck,
        {},
        createElement(
          Slide,
          { slideId: "before" },
          createElement(Marker, {
            scope: leftScope,
            testId: "left-before",
            x: 20,
          }),
          createElement(Marker, {
            scope: rightScope,
            testId: "right-before",
            x: 160,
          }),
        ),
        createElement(
          Slide,
          {
            slideId: "after",
            transition: transitions.morph({ fallback: transitions.fade() }),
          },
          createElement(Marker, {
            scope: rightScope,
            testId: "right-after",
            x: 300,
          }),
          createElement(Marker, {
            scope: leftScope,
            testId: "left-after",
            x: 440,
          }),
        ),
      ),
    );
    const file = await validatePptx(result.pptx, {
      allowedExtensions: ["p14", "p159"],
    });

    assert.deepEqual(file.errors, []);
    const before = result.report.slides[0]?.objects ?? [];
    const after = result.report.slides[1]?.objects ?? [];
    const leftBefore = before.find(({ testId }) => testId === "left-before");
    const leftAfter = after.find(({ testId }) => testId === "left-after");
    const rightBefore = before.find(({ testId }) => testId === "right-before");
    const rightAfter = after.find(({ testId }) => testId === "right-after");
    assert.equal(leftBefore?.name, leftAfter?.name);
    assert.equal(rightBefore?.name, rightAfter?.name);
    assert.notEqual(leftBefore?.name, rightBefore?.name);
    assert.notEqual(leftAfter?.name, rightAfter?.name);
    assert.notEqual(
      file.slides[0]?.objects[0]?.name,
      file.slides[0]?.objects[1]?.name,
    );
  });

  it("并发构建各自分配连续性名称而不互相污染", async () => {
    const sharedScope = createMorphScope();
    const shared = sharedScope.handle("hero");
    const other = createMorphScope().handle("hero");
    const declarations = (leading?: typeof other) =>
      createElement(
        Deck,
        {},
        createElement(
          Slide,
          { slideId: "before" },
          ...(leading === undefined
            ? []
            : [
                createElement(Shape, {
                  morphKey: leading,
                  x: 0,
                  y: 0,
                  width: 20,
                  height: 20,
                }),
              ]),
          createElement(Shape, {
            morphKey: shared,
            x: 40,
            y: 0,
            width: 20,
            height: 20,
          }),
        ),
        createElement(
          Slide,
          {
            slideId: "after",
            transition: transitions.morph({ fallback: transitions.fade() }),
          },
          createElement(Shape, {
            morphKey: shared,
            x: 40,
            y: 40,
            width: 20,
            height: 20,
          }),
        ),
      );
    const [first, second] = await Promise.all([
      build(declarations()),
      build(declarations(other)),
    ]);

    assert.equal(
      first.report.slides[0]?.objects[0]?.name,
      first.report.slides[1]?.objects[0]?.name,
    );
    assert.equal(
      second.report.slides[0]?.objects[1]?.name,
      second.report.slides[1]?.objects[0]?.name,
    );
    assert.notEqual(
      first.report.slides[0]?.objects[0]?.name,
      second.report.slides[0]?.objects[1]?.name,
    );
    assert.deepEqual((await validatePptx(first.pptx)).errors, []);
    assert.deepEqual((await validatePptx(second.pptx)).errors, []);
  });

  it("保留文本、形状、图片、组合、表格和图表的同类对象匹配", async () => {
    const scope = createMorphScope();
    const handles = {
      shape: scope.handle("shape"),
      text: scope.handle("text"),
      image: scope.handle("image"),
      group: scope.handle("group"),
      table: scope.handle("table"),
      chart: scope.handle("chart"),
    };
    const image = new Uint8Array(
      readFileSync(new URL("./assets/wide-20000x1.png", import.meta.url)),
    );
    const objects = (suffix: string, offset: number) => [
      createElement(Shape, {
        morphKey: handles.shape,
        testId: `shape-${suffix}`,
        x: offset,
        y: 10,
        width: 60,
        height: 40,
      }),
      createElement(
        Text,
        {
          morphKey: handles.text,
          testId: `text-${suffix}`,
          x: offset,
          y: 60,
          width: 100,
          height: 30,
          style: { fontFamily: "Arial" },
        },
        "Editable text",
      ),
      createElement(Image, {
        morphKey: handles.image,
        testId: `image-${suffix}`,
        src: image,
        x: offset,
        y: 100,
        width: 60,
        height: 40,
      }),
      createElement(
        Group,
        {
          morphKey: handles.group,
          testId: `group-${suffix}`,
        },
        createElement(Shape, {
          x: offset + 80,
          y: 100,
          width: 40,
          height: 30,
        }),
      ),
      createElement(
        Table,
        {
          morphKey: handles.table,
          testId: `table-${suffix}`,
          columns: [80],
        },
        createElement(
          Row,
          { style: { height: 30 } },
          createElement(Cell, null, `cell-${suffix}`),
        ),
      ),
      createElement(Chart, {
        morphKey: handles.chart,
        testId: `chart-${suffix}`,
        kind: "column",
        x: offset,
        y: 200,
        width: 160,
        height: 100,
        data: { categories: ["A"], series: [{ name: "S", values: [1] }] },
      }),
    ];
    const result = await build(
      createElement(
        Deck,
        {},
        createElement(Slide, { slideId: "before" }, ...objects("before", 20)),
        createElement(
          Slide,
          {
            slideId: "after",
            transition: transitions.morph({ fallback: transitions.fade() }),
          },
          ...objects("after", 320),
        ),
      ),
    );
    const file = await validatePptx(result.pptx, {
      allowedExtensions: ["p14", "p159"],
    });

    assert.deepEqual(file.errors, []);
    for (const handle of Object.keys(handles)) {
      const before = result.report.slides[0]?.objects.find(
        ({ testId }) => testId === `${handle}-before`,
      );
      const after = result.report.slides[1]?.objects.find(
        ({ testId }) => testId === `${handle}-after`,
      );
      assert.equal(before?.name, after?.name, `${handle} report name`);
      assert.ok(before?.name.startsWith("!!"), `${handle} native Morph name`);
    }
    assert.deepEqual(
      file.slides[1]?.objects.map(({ kind }) => kind),
      ["shape", "shape", "image", "group", "table", "chart"],
    );
    assert.equal(
      file.slides[0]?.objects[4]?.table?.cells[0]?.[0],
      "cell-before",
    );
  });

  it("允许未配对的进入、离开和替换，普通名称不充当连续性句柄", async () => {
    const scope = createMorphScope();
    const departing = scope.handle("departing");
    const entering = scope.handle("entering");
    const replacedBefore = scope.handle("before");
    const replacedAfter = scope.handle("after");
    const result = await build(
      createElement(
        Deck,
        {},
        createElement(
          Slide,
          { slideId: "before" },
          createElement(Shape, {
            morphKey: departing,
            x: 10,
            y: 10,
            width: 60,
            height: 40,
          }),
          createElement(Shape, {
            morphKey: replacedBefore,
            x: 80,
            y: 10,
            width: 60,
            height: 40,
          }),
          createElement(Shape, {
            name: "same ordinary name",
            x: 150,
            y: 10,
            width: 60,
            height: 40,
          }),
        ),
        createElement(
          Slide,
          {
            slideId: "after",
            transition: transitions.morph({ fallback: transitions.fade() }),
          },
          createElement(Shape, {
            morphKey: entering,
            x: 10,
            y: 80,
            width: 60,
            height: 40,
          }),
          createElement(Shape, {
            morphKey: replacedAfter,
            x: 80,
            y: 80,
            width: 60,
            height: 40,
          }),
          createElement(Shape, {
            name: "same ordinary name",
            x: 150,
            y: 80,
            width: 60,
            height: 40,
          }),
        ),
      ),
    );
    const file = await validatePptx(result.pptx, {
      allowedExtensions: ["p14", "p159"],
    });

    assert.deepEqual(file.errors, []);
    assert.notEqual(
      file.slides[0]?.objects[0]?.name,
      file.slides[1]?.objects[0]?.name,
    );
    assert.notEqual(
      file.slides[0]?.objects[1]?.name,
      file.slides[1]?.objects[1]?.name,
    );
    assert.equal(file.slides[0]?.objects[2]?.name, "same ordinary name");
    assert.equal(file.slides[1]?.objects[2]?.name, "same ordinary name");
    assert.ok(!file.slides[0]?.objects[2]?.name.startsWith("!!"));
  });

  it("定位同页重复句柄、透明 Box、Group 内成员和异类配对", async () => {
    const scope = createMorphScope();
    const key = scope.handle("object");
    const image = new Uint8Array(
      readFileSync(new URL("./assets/wide-20000x1.png", import.meta.url)),
    );
    function MotionWrappedShape() {
      const scope = useMotionScope();
      return createElement(
        MotionScope,
        { scope },
        createElement(Shape, {
          morphKey: key,
          x: 0,
          y: 0,
          width: 40,
          height: 30,
        }),
      );
    }
    function MotionWrappedBox() {
      const scope = useMotionScope();
      return createElement(
        MotionScope,
        { scope },
        createElement(
          Box,
          { style: { width: 100, height: 80 } },
          createElement(MotionWrappedShape),
        ),
      );
    }
    const invalid = [
      {
        code: "invalid-morph-key",
        deck: createElement(
          Deck,
          {},
          createElement(
            Slide,
            { slideId: "invalid-handle" },
            createElement(Shape, {
              morphKey: "ordinary string" as never,
              x: 0,
              y: 0,
              width: 40,
              height: 30,
            }),
          ),
        ),
      },
      {
        code: "morph-object-name-conflict",
        deck: createElement(
          Deck,
          {},
          createElement(
            Slide,
            { slideId: "name-conflict" },
            createElement(Shape, {
              morphKey: key,
              name: "ordinary object name",
              x: 0,
              y: 0,
              width: 40,
              height: 30,
            } as never),
          ),
        ),
      },
      {
        code: "duplicate-morph-key",
        deck: createElement(
          Deck,
          {},
          createElement(
            Slide,
            { slideId: "duplicate" },
            createElement(Shape, {
              morphKey: key,
              x: 0,
              y: 0,
              width: 40,
              height: 30,
            }),
            createElement(Shape, {
              morphKey: key,
              x: 50,
              y: 0,
              width: 40,
              height: 30,
            }),
          ),
        ),
      },
      {
        code: "morph-needs-object",
        deck: createElement(
          Deck,
          {},
          createElement(
            Slide,
            { slideId: "transparent" },
            createElement(
              Box,
              { morphKey: key } as never,
              createElement(Shape, {
                x: 0,
                y: 0,
                width: 40,
                height: 30,
              }),
            ),
          ),
        ),
      },
      {
        code: "morph-inside-group",
        deck: createElement(
          Deck,
          {},
          createElement(
            Slide,
            { slideId: "nested" },
            createElement(
              Group,
              null,
              createElement(Shape, {
                morphKey: key,
                x: 0,
                y: 0,
                width: 40,
                height: 30,
              }),
            ),
          ),
        ),
      },
      {
        code: "morph-inside-group",
        deck: createElement(
          Deck,
          {},
          createElement(
            Slide,
            { slideId: "before" },
            createElement(
              Group,
              null,
              createElement(
                Box,
                { style: { width: 100, height: 80 } },
                createElement(Shape, {
                  morphKey: key,
                  x: 0,
                  y: 0,
                  width: 40,
                  height: 30,
                }),
              ),
            ),
          ),
          createElement(
            Slide,
            {
              slideId: "after",
              transition: transitions.morph({ fallback: transitions.fade() }),
            },
            createElement(Shape, {
              morphKey: key,
              x: 240,
              y: 20,
              width: 40,
              height: 30,
            }),
          ),
        ),
      },
      {
        code: "morph-inside-group",
        deck: createElement(
          Deck,
          {},
          createElement(
            Slide,
            { slideId: "nested-boxes" },
            createElement(
              Group,
              null,
              createElement(
                Box,
                { style: { width: 100, height: 80 } },
                createElement(
                  Box,
                  { style: { width: 80, height: 60 } },
                  createElement(Shape, {
                    morphKey: key,
                    x: 0,
                    y: 0,
                    width: 40,
                    height: 30,
                  }),
                ),
              ),
            ),
          ),
        ),
      },
      {
        code: "morph-inside-group",
        deck: createElement(
          Deck,
          {},
          createElement(
            Slide,
            { slideId: "mixed-wrappers" },
            createElement(Group, null, createElement(MotionWrappedBox)),
          ),
        ),
      },
      {
        code: "morph-object-kind-mismatch",
        deck: createElement(
          Deck,
          {},
          createElement(
            Slide,
            { slideId: "before" },
            createElement(Shape, {
              morphKey: key,
              x: 0,
              y: 0,
              width: 40,
              height: 30,
            }),
          ),
          createElement(
            Slide,
            {
              slideId: "after",
              transition: transitions.morph({ fallback: transitions.fade() }),
            },
            createElement(Image, {
              morphKey: key,
              src: image,
              x: 0,
              y: 0,
              width: 40,
              height: 30,
            }),
          ),
        ),
      },
    ];

    for (const { code, deck } of invalid) {
      await assert.rejects(build(deck), (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.ok(
          error.diagnostics.some((diagnostic) => diagnostic.code === code),
          `${code}: ${JSON.stringify(error.diagnostics)}`,
        );
        return true;
      });
    }
  });
});
