import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

import {
  build,
  BuildError,
  Deck,
  createBuilder,
  font,
  P,
  Shape,
  shapeCapability,
  Slide,
  Span,
  Text,
  targets,
} from "@react-pptx/react-pptx";
import { validatePptx } from "@react-pptx/validation";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import {
  createElement,
  useEffect,
  useLayoutEffect,
  type ComponentType,
} from "react";
import { describe, it } from "vitest";

function deck(shape: ReturnType<typeof createElement>) {
  return createElement(
    Deck,
    { width: 960, height: 540 },
    createElement(Slide, { slideId: "cover" }, shape),
  );
}

describe("公开构建入口", () => {
  it("生成固定区域文字并保留中英与 Unicode 内容", async () => {
    const content = "固定区域文字 — React + 简体中文 😀";
    const result = await build(
      deck(
        createElement(Text, {
          testId: "mixed-text",
          x: 32,
          y: 24,
          width: 360,
          height: 96,
          style: { fontFamily: "Aptos" },
          children: content,
        }),
      ),
    );

    assert.equal(result.report.slides[0]?.objects[0]?.kind, "text");
    assert.equal(result.report.slides[0]?.objects[0]?.text, content);
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(file.slides[0]?.objects[0]?.text, {
      text: content,
      paragraphs: [
        {
          text: content,
          runs: [
            {
              text: content,
              fontFamily: "Aptos",
              fontSizePt: 18,
              bold: false,
              italic: false,
              color: "#000000",
            },
          ],
        },
      ],
    });
    assert.ok(result.pptx.byteLength > 0);
  });

  it("把分段样式、固定行距和列表写成原生段落语义", async () => {
    const result = await build(
      deck(
        createElement(
          Text,
          {
            testId: "formatted-text",
            x: 24,
            y: 32,
            width: 420,
            height: 180,
            style: { lineHeight: "28px" },
          },
          createElement(
            P,
            {
              style: {
                align: "center",
                lineHeight: "24pt",
                list: { type: "bullet", character: "•" },
              },
            },
            "默认片段 ",
            createElement(
              Span,
              {
                style: {
                  fontFamily: font("Example Sans", [
                    { weight: 700, data: new Uint8Array([1, 2, 3]) },
                  ]),
                  fontSize: "12pt",
                  fontWeight: "bold",
                  color: "#334455",
                },
              },
              "中英 Text",
            ),
          ),
          createElement(
            P,
            { style: { list: { type: "number", startAt: 3 } } },
            "第二段",
          ),
        ),
      ),
    );
    assert.equal(
      result.report.slides[0]?.objects[0]?.text,
      "默认片段 中英 Text\n第二段",
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(file.slides[0]?.objects[0]?.text?.paragraphs, [
      {
        text: "默认片段 中英 Text",
        alignment: "ctr",
        lineSpacingPoints: 24,
        list: { type: "bullet", character: "•" },
        runs: [
          {
            text: "默认片段 ",
            fontFamily: "Aptos",
            fontSizePt: 18,
            bold: false,
            italic: false,
            color: "#000000",
          },
          {
            text: "中英 Text",
            fontFamily: "Example Sans",
            fontSizePt: 12,
            bold: true,
            italic: false,
            color: "#334455",
          },
        ],
      },
      {
        text: "第二段",
        lineSpacingPoints: 21,
        list: { type: "number", style: "arabicPeriod", startAt: 3 },
        runs: [
          {
            text: "第二段",
            fontFamily: "Aptos",
            fontSizePt: 18,
            bold: false,
            italic: false,
            color: "#000000",
          },
        ],
      },
    ]);
  });

  it("同名字族、字重和样式引用冲突资源时报告来源", async () => {
    const element = deck(
      createElement(
        Text,
        { testId: "font-conflict", x: 12, y: 12, width: 300, height: 80 },
        createElement(
          P,
          {},
          createElement(
            Span,
            {
              style: {
                fontFamily: font("Collision Sans", [
                  { data: new Uint8Array([1, 2]) },
                ]),
              },
            },
            "first",
          ),
          createElement(
            Span,
            {
              style: {
                fontFamily: font("Collision Sans", [
                  { data: new Uint8Array([3, 4]) },
                ]),
              },
            },
            "second",
          ),
        ),
      ),
    );

    await assert.rejects(build(element), (error: unknown) => {
      assert.ok(error instanceof BuildError);
      assert.equal(error.diagnostics[0]?.code, "conflicting-resource");
      assert.match(error.diagnostics[0]?.path ?? "", /text\[font-conflict\]/u);
      return true;
    });
  });

  it("固定区域可只声明字体名，报告度量限制且默认不嵌字体", async () => {
    const result = await build(
      deck(
        createElement(
          Text,
          {
            testId: "missing-metrics",
            x: 0,
            y: 0,
            width: 280,
            height: 60,
            style: { fontFamily: "Local Only" },
          },
          "不要求字体文件的固定文字",
        ),
      ),
    );

    assert.deepEqual(
      result.report.diagnostics.map(({ severity, path, code }) => ({
        severity,
        path,
        code,
      })),
      [
        {
          severity: "warning",
          path: "deck/slide[cover]/text[missing-metrics]",
          code: "font-measurement-unavailable",
        },
      ],
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.equal(
      file.slides[0]?.objects[0]?.text?.paragraphs[0]?.runs[0]?.fontFamily,
      "Local Only",
    );
    const parts = unzipSync(result.pptx);
    assert.deepEqual(
      Object.keys(parts).filter((path) => path.startsWith("ppt/fonts/")),
      [],
    );
    assert.doesNotMatch(
      strFromU8(parts["ppt/presentation.xml"]!),
      /embeddedFont/u,
    );
  });

  it("独立裁剪进程可在未加载文本框实现时生成纯形状和带文字形状", async () => {
    const directory = await mkdtemp(join(tmpdir(), "react-pptx-text-trim-"));
    try {
      const loader = join(directory, "reject-text-box.mjs");
      await writeFile(
        loader,
        [
          "export async function resolve(specifier, context, nextResolve) {",
          "  const resolved = await nextResolve(specifier, context);",
          '  if (resolved.url.endsWith("/capabilities/src/text.ts")) { throw new Error("text-box implementation was loaded"); }',
          "  return resolved;",
          "}",
        ].join("\n"),
      );
      const child = spawnSync(
        process.execPath,
        [
          "--conditions=source",
          "--loader",
          loader,
          "--input-type=module",
          "--eval",
          [
            'import { createElement } from "react";',
            'import { BuildError } from "@react-pptx/core";',
            'import { Text as TextDeclaration } from "@react-pptx/capabilities/text-elements";',
            'import { build, Deck, Shape, Slide } from "@react-pptx/react-pptx/shape";',
            'import { validatePptx } from "@react-pptx/validation";',
            'const result = await build(createElement(Deck, { width: 120, height: 96 }, createElement(Slide, { slideId: "trimmed" }, createElement(Shape, { testId: "plain", x: 0, y: 0, width: 20, height: 20 }), createElement(Shape, { testId: "labeled", x: 24, y: 0, width: 60, height: 20, text: "形状文字 — Shape text" }), createElement(Shape, { testId: "empty-label", x: 90, y: 0, width: 20, height: 20, text: "" }))));',
            "const file = await validatePptx(result.pptx);",
            'if (file.errors.length !== 0 || result.report.slides[0]?.objects[1]?.text !== "形状文字 — Shape text" || file.slides[0]?.objects[1]?.text?.text !== "形状文字 — Shape text" || file.slides[0]?.objects[2]?.text !== undefined) throw new Error(JSON.stringify(file.errors));',
            'function NestedText() { return createElement(TextDeclaration, { testId: "missing-text", x: 0, y: 0, width: 20, height: 20 }, "not available"); }',
            "let missingText;",
            'try { await build(createElement(Deck, {}, createElement(Slide, { slideId: "nested" }, createElement(NestedText)))); } catch (error) { missingText = error; }',
            'if (!(missingText instanceof BuildError) || missingText.diagnostics[0]?.code !== "missing-capability" || !missingText.diagnostics[0]?.path.includes("text[missing-text]")) throw new Error("nested Text did not produce a located missing-capability error");',
            'process.stdout.write(JSON.stringify({ kinds: file.slides[0]?.objects.map(({ name, text }) => [name, text?.text ?? ""]), missing: missingText.diagnostics[0], loaded: false }));',
          ].join("\n"),
        ],
        { cwd: process.cwd(), encoding: "utf8" },
      );
      assert.equal(child.status, 0, child.stderr || child.stdout);
      assert.deepEqual(JSON.parse(child.stdout), {
        kinds: [
          ["plain", ""],
          ["labeled", "形状文字 — Shape text"],
          ["empty-label", ""],
        ],
        missing: {
          severity: "error",
          path: "deck/slide[nested]/text[missing-text]",
          code: "missing-capability",
          message: "没有注册 rp:text 的能力模块",
        },
        loaded: false,
      });
    } finally {
      await rm(directory, { recursive: true, force: true });
    }
  });

  it("把真实 React 形状生成 Transitional PPTX 并保留只读身份观察", async () => {
    const result = await build(
      deck(
        createElement(Shape, {
          testId: "badge",
          geometry: "ellipse",
          x: 32,
          y: 48,
          width: 120,
          height: 72,
          fill: "#1A2B3C",
        }),
      ),
    );

    assert.ok(result.pptx instanceof Uint8Array);
    assert.ok(result.pptx.byteLength > 0);
    assert.equal(
      result.report.target.format,
      "ECMA-376 5th edition Transitional",
    );
    assert.equal(result.report.evidence, "generation");
    assert.deepEqual(result.report.slides[0]?.objects, [
      {
        id: 2,
        kind: "shape",
        name: "badge",
        testId: "badge",
        geometry: "ellipse",
        box: { x: 32, y: 48, width: 120, height: 72 },
      },
    ]);
    assert.ok(Object.isFrozen(result.report));
    assert.ok(Object.isFrozen(result.report.slides));
    assert.ok(Object.isFrozen(result.report.slides[0]?.objects[0]));

    const file = await validatePptx(result.pptx);
    assert.equal(file.profile.standard, result.report.target.format);
    assert.deepEqual(file.errors, []);
    assert.equal(file.presentation.slideCount, 1);
    assert.equal(file.slides[0]?.id, "cover");
    assert.deepEqual(file.slides[0]?.objects[0], {
      id: 2,
      name: "badge",
      geometry: "ellipse",
      boxEmu: { x: 304800, y: 457200, width: 1143000, height: 685800 },
    });
  });

  it("从 PPTX 实际读回时保留页面标识中的换行、回车和制表符", async () => {
    const slideIds = [
      "alpha\nbeta",
      "alpha\rbeta",
      "alpha\tbeta",
      "alpha beta",
    ] as const;
    const result = await build(
      createElement(
        Deck,
        { width: 960, height: 540 },
        createElement(
          Slide,
          { slideId: slideIds[0] },
          createElement(Shape, {
            testId: "object-name",
            name: "object\nname\r\tend",
            x: 0,
            y: 0,
            width: 40,
            height: 30,
          }),
          createElement(Shape, {
            testId: "test\nid",
            x: 50,
            y: 0,
            width: 40,
            height: 30,
          }),
        ),
        ...slideIds
          .slice(1)
          .map((slideId) => createElement(Slide, { slideId })),
      ),
    );

    const file = await validatePptx(result.pptx);

    assert.deepEqual(
      result.report.slides.map(({ id }) => id),
      slideIds,
    );
    assert.deepEqual(file.errors, []);
    assert.deepEqual(
      file.slides.map(({ id }) => id),
      slideIds,
    );
    assert.deepEqual(
      file.slides[0]?.objects.map(({ name }) => name),
      ["object\nname\r\tend", "test\nid"],
    );
  });

  it("拒绝 XML 1.0 无法表示的对象名称和页面标识并定位来源", async () => {
    const invalidObjectName = deck(
      createElement(Shape, {
        testId: "invalid-name",
        name: "hello\u0000world",
        x: 0,
        y: 0,
        width: 40,
        height: 30,
      }),
    );
    const invalidTestId = deck(
      createElement(Shape, {
        testId: "hello\u0000world",
        x: 0,
        y: 0,
        width: 40,
        height: 30,
      }),
    );
    const invalidText = deck(
      createElement(
        Text,
        {
          testId: "invalid-text",
          x: 0,
          y: 0,
          width: 40,
          height: 30,
        },
        "hello\u0000world",
      ),
    );
    const invalidSlideId = createElement(
      Deck,
      { width: 960, height: 540 },
      createElement(Slide, { slideId: "alpha\u0000beta" }),
    );

    for (const [element, path] of [
      [invalidObjectName, /slide\[cover\].*shape\[invalid-name\]/u],
      [invalidTestId, /slide\[cover\].*shape\[/u],
      [invalidText, /slide\[cover\].*text\[invalid-text\]/u],
      [invalidSlideId, /slide\[1\]/u],
    ] as const) {
      await assert.rejects(build(element), (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "invalid-xml-character");
        assert.match(error.diagnostics[0]?.path ?? "", path);
        return true;
      });
    }
  });

  it("拒绝无法在 EMU 闭区间内表达的形状宽高", async () => {
    for (const size of [
      { width: 0.00001, height: 20 },
      { width: 40, height: 0.00001 },
      { width: 1_000_000, height: 20 },
      { width: 40, height: 1_000_000 },
    ]) {
      await assert.rejects(
        build(
          deck(
            createElement(Shape, {
              testId: "too-small",
              x: 0,
              y: 0,
              ...size,
            }),
          ),
        ),
        (error: unknown) => {
          assert.ok(error instanceof BuildError);
          assert.equal(error.diagnostics[0]?.code, "invalid-coordinate");
          assert.match(
            error.diagnostics[0]?.path ?? "",
            /slide\[cover\].*shape\[too-small\]/u,
          );
          return true;
        },
      );
    }
  });

  it("独立读取器接受形状宽高的最终 EMU 闭区间边界", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 5376, height: 5376 },
        createElement(
          Slide,
          { slideId: "extent-boundary" },
          createElement(Shape, {
            testId: "extent-boundary",
            x: 0,
            y: 0,
            // 该像素值按标准换算后舍入到形状尺寸上界。
            width: 225457.60073490813,
            height: 0.00006,
          }),
        ),
      ),
    );

    const file = await validatePptx(result.pptx);

    assert.deepEqual(file.errors, []);
    assert.equal(file.slides[0]?.objects[0]?.boxEmu.width, 2_147_483_647);
    assert.equal(file.slides[0]?.objects[0]?.boxEmu.height, 1);
  });

  it("独立校验器拒绝超过格式上限的形状尺寸", async () => {
    const result = await build(
      deck(
        createElement(Shape, {
          testId: "oversized",
          x: 0,
          y: 0,
          width: 40,
          height: 30,
        }),
      ),
    );
    const parts = unzipSync(result.pptx);
    const slidePart = "ppt/slides/slide1.xml";
    const slide = parts[slidePart];
    assert.ok(slide);
    const slideXml = strFromU8(slide);
    const invalidSlideXml = slideXml.replace(
      '<a:ext cx="381000" cy="285750"',
      '<a:ext cx="2147483648" cy="285750"',
    );
    assert.notEqual(invalidSlideXml, slideXml);
    parts[slidePart] = strToU8(invalidSlideXml);

    const file = await validatePptx(zipSync(parts));

    assert.ok(
      file.errors.some(
        ({ part, code }) =>
          part === slidePart && code === "invalid-shape-transform",
      ),
    );
  });

  it("按最终 EMU 限制页面尺寸并保留边界内的小数像素", async () => {
    const boundary = await build(
      createElement(
        Deck,
        { width: 95.99995, height: 5376.00005 },
        createElement(Slide, { slideId: "boundary" }),
      ),
    );
    const boundaryFile = await validatePptx(boundary.pptx);
    assert.deepEqual(boundaryFile.errors, []);
    assert.equal(boundaryFile.presentation.widthEmu, 914400);
    assert.equal(boundaryFile.presentation.heightEmu, 51206400);

    for (const size of [
      { width: 1, height: 1 },
      { width: 95.9999, height: 100 },
      { width: 100, height: 95.9999 },
      { width: 5376.0001, height: 100 },
      { width: 100, height: 5376.0001 },
      { width: 10000, height: 720 },
    ]) {
      await assert.rejects(
        build(
          createElement(
            Deck,
            size,
            createElement(Slide, { slideId: "invalid-size" }),
          ),
        ),
        (error: unknown) => {
          assert.ok(error instanceof BuildError);
          assert.equal(error.diagnostics[0]?.code, "invalid-dimension");
          assert.equal(error.diagnostics[0]?.path, "deck");
          return true;
        },
      );
    }
  });

  it("独立校验器拒绝超出格式范围的实际页面尺寸", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 96, height: 5376 },
        createElement(Slide, { slideId: "size" }),
      ),
    );
    const parts = unzipSync(result.pptx);
    const presentationPart = "ppt/presentation.xml";
    const presentation = parts[presentationPart];
    assert.ok(presentation);
    const presentationXml = strFromU8(presentation);
    const invalidPresentationXml = presentationXml
      .replace('cx="914400"', 'cx="914399"')
      .replace('cy="51206400"', 'cy="51206401"');
    assert.notEqual(invalidPresentationXml, presentationXml);
    parts[presentationPart] = strToU8(invalidPresentationXml);

    const file = await validatePptx(zipSync(parts));

    assert.ok(
      file.errors.some(
        ({ part, code }) =>
          part === presentationPart && code === "invalid-slide-size",
      ),
    );
  });

  it("独立校验器拒绝低于格式下限的母版版式标识", async () => {
    const result = await build(
      deck(
        createElement(Shape, {
          testId: "badge",
          geometry: "ellipse",
          x: 32,
          y: 48,
          width: 120,
          height: 72,
        }),
      ),
    );
    const parts = unzipSync(result.pptx);
    const masterPart = "ppt/slideMasters/slideMaster1.xml";
    const master = parts[masterPart];
    assert.ok(master);
    const masterXml = strFromU8(master);
    const invalidMasterXml = masterXml.replace(
      /(<p:sldLayoutId id=")\d+(")/u,
      (_match, prefix: string, suffix: string) =>
        `${prefix}2147483647${suffix}`,
    );
    assert.notEqual(invalidMasterXml, masterXml);
    parts[masterPart] = strToU8(invalidMasterXml);

    const file = await validatePptx(zipSync(parts));

    assert.ok(
      file.errors.some(
        ({ part, code }) =>
          part === masterPart && code === "invalid-slide-layout-id",
      ),
    );
  });

  it("等待真实组件的布局与被动 Effect 提交并完成 root 卸载清理", async () => {
    const events: string[] = [];
    function EffectShape() {
      useLayoutEffect(() => {
        events.push("layout-commit");
        return () => {
          events.push("layout-cleanup");
        };
      }, []);
      useEffect(() => {
        events.push("passive-commit");
        return () => {
          events.push("passive-cleanup");
        };
      }, []);
      return createElement(Shape, {
        testId: "from-component",
        x: 8,
        y: 9,
        width: 40,
        height: 30,
      });
    }

    const result = await build(deck(createElement(EffectShape)));

    assert.deepEqual(
      result.report.slides[0]?.objects.map(({ testId }) => testId),
      ["from-component"],
    );
    assert.deepEqual(events.sort(), [
      "layout-cleanup",
      "layout-commit",
      "passive-cleanup",
      "passive-commit",
    ]);
  });

  it("在声明路径报告未知形状参数与裁剪后缺失的能力", async () => {
    const UncheckedShape = Shape as ComponentType<Record<string, unknown>>;
    const unknownParameter = deck(
      createElement(UncheckedShape, {
        testId: "unknown-prop",
        x: 0,
        y: 0,
        width: 20,
        height: 20,
        mystery: true,
      }),
    );

    await assert.rejects(build(unknownParameter), (error: unknown) => {
      assert.ok(error instanceof BuildError);
      assert.equal(error.diagnostics[0]?.code, "unknown-parameter");
      assert.match(
        error.diagnostics[0]?.path ?? "",
        /slide\[cover\].*shape\[unknown-prop\]/u,
      );
      return true;
    });

    const clippedBuild = createBuilder([]);
    await assert.rejects(
      clippedBuild(
        deck(
          createElement(Shape, {
            testId: "trimmed-shape",
            x: 0,
            y: 0,
            width: 20,
            height: 20,
          }),
        ),
      ),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "missing-capability");
        assert.match(
          error.diagnostics[0]?.path ?? "",
          /slide\[cover\].*shape\[trimmed-shape\]/u,
        );
        return true;
      },
    );
  });

  it("显式形状模块沿同一链路遵守基础 Transitional 目标", async () => {
    const result = await createBuilder([shapeCapability])(
      deck(
        createElement(Shape, {
          testId: "triangle",
          geometry: "triangle",
          x: 0,
          y: 0,
          width: 80,
          height: 60,
          fill: "#FFAA00",
        }),
      ),
      { target: targets.baselineOnly },
    );

    assert.deepEqual(result.report.target.extensions, []);
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.equal(file.slides[0]?.objects[0]?.geometry, "triangle");
  });

  it("越界保留原始形状并在生成证据中返回警告", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 100, height: 100 },
        createElement(
          Slide,
          { slideId: "cover" },
          createElement(Shape, {
            testId: "overflowing",
            x: 90,
            y: 70,
            width: 30,
            height: 20,
          }),
        ),
      ),
    );

    assert.deepEqual(result.report.diagnostics, [
      {
        severity: "warning",
        path: "deck/slide[cover]/shape[overflowing]",
        code: "overflow",
        message: "对象超出页面画布，构建保留原始声明",
      },
    ]);
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(file.slides[0]?.objects[0]?.boxEmu, {
      x: 857250,
      y: 666750,
      width: 285750,
      height: 190500,
    });
  });

  it("明确拒绝声明外形状和不能写入的尺寸", async () => {
    await assert.rejects(
      build(
        deck(
          createElement(Shape, {
            testId: "unknown-geometry",
            geometry: "madeUpShape",
            x: 0,
            y: 0,
            width: 20,
            height: 20,
          }),
        ),
      ),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "unsupported-geometry");
        return true;
      },
    );

    await assert.rejects(
      build(
        deck(
          createElement(Shape, {
            testId: "zero-size",
            x: 0,
            y: 0,
            width: 0,
            height: 20,
          }),
        ),
      ),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "invalid-coordinate");
        return true;
      },
    );
  });
});
