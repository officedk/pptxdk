import assert from "node:assert/strict";

import {
  build,
  BuildError,
  Deck,
  createBuilder,
  Shape,
  shapeCapability,
  Slide,
  targets,
} from "@react-pptx/react-pptx";
import { validatePptx } from "@react-pptx/validation";
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
        { width: 100, height: 80 },
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
