import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import {
  Box,
  build,
  BuildError,
  Deck,
  font,
  Group,
  Image,
  Shape,
  Slide,
  Text,
} from "@react-pptx/react-pptx";
import { validatePptx } from "@react-pptx/validation";
import { createElement } from "react";
import { describe, it } from "vitest";

const cjkFont = font("Noto Sans CJK SC", [
  {
    weight: 400,
    data: new Uint8Array(
      readFileSync(
        new URL("./assets/fonts/NotoSansSC-Regular.woff2", import.meta.url),
      ),
    ),
  },
]);
const wideImage = new Uint8Array(
  readFileSync(new URL("./assets/wide-20000x1.png", import.meta.url)),
);

describe("公开布局入口", () => {
  it("显式 display:block 保留块布局，块容器仍拒绝弹性专属属性", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 320, height: 240 },
        createElement(
          Slide,
          { slideId: "explicit-block-layout" },
          createElement(
            Box,
            {
              testId: "block",
              style: { display: "block", width: 100, height: 100 },
            },
            createElement(Shape, { testId: "member", height: 20 }),
          ),
        ),
      ),
    );

    assert.deepEqual(result.report.slides[0]?.objects[0]?.box, {
      x: 0,
      y: 0,
      width: 100,
      height: 20,
    });
    assert.deepEqual(result.report.slides[0]?.layouts?.[0]?.box, {
      x: 0,
      y: 0,
      width: 100,
      height: 100,
    });
    await assert.rejects(
      build(
        createElement(
          Deck,
          { width: 320, height: 240 },
          createElement(
            Slide,
            { slideId: "invalid-block-flex-style" },
            createElement(Box, {
              testId: "block",
              style: { display: "block", justifyContent: "center" },
            }),
          ),
        ),
      ),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some(
          ({ code, path }) =>
            code === "invalid-layout-context" &&
            path === "deck/slide[invalid-block-flex-style]/box[block]",
        ),
    );
  });

  it("块容器纵向排列并横向拉伸，画布外的显式几何保持原位", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 320, height: 240 },
        createElement(
          Slide,
          { slideId: "block-layout" },
          createElement(
            Box,
            { testId: "column", style: { width: 200, height: 10 } },
            createElement(Shape, {
              testId: "first",
              height: 20,
            }),
            createElement(Shape, {
              testId: "second",
              height: 10,
            }),
          ),
          createElement(Shape, {
            testId: "explicit",
            x: 24,
            y: 32,
            width: 50,
            height: 40,
          }),
        ),
      ),
    );

    assert.deepEqual(
      result.report.slides[0]?.objects.map(({ testId, box }) => ({
        testId,
        box,
      })),
      [
        {
          testId: "first",
          box: { x: 0, y: 0, width: 200, height: 20 },
        },
        {
          testId: "second",
          box: { x: 0, y: 20, width: 200, height: 10 },
        },
        {
          testId: "explicit",
          box: { x: 24, y: 32, width: 50, height: 40 },
        },
      ],
    );
    assert.equal(result.report.slides[0]?.layouts?.[0]?.testId, "column");
    assert.deepEqual(result.report.slides[0]?.layouts?.[0]?.box, {
      x: 0,
      y: 0,
      width: 200,
      height: 10,
    });

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(
      file.slides[0]?.objects.map(({ boxEmu }) => ({
        x: boxEmu.x / 9525,
        y: boxEmu.y / 9525,
        width: boxEmu.width / 9525,
        height: boxEmu.height / 9525,
      })),
      [
        { x: 0, y: 0, width: 200, height: 20 },
        { x: 0, y: 20, width: 200, height: 10 },
        { x: 24, y: 32, width: 50, height: 40 },
      ],
    );
  });

  it("auto 尺寸会覆盖原生对象的显式尺寸并按内容排版", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 320, height: 240 },
        createElement(
          Slide,
          { slideId: "auto-layout-size" },
          createElement(
            Box,
            { style: { width: 100, height: 100 } },
            createElement(Shape, {
              testId: "auto-text-shape",
              width: 20,
              height: 10,
              text: "中文中文中文",
              textStyle: { fontFamily: cjkFont, fontSize: 20 },
              style: { width: "auto", height: "auto" },
            }),
          ),
        ),
      ),
    );

    assert.deepEqual(result.report.slides[0]?.objects[0]?.box, {
      x: 0,
      y: 0,
      width: 100,
      height: 57.92,
    });
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(file.slides[0]?.objects[0]?.boxEmu, {
      x: 0,
      y: 0,
      width: 100 * 9525,
      height: 57.92 * 9525,
    });
  });

  it("换行弹性容器默认拉伸行，显式 flex-start 仍从行首排列", async () => {
    const buildWrapped = (alignContent?: "flex-start") =>
      build(
        createElement(
          Deck,
          { width: 320, height: 240 },
          createElement(
            Slide,
            {
              slideId:
                alignContent === undefined ? "wrap-stretch" : "wrap-start",
            },
            createElement(
              Box,
              {
                style: {
                  display: "flex",
                  flexWrap: "wrap",
                  width: 100,
                  height: 100,
                  ...(alignContent === undefined ? {} : { alignContent }),
                },
              },
              createElement(Shape, { testId: "first", width: 60, height: 20 }),
              createElement(Shape, { testId: "second", width: 60, height: 20 }),
            ),
          ),
        ),
      );

    const stretched = await buildWrapped();
    assert.deepEqual(
      stretched.report.slides[0]?.objects.map(({ box }) => box?.y),
      [0, 50],
    );
    const stretchedFile = await validatePptx(stretched.pptx);
    assert.deepEqual(stretchedFile.errors, []);
    assert.deepEqual(
      stretchedFile.slides[0]?.objects.map(({ boxEmu }) => boxEmu.y / 9525),
      [0, 50],
    );

    const rowStart = await buildWrapped("flex-start");
    assert.deepEqual(
      rowStart.report.slides[0]?.objects.map(({ box }) => box?.y),
      [0, 20],
    );
    const rowStartFile = await validatePptx(rowStart.pptx);
    assert.deepEqual(rowStartFile.errors, []);
    assert.deepEqual(
      rowStartFile.slides[0]?.objects.map(({ boxEmu }) => boxEmu.y / 9525),
      [0, 20],
    );
  });

  it("弹性子项按内容最小宽度收缩，minWidth: 0 解除该下限", async () => {
    const makeText = (testId: string, minWidth?: number) =>
      createElement(Text, {
        testId,
        width: 100,
        height: "auto",
        style: {
          fontFamily: cjkFont,
          fontSize: 20,
          ...(minWidth === undefined ? {} : { minWidth }),
        },
        children: "中文中文中文",
      });
    const result = await build(
      createElement(
        Deck,
        { width: 320, height: 500 },
        createElement(
          Slide,
          { slideId: "flex-min-content" },
          createElement(
            Box,
            {
              testId: "constrained",
              style: { display: "flex", width: 30, height: 200 },
            },
            makeText("constrained-a"),
            makeText("constrained-b"),
          ),
          createElement(
            Box,
            {
              testId: "released",
              style: { display: "flex", width: 30, height: 200 },
            },
            makeText("released-a", 0),
            makeText("released-b", 0),
          ),
        ),
      ),
    );

    assert.deepEqual(
      result.report.slides[0]?.objects.map(({ testId, box }) => ({
        testId,
        box,
      })),
      [
        {
          testId: "constrained-a",
          box: { x: 0, y: 0, width: 20, height: 200 },
        },
        {
          testId: "constrained-b",
          box: { x: 20, y: 0, width: 20, height: 200 },
        },
        {
          testId: "released-a",
          box: { x: 0, y: 200, width: 15, height: 200 },
        },
        {
          testId: "released-b",
          box: { x: 15, y: 200, width: 15, height: 200 },
        },
      ],
    );
    assert.deepEqual(
      result.report.diagnostics.map(({ code, path }) => ({ code, path })),
      [
        {
          code: "overset",
          path: "deck/slide[flex-min-content]/box[released]/text[released-a]",
        },
        {
          code: "overset",
          path: "deck/slide[flex-min-content]/box[released]/text[released-b]",
        },
      ],
    );

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(
      file.slides[0]?.objects.map(({ boxEmu }) => ({
        x: boxEmu.x / 9525,
        y: boxEmu.y / 9525,
        width: boxEmu.width / 9525,
        height: boxEmu.height / 9525,
      })),
      [
        { x: 0, y: 0, width: 20, height: 200 },
        { x: 20, y: 0, width: 20, height: 200 },
        { x: 0, y: 200, width: 15, height: 200 },
        { x: 15, y: 200, width: 15, height: 200 },
      ],
    );
  });

  it("嵌套自动最小宽度采用后代尺寸与 minWidth 贡献，显式零仍可收缩", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 320, height: 240 },
        createElement(
          Slide,
          { slideId: "nested-minimums" },
          createElement(
            Box,
            {
              testId: "fixed-parent",
              style: { display: "flex", width: 100, height: 60 },
            },
            createElement(
              Box,
              { testId: "fixed-content" },
              createElement(Shape, {
                testId: "fixed-shape",
                width: 200,
                height: 20,
              }),
            ),
          ),
          createElement(
            Box,
            {
              testId: "min-parent",
              style: { display: "flex", width: 100, height: 60 },
            },
            createElement(
              Box,
              { testId: "min-content" },
              createElement(Shape, {
                testId: "min-shape",
                width: 20,
                height: 20,
                style: { minWidth: 150 },
              }),
            ),
          ),
          createElement(
            Box,
            {
              testId: "released-parent",
              style: { display: "flex", width: 100, height: 60 },
            },
            createElement(
              Box,
              { testId: "released-content", style: { minWidth: 0 } },
              createElement(Shape, {
                testId: "released-shape",
                width: 200,
                height: 20,
              }),
            ),
          ),
          createElement(
            Box,
            {
              testId: "margin-parent",
              style: { display: "flex", width: 100, height: 60 },
            },
            createElement(
              Box,
              { testId: "margin-content" },
              createElement(Shape, {
                testId: "margin-shape",
                width: 100,
                height: 20,
                style: { marginLeft: 10, marginRight: 15 },
              }),
            ),
          ),
        ),
      ),
    );

    const layouts = result.report.slides[0]?.layouts ?? [];
    const layoutWidth = (testId: string) =>
      layouts.find((layout) => layout.testId === testId)?.box.width;
    assert.equal(layoutWidth("fixed-content"), 200);
    assert.equal(layoutWidth("min-content"), 150);
    assert.equal(layoutWidth("released-content"), 100);
    assert.equal(layoutWidth("margin-content"), 125);
    assert.deepEqual(
      result.report.slides[0]?.objects.map(({ testId, box }) => ({
        testId,
        x: box?.x,
        width: box?.width,
      })),
      [
        { testId: "fixed-shape", x: 0, width: 200 },
        { testId: "min-shape", x: 0, width: 150 },
        { testId: "released-shape", x: 0, width: 200 },
        { testId: "margin-shape", x: 10, width: 100 },
      ],
    );

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(
      file.slides[0]?.objects.map(({ boxEmu }) => ({
        x: boxEmu.x / 9525,
        width: boxEmu.width / 9525,
      })),
      [
        { x: 0, width: 200 },
        { x: 0, width: 150 },
        { x: 0, width: 200 },
        { x: 10, width: 100 },
      ],
    );
  });

  it("嵌套布局容器保留实例身份，复用同一元素时仍各自求解", async () => {
    const card = createElement(
      Box,
      { testId: "card", style: { flex: 1, height: 100 } },
      createElement(Shape, { testId: "card-shape", height: 20 }),
    );
    const result = await build(
      createElement(
        Deck,
        { width: 320, height: 240 },
        createElement(
          Slide,
          { slideId: "reused-layout" },
          createElement(
            Box,
            {
              testId: "row",
              style: { display: "flex", width: 300, height: 120 },
            },
            card,
            card,
          ),
        ),
      ),
    );

    assert.deepEqual(
      result.report.slides[0]?.objects.map(({ testId, box }) => ({
        testId,
        box,
      })),
      [
        {
          testId: "card-shape",
          box: { x: 0, y: 0, width: 150, height: 20 },
        },
        {
          testId: "card-shape",
          box: { x: 150, y: 0, width: 150, height: 20 },
        },
      ],
    );
    const layouts = result.report.slides[0]?.layouts ?? [];
    assert.deepEqual(
      layouts.map(({ testId }) => testId),
      ["row", "card", "card"],
    );
    const row = layouts.find(({ testId }) => testId === "row");
    const cards = layouts.filter(({ testId }) => testId === "card");
    assert.equal(row?.box.width, 300);
    assert.equal(cards.length, 2);
    assert.deepEqual(
      row?.children.map(({ kind }) => kind),
      ["layout", "layout"],
    );
    assert.notEqual(row?.children[0]?.id, row?.children[1]?.id);

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(
      file.slides[0]?.objects.map(({ boxEmu }) => ({
        x: boxEmu.x / 9525,
        y: boxEmu.y / 9525,
        width: boxEmu.width / 9525,
        height: boxEmu.height / 9525,
      })),
      [
        { x: 0, y: 0, width: 150, height: 20 },
        { x: 150, y: 0, width: 150, height: 20 },
      ],
    );
  });

  it("原生组合保留成员结构并求解组合内部的透明布局容器", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 320, height: 240 },
        createElement(
          Slide,
          { slideId: "group-with-layout" },
          createElement(
            Group,
            { testId: "native-group" },
            createElement(
              Box,
              {
                testId: "inner-box",
                style: { width: 100, height: 100 },
              },
              createElement(Shape, { testId: "group-member", height: 20 }),
            ),
          ),
        ),
      ),
    );

    const group = result.report.slides[0]?.objects[0];
    const member = group?.members?.[0];
    assert.equal(group?.kind, "group");
    assert.equal(group?.testId, "native-group");
    assert.deepEqual(member?.box, {
      x: 0,
      y: 0,
      width: 100,
      height: 20,
    });
    const layout = result.report.slides[0]?.layouts?.[0];
    assert.equal(layout?.testId, "inner-box");
    assert.deepEqual(layout?.box, {
      x: 0,
      y: 0,
      width: 100,
      height: 100,
    });
    assert.deepEqual(layout?.children, [{ kind: "object", id: member?.id }]);

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.equal(file.slides[0]?.objects[0]?.kind, "group");
    assert.equal(
      file.slides[0]?.objects[0]?.members?.[0]?.name,
      "group-member",
    );
    assert.deepEqual(file.slides[0]?.objects[0]?.members?.[0]?.boxEmu, {
      x: 0,
      y: 0,
      width: 100 * 9525,
      height: 20 * 9525,
    });
  });

  it("拒绝块容器上的弹性专属属性，并在块父级忽略弹性子项属性", async () => {
    await assert.rejects(
      build(
        createElement(
          Deck,
          { width: 320, height: 240 },
          createElement(
            Slide,
            { slideId: "invalid-context" },
            createElement(Box, { style: { justifyContent: "center" } }),
          ),
        ),
      ),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some(({ code }) => code === "invalid-layout-context"),
    );

    const result = await build(
      createElement(
        Deck,
        { width: 320, height: 240 },
        createElement(
          Slide,
          { slideId: "ignored-item-style" },
          createElement(
            Box,
            { style: { width: 100 } },
            createElement(Shape, {
              testId: "block-item",
              width: 30,
              height: 10,
              style: { flex: 1, alignSelf: "baseline" },
            }),
          ),
        ),
      ),
    );
    assert.deepEqual(result.report.slides[0]?.objects[0]?.box, {
      x: 0,
      y: 0,
      width: 30,
      height: 10,
    });
  });

  it("容器和子对象越界时保留求解尺寸并返回警告", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 320, height: 240 },
        createElement(
          Slide,
          { slideId: "overflow" },
          createElement(
            Box,
            { testId: "wide-container", style: { width: 400, height: 50 } },
            createElement(Shape, { testId: "wide-shape", height: 10 }),
          ),
        ),
      ),
    );

    assert.equal(result.report.slides[0]?.layouts?.[0]?.box.width, 400);
    assert.equal(result.report.slides[0]?.objects[0]?.box?.width, 400);
    assert.deepEqual(
      result.report.diagnostics.map(({ path, code }) => ({ path, code })),
      [
        { path: "deck/slide[overflow]/box[wide-container]", code: "overflow" },
        {
          path: "deck/slide[overflow]/box[wide-container]/shape[wide-shape]",
          code: "overflow",
        },
      ],
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.equal(file.slides[0]?.objects[0]?.boxEmu.width, 400 * 9525);
  });

  it("图片和形状都按弹性对齐位置写入原生对象", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 320, height: 240 },
        createElement(
          Slide,
          { slideId: "aligned-objects" },
          createElement(
            Box,
            {
              testId: "toolbar",
              style: {
                display: "flex",
                width: 100,
                height: 100,
                justifyContent: "space-between",
                alignItems: "center",
              },
            },
            createElement(Image, {
              testId: "logo",
              src: wideImage,
              width: 30,
              height: 20,
            }),
            createElement(Shape, {
              testId: "action",
              width: 20,
              height: 20,
            }),
          ),
        ),
      ),
    );

    assert.deepEqual(
      result.report.slides[0]?.objects.map(({ kind, testId, box }) => ({
        kind,
        testId,
        box,
      })),
      [
        {
          kind: "image",
          testId: "logo",
          box: { x: 0, y: 40, width: 30, height: 20 },
        },
        {
          kind: "shape",
          testId: "action",
          box: { x: 80, y: 40, width: 20, height: 20 },
        },
      ],
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(
      file.slides[0]?.objects.map(({ boxEmu }) => ({
        x: boxEmu.x / 9525,
        y: boxEmu.y / 9525,
        width: boxEmu.width / 9525,
        height: boxEmu.height / 9525,
      })),
      [
        { x: 0, y: 40, width: 30, height: 20 },
        { x: 80, y: 40, width: 20, height: 20 },
      ],
    );
  });

  it("baseline 按有度量字体的文字基线对齐，并拒绝无依据的文字组合", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 320, height: 240 },
        createElement(
          Slide,
          { slideId: "text-baseline" },
          createElement(
            Box,
            {
              style: {
                display: "flex",
                alignItems: "baseline",
                width: 300,
                height: 100,
              },
            },
            createElement(
              Text,
              {
                testId: "small-text",
                width: 100,
                height: 60,
                style: { fontFamily: cjkFont, fontSize: 20 },
              },
              "中",
            ),
            createElement(
              Text,
              {
                testId: "large-text",
                width: 100,
                height: 60,
                style: { fontFamily: cjkFont, fontSize: 40 },
              },
              "中",
            ),
          ),
        ),
      ),
    );

    assert.deepEqual(
      result.report.slides[0]?.objects.map(({ testId, box }) => ({
        testId,
        y: box?.y,
      })),
      [
        { testId: "small-text", y: 23.2 },
        { testId: "large-text", y: 0 },
      ],
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(
      file.slides[0]?.objects.map(({ boxEmu }) => boxEmu.y / 9525),
      [23.2, 0],
    );

    await assert.rejects(
      build(
        createElement(
          Deck,
          { width: 320, height: 240 },
          createElement(
            Slide,
            { slideId: "unmeasured-baseline" },
            createElement(
              Box,
              {
                style: {
                  display: "flex",
                  alignItems: "baseline",
                  width: 300,
                  height: 100,
                },
              },
              createElement(
                Text,
                { testId: "unmeasured-text", width: 100, height: 60 },
                "text",
              ),
            ),
          ),
        ),
      ),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some(
          ({ code, path }) =>
            code === "baseline-measurement-unavailable" &&
            path ===
              "deck/slide[unmeasured-baseline]/box[1]/text[unmeasured-text]",
        ),
    );

    await assert.rejects(
      build(
        createElement(
          Deck,
          { width: 320, height: 240 },
          createElement(
            Slide,
            { slideId: "wrapped-text-baseline" },
            createElement(
              Box,
              {
                style: {
                  display: "flex",
                  flexWrap: "wrap",
                  alignItems: "baseline",
                  width: 100,
                  height: 100,
                },
              },
              createElement(
                Text,
                {
                  testId: "wrapped-text",
                  width: 60,
                  height: 20,
                  style: { fontFamily: cjkFont, fontSize: 20 },
                },
                "中",
              ),
            ),
          ),
        ),
      ),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some(
          ({ code }) => code === "unsupported-baseline-context",
        ),
    );
  });

  it("纯盒子与列方向 baseline 使用 Yoga 合成边缘基线", async () => {
    const wrapped = await build(
      createElement(
        Deck,
        { width: 320, height: 240 },
        createElement(
          Slide,
          { slideId: "synthetic-wrap-baseline" },
          createElement(
            Box,
            {
              style: {
                display: "flex",
                flexWrap: "wrap",
                alignItems: "baseline",
                width: 100,
                height: 100,
              },
            },
            createElement(Shape, {
              testId: "wrap-first",
              width: 60,
              height: 20,
            }),
            createElement(Shape, {
              testId: "wrap-second",
              width: 60,
              height: 20,
            }),
          ),
        ),
      ),
    );
    assert.deepEqual(
      wrapped.report.slides[0]?.objects.map(({ box }) => box?.y),
      [0, 50],
    );
    assert.deepEqual((await validatePptx(wrapped.pptx)).errors, []);

    const column = await build(
      createElement(
        Deck,
        { width: 320, height: 240 },
        createElement(
          Slide,
          { slideId: "synthetic-column-baseline" },
          createElement(
            Box,
            {
              style: {
                display: "flex",
                flexDirection: "column",
                alignItems: "baseline",
                width: 120,
                height: 120,
              },
            },
            createElement(
              Text,
              {
                testId: "column-small-text",
                width: 100,
                height: 40,
                style: { fontFamily: "Aptos", fontSize: 20 },
              },
              "中",
            ),
            createElement(Image, {
              testId: "column-image",
              src: wideImage,
              width: 30,
              height: 20,
            }),
            createElement(Shape, {
              testId: "column-shape",
              width: 80,
              height: 20,
            }),
          ),
        ),
      ),
    );
    assert.equal(
      column.report.diagnostics.some(
        ({ code }) => code === "baseline-measurement-unavailable",
      ),
      false,
    );
    assert.deepEqual((await validatePptx(column.pptx)).errors, []);
  });

  it("自适应 baseline 行高包含下移文字框并保持后续块位置", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 320, height: 240 },
        createElement(
          Slide,
          { slideId: "natural-baseline-height" },
          createElement(
            Box,
            { style: { width: 300 } },
            createElement(
              Box,
              {
                testId: "baseline-row",
                style: {
                  display: "flex",
                  alignItems: "baseline",
                  width: 300,
                },
              },
              createElement(
                Text,
                {
                  testId: "small-text",
                  width: 100,
                  height: 60,
                  style: { fontFamily: cjkFont, fontSize: 20 },
                },
                "中",
              ),
              createElement(
                Text,
                {
                  testId: "large-text",
                  width: 100,
                  height: 60,
                  style: { fontFamily: cjkFont, fontSize: 40 },
                },
                "中",
              ),
            ),
            createElement(Shape, {
              testId: "following-block",
              width: 100,
              height: 20,
            }),
          ),
        ),
      ),
    );

    const layouts = result.report.slides[0]?.layouts ?? [];
    assert.equal(
      layouts.find(({ testId }) => testId === "baseline-row")?.box.height,
      83.2,
    );
    assert.deepEqual(
      result.report.slides[0]?.objects.map(({ testId, box }) => ({
        testId,
        y: box?.y,
      })),
      [
        { testId: "small-text", y: 23.2 },
        { testId: "large-text", y: 0 },
        { testId: "following-block", y: 83.2 },
      ],
    );

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(
      file.slides[0]?.objects.map(({ boxEmu }) => boxEmu.y / 9525),
      [23.2, 0, 83.2],
    );
  });

  it("弹性列保留文字主轴内容高度，minHeight: 0 允许收缩", async () => {
    const columnText = (testId: string, minHeight?: number) =>
      createElement(Text, {
        testId,
        width: 20,
        height: "auto",
        style: {
          fontFamily: cjkFont,
          fontSize: 20,
          ...(minHeight === undefined ? {} : { minHeight }),
        },
        children: "中文中文中文",
      });
    const result = await build(
      createElement(
        Deck,
        { width: 320, height: 700 },
        createElement(
          Slide,
          { slideId: "column-min-content" },
          createElement(
            Box,
            {
              testId: "constrained-column",
              style: {
                display: "flex",
                flexDirection: "column",
                width: 100,
                height: 300,
              },
            },
            columnText("column-a"),
            columnText("column-b"),
          ),
          createElement(
            Box,
            {
              testId: "released-column",
              style: {
                display: "flex",
                flexDirection: "column",
                width: 100,
                height: 300,
              },
            },
            columnText("released-column-a", 0),
            columnText("released-column-b", 0),
          ),
        ),
      ),
    );

    assert.deepEqual(
      result.report.slides[0]?.objects.map(({ testId, box }) => ({
        testId,
        box,
      })),
      [
        {
          testId: "column-a",
          box: { x: 0, y: 0, width: 20, height: 173.76 },
        },
        {
          testId: "column-b",
          box: { x: 0, y: 173.76, width: 20, height: 173.76 },
        },
        {
          testId: "released-column-a",
          box: { x: 0, y: 300, width: 20, height: 150 },
        },
        {
          testId: "released-column-b",
          box: { x: 0, y: 450, width: 20, height: 150 },
        },
      ],
    );
    assert.deepEqual(
      result.report.diagnostics.map(({ code, path }) => ({ code, path })),
      [
        {
          code: "overset",
          path: "deck/slide[column-min-content]/box[released-column]/text[released-column-a]",
        },
        {
          code: "overset",
          path: "deck/slide[column-min-content]/box[released-column]/text[released-column-b]",
        },
      ],
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(
      file.slides[0]?.objects.map(({ boxEmu }) => ({
        y: boxEmu.y / 9525,
        height: boxEmu.height / 9525,
      })),
      [
        { y: 0, height: 173.76 },
        { y: 173.76, height: 173.76 },
        { y: 300, height: 150 },
        { y: 450, height: 150 },
      ],
    );
  });
});
