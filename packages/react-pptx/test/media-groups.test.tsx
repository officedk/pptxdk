import assert from "node:assert/strict";
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { pathToFileURL } from "node:url";

import {
  Box,
  build,
  BuildError,
  createBuilder,
  Deck,
  font,
  Group,
  groupCapability,
  Image,
  imageCapability,
  P,
  Shape,
  shapeCapability,
  Slide,
  Span,
  Text,
} from "@react-pptx/react-pptx";
import { validatePptx } from "@react-pptx/validation";
import { strFromU8, unzipSync, zipSync } from "fflate";
import { createElement, type ComponentType } from "react";
import { describe, it, vi } from "vitest";

const png = new Uint8Array(
  Buffer.from(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/n1sAAAAASUVORK5CYII=",
    "base64",
  ),
);
const jpeg = new Uint8Array(
  Buffer.from(
    "/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAP//////////////////////////////////////////////////////////////////////////////////////2wBDAf//////////////////////////////////////////////////////////////////////////////////////wAARCAABAAEDASIAAhEBAxEB/8QAFQABAQAAAAAAAAAAAAAAAAAAAAX/xAAUEAEAAAAAAAAAAAAAAAAAAAAA/9oADAMBAAIQAxAAAAFR/8QAFBABAAAAAAAAAAAAAAAAAAAAAP/aAAgBAQABBQJ//8QAFBABAAAAAAAAAAAAAAAAAAAAAP/aAAgBAwEBPwF//8QAFBABAAAAAAAAAAAAAAAAAAAAAP/aAAgBAgEBPwF//8QAFBABAAAAAAAAAAAAAAAAAAAAAP/aAAgBAQAGPwJ//8QAFBABAAAAAAAAAAAAAAAAAAAAAP/aAAgBAQABPyF//9oADAMBAAIRAxEAPwD/xAAUEAEAAAAAAAAAAAAAAAAAAAAA/9oACAEBAAE/IV//2Q==",
    "base64",
  ),
);
const gif = new Uint8Array(
  Buffer.from(
    "47494638396101000100800000000000ffffff2c00000000010001000002024401003b",
    "hex",
  ),
);
const widePng = new Uint8Array(
  readFileSync(new URL("./assets/wide-20000x1.png", import.meta.url)),
);

describe("images and native groups", () => {
  it("preserves text paragraphs, shape text, media relationships, and member geometry in nested native groups", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 960, height: 540 },
        createElement(
          Slide,
          { slideId: "mixed-group" },
          createElement(
            Group,
            { testId: "outer-group" },
            createElement(Shape, {
              testId: "shape-text",
              geometry: "hexagon",
              x: 20,
              y: 24,
              width: 120,
              height: 72,
              text: "Shape body",
            }),
            createElement(
              Text,
              {
                testId: "paragraph-text",
                x: 160,
                y: 24,
                width: 360,
                height: 120,
                style: { fontFamily: font("Aptos"), fontSize: "14pt" },
              },
              createElement(
                P,
                { style: { list: { type: "bullet", character: "•" } } },
                "Group paragraph ",
                createElement(Span, { style: { fontWeight: "bold" } }, "style"),
              ),
            ),
            createElement(Image, {
              testId: "group-image",
              src: png,
              x: 20,
              y: 120,
              width: 120,
              height: 72,
              fit: "cover",
            }),
            createElement(
              Group,
              { testId: "inner-group" },
              createElement(Shape, {
                testId: "inner-shape",
                x: 190,
                y: 80,
                width: 40,
                height: 20,
              }),
            ),
          ),
        ),
      ),
    );

    const outerObservation = result.report.slides[0]?.objects[0];
    assert.equal(outerObservation?.kind, "group");
    assert.deepEqual(
      outerObservation?.members?.map(({ name }) => name),
      ["shape-text", "paragraph-text", "group-image", "inner-group"],
    );

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const outer = file.slides[0]?.objects[0];
    assert.equal(outer?.kind, "group");
    assert.equal(outer?.name, "outer-group");
    assert.deepEqual(outer?.boxEmu, {
      x: 190_500,
      y: 228_600,
      width: 4_762_500,
      height: 1_600_200,
    });
    assert.deepEqual(outer?.groupTransform, {
      childOffset: { x: 190_500, y: 228_600 },
      childExtent: { width: 4_762_500, height: 1_600_200 },
    });

    const members = outer?.members ?? [];
    const shapeText = members.find(({ name }) => name === "shape-text");
    assert.equal(shapeText?.kind, "shape");
    assert.equal(shapeText?.geometry, "hexagon");
    assert.equal(shapeText?.text?.text, "Shape body");
    assert.deepEqual(shapeText?.boxEmu, {
      x: 190_500,
      y: 228_600,
      width: 1_143_000,
      height: 685_800,
    });

    const paragraphText = members.find(({ name }) => name === "paragraph-text");
    assert.equal(paragraphText?.text?.text, "Group paragraph style");
    assert.deepEqual(paragraphText?.text?.paragraphs[0]?.list, {
      type: "bullet",
      character: "•",
    });
    assert.deepEqual(
      paragraphText?.text?.paragraphs[0]?.runs.map(({ text, bold }) => ({
        text,
        bold,
      })),
      [
        { text: "Group paragraph ", bold: false },
        { text: "style", bold: true },
      ],
    );
    assert.deepEqual(paragraphText?.boxEmu, {
      x: 1_524_000,
      y: 228_600,
      width: 3_429_000,
      height: 1_143_000,
    });

    const image = members.find(({ kind }) => kind === "image");
    assert.equal(image?.name, "group-image");
    assert.equal(image?.image?.contentType, "image/png");
    assert.ok(image?.image?.relationshipId);
    assert.deepEqual(image?.boxEmu, {
      x: 190_500,
      y: 1_143_000,
      width: 1_143_000,
      height: 685_800,
    });
    assert.deepEqual(image?.crop, { top: 20_000, bottom: 20_000 });
    assert.deepEqual(unzipSync(result.pptx)[image?.image?.part ?? ""], png);

    const inner = members.find(({ kind }) => kind === "group");
    assert.equal(inner?.name, "inner-group");
    assert.deepEqual(inner?.boxEmu, {
      x: 1_809_750,
      y: 762_000,
      width: 381_000,
      height: 190_500,
    });
    const innerShape = inner?.members?.[0];
    assert.equal(innerShape?.name, "inner-shape");
    assert.ok(Number.isSafeInteger(innerShape?.id));
    assert.deepEqual(innerShape?.boxEmu, inner?.boxEmu);
    const objectIds = [
      outer?.id,
      shapeText?.id,
      paragraphText?.id,
      image?.id,
      inner?.id,
      innerShape?.id,
    ];
    assert.ok(objectIds.every((id) => Number.isSafeInteger(id)));
    assert.equal(new Set(objectIds).size, objectIds.length);
    assert.equal(
      Object.keys(unzipSync(result.pptx)).filter((part) =>
        part.startsWith("ppt/media/"),
      ).length,
      1,
    );
  });

  it("reads image resources through public build and establishes image relationships in actual PPTX", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 960, height: 540 },
        createElement(
          Slide,
          { slideId: "media" },
          createElement(Image, {
            testId: "cover-photo",
            src: png,
            alt: "封面照片",
            x: 32,
            y: 48,
            width: 120,
            height: 72,
            fit: "cover",
          }),
        ),
      ),
    );

    assert.deepEqual(result.report.slides[0]?.objects, [
      {
        id: 2,
        kind: "image",
        name: "cover-photo",
        testId: "cover-photo",
        box: { x: 32, y: 48, width: 120, height: 72 },
      },
    ]);

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const image = file.slides[0]?.objects[0];
    assert.equal(image?.kind, "image");
    assert.deepEqual(image?.boxEmu, {
      x: 304800,
      y: 457200,
      width: 1143000,
      height: 685800,
    });
    assert.equal(image?.image?.contentType, "image/png");
    assert.deepEqual(image?.crop, { top: 20000, bottom: 20000 });
    assert.ok(image?.image?.relationshipId);
    assert.ok(image?.image?.part);
    assert.deepEqual(unzipSync(result.pptx)[image.image.part], png);

    const damagedParts = unzipSync(result.pptx);
    delete damagedParts[image.image.part];
    const damagedFile = await validatePptx(zipSync(damagedParts));
    assert.ok(
      damagedFile.errors.some(({ code }) => code === "invalid-image-part"),
    );
  });

  it("reads image resources from local paths and file URLs", async () => {
    const directory = mkdtempSync(join(tmpdir(), "react-pptx-image-"));
    const path = join(directory, "pixel.png");
    writeFileSync(path, png);
    try {
      for (const source of [path, pathToFileURL(path)]) {
        const result = await build(
          createElement(
            Deck,
            { width: 960, height: 540 },
            createElement(
              Slide,
              { slideId: "local-image" },
              createElement(Image, {
                testId: "local-image",
                src: source,
              }),
            ),
          ),
        );
        const file = await validatePptx(result.pptx);
        assert.deepEqual(file.errors, []);
        assert.deepEqual(result.report.slides[0]?.objects[0]?.box, {
          x: 0,
          y: 0,
          width: 1,
          height: 1,
        });
        assert.equal(
          file.slides[0]?.objects[0]?.image?.contentType,
          "image/png",
        );
      }
    } finally {
      rmSync(directory, { recursive: true, force: true });
    }
  });

  it("reads intrinsic dimensions of JPEG and GIF and preserves their native resource types", async () => {
    for (const [src, contentType] of [
      [jpeg, "image/jpeg"],
      [gif, "image/gif"],
    ] as const) {
      const result = await build(
        createElement(
          Deck,
          { width: 960, height: 540 },
          createElement(
            Slide,
            { slideId: contentType },
            createElement(Image, { testId: contentType, src }),
          ),
        ),
      );
      const file = await validatePptx(result.pptx);
      assert.deepEqual(file.errors, []);
      assert.deepEqual(result.report.slides[0]?.objects[0]?.box, {
        x: 0,
        y: 0,
        width: 1,
        height: 1,
      });
      const image = file.slides[0]?.objects[0];
      assert.equal(image?.image?.contentType, contentType);
      assert.deepEqual(unzipSync(result.pptx)[image?.image?.part ?? ""], src);
    }
  });

  it("reuses media parts by content and computes image dimensions for aspect-ratio-preserving modes", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 960, height: 540 },
        createElement(
          Slide,
          { slideId: "reuse" },
          createElement(Image, {
            testId: "contained",
            src: png,
            x: 32,
            y: 48,
            width: 120,
            height: 72,
            fit: "contain",
          }),
          createElement(Image, {
            testId: "duplicate",
            src: png,
            x: 200,
            y: 48,
            width: 120,
            height: 72,
          }),
        ),
      ),
    );

    assert.deepEqual(
      result.report.slides[0]?.objects.map(({ box }) => box),
      [
        { x: 56, y: 48, width: 72, height: 72 },
        { x: 200, y: 48, width: 120, height: 72 },
      ],
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const images = file.slides[0]?.objects ?? [];
    assert.equal(images[0]?.image?.part, images[1]?.image?.part);
    assert.equal(
      images[0]?.image?.relationshipId,
      images[1]?.image?.relationshipId,
    );
    assert.equal(
      Object.keys(unzipSync(result.pptx)).filter((part) =>
        part.startsWith("ppt/media/"),
      ).length,
      1,
    );
  });

  it("writes final dimensions of contain to native image transform", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 960, height: 540 },
        createElement(
          Slide,
          { slideId: "contain-transform" },
          createElement(Image, {
            testId: "contained-square",
            src: png,
            x: 20,
            y: 30,
            width: 100,
            height: 80,
            fit: "contain",
          }),
        ),
      ),
    );

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(result.report.slides[0]?.objects[0]?.box, {
      x: 30,
      y: 30,
      width: 80,
      height: 80,
    });
    const slideXml = strFromU8(
      unzipSync(result.pptx)["ppt/slides/slide1.xml"]!,
    );
    assert.match(
      slideXml,
      /<p:spPr><a:xfrm><a:off x="285750" y="285750"\/><a:ext cx="762000" cy="762000"\/><\/a:xfrm>/u,
    );
  });

  it("rejects unrepresentable tiny image dimensions derived from contain", async () => {
    await assert.rejects(
      build(
        createElement(
          Deck,
          { width: 960, height: 540 },
          createElement(
            Slide,
            { slideId: "tiny-contain" },
            createElement(Image, {
              testId: "tiny-contain",
              src: widePng,
              width: 1,
              height: 1,
              fit: "contain",
            }),
          ),
        ),
      ),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "invalid-coordinate");
        assert.match(
          error.diagnostics[0]?.path ?? "",
          /image\[tiny-contain\]/u,
        );
        return true;
      },
    );
  });

  it("crops by explicit fraction and rejects out-of-bounds crop values", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 960, height: 540 },
        createElement(
          Slide,
          { slideId: "crop" },
          createElement(Image, {
            testId: "cropped",
            src: png,
            x: 0,
            y: 0,
            width: 120,
            height: 72,
            fit: "contain",
            crop: { left: 0.1, right: 0.2 },
          }),
        ),
      ),
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(file.slides[0]?.objects[0]?.crop, {
      left: 10000,
      right: 20000,
    });

    for (const [index, crop] of [
      { left: 0.6, right: 0.5 },
      { top: Number.POSITIVE_INFINITY },
      { bottom: -0.1 },
    ].entries()) {
      await assert.rejects(
        build(
          createElement(
            Deck,
            { width: 960, height: 540 },
            createElement(
              Slide,
              { slideId: `bad-crop-${index}` },
              createElement(Image, {
                testId: `bad-crop-${index}`,
                src: png,
                width: 120,
                height: 72,
                crop,
              }),
            ),
          ),
        ),
        (error: unknown) => {
          assert.ok(error instanceof BuildError);
          assert.equal(error.diagnostics[0]?.code, "invalid-crop");
          assert.match(error.diagnostics[0]?.path ?? "", /image\[bad-crop-/u);
          return true;
        },
      );
    }
  });

  it("rejects invalid fit values, missing files, and unrecognizable image content", async () => {
    const invalidFit = Image as ComponentType<Record<string, unknown>>;
    for (const [props, code] of [
      [
        { testId: "bad-fit", src: png, width: 120, height: 72, fit: "stretch" },
        "unsupported-fit",
      ],
      [
        {
          testId: "missing",
          src: new URL("file:///tmp/react-pptx-no-image.png"),
        },
        "image-read-failed",
      ],
      [
        { testId: "unknown-format", src: new Uint8Array([1, 2, 3, 4]) },
        "unsupported-image-format",
      ],
    ] as const) {
      await assert.rejects(
        build(
          createElement(
            Deck,
            { width: 960, height: 540 },
            createElement(
              Slide,
              { slideId: "invalid-image" },
              createElement(invalidFit, props),
            ),
          ),
        ),
        (error: unknown) => {
          assert.ok(error instanceof BuildError);
          assert.equal(error.diagnostics[0]?.code, code);
          assert.match(
            error.diagnostics[0]?.path ?? "",
            /slide\[invalid-image\]/u,
          );
          return true;
        },
      );
    }
  });

  it("rejects PNG missing IDAT chunks", async () => {
    const noIdatPng = new Uint8Array(
      Buffer.from(
        "89504e470d0a1a0a0000000d49484452000000010000000108000000003a7e9b550000000049454e44ae426082",
        "hex",
      ),
    );
    await assert.rejects(
      build(
        createElement(
          Deck,
          { width: 960, height: 540 },
          createElement(
            Slide,
            { slideId: "no-idat" },
            createElement(Image, {
              testId: "no-idat",
              src: noIdatPng,
            }),
          ),
        ),
      ),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "unsupported-image-format");
        assert.match(error.diagnostics[0]?.path ?? "", /image\[no-idat\]/u);
        return true;
      },
    );
  });

  it("rejects JPEG missing scan data and EOI marker", async () => {
    const truncatedJpeg = new Uint8Array(
      Buffer.from("ffd8ffc00011080001000103011100021100031100", "hex"),
    );
    await assert.rejects(
      build(
        createElement(
          Deck,
          { width: 960, height: 540 },
          createElement(
            Slide,
            { slideId: "truncated-jpeg" },
            createElement(Image, {
              testId: "truncated-jpeg",
              src: truncatedJpeg,
            }),
          ),
        ),
      ),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "unsupported-image-format");
        assert.match(
          error.diagnostics[0]?.path ?? "",
          /image\[truncated-jpeg\]/u,
        );
        return true;
      },
    );
  });

  it("rejects GIF missing image data block", async () => {
    const emptyGif = new Uint8Array(
      Buffer.from("474946383961010001000000003b", "hex"),
    );
    await assert.rejects(
      build(
        createElement(
          Deck,
          { width: 960, height: 540 },
          createElement(
            Slide,
            { slideId: "empty-gif" },
            createElement(Image, {
              testId: "empty-gif",
              src: emptyGif,
            }),
          ),
        ),
      ),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "unsupported-image-format");
        assert.match(error.diagnostics[0]?.path ?? "", /image\[empty-gif\]/u);
        return true;
      },
    );
  });

  it("preserves identities, extents, and relative coordinates of nested native group members", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 960, height: 540 },
        createElement(
          Slide,
          { slideId: "groups" },
          createElement(
            Group,
            { testId: "outer" },
            createElement(Shape, {
              testId: "left",
              x: 20,
              y: 30,
              width: 40,
              height: 30,
            }),
            createElement(
              Group,
              { testId: "inner" },
              createElement(Image, {
                testId: "photo",
                src: png,
                x: 100,
                y: 50,
                width: 80,
                height: 40,
              }),
              createElement(Shape, {
                testId: "right",
                x: 150,
                y: 90,
                width: 20,
                height: 10,
              }),
            ),
          ),
        ),
      ),
    );

    assert.deepEqual(result.report.slides[0]?.objects, [
      {
        id: 2,
        kind: "group",
        name: "outer",
        testId: "outer",
        box: { x: 20, y: 30, width: 160, height: 70 },
        members: [
          {
            id: 3,
            kind: "shape",
            name: "left",
            testId: "left",
            geometry: "rect",
            box: { x: 20, y: 30, width: 40, height: 30 },
          },
          {
            id: 4,
            kind: "group",
            name: "inner",
            testId: "inner",
            box: { x: 100, y: 50, width: 80, height: 50 },
            members: [
              {
                id: 5,
                kind: "image",
                name: "photo",
                testId: "photo",
                box: { x: 100, y: 50, width: 80, height: 40 },
              },
              {
                id: 6,
                kind: "shape",
                name: "right",
                testId: "right",
                geometry: "rect",
                box: { x: 150, y: 90, width: 20, height: 10 },
              },
            ],
          },
        ],
      },
    ]);

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const outer = file.slides[0]?.objects[0];
    assert.equal(outer?.kind, "group");
    assert.equal(outer?.id, 2);
    assert.deepEqual(outer?.boxEmu, {
      x: 190500,
      y: 285750,
      width: 1524000,
      height: 666750,
    });
    assert.deepEqual(outer?.groupTransform, {
      childOffset: { x: 190500, y: 285750 },
      childExtent: { width: 1524000, height: 666750 },
    });
    assert.deepEqual(
      outer?.members?.map(({ id, kind, name, boxEmu }) => ({
        id,
        kind,
        name,
        boxEmu,
      })),
      [
        {
          id: 3,
          kind: "shape",
          name: "left",
          boxEmu: { x: 190500, y: 285750, width: 381000, height: 285750 },
        },
        {
          id: 4,
          kind: "group",
          name: "inner",
          boxEmu: {
            x: 952500,
            y: 476250,
            width: 762000,
            height: 476250,
          },
        },
      ],
    );
    const inner = outer?.members?.[1];
    assert.deepEqual(
      [
        (inner?.boxEmu.x ?? 0) - (outer?.groupTransform?.childOffset.x ?? 0),
        (inner?.boxEmu.y ?? 0) - (outer?.groupTransform?.childOffset.y ?? 0),
        (inner?.members?.[1]?.boxEmu.x ?? 0) -
          (inner?.groupTransform?.childOffset.x ?? 0),
        (inner?.members?.[1]?.boxEmu.y ?? 0) -
          (inner?.groupTransform?.childOffset.y ?? 0),
      ],
      [762000, 190500, 476250, 381000],
    );
    assert.deepEqual(
      [
        outer?.id,
        ...(outer?.members?.map(({ id }) => id) ?? []),
        ...(inner?.members?.map(({ id }) => id) ?? []),
      ],
      [2, 3, 4, 5, 6],
    );
    assert.equal(inner?.members?.[0]?.image?.contentType, "image/png");
  });

  it("restricts Group to native groups and reports errors accurately when member capabilities are stripped", async () => {
    const positionedGroup = Group as ComponentType<Record<string, unknown>>;
    await assert.rejects(
      build(
        createElement(
          Deck,
          { width: 960, height: 540 },
          createElement(
            Slide,
            { slideId: "group-props" },
            createElement(
              positionedGroup,
              { testId: "not-a-box", x: 20 },
              createElement(Shape, {
                testId: "member",
                x: 20,
                y: 30,
                width: 40,
                height: 30,
              }),
            ),
          ),
        ),
      ),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "unknown-parameter");
        assert.match(error.diagnostics[0]?.path ?? "", /group\[not-a-box\]/u);
        return true;
      },
    );

    await assert.rejects(
      build(
        createElement(
          Deck,
          { width: 960, height: 540 },
          createElement(
            Slide,
            { slideId: "empty-group" },
            createElement(Group, { testId: "empty" }),
          ),
        ),
      ),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "empty-group");
        assert.match(error.diagnostics[0]?.path ?? "", /group\[empty\]/u);
        return true;
      },
    );

    const clippedBuild = createBuilder([shapeCapability, groupCapability]);
    await assert.rejects(
      clippedBuild(
        createElement(
          Deck,
          { width: 960, height: 540 },
          createElement(
            Slide,
            { slideId: "clipped-group" },
            createElement(
              Group,
              { testId: "clipped" },
              createElement(Image, { testId: "missing-image", src: png }),
            ),
          ),
        ),
      ),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.ok(
          error.diagnostics.some(
            ({ code, path }) =>
              code === "missing-capability" &&
              /group\[clipped\]\/image\[missing-image\]/u.test(path),
          ),
        );
        return true;
      },
    );
  });

  it("executes imageCapability.parse and reads image resource exactly once per Image node during build", async () => {
    const parseSpy = vi.spyOn(imageCapability, "parse");
    const measureSpy = vi.spyOn(imageCapability, "measureLayout");
    const compileSpy = vi.spyOn(imageCapability, "compile");

    try {
      const result = await build(
        createElement(
          Deck,
          { width: 960, height: 540 },
          createElement(
            Slide,
            { slideId: "single-image-read" },
            createElement(
              Box,
              { style: { display: "flex", width: 400 } },
              createElement(Image, {
                testId: "single-read-image",
                src: png,
                style: { opacity: 0.8 },
              }),
            ),
          ),
        ),
      );

      assert.equal(
        parseSpy.mock.calls.length,
        1,
        `Expected imageCapability.parse to be called exactly 1 time, got ${parseSpy.mock.calls.length}`,
      );

      const parsedResult = parseSpy.mock.results[0]?.value;
      assert.ok(parsedResult);
      assert.equal(parsedResult.contract.nativeKind, "pic");
      assert.equal(parsedResult.parsed.image.width, 1);
      assert.equal(parsedResult.parsed.image.height, 1);

      // Verify measureLayout and compile reused the parsed image object directly
      assert.ok(measureSpy.mock.calls.length >= 1);
      for (const call of measureSpy.mock.calls) {
        assert.equal(call[0]?.image, parsedResult.parsed.image);
      }
      assert.equal(compileSpy.mock.calls.length, 1);
      assert.equal(
        compileSpy.mock.calls[0]?.[0]?.image,
        parsedResult.parsed.image,
      );

      const validation = await validatePptx(result.pptx);
      assert.deepEqual(validation.errors, []);
    } finally {
      parseSpy.mockRestore();
      measureSpy.mockRestore();
      compileSpy.mockRestore();
    }
  });

  it("rejects unknown parameters and unknown style keys on Image during parse", async () => {
    await assert.rejects(
      build(
        createElement(
          Deck,
          { width: 960, height: 540 },
          createElement(
            Slide,
            { slideId: "invalid-image-param" },
            // @ts-expect-error -- testing unknown parameter
            createElement(Image, { src: png, unknownParam: "forbidden" }),
          ),
        ),
      ),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "unknown-parameter");
        return true;
      },
    );

    await assert.rejects(
      build(
        createElement(
          Deck,
          { width: 960, height: 540 },
          createElement(
            Slide,
            { slideId: "invalid-image-style" },
            createElement(Image, {
              src: png,
              // @ts-expect-error -- testing unknown style property
              style: { unknownStyleKey: "forbidden" },
            }),
          ),
        ),
      ),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "unknown-style");
        return true;
      },
    );

    await assert.rejects(
      build(
        createElement(
          Deck,
          { width: 960, height: 540 },
          createElement(
            Slide,
            { slideId: "invalid-image-padding" },
            createElement(Image, {
              src: png,
              style: { padding: 10 },
            }),
          ),
        ),
      ),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "invalid-layout-context");
        return true;
      },
    );
  });
});
