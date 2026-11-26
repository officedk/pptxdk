import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import type { LinearGradient } from "@pptxdk/capabilities";
import {
  deck,
  group,
  image,
  paragraph,
  shape,
  slide,
  span,
} from "@pptxdk/capabilities/document";
import { Group as ReactGroup } from "@pptxdk/capabilities/group-elements";
import { Image as ReactImage } from "@pptxdk/capabilities/image-elements";
import { Shape as ReactShape } from "@pptxdk/capabilities/shape-elements";
import { BuildError, createMorphScope } from "@pptxdk/core";
import { createResolver } from "@pptxdk/core/build";
import { emit } from "@pptxdk/emit";
import { Deck, Slide } from "@pptxdk/react";
import { validatePptx } from "@pptxdk/validation";
import { strFromU8, unzipSync } from "fflate";
import { createElement } from "react";
import { describe, it } from "vitest";

import { build, build as buildReact } from "./test-helpers.ts";

function slideXml(result: { readonly pptx: Uint8Array }, index = 1): string {
  return strFromU8(unzipSync(result.pptx)[`ppt/slides/slide${index}.xml`]!);
}

const png = new Uint8Array(
  Buffer.from(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/n1sAAAAASUVORK5CYII=",
    "base64",
  ),
);
const widePng = new Uint8Array(
  readFileSync(new URL("./assets/wide-20000x1.png", import.meta.url)),
);

describe("native content in the common document", () => {
  it("creates editable shapes with geometry, rich text, and value-isolated visuals", async () => {
    const stops = [
      { color: "#AA0000", position: 0 },
      { color: "#00AA00", position: 1 },
    ];
    const gradient: LinearGradient = {
      type: "linear-gradient",
      angle: 45,
      stops,
    };
    const style = {
      position: "absolute" as const,
      left: 20,
      top: 24,
      width: 120,
      height: 72,
      background: gradient,
      border: {
        color: "#123456",
        width: 2,
        style: "dashed" as const,
      },
      flipH: true,
      rotate: -15,
      boxShadow: {
        color: { color: "#000000", alpha: 0.4 },
        blur: 3,
        distance: 2,
        angle: 125,
      },
      softEdge: 1,
    };
    const body = paragraph().append(
      span("Shape ").style({ fontWeight: "bold" }),
      "body",
    );
    const hero = createMorphScope().handle("shape-hero");
    const node = shape({
      id: "editable-shape",
      morphKey: hero,
      geometry: "hexagon",
      style,
    }).append(body);
    const document = deck({ width: 960, height: 540 }).append(
      slide({ id: "shape" }).append(node),
    );

    const first = await build(document);
    assert.deepEqual((await validatePptx(first.pptx)).errors, []);
    const firstObject = first.report.slides[0]?.objects[0];
    assert.equal(firstObject?.kind, "shape");
    assert.equal(firstObject?.authorId, "editable-shape");
    assert.equal(firstObject?.text, "Shape body");
    assert.match(firstObject?.name ?? "", /^!!morph-\d+$/u);
    const firstXml = slideXml(first);
    assert.match(firstXml, /<a:prstGeom prst="hexagon">/u);
    assert.match(firstXml, /<a:gradFill/u);
    assert.match(firstXml, /<a:outerShdw/u);
    assert.match(firstXml, /<a:effectLst/u);
    assert.match(firstXml, /Shape /u);
    assert.match(firstXml, /b="1"/u);
    assert.match(firstXml, /AA0000/u);
    assert.doesNotMatch(firstXml, /0000FF/u);

    stops[0]!.color = "#0000FF";
    style.border.color = "#654321";
    const isolated = await build(document);
    assert.equal(isolated.report.slides[0]?.objects[0]?.name, firstObject.name);
    assert.match(slideXml(isolated), /AA0000/u);
    assert.match(slideXml(isolated), /123456/u);
    assert.doesNotMatch(slideXml(isolated), /0000FF|654321/u);

    node.style({
      border: { color: "#00AA00", width: 2, style: "dashed" },
      background: "#00AA00",
    });
    const updated = await build(document);
    assert.deepEqual((await validatePptx(updated.pptx)).errors, []);
    assert.match(slideXml(updated), /00AA00/u);
    assert.match(slideXml(updated), /<a:solidFill><a:srgbClr val="00AA00"\/>/u);
    assert.doesNotMatch(slideXml(updated), /123456/u);
    assert.doesNotMatch(slideXml(updated), /AA0000/u);
    assert.match(firstXml, /AA0000/u);
  });

  it("creates editable native groups whose members, local theme, and appearance rebuild", async () => {
    const localTheme = { colors: { ink: "#112233" } };
    const member = shape({
      id: "group-shape",
      name: "group-shape",
      style: {
        position: "absolute",
        left: 20,
        top: 24,
        width: 120,
        height: 72,
        background: "ink",
        border: { color: "#000000", width: 1 },
      },
    });
    const badge = image({
      id: "group-image",
      name: "group-image",
      src: png,
      style: {
        position: "absolute",
        left: 160,
        top: 24,
        width: 40,
        height: 40,
      },
    });
    const node = group({
      id: "editable-group",
      style: { rotate: 12 },
      theme: localTheme,
    }).append(member, badge);
    const document = deck({
      width: 960,
      height: 540,
      theme: { colors: { ink: "#223344" } },
    }).append(slide({ id: "group" }).append(node));

    const first = await build(document);
    assert.deepEqual((await validatePptx(first.pptx)).errors, []);
    const firstGroup = first.report.slides[0]?.objects[0];
    assert.equal(firstGroup?.kind, "group");
    assert.deepEqual(
      firstGroup?.members?.map(({ name }) => name),
      ["group-shape", "group-image"],
    );
    const firstXml = slideXml(first);
    assert.match(firstXml, /<p:grpSp>/u);
    assert.match(firstXml, /rot="720000"/u);
    assert.match(firstXml, /112233/u);

    localTheme.colors.ink = "#0000FF";
    const isolated = await build(document);
    assert.match(slideXml(isolated), /112233/u);
    assert.doesNotMatch(slideXml(isolated), /0000FF/u);

    member.style({ border: { color: "#445566", width: 1 } });
    node.set({ theme: { colors: { ink: "#778899" } } });
    const updated = await build(document);
    assert.deepEqual((await validatePptx(updated.pptx)).errors, []);
    const updatedXml = slideXml(updated);
    assert.match(updatedXml, /445566/u);
    assert.match(updatedXml, /778899/u);
    assert.doesNotMatch(updatedXml, /112233/u);
  });

  it("creates editable slide backgrounds and local themes through the common document", async () => {
    const node = slide({
      id: "appearance",
      background: "#112233",
      theme: { colors: { ink: "#223344" } },
    }).append(
      shape({
        id: "theme-shape",
        style: {
          position: "absolute",
          left: 20,
          top: 24,
          width: 120,
          height: 72,
          background: "ink",
        },
      }),
    );
    const document = deck({
      width: 960,
      height: 540,
      theme: { colors: { ink: "#000000" } },
    }).append(node);

    const first = await build(document);
    assert.deepEqual((await validatePptx(first.pptx)).errors, []);
    const firstXml = slideXml(first);
    assert.match(
      firstXml,
      /<p:bg><p:bgPr><a:solidFill><a:srgbClr val="112233"\/><\/a:solidFill><a:effectLst\/><\/p:bgPr><\/p:bg>/u,
    );
    assert.match(firstXml, /223344/u);

    node.set({ background: "#445566" });
    node.set({ theme: { colors: { ink: "#556677" } } });
    const updated = await build(document);
    assert.deepEqual((await validatePptx(updated.pptx)).errors, []);
    const updatedXml = slideXml(updated);
    assert.match(updatedXml, /445566/u);
    assert.match(updatedXml, /556677/u);
    assert.doesNotMatch(updatedXml, /112233|223344/u);

    const backgroundBytes = new Uint8Array(png);
    node.set({ background: { image: backgroundBytes } });
    const imageBackground = await build(document);
    const imageFile = await validatePptx(imageBackground.pptx);
    assert.deepEqual(imageFile.errors, []);
    const imageXml = slideXml(imageBackground);
    assert.match(imageXml, /<p:bg><p:bgPr><a:blipFill>/u);
    assert.match(imageXml, /<a:blip r:embed="rId\d+"\/>/u);
    backgroundBytes[0] = 0;
    const isolated = await build(document);
    assert.deepEqual((await validatePptx(isolated.pptx)).errors, []);
    assert.match(slideXml(isolated), /<a:blipFill>/u);
    assert.doesNotMatch(slideXml(isolated), /445566/u);
  });

  it("creates editable images with media relationships and value-isolated bytes", async () => {
    const source = new Uint8Array(png);
    const node = image({
      id: "photo",
      src: source,
      alt: "Picture",
      fit: "cover",
      style: {
        position: "absolute",
        left: 32,
        top: 48,
        width: 120,
        height: 72,
        border: { color: "#123456", width: 1 },
        opacity: 0.5,
        softEdge: 2,
        boxShadow: {
          color: { color: "#000000", alpha: 0.3 },
          blur: 2,
          distance: 1,
          angle: 45,
        },
      },
    });
    const document = deck({ width: 960, height: 540 }).append(
      slide({ id: "image" }).append(node),
    );

    const first = await build(document);
    const firstFile = await validatePptx(first.pptx);
    assert.deepEqual(firstFile.errors, []);
    const firstObject = first.report.slides[0]?.objects[0];
    assert.equal(firstObject?.kind, "image");
    assert.equal(firstObject?.authorId, "photo");
    assert.deepEqual(firstObject?.box, {
      x: 32,
      y: 48,
      width: 120,
      height: 72,
    });
    const firstImage = firstFile.slides[0]?.objects[0]?.image;
    assert.equal(firstImage?.contentType, "image/png");
    assert.ok(firstImage?.part);
    assert.deepEqual(unzipSync(first.pptx)[firstImage.part], png);
    const firstXml = slideXml(first);
    assert.match(firstXml, /<p:pic>/u);
    assert.match(firstXml, /<a:srcRect t="20000" b="20000"\/>/u);
    assert.match(firstXml, /<a:alphaModFix amt="50000"\/>/u);
    assert.match(firstXml, /<a:outerShdw/u);

    source[0] = 0;
    const isolated = await build(document);
    assert.deepEqual((await validatePptx(isolated.pptx)).errors, []);
    assert.deepEqual(unzipSync(isolated.pptx)[firstImage.part], png);

    node.set({ src: widePng, fit: "fill", crop: { left: 0.25, right: 0.25 } });
    node.style({
      opacity: 0.25,
      border: undefined,
      softEdge: undefined,
      boxShadow: undefined,
    });
    const updated = await build(document);
    const updatedFile = await validatePptx(updated.pptx);
    assert.deepEqual(updatedFile.errors, []);
    const updatedImage = updatedFile.slides[0]?.objects[0]?.image;
    assert.ok(updatedImage?.part);
    assert.deepEqual(unzipSync(updated.pptx)[updatedImage.part], widePng);
    const updatedXml = slideXml(updated);
    assert.match(updatedXml, /<a:srcRect l="25000" r="25000"\/>/u);
    assert.match(updatedXml, /<a:alphaModFix amt="25000"\/>/u);
    assert.doesNotMatch(updatedXml, /<a:outerShdw/u);
    assert.deepEqual(unzipSync(first.pptx)[firstImage.part], png);
  });

  it("keeps native capability registration explicit for trimmed builds", async () => {
    const node = shape({
      id: "trimmed-shape",
      style: { position: "absolute", left: 10, top: 10, width: 40, height: 30 },
    });
    const document = deck({ width: 960, height: 540 }).append(
      slide({ id: "trimmed" }).append(node),
    );

    await assert.rejects(
      async () => emit(createResolver([])(document)),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.ok(
          error.diagnostics.some(({ code }) => code === "missing-capability"),
        );
        return true;
      },
    );
  });

  it("matches the React entry for representative native objects", async () => {
    const commandDocument = deck({ width: 960, height: 540 }).append(
      slide({ id: "native" }).append(
        shape("Shape", {
          id: "shape",
          geometry: "ellipse",
          style: {
            position: "absolute",
            left: 12,
            top: 16,
            width: 120,
            height: 80,
            background: "#336699",
            border: { color: "#000000", width: 1 },
          },
        }),
        image({
          id: "image",
          src: png,
          fit: "cover",
          style: {
            position: "absolute",
            left: 150,
            top: 16,
            width: 80,
            height: 60,
            border: { color: "#AA0000", width: 1 },
          },
        }),
        group({ id: "group" }).append(
          shape({
            id: "group-shape",
            style: {
              position: "absolute",
              left: 260,
              top: 16,
              width: 60,
              height: 40,
              background: "#00AA00",
            },
          }),
        ),
      ),
    );
    const reactElement = createElement(
      Deck,
      { width: 960, height: 540 },
      createElement(
        Slide,
        { id: "native" },
        createElement(
          ReactShape,
          {
            id: "shape",
            geometry: "ellipse",
            style: {
              position: "absolute",
              left: 12,
              top: 16,
              width: 120,
              height: 80,
              background: "#336699",
              border: { color: "#000000", width: 1 },
            },
          },
          "Shape",
        ),
        createElement(ReactImage, {
          id: "image",
          src: png,
          fit: "cover",
          style: {
            position: "absolute",
            left: 150,
            top: 16,
            width: 80,
            height: 60,
            border: { color: "#AA0000", width: 1 },
          },
        }),
        createElement(
          ReactGroup,
          { id: "group" },
          createElement(ReactShape, {
            id: "group-shape",
            style: {
              position: "absolute",
              left: 260,
              top: 16,
              width: 60,
              height: 40,
              background: "#00AA00",
            },
          }),
        ),
      ),
    );

    const commandResult = await build(commandDocument);
    const reactResult = await buildReact(reactElement);
    assert.deepEqual(
      commandResult.report.slides[0]?.objects,
      reactResult.report.slides[0]?.objects,
    );
    assert.deepEqual((await validatePptx(commandResult.pptx)).errors, []);
    assert.deepEqual((await validatePptx(reactResult.pptx)).errors, []);
  });
});
