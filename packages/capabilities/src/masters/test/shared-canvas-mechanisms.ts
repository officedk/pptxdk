import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import { Box } from "@pptxdk/capabilities/box-elements";
import {
  box,
  deck,
  defineSlideLayout,
  font,
  group,
  image,
  shape,
  slide,
  slideLayout,
  slideMaster,
  text,
} from "@pptxdk/capabilities/document";
import { Group } from "@pptxdk/capabilities/group-elements";
import { Image } from "@pptxdk/capabilities/image-elements";
import { Shape } from "@pptxdk/capabilities/shape-elements";
import { Text } from "@pptxdk/capabilities/text-elements";
import { CommonNode, DeckContainer } from "@pptxdk/core";
import { emit } from "@pptxdk/emit";
import { createImporter } from "@pptxdk/import";
import { paint } from "@pptxdk/preview";
import { Deck, Slide, SlideLayout, SlideMaster } from "@pptxdk/react";
import { validatePptx } from "@pptxdk/validation";
import { strFromU8, unzipSync } from "fflate";
import { PNG } from "pngjs";
import { createElement } from "react";

import type { CapabilityScenario } from "../../../test/acceptance-runner.ts";
import type { ExhaustiveMechanismVerificationMap } from "../../../test/mechanism-verification.ts";

function solidPng(r: number, g: number, b: number): Uint8Array {
  const img = new PNG({ width: 8, height: 8 });
  for (let i = 0; i < img.data.length; i += 4) {
    img.data[i] = r;
    img.data[i + 1] = g;
    img.data[i + 2] = b;
    img.data[i + 3] = 255;
  }
  return new Uint8Array(PNG.sync.write(img));
}

function readPngPixelHex(
  pngBytes: Uint8Array | undefined,
  x: number,
  y: number,
): string {
  assert.ok(pngBytes instanceof Uint8Array);
  const decoded = PNG.sync.read(Buffer.from(pngBytes));
  const offset = (y * decoded.width + x) * 4;
  const r = decoded.data[offset] ?? 0;
  const g = decoded.data[offset + 1] ?? 0;
  const b = decoded.data[offset + 2] ?? 0;
  return `#${[r, g, b]
    .map((c) => c.toString(16).padStart(2, "0"))
    .join("")
    .toUpperCase()}`;
}

const jostFont = font("Jost*", [
  {
    weight: 400,
    data: readFileSync(
      new URL("../../../test/assets/fonts/Jost-400-Book.ttf", import.meta.url),
    ),
  },
]);

const logoPng = solidPng(0xef, 0x44, 0x44);

const withRibbonLayout = defineSlideLayout({
  id: "with-ribbon",
  name: "With Ribbon",
  type: "obj",
  preserve: true,
});

const suppressMasterLayout = defineSlideLayout({
  id: "no-master-sp",
  name: "Suppress Master Shapes",
  type: "blank",
  preserve: true,
  showMasterShapes: false,
});

export const masterSharedCanvasScenario: CapabilityScenario = {
  id: "mechanism-master-shared-canvas",
  capability: "masters",
  description:
    "Shared non-placeholder Shape, Image (Logo), Box, Text, and Group objects owned by SlideMaster and SlideLayout with per-part resource registration and showMasterShapes suppression.",
  visualCheckpoints: [
    "Page 1 shows master red Logo (#EF4444), master navy top bar (#1E293B), layout emerald group block (#10B981), and slide amber card (#F59E0B).",
    "Page 2 sets Slide.showMasterShapes=false, hiding both master Logo/bar and layout emerald group block while retaining slide sky card (#0284C7).",
    "Page 3 uses layout showMasterShapes=false, hiding master Logo/bar while retaining layout orange banner (#EA580C).",
  ],
  nativeCheckpoints: [
    "Master Logo Image writes r:embed into slideMaster1.xml.rels; master/layout objects stay in slideMaster1.xml / slideLayout1..2.xml without flattening into slides; slide2 writes showMasterSp='0'.",
  ],
  jsx() {
    return createElement(
      Deck,
      { width: 640, height: 360 },
      createElement(
        SlideMaster,
        { id: "brand-master", name: "Brand Master", preserve: true },
        createElement(Shape, {
          id: "master-bar",
          style: {
            position: "absolute",
            left: 0,
            top: 0,
            width: 640,
            height: 48,
            background: "#1E293B",
          },
        }),
        createElement(Image, {
          id: "logo",
          src: logoPng,
          style: {
            position: "absolute",
            left: 24,
            top: 8,
            width: 32,
            height: 32,
          },
        }),
        createElement(
          Box,
          {
            id: "master-badge",
            style: {
              position: "absolute",
              left: 72,
              top: 10,
              width: 180,
              height: 28,
            },
          },
          createElement(
            Text,
            {
              id: "master-brand-text",
              style: {
                fontFamily: jostFont,
                fontSize: 16,
                color: "#F8FAFC",
              },
            },
            "ACME CORP",
          ),
        ),
        createElement(
          SlideLayout,
          { layout: withRibbonLayout },
          createElement(
            Group,
            { id: "layout-group" },
            createElement(Shape, {
              id: "layout-pill",
              style: {
                position: "absolute",
                left: 400,
                top: 80,
                width: 180,
                height: 80,
                background: "#10B981",
              },
            }),
          ),
        ),
        createElement(
          SlideLayout,
          { layout: suppressMasterLayout },
          createElement(Shape, {
            id: "layout-banner",
            style: {
              position: "absolute",
              left: 40,
              top: 80,
              width: 200,
              height: 80,
              background: "#EA580C",
            },
          }),
        ),
      ),
      createElement(
        Slide,
        { id: "s1-all", layout: withRibbonLayout },
        createElement(Shape, {
          id: "s1-card",
          style: {
            position: "absolute",
            left: 40,
            top: 180,
            width: 160,
            height: 80,
            background: "#F59E0B",
          },
        }),
      ),
      createElement(
        Slide,
        {
          id: "s2-hide-all",
          layout: withRibbonLayout,
          showMasterShapes: false,
        },
        createElement(Shape, {
          id: "s2-card",
          style: {
            position: "absolute",
            left: 40,
            top: 180,
            width: 160,
            height: 80,
            background: "#0284C7",
          },
        }),
      ),
      createElement(Slide, {
        id: "s3-layout-only",
        layout: suppressMasterLayout,
      }),
    );
  },
  document() {
    return deck({ width: 640, height: 360 }).append(
      slideMaster({
        id: "brand-master",
        name: "Brand Master",
        preserve: true,
      }).append(
        shape({
          id: "master-bar",
          style: {
            position: "absolute",
            left: 0,
            top: 0,
            width: 640,
            height: 48,
            background: "#1E293B",
          },
        }),
        image({
          id: "logo",
          src: logoPng,
          style: {
            position: "absolute",
            left: 24,
            top: 8,
            width: 32,
            height: 32,
          },
        }),
        box({
          id: "master-badge",
          style: {
            position: "absolute",
            left: 72,
            top: 10,
            width: 180,
            height: 28,
          },
        }).append(
          text("ACME CORP", {
            id: "master-brand-text",
            style: {
              fontFamily: jostFont,
              fontSize: 16,
              color: "#F8FAFC",
            },
          }),
        ),
        slideLayout(withRibbonLayout).append(
          group({ id: "layout-group" }).append(
            shape({
              id: "layout-pill",
              style: {
                position: "absolute",
                left: 400,
                top: 80,
                width: 180,
                height: 80,
                background: "#10B981",
              },
            }),
          ),
        ),
        slideLayout(suppressMasterLayout).append(
          shape({
            id: "layout-banner",
            style: {
              position: "absolute",
              left: 40,
              top: 80,
              width: 200,
              height: 80,
              background: "#EA580C",
            },
          }),
        ),
      ),
      slide({ id: "s1-all", layout: withRibbonLayout }).append(
        shape({
          id: "s1-card",
          style: {
            position: "absolute",
            left: 40,
            top: 180,
            width: 160,
            height: 80,
            background: "#F59E0B",
          },
        }),
      ),
      slide({
        id: "s2-hide-all",
        layout: withRibbonLayout,
        showMasterShapes: false,
      }).append(
        shape({
          id: "s2-card",
          style: {
            position: "absolute",
            left: 40,
            top: 180,
            width: 160,
            height: 80,
            background: "#0284C7",
          },
        }),
      ),
      slide({ id: "s3-layout-only", layout: suppressMasterLayout }),
    );
  },
  assert({ validation }) {
    assert.equal(validation.errors.length, 0);
    assert.equal(validation.masters?.[0]?.objects?.length, 3);
    assert.equal(validation.masters?.[0]?.layouts[0]?.objects?.length, 1);
    assert.equal(validation.masters?.[0]?.layouts[1]?.objects?.length, 1);
    assert.equal(validation.slides[0]?.objects.length, 1);
    assert.equal(validation.slides[1]?.showMasterShapes, false);
  },
};

export const masterSharedCanvasMechanismVerifications: Pick<
  ExhaustiveMechanismVerificationMap,
  "master-shared-canvas"
> = {
  "master-shared-canvas": {
    scenario: masterSharedCanvasScenario,
    assertNative({ stage, pptx, validation }) {
      assert.equal(validation.errors.length, 0, `${stage} errors`);
      const master = validation.masters?.[0];
      assert.ok(master);
      assert.equal(master.objects?.length, 3, `${stage} master object count`);
      const logoObj = master.objects?.find((obj) => obj.kind === "image");
      assert.ok(logoObj?.image?.part?.startsWith("ppt/media/"));
      assert.equal(master.layouts[0]?.objects?.[0]?.kind, "group");
      assert.equal(master.layouts[1]?.objects?.[0]?.kind, "shape");
      assert.equal(validation.slides[0]?.objects.length, 1);
      assert.equal(validation.slides[1]?.showMasterShapes, false);
      assert.equal(validation.slides[2]?.objects.length, 0);

      const files = unzipSync(pptx);
      const mediaFiles = Object.keys(files).filter((k) =>
        k.startsWith("ppt/media/"),
      );
      assert.equal(mediaFiles.length, 1, `${stage} deduplicated media count`);
      const masterRels = strFromU8(
        files["ppt/slideMasters/_rels/slideMaster1.xml.rels"]!,
      );
      assert.match(masterRels, /relationships\/image/u);
      const slide1Xml = strFromU8(files["ppt/slides/slide1.xml"]!);
      assert.equal(slide1Xml.includes("ACME CORP"), false);
      assert.equal(slide1Xml.includes("<p:pic"), false);
      const slide2Xml = strFromU8(files["ppt/slides/slide2.xml"]!);
      assert.match(slide2Xml, /showMasterSp="0"/u);
    },
    assertPaint({ initial, roundtrip }) {
      assert.equal(initial.pages.length, 3);
      for (const page of initial.pages) {
        assert.deepEqual(page.errors, []);
        assert.ok(page.png instanceof Uint8Array);
      }
      // Page 1: master red Logo (40, 24), master navy bar (300, 20), layout emerald pill (450, 120), slide amber card (100, 220)
      assert.equal(readPngPixelHex(initial.pages[0]!.png, 40, 24), "#EF4444");
      assert.equal(readPngPixelHex(initial.pages[0]!.png, 300, 20), "#1E293B");
      assert.equal(readPngPixelHex(initial.pages[0]!.png, 450, 120), "#10B981");
      assert.equal(readPngPixelHex(initial.pages[0]!.png, 100, 220), "#F59E0B");

      // Page 2: showMasterShapes=false hides master bar/logo and layout pill -> white at (40,24) and (450,120)
      assert.equal(readPngPixelHex(initial.pages[1]!.png, 40, 24), "#FFFFFF");
      assert.equal(readPngPixelHex(initial.pages[1]!.png, 450, 120), "#FFFFFF");
      assert.equal(readPngPixelHex(initial.pages[1]!.png, 100, 220), "#0284C7");

      // Page 3: layout showMasterShapes=false hides master bar/logo (#FFFFFF at 40,24) while retaining layout banner (#EA580C at 100,120)
      assert.equal(readPngPixelHex(initial.pages[2]!.png, 40, 24), "#FFFFFF");
      assert.equal(readPngPixelHex(initial.pages[2]!.png, 100, 120), "#EA580C");
      assert.equal(roundtrip.pages.length, 3);
    },
  },
};

export async function verifyMasterLogoSingleEditPropagates(
  capabilities: Parameters<typeof createImporter>[0],
  resolveDoc: (doc: ReturnType<typeof deck>) => Parameters<typeof emit>[0],
): Promise<void> {
  const layoutDef = defineSlideLayout({
    id: "logo-layout",
    name: "Logo Layout",
    type: "blank",
  });
  const initialDoc = deck({ width: 640, height: 360 }).append(
    slideMaster({ id: "brand", name: "Brand" }).append(
      shape({
        id: "bar",
        style: {
          position: "absolute",
          left: 0,
          top: 0,
          width: 640,
          height: 48,
          background: "#1E293B",
        },
      }),
      image({
        id: "logo",
        src: logoPng,
        style: {
          position: "absolute",
          left: 24,
          top: 8,
          width: 32,
          height: 32,
        },
      }),
      slideLayout(layoutDef),
    ),
    slide({ id: "s1", layout: layoutDef }).append(
      shape({
        id: "s1-card",
        style: {
          position: "absolute",
          left: 40,
          top: 180,
          width: 160,
          height: 80,
          background: "#F59E0B",
        },
      }),
    ),
    slide({ id: "s2", layout: layoutDef }).append(
      shape({
        id: "s2-card",
        style: {
          position: "absolute",
          left: 240,
          top: 180,
          width: 160,
          height: 80,
          background: "#0284C7",
        },
      }),
    ),
  );
  const initialPptx = emit(resolveDoc(initialDoc)).pptx;
  const imported = createImporter(capabilities)(initialPptx);
  assert.deepEqual(imported.report.diagnostics, []);

  // Move master Logo once on the imported document (left: 24 -> 540) and confirm BOTH slides update while local slide cards stay untouched.
  assert.ok(imported.document instanceof DeckContainer);
  const masterNode = imported.document.masters.find(
    (master) => master.props["name"] === "Brand",
  );
  assert.ok(masterNode, "imported document must retain Brand master");
  const logoNode = masterNode.children.find(
    (child): child is CommonNode =>
      child instanceof CommonNode && child.kind === "image",
  );
  assert.ok(logoNode, "imported master must retain Logo Image node");
  logoNode.style({ left: 540 });
  const movedResolution = resolveDoc(
    imported.document as ReturnType<typeof deck>,
  );
  const movedPaint = paint(movedResolution);
  const movedEmit = emit(movedResolution);
  const validation = await validatePptx(movedEmit.pptx);
  assert.deepEqual(validation.errors, []);
  assert.equal(movedPaint.pages.length, 2);
  for (const page of movedPaint.pages) {
    assert.deepEqual(page.errors, []);
    assert.equal(readPngPixelHex(page.png, 556, 24), "#EF4444");
    assert.equal(readPngPixelHex(page.png, 40, 24), "#1E293B");
  }
  assert.equal(readPngPixelHex(movedPaint.pages[0]!.png, 100, 220), "#F59E0B");
  assert.equal(readPngPixelHex(movedPaint.pages[1]!.png, 300, 220), "#0284C7");
}
