import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import { validatePptx } from "@pptxdk/validation";
import { strFromU8, unzipSync } from "fflate";
import * as fontkit from "fontkit";
import { build, BuildError, font } from "pptxdk";
import { Box, Chart, Deck, Group, Image, Slide, Text } from "pptxdk/react";
import {
  box as authorBox,
  chart as authorChart,
  deck as authorDeck,
  group as authorGroup,
  image as authorImage,
  slide as authorSlide,
  text as authorText,
} from "pptxdk/vanilla";
import { describe, it } from "vitest";

const jostBoldBytes = new Uint8Array(
  readFileSync(new URL("./assets/fonts/Jost-700-Bold.ttf", import.meta.url)),
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
    data: jostBoldBytes,
  },
]);
const heroImage = new Uint8Array(
  readFileSync(new URL("./assets/wide-20000x1.png", import.meta.url)),
);
const mixedTheme = {
  name: "Mixed layout",
  colors: {
    dk1: "#131A2A",
    lt1: "#FFFFFF",
    dk2: "#42506B",
    lt2: "#EFF3F8",
    accent1: "#2F7D7A",
    accent2: "#E3B448",
    accent3: "#C85C5C",
  },
  fonts: { major: jost, minor: jost },
};
const revenueChartData = {
  categories: ["Q1", "Q2", "Q3", "Q4"],
  series: [{ name: "Revenue", values: [12, 18, 24, 31] }],
};
const featureBody =
  "Reusable components reflow by declared layout as the available region changes, and following content follows their measured height.";
const sidebarTitle = "Quarterly performance overview";
const sidebarSubtitle =
  "Revenue grew in every region this quarter, led by managed services.";
const heroTitle = "A layered hero that keeps every object editable";
const heroBody =
  "Image, shade, and copy share one aligned region. The measured copy grows upward from the bottom edge.";

function SidebarScene() {
  return (
    <Slide id="mixed-sidebar">
      <Box
        id="frame"
        style={{
          display: "flex",
          flexDirection: "row",
          width: 960,
          height: 540,
          padding: 32,
          gap: 24,
          backgroundColor: "lt2",
        }}
      >
        <Box
          id="sidebar"
          style={{
            display: "flex",
            flexDirection: "column",
            width: 240,
            gap: 12,
            padding: 20,
            backgroundColor: "dk1",
          }}
        >
          <Text
            id="brand"
            style={{
              fontFamily: "major",
              fontSize: 24,
              fontWeight: "bold",
              color: "lt1",
              height: "auto",
            }}
          >
            Northwind
          </Text>
          <Text
            id="nav"
            style={{
              fontFamily: "minor",
              fontSize: 14,
              color: "lt1",
              lineHeight: 22,
              height: "auto",
            }}
          >
            Overview / Pipeline / Reports
          </Text>
        </Box>
        <Box
          id="main"
          style={{
            display: "grid",
            flex: 1,
            minWidth: 0,
            gap: 16,
            gridTemplateColumns: "1fr",
            gridTemplateRows: "auto auto 40px minmax(0, 1fr) auto",
          }}
        >
          <Text
            id="page-title"
            style={{
              fontFamily: "major",
              fontSize: 32,
              fontWeight: "bold",
              color: "dk1",
              height: "auto",
            }}
          >
            {sidebarTitle}
          </Text>
          <Text
            id="page-subtitle"
            style={{
              fontFamily: "minor",
              fontSize: 15,
              color: "dk2",
              lineHeight: 22,
              height: "auto",
            }}
          >
            {sidebarSubtitle}
          </Text>
          <Box
            id="toolbar"
            style={{
              height: 40,
              display: "flex",
              alignItems: "center",
              padding: [0, 12],
              backgroundColor: "accent1",
              borderRadius: 6,
            }}
          >
            <Text
              id="toolbar-label"
              style={{
                fontFamily: "minor",
                fontSize: 13,
                fontWeight: "bold",
                color: "lt1",
                height: "auto",
              }}
            >
              All regions / Last 12 months
            </Text>
          </Box>
          <Chart
            id="revenue-chart"
            style={{ minHeight: 0 }}
            groups={[
              {
                kind: "column",
                data: {
                  kind: "shared-categories",
                  categories: revenueChartData.categories,
                },
                series: revenueChartData.series,
              },
            ]}
          />
          <Text
            id="footer"
            style={{
              fontFamily: "minor",
              fontSize: 12,
              color: "dk2",
              height: "auto",
            }}
          >
            Source: finance system, unaudited.
          </Text>
        </Box>
      </Box>
    </Slide>
  );
}

function HeroScene() {
  return (
    <Slide id="mixed-hero">
      <Box
        id="hero"
        style={{ position: "relative", width: 720, height: 360, margin: 40 }}
      >
        <Image
          id="hero-photo"
          src={heroImage}
          style={{
            position: "absolute",
            left: 0,
            top: 0,
            width: 720,
            height: 360,
            zIndex: 0,
          }}
        />
        <Box
          id="hero-overlay"
          style={{
            position: "absolute",
            left: 0,
            top: 0,
            width: 720,
            height: 360,
            display: "flex",
            flexDirection: "column",
            justifyContent: "flex-end",
            alignItems: "flex-start",
            gap: 10,
            padding: 32,
            backgroundColor: "#0B1220",
            opacity: 0.72,
            zIndex: 1,
          }}
        >
          <Text
            id="hero-kicker"
            style={{
              fontFamily: "minor",
              fontSize: 12,
              fontWeight: "bold",
              color: "accent2",
              letterSpacing: 2,
              height: "auto",
            }}
          >
            FIELD NOTES
          </Text>
          <Text
            id="hero-title"
            style={{
              fontFamily: "major",
              fontSize: 34,
              fontWeight: "bold",
              color: "lt1",
              height: "auto",
              width: "100%",
            }}
          >
            {heroTitle}
          </Text>
          <Text
            id="hero-body"
            style={{
              fontFamily: "minor",
              fontSize: 15,
              color: "lt1",
              lineHeight: 22,
              height: "auto",
              width: "100%",
            }}
          >
            {heroBody}
          </Text>
          <Box
            id="hero-cta"
            style={{
              alignSelf: "flex-start",
              backgroundColor: "accent2",
              padding: [8, 16],
              borderRadius: 12,
            }}
          >
            <Text
              id="hero-cta-label"
              style={{
                fontFamily: "minor",
                fontSize: 13,
                fontWeight: "bold",
                color: "dk1",
                height: "auto",
              }}
            >
              Open report
            </Text>
          </Box>
        </Box>
      </Box>
    </Slide>
  );
}

function FeatureCard({ id }: { readonly id: string }) {
  return (
    <Box
      id={id}
      style={{
        display: "flex",
        flexDirection: "row",
        flexWrap: "wrap",
        gap: 12,
        padding: 16,
        backgroundColor: "lt1",
        borderRadius: 12,
      }}
    >
      <Image id="thumb" src={heroImage} width={120} height={80} />
      <Box
        id="copy"
        style={{
          display: "flex",
          flexDirection: "column",
          gap: 6,
          flexGrow: 1,
          flexBasis: 180,
          minWidth: 0,
        }}
      >
        <Text
          id="title"
          style={{
            fontFamily: "major",
            fontSize: 20,
            fontWeight: "bold",
            color: "dk1",
            height: "auto",
          }}
        >
          Adaptive field report
        </Text>
        <Text
          id="body"
          style={{
            fontFamily: "minor",
            fontSize: 14,
            color: "dk2",
            lineHeight: 20,
            height: "auto",
          }}
        >
          {featureBody}
        </Text>
      </Box>
    </Box>
  );
}

function ReflowScene() {
  return (
    <Slide id="mixed-reflow">
      <Box
        id="reflow-row"
        style={{
          display: "flex",
          flexDirection: "row",
          width: 960,
          height: 540,
          padding: 32,
          gap: 24,
          backgroundColor: "lt2",
        }}
      >
        <Box
          id="narrow-region"
          style={{
            display: "flex",
            flexDirection: "column",
            width: 280,
            gap: 10,
          }}
        >
          <FeatureCard id="narrow-card" />
          <Text
            id="narrow-after"
            style={{
              fontFamily: "minor",
              fontSize: 12,
              color: "dk2",
              height: "auto",
            }}
          >
            After narrow card
          </Text>
        </Box>
        <Box
          id="wide-region"
          style={{
            display: "flex",
            flexDirection: "column",
            width: 460,
            gap: 10,
          }}
        >
          <FeatureCard id="wide-card" />
          <Text
            id="wide-after"
            style={{
              fontFamily: "minor",
              fontSize: 12,
              color: "dk2",
              height: "auto",
            }}
          >
            After wide card
          </Text>
        </Box>
      </Box>
    </Slide>
  );
}

function MixedDeck() {
  return (
    <Deck width={960} height={540} theme={mixedTheme}>
      <SidebarScene />
      <HeroScene />
      <ReflowScene />
    </Deck>
  );
}

function authorSidebarScene() {
  return authorSlide({ id: "mixed-sidebar" }).append(
    authorBox({
      id: "frame",
      style: {
        display: "flex",
        flexDirection: "row",
        width: 960,
        height: 540,
        padding: 32,
        gap: 24,
        backgroundColor: "lt2",
      },
    }).append(
      authorBox({
        id: "sidebar",
        style: {
          display: "flex",
          flexDirection: "column",
          width: 240,
          gap: 12,
          padding: 20,
          backgroundColor: "dk1",
        },
      }).append(
        authorText("Northwind", {
          id: "brand",
          style: {
            fontFamily: "major",
            fontSize: 24,
            fontWeight: "bold",
            color: "lt1",
            height: "auto",
          },
        }),
        authorText("Overview / Pipeline / Reports", {
          id: "nav",
          style: {
            fontFamily: "minor",
            fontSize: 14,
            color: "lt1",
            lineHeight: 22,
            height: "auto",
          },
        }),
      ),
      authorBox({
        id: "main",
        style: {
          display: "grid",
          flex: 1,
          minWidth: 0,
          gap: 16,
          gridTemplateColumns: "1fr",
          gridTemplateRows: "auto auto 40px minmax(0, 1fr) auto",
        },
      }).append(
        authorText(sidebarTitle, {
          id: "page-title",
          style: {
            fontFamily: "major",
            fontSize: 32,
            fontWeight: "bold",
            color: "dk1",
            height: "auto",
          },
        }),
        authorText(sidebarSubtitle, {
          id: "page-subtitle",
          style: {
            fontFamily: "minor",
            fontSize: 15,
            color: "dk2",
            lineHeight: 22,
            height: "auto",
          },
        }),
        authorBox({
          id: "toolbar",
          style: {
            height: 40,
            display: "flex",
            alignItems: "center",
            padding: [0, 12],
            backgroundColor: "accent1",
            borderRadius: 6,
          },
        }).append(
          authorText("All regions / Last 12 months", {
            id: "toolbar-label",
            style: {
              fontFamily: "minor",
              fontSize: 13,
              fontWeight: "bold",
              color: "lt1",
              height: "auto",
            },
          }),
        ),
        authorChart({
          id: "revenue-chart",
          style: { minHeight: 0 },
          groups: [
            {
              kind: "column",
              data: {
                kind: "shared-categories",
                categories: revenueChartData.categories,
              },
              series: revenueChartData.series,
            },
          ],
        }),
        authorText("Source: finance system, unaudited.", {
          id: "footer",
          style: {
            fontFamily: "minor",
            fontSize: 12,
            color: "dk2",
            height: "auto",
          },
        }),
      ),
    ),
  );
}

function authorHeroScene() {
  return authorSlide({ id: "mixed-hero" }).append(
    authorBox({
      id: "hero",
      style: { position: "relative", width: 720, height: 360, margin: 40 },
    }).append(
      authorImage({
        id: "hero-photo",
        src: heroImage,
        style: {
          position: "absolute",
          left: 0,
          top: 0,
          width: 720,
          height: 360,
          zIndex: 0,
        },
      }),
      authorBox({
        id: "hero-overlay",
        style: {
          position: "absolute",
          left: 0,
          top: 0,
          width: 720,
          height: 360,
          display: "flex",
          flexDirection: "column",
          justifyContent: "flex-end",
          alignItems: "flex-start",
          gap: 10,
          padding: 32,
          backgroundColor: "#0B1220",
          opacity: 0.72,
          zIndex: 1,
        },
      }).append(
        authorText("FIELD NOTES", {
          id: "hero-kicker",
          style: {
            fontFamily: "minor",
            fontSize: 12,
            fontWeight: "bold",
            color: "accent2",
            letterSpacing: 2,
            height: "auto",
          },
        }),
        authorText(heroTitle, {
          id: "hero-title",
          style: {
            fontFamily: "major",
            fontSize: 34,
            fontWeight: "bold",
            color: "lt1",
            height: "auto",
            width: "100%",
          },
        }),
        authorText(heroBody, {
          id: "hero-body",
          style: {
            fontFamily: "minor",
            fontSize: 15,
            color: "lt1",
            lineHeight: 22,
            height: "auto",
            width: "100%",
          },
        }),
        authorBox({
          id: "hero-cta",
          style: {
            alignSelf: "flex-start",
            backgroundColor: "accent2",
            padding: [8, 16],
            borderRadius: 12,
          },
        }).append(
          authorText("Open report", {
            id: "hero-cta-label",
            style: {
              fontFamily: "minor",
              fontSize: 13,
              fontWeight: "bold",
              color: "dk1",
              height: "auto",
            },
          }),
        ),
      ),
    ),
  );
}

function authorFeatureCard(id: string) {
  return authorBox({
    id,
    style: {
      display: "flex",
      flexDirection: "row",
      flexWrap: "wrap",
      gap: 12,
      padding: 16,
      backgroundColor: "lt1",
      borderRadius: 12,
    },
  }).append(
    authorImage({ id: "thumb", src: heroImage, width: 120, height: 80 }),
    authorBox({
      id: "copy",
      style: {
        display: "flex",
        flexDirection: "column",
        gap: 6,
        flexGrow: 1,
        flexBasis: 180,
        minWidth: 0,
      },
    }).append(
      authorText("Adaptive field report", {
        id: "title",
        style: {
          fontFamily: "major",
          fontSize: 20,
          fontWeight: "bold",
          color: "dk1",
          height: "auto",
        },
      }),
      authorText(featureBody, {
        id: "body",
        style: {
          fontFamily: "minor",
          fontSize: 14,
          color: "dk2",
          lineHeight: 20,
          height: "auto",
        },
      }),
    ),
  );
}

function authorReflowScene() {
  return authorSlide({ id: "mixed-reflow" }).append(
    authorBox({
      id: "reflow-row",
      style: {
        display: "flex",
        flexDirection: "row",
        width: 960,
        height: 540,
        padding: 32,
        gap: 24,
        backgroundColor: "lt2",
      },
    }).append(
      authorBox({
        id: "narrow-region",
        style: {
          display: "flex",
          flexDirection: "column",
          width: 280,
          gap: 10,
        },
      }).append(
        authorFeatureCard("narrow-card"),
        authorText("After narrow card", {
          id: "narrow-after",
          style: {
            fontFamily: "minor",
            fontSize: 12,
            color: "dk2",
            height: "auto",
          },
        }),
      ),
      authorBox({
        id: "wide-region",
        style: {
          display: "flex",
          flexDirection: "column",
          width: 460,
          gap: 10,
        },
      }).append(
        authorFeatureCard("wide-card"),
        authorText("After wide card", {
          id: "wide-after",
          style: {
            fontFamily: "minor",
            fontSize: 12,
            color: "dk2",
            height: "auto",
          },
        }),
      ),
    ),
  );
}

function authorMixedDeck() {
  return authorDeck({ width: 960, height: 540, theme: mixedTheme }).append(
    authorSidebarScene(),
    authorHeroScene(),
    authorReflowScene(),
  );
}

function resultSummary(result: Awaited<ReturnType<typeof build>>) {
  return result.report.slides.map((slide) =>
    (slide.objects ?? []).map((object) => ({
      kind: object.kind,
      authorId: object.authorId,
      box: object.box,
      text: object.text,
    })),
  );
}

function nativeSummary(file: Awaited<ReturnType<typeof validatePptx>>) {
  return file.slides.map((slide) =>
    (slide.objects ?? []).map((object) => ({
      kind: object.kind,
      box: {
        x: object.boxEmu.x / 9525,
        y: object.boxEmu.y / 9525,
        width: object.boxEmu.width / 9525,
        height: object.boxEmu.height / 9525,
      },
      text: object.text?.text,
    })),
  );
}

function objectBox(
  result: Awaited<ReturnType<typeof build>>,
  slideIndex: number,
  authorId: string,
) {
  return result.report.slides[slideIndex]?.objects.find(
    (object) => object.authorId === authorId,
  )?.box;
}

describe("mixed layout, layering, and component reflow", () => {
  it("builds the same mixed intent through React and document authoring", async () => {
    const reactResult = await build(<MixedDeck />);
    const documentResult = await build(authorMixedDeck());

    assert.deepEqual(reactResult.report.diagnostics, []);
    assert.deepEqual(documentResult.report.diagnostics, []);
    assert.deepEqual(resultSummary(documentResult), resultSummary(reactResult));

    const sidebar = reactResult.report.slides[0];
    assert.deepEqual(
      sidebar?.layouts?.find(({ authorId }) => authorId === "main")?.box,
      { x: 296, y: 32, width: 632, height: 476 },
    );
    assert.deepEqual(objectBox(reactResult, 0, "sidebar"), {
      x: 32,
      y: 32,
      width: 240,
      height: 476,
    });
    assert.deepEqual(objectBox(reactResult, 0, "page-title"), {
      x: 296,
      y: 32,
      width: 632,
      height: 46.24,
    });
    assert.deepEqual(objectBox(reactResult, 0, "toolbar"), {
      x: 296,
      y: 132.24,
      width: 632,
      height: 40,
    });
    assert.deepEqual(objectBox(reactResult, 0, "revenue-chart"), {
      x: 296,
      y: 188.24,
      width: 632,
      height: 286.41,
    });
    assert.deepEqual(objectBox(reactResult, 0, "footer"), {
      x: 296,
      y: 490.65,
      width: 632,
      height: 17.35,
    });
    assert.ok(
      Math.abs(
        (objectBox(reactResult, 0, "revenue-chart")?.y ?? 0) +
          (objectBox(reactResult, 0, "revenue-chart")?.height ?? 0) +
          16 -
          (objectBox(reactResult, 0, "footer")?.y ?? 0),
      ) < 0.01,
    );

    assert.deepEqual(
      reactResult.report.slides[1]?.objects.map(({ authorId }) => authorId),
      [
        "hero-photo",
        "hero-overlay",
        "hero-kicker",
        "hero-title",
        "hero-body",
        "hero-cta",
        "hero-cta-label",
      ],
    );
    assert.deepEqual(objectBox(reactResult, 1, "hero-photo"), {
      x: 40,
      y: 40,
      width: 720,
      height: 360,
    });
    assert.deepEqual(objectBox(reactResult, 1, "hero-overlay"), {
      x: 40,
      y: 40,
      width: 720,
      height: 360,
    });
    assert.deepEqual(objectBox(reactResult, 1, "hero-title"), {
      x: 72,
      y: 170.95,
      width: 656,
      height: 98.27,
    });
    assert.deepEqual(objectBox(reactResult, 1, "hero-body"), {
      x: 72,
      y: 279.21,
      width: 656,
      height: 44,
    });
    assert.equal(
      (objectBox(reactResult, 1, "hero-cta")?.y ?? 0) +
        (objectBox(reactResult, 1, "hero-cta")?.height ?? 0),
      368,
    );

    const reflow = reactResult.report.slides[2]?.objects ?? [];
    const cardBox = (authorId: string) =>
      reflow.find((object) => object.authorId === authorId)?.box;
    assert.deepEqual(cardBox("narrow-card"), {
      x: 32,
      y: 32,
      width: 280,
      height: 238.91,
    });
    assert.deepEqual(cardBox("wide-card"), {
      x: 336,
      y: 32,
      width: 460,
      height: 126.91,
    });
    assert.equal(
      (cardBox("narrow-card")?.height ?? 0) >
        (cardBox("wide-card")?.height ?? 0),
      true,
    );
    const bodies = reflow.filter(({ authorId }) => authorId === "body");
    assert.deepEqual(
      bodies.map(({ box }) => box),
      [
        { x: 48, y: 174.91, width: 248, height: 80 },
        { x: 484, y: 82.91, width: 296, height: 60 },
      ],
    );
    assert.ok((bodies[0]?.box?.height ?? 0) > (bodies[1]?.box?.height ?? 0));
    assert.deepEqual(
      [
        reflow.find(({ authorId }) => authorId === "narrow-after")?.box?.y,
        reflow.find(({ authorId }) => authorId === "wide-after")?.box?.y,
      ],
      [280.91, 168.91],
    );
    assert.equal(
      reflow.filter(({ authorId }) => authorId === "thumb").length,
      2,
    );

    const reactFile = await validatePptx(reactResult.pptx);
    const documentFile = await validatePptx(documentResult.pptx);
    assert.deepEqual(reactFile.errors, []);
    assert.deepEqual(documentFile.errors, []);
    assert.deepEqual(nativeSummary(documentFile), nativeSummary(reactFile));
    const nativeKinds = reactFile.slides.map((slide) =>
      (slide.objects ?? []).map(({ kind }) => kind),
    );
    assert.equal(nativeKinds[0]?.filter((kind) => kind === "chart").length, 1);
    assert.deepEqual(nativeKinds[1]?.slice(0, 2), ["image", "shape"]);
    assert.equal(nativeKinds[2]?.filter((kind) => kind === "image").length, 2);
    assert.equal(
      reactFile.slides[0]?.objects.find(({ kind }) => kind === "chart")?.chart
        ?.series[0]?.values?.[3],
      31,
    );
    assert.equal(
      reactFile.slides[1]?.objects
        .flatMap((object) => object.text?.text ?? [])
        .includes(heroTitle),
      true,
    );
    assert.equal(
      reactFile.slides[2]?.objects.filter(
        ({ text }) => text?.text === featureBody,
      ).length,
      2,
    );
  }, 45_000);

  it("reports real overflow and text fitting failures without rewriting content", async () => {
    const longText =
      "Measured copy that cannot fit the declared height keeps every word in the native text body instead of shrinking or paginating the presentation content.";
    const result = await build(
      <Deck width={960} height={540} theme={mixedTheme}>
        <Slide id="mixed-overset">
          <Text
            id="overset-copy"
            style={{
              width: 240,
              height: 40,
              fontFamily: "minor",
              fontSize: 18,
              color: "dk1",
            }}
          >
            {longText}
          </Text>
        </Slide>
        <Slide id="mixed-overflow">
          <Text
            id="off-canvas-copy"
            style={{
              position: "absolute",
              left: 900,
              top: 500,
              width: 200,
              height: 60,
              fontFamily: "minor",
              fontSize: 16,
              color: "dk1",
            }}
          >
            Off-canvas copy remains editable.
          </Text>
        </Slide>
      </Deck>,
    );

    assert.deepEqual(
      result.report.diagnostics.map(({ path, code }) => ({ path, code })),
      [
        {
          path: "deck/slide[mixed-overset]/text[overset-copy]",
          code: "overset",
        },
        {
          path: "deck/slide[mixed-overflow]/text[off-canvas-copy]",
          code: "overflow",
        },
      ],
    );
    assert.deepEqual(objectBox(result, 0, "overset-copy"), {
      x: 0,
      y: 0,
      width: 240,
      height: 40,
    });
    assert.deepEqual(objectBox(result, 1, "off-canvas-copy"), {
      x: 900,
      y: 500,
      width: 200,
      height: 60,
    });

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.equal(file.slides[0]?.objects[0]?.text?.text, longText);
    assert.deepEqual(
      {
        x: (file.slides[1]?.objects[0]?.boxEmu.x ?? 0) / 9525,
        y: (file.slides[1]?.objects[0]?.boxEmu.y ?? 0) / 9525,
      },
      { x: 900, y: 500 },
    );
    const slideXml = strFromU8(
      unzipSync(result.pptx)["ppt/slides/slide1.xml"]!,
    );
    assert.match(slideXml, /<a:noAutofit\/>/u);
    assert.doesNotMatch(slideXml, /normAutofit|spAutoFit/u);

    await assert.rejects(
      build(
        <Deck width={960} height={540}>
          <Slide id="mixed-invalid">
            <Box
              id="invalid-static-offset"
              style={{ left: 20, width: 100, height: 40 }}
            />
          </Slide>
        </Deck>,
      ),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics[0]?.code === "invalid-layout-context",
    );
  });

  it("keeps intrinsic content width in the native file while constrained text wraps and flows", async () => {
    const label = "Open report";
    const labelStyle = {
      fontFamily: "minor",
      fontSize: 13,
      fontWeight: "bold",
      color: "dk1",
      height: "auto",
    } as const;
    const columnStyle = {
      display: "flex",
      flexDirection: "column",
      alignItems: "flex-start",
      width: 400,
      padding: 20,
      gap: 6,
    } as const;

    const reactResult = await build(
      <Deck width={400} height={300} theme={mixedTheme}>
        <Slide id="intrinsic-precision">
          <Box id="column" style={columnStyle}>
            <Box
              id="cta"
              style={{
                backgroundColor: "accent2",
                padding: [8, 16],
                borderRadius: 12,
              }}
            >
              <Text id="cta-label" style={labelStyle}>
                {label}
              </Text>
            </Box>
            <Box
              id="after-cta"
              style={{ width: 20, height: 10, backgroundColor: "accent1" }}
            />
            <Box id="narrow" style={{ width: 60, backgroundColor: "lt1" }}>
              <Text id="narrow-label" style={labelStyle}>
                {label}
              </Text>
            </Box>
            <Box
              id="after-narrow"
              style={{ width: 20, height: 10, backgroundColor: "accent1" }}
            />
            <Text
              id="overset-label"
              style={{ ...labelStyle, width: 60, height: 10 }}
            >
              {label}
            </Text>
          </Box>
        </Slide>
      </Deck>,
    );
    const documentResult = await build(
      authorDeck({ width: 400, height: 300, theme: mixedTheme }).append(
        authorSlide({ id: "intrinsic-precision" }).append(
          authorBox({ id: "column", style: columnStyle }).append(
            authorBox({
              id: "cta",
              style: {
                backgroundColor: "accent2",
                padding: [8, 16],
                borderRadius: 12,
              },
            }).append(
              authorText(label, { id: "cta-label", style: labelStyle }),
            ),
            authorBox({
              id: "after-cta",
              style: { width: 20, height: 10, backgroundColor: "accent1" },
            }),
            authorBox({
              id: "narrow",
              style: { width: 60, backgroundColor: "lt1" },
            }).append(
              authorText(label, { id: "narrow-label", style: labelStyle }),
            ),
            authorBox({
              id: "after-narrow",
              style: { width: 20, height: 10, backgroundColor: "accent1" },
            }),
            authorText(label, {
              id: "overset-label",
              style: { ...labelStyle, width: 60, height: 10 },
            }),
          ),
        ),
      ),
    );

    const face = fontkit.create(Buffer.from(jostBoldBytes));
    if (face.type !== "TTF") throw new Error("expected a static TrueType font");
    const measuredWidth = Math.round(
      (face.layout(label).advanceWidth / face.unitsPerEm) * 13 * 9525,
    );
    // Jost Bold 13px "Open report" shaped advances: natural 717071 EMU, native 1/8pt
    // grid 717550 EMU, and 9.8pt at the 1/20pt grid 721360 EMU. The last is the
    // conservative candidate the native file must cover.
    const nativeQuantizedWidth = 721360;

    for (const [entry, result] of [
      ["react", reactResult],
      ["document", documentResult],
    ] as const) {
      assert.deepEqual(
        result.report.diagnostics.map(({ code, path }) => ({ code, path })),
        [
          {
            code: "overset",
            path: "deck/slide[intrinsic-precision]/box[column]/text[overset-label]",
          },
        ],
        entry,
      );
      const file = await validatePptx(result.pptx);
      assert.deepEqual(file.errors, [], entry);
      const objectIds = new Map(
        (result.report.slides[0]?.objects ?? []).flatMap((object) =>
          object.authorId === undefined
            ? []
            : [[object.authorId, object.id] as const],
        ),
      );
      const nativeBox = (authorId: string) => {
        const id = objectIds.get(authorId);
        assert.ok(id !== undefined, `${entry}: missing object ${authorId}`);
        const object = file.slides[0]?.objects.find(
          (candidate) => candidate.id === id,
        );
        assert.ok(object, `${entry}: missing native object ${authorId}`);
        return object.boxEmu;
      };

      const labelWidth = nativeBox("cta-label").width;
      assert.ok(
        labelWidth >= measuredWidth,
        `${entry}: intrinsic width ${labelWidth} EMU is shorter than measured ${measuredWidth} EMU`,
      );
      assert.ok(
        labelWidth >= nativeQuantizedWidth,
        `${entry}: intrinsic width ${labelWidth} EMU is shorter than the native-quantized candidate ${nativeQuantizedWidth} EMU`,
      );
      assert.ok(
        Math.abs(nativeBox("cta").width - labelWidth - 32 * 9525) <= 1,
        `${entry}: CTA parent lost its 32px horizontal padding`,
      );
      assert.ok(
        Math.abs(
          nativeBox("cta").height - nativeBox("cta-label").height - 16 * 9525,
        ) <= 1,
        `${entry}: CTA parent lost its 16px vertical padding`,
      );
      assert.ok(
        Math.abs(
          nativeBox("after-cta").y -
            (nativeBox("cta").y + nativeBox("cta").height + 6 * 9525),
        ) <= 1,
        `${entry}: content after the CTA did not follow its measured height`,
      );

      assert.equal(nativeBox("narrow").width, 60 * 9525, entry);
      assert.ok(
        nativeBox("narrow-label").height > nativeBox("cta-label").height,
        `${entry}: text constrained to a 60px parent did not wrap to a taller box`,
      );
      assert.ok(
        nativeBox("after-narrow").y >=
          nativeBox("narrow").y + nativeBox("narrow").height,
        `${entry}: content after the narrow card overlapped it`,
      );
      assert.equal(nativeBox("overset-label").height, 10 * 9525, entry);
    }
  });

  it("writes group extents from the union of precise member geometry", async () => {
    const memberWidth = 74.39453125;
    const memberWidthEmu = Math.round(memberWidth * 9525);
    const reactResult = await build(
      <Deck width={400} height={300} theme={mixedTheme}>
        <Slide id="group-precision">
          <Group id="top-group">
            <Box
              id="top-member"
              style={{
                width: memberWidth,
                height: 20,
                backgroundColor: "accent1",
              }}
            />
          </Group>
          <Group id="nested-outer">
            <Group id="nested-inner">
              <Box
                id="nested-member"
                style={{
                  width: memberWidth,
                  height: 20,
                  backgroundColor: "accent2",
                }}
              />
            </Group>
          </Group>
          <Group id="fixed-group" style={{ width: 100, height: 20 }}>
            <Box
              id="fixed-member"
              style={{
                width: memberWidth,
                height: 20,
                backgroundColor: "accent3",
              }}
            />
          </Group>
        </Slide>
      </Deck>,
    );
    const documentResult = await build(
      authorDeck({ width: 400, height: 300, theme: mixedTheme }).append(
        authorSlide({ id: "group-precision" }).append(
          authorGroup({ id: "top-group" }).append(
            authorBox({
              id: "top-member",
              style: {
                width: memberWidth,
                height: 20,
                backgroundColor: "accent1",
              },
            }),
          ),
          authorGroup({ id: "nested-outer" }).append(
            authorGroup({ id: "nested-inner" }).append(
              authorBox({
                id: "nested-member",
                style: {
                  width: memberWidth,
                  height: 20,
                  backgroundColor: "accent2",
                },
              }),
            ),
          ),
          authorGroup({
            id: "fixed-group",
            style: { width: 100, height: 20 },
          }).append(
            authorBox({
              id: "fixed-member",
              style: {
                width: memberWidth,
                height: 20,
                backgroundColor: "accent3",
              },
            }),
          ),
        ),
      ),
    );

    for (const [entry, result] of [
      ["react", reactResult],
      ["document", documentResult],
    ] as const) {
      assert.deepEqual(result.report.diagnostics, [], entry);
      const file = await validatePptx(result.pptx);
      assert.deepEqual(file.errors, [], entry);
      const reportIds = new Map<string, number>();
      const collectReportIds = (
        objects: readonly {
          readonly authorId?: string;
          readonly id: number;
          readonly members?: readonly unknown[];
        }[],
      ) => {
        for (const object of objects) {
          if (object.authorId !== undefined)
            reportIds.set(object.authorId, object.id);
          if (object.members !== undefined)
            collectReportIds(
              object.members as readonly {
                readonly authorId?: string;
                readonly id: number;
                readonly members?: readonly unknown[];
              }[],
            );
        }
      };
      collectReportIds(result.report.slides[0]?.objects ?? []);
      const nativeObjects = new Map<
        number,
        { readonly boxEmu: { readonly width: number; readonly height: number } }
      >();
      const collectNative = (
        objects: readonly {
          readonly id: number;
          readonly boxEmu: { readonly width: number; readonly height: number };
          readonly members?: readonly unknown[];
        }[],
      ) => {
        for (const object of objects) {
          nativeObjects.set(object.id, object);
          if (object.members !== undefined)
            collectNative(
              object.members as readonly {
                readonly id: number;
                readonly boxEmu: {
                  readonly width: number;
                  readonly height: number;
                };
                readonly members?: readonly unknown[];
              }[],
            );
        }
      };
      collectNative(file.slides[0]?.objects ?? []);
      const nativeBox = (authorId: string) => {
        const id = reportIds.get(authorId);
        assert.ok(id !== undefined, `${entry}: missing object ${authorId}`);
        const object = nativeObjects.get(id);
        assert.ok(object, `${entry}: missing native object ${authorId}`);
        return object.boxEmu;
      };

      assert.equal(nativeBox("top-member").width, memberWidthEmu, entry);
      assert.equal(nativeBox("top-group").width, memberWidthEmu, entry);
      assert.equal(nativeBox("nested-member").width, memberWidthEmu, entry);
      assert.equal(nativeBox("nested-inner").width, memberWidthEmu, entry);
      assert.equal(nativeBox("nested-outer").width, memberWidthEmu, entry);
      assert.equal(nativeBox("top-group").height, 20 * 9525, entry);
      assert.equal(nativeBox("fixed-group").width, 100 * 9525, entry);
      assert.equal(nativeBox("fixed-member").width, memberWidthEmu, entry);
    }
  });

  it("uses the rounded-up tenth-point effective size only for non-integral serialized sizes", async () => {
    const cases = [
      { id: "fs-9_74", points: 9.74, hundredths: 974, widthEmu: 721360 },
      { id: "fs-9_75", points: 9.75, hundredths: 975, widthEmu: 721360 },
      { id: "fs-9_8", points: 9.8, hundredths: 980, widthEmu: 721360 },
      { id: "fs-9_9", points: 9.9, hundredths: 990, widthEmu: 728663 },
      { id: "fs-10", points: 10, hundredths: 1000, widthEmu: 735457 },
    ] as const;
    const textStyle = {
      fontFamily: "minor",
      fontWeight: "bold",
      height: "auto",
    } as const;
    const reactResult = await build(
      <Deck width={400} height={400} theme={mixedTheme}>
        <Slide id="native-font-size">
          <Box
            id="column"
            style={{
              display: "flex",
              flexDirection: "column",
              alignItems: "flex-start",
              gap: 4,
              width: 400,
            }}
          >
            {cases.map(({ id, points }) => (
              <Text
                key={id}
                id={id}
                style={{ ...textStyle, fontSize: `${points}pt` }}
              >
                Open report
              </Text>
            ))}
            <Box id="narrow" style={{ width: 60, backgroundColor: "lt1" }}>
              <Text id="fs-narrow" style={{ ...textStyle, fontSize: "9.74pt" }}>
                Open report
              </Text>
            </Box>
          </Box>
        </Slide>
      </Deck>,
    );
    const documentResult = await build(
      authorDeck({ width: 400, height: 400, theme: mixedTheme }).append(
        authorSlide({ id: "native-font-size" }).append(
          authorBox({
            id: "column",
            style: {
              display: "flex",
              flexDirection: "column",
              alignItems: "flex-start",
              gap: 4,
              width: 400,
            },
          }).append(
            ...cases.map(({ id, points }) =>
              authorText("Open report", {
                id,
                style: { ...textStyle, fontSize: `${points}pt` },
              }),
            ),
            authorBox({
              id: "narrow",
              style: { width: 60, backgroundColor: "lt1" },
            }).append(
              authorText("Open report", {
                id: "fs-narrow",
                style: { ...textStyle, fontSize: "9.74pt" },
              }),
            ),
          ),
        ),
      ),
    );

    for (const [entry, result] of [
      ["react", reactResult],
      ["document", documentResult],
    ] as const) {
      assert.deepEqual(result.report.diagnostics, [], entry);
      const file = await validatePptx(result.pptx);
      assert.deepEqual(file.errors, [], entry);
      const objectIds = new Map(
        (result.report.slides[0]?.objects ?? []).flatMap((object) =>
          object.authorId === undefined
            ? []
            : [[object.authorId, object.id] as const],
        ),
      );
      const nativeObject = (authorId: string) => {
        const id = objectIds.get(authorId);
        assert.ok(id !== undefined, `${entry}: missing object ${authorId}`);
        const object = file.slides[0]?.objects.find(
          (candidate) => candidate.id === id,
        );
        assert.ok(object, `${entry}: missing native object ${authorId}`);
        return object;
      };
      const slideXml = strFromU8(
        unzipSync(result.pptx)["ppt/slides/slide1.xml"]!,
      );
      const shapeXml = (id: number) => {
        const match = slideXml.match(
          new RegExp(
            `<p:sp>(?:(?!</p:sp>).)*?<p:cNvPr id="${id}"[^>]*>(?:(?!</p:sp>).)*?</p:sp>`,
            "s",
          ),
        );
        assert.ok(match, `${entry}: missing shape ${id}`);
        return match[0];
      };

      for (const { id, hundredths, widthEmu } of cases) {
        const object = nativeObject(id);
        assert.equal(object.boxEmu.width, widthEmu, `${entry}: ${id} width`);
        const size = shapeXml(object.id).match(/<a:rPr sz="(\d+)"/u);
        assert.equal(Number(size?.[1]), hundredths, `${entry}: ${id} sz`);
      }
      assert.equal(nativeObject("fs-narrow").boxEmu.width, 60 * 9525, entry);
      assert.ok(
        nativeObject("fs-narrow").boxEmu.height >
          nativeObject("fs-9_74").boxEmu.height,
        `${entry}: narrow text did not wrap`,
      );
    }
  });
});
