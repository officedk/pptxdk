import assert from "node:assert/strict";

import { shapeCapability } from "@pptxdk/capabilities/shape";
import { validatePptx } from "@pptxdk/validation";
import { strFromU8, unzipSync } from "fflate";
import {
  after,
  build,
  createMotionScopeHandle,
  emit,
  fade,
  transitions,
  type MotionObservation,
} from "pptxdk";
import { createResolver } from "pptxdk/build";
import {
  Box,
  Cell,
  Chart,
  Deck,
  expandDocument,
  Group,
  Image,
  MotionScope,
  P,
  Row,
  Shape,
  Slide,
  Span,
  Step,
  Table,
  Text,
  useMotionScope,
} from "pptxdk/react";
import {
  box,
  cell,
  chart,
  deck,
  group,
  image,
  motionScope,
  paragraph,
  row,
  shape,
  slide,
  span,
  step,
  table,
  text,
} from "pptxdk/vanilla";
import { describe, it } from "vitest";

const png = new Uint8Array(
  Buffer.from(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/n1sAAAAASUVORK5CYII=",
    "base64",
  ),
);

function plainDeck() {
  const lead = createMotionScopeHandle();
  const follower = createMotionScopeHandle();
  const ready = lead.cue("ready");
  return deck({
    width: 960,
    height: 540,
    theme: { colors: { ink: "#123456", accent1: "#2F6F8F" } },
  }).append(
    slide({
      id: "overview",
      background: "#F7F9FB",
      theme: { colors: { ink: "#0B5C7A" } },
    }).append(
      box({
        id: "layout",
        style: {
          width: 760,
          height: 180,
          padding: 12,
          backgroundColor: "#FFFFFF",
        },
      }).append(
        text(
          paragraph(
            span("Same ", { style: { fontWeight: "bold" } }),
            span("intent", { style: { color: "ink" } }),
          ),
          {
            id: "headline",
            x: 12,
            y: 12,
            width: 720,
            height: 48,
            style: { fontSize: 24 },
          },
        ),
        shape({
          id: "panel",
          x: 12,
          y: 72,
          width: 240,
          height: 80,
          text: "Native panel",
          fill: "#2F6F8F",
        }),
      ),
      image({
        id: "logo",
        src: png,
        x: 800,
        y: 20,
        width: 100,
        height: 80,
        fit: "cover",
      }),
      group({ id: "badge" }).append(
        shape({
          id: "badge-shape",
          x: 800,
          y: 120,
          width: 100,
          height: 60,
          text: "Badge",
          fill: "accent1",
        }),
      ),
      table({
        id: "metrics",
        x: 20,
        y: 220,
        width: 420,
        columns: [140, "1fr"],
        style: { fontSize: 14 },
      }).append(
        row({ style: { height: 28 } }).append(cell("Metric"), cell("Value")),
        row({ style: { height: 28 } }).append(cell("North"), cell("12")),
      ),
      chart({
        id: "revenue",
        x: 480,
        y: 220,
        width: 400,
        height: 220,
        groups: [
          {
            kind: "column",
            data: { kind: "shared-categories", categories: ["Q1", "Q2"] },
            series: [
              { name: "North", values: [12, 18] },
              { name: "South", values: [9, 15] },
            ],
          },
        ],
      }),
      step({ n: 1 }).append(
        shape({
          id: "fade-in",
          x: 20,
          y: 460,
          width: 120,
          height: 40,
          enter: fade({ duration: 200 }),
        }),
        motionScope(lead, { publishEnd: ready }).append(
          shape({
            id: "scope-lead",
            x: 160,
            y: 460,
            width: 120,
            height: 40,
            enter: fade({ duration: 100 }),
          }),
        ),
        motionScope(follower, { startAt: after(ready, 20) }).append(
          shape({
            id: "scope-follow",
            x: 300,
            y: 460,
            width: 120,
            height: 40,
            enter: fade({ duration: 80 }),
          }),
        ),
      ),
      shape({
        id: "next",
        x: 440,
        y: 460,
        width: 120,
        height: 40,
        text: "Details",
        fill: "accent1",
        link: { id: "details" },
      }),
    ),
    slide({ id: "details", transition: transitions.fade() }).append(
      text("Details", {
        id: "details-title",
        x: 20,
        y: 20,
        width: 400,
        height: 48,
        style: { fontSize: 24, color: "ink" },
      }),
    ),
  );
}

function ReactDeck() {
  const lead = useMotionScope();
  const follower = useMotionScope();
  return (
    <Deck
      width={960}
      height={540}
      theme={{ colors: { ink: "#123456", accent1: "#2F6F8F" } }}
    >
      <Slide
        id="overview"
        background="#F7F9FB"
        theme={{ colors: { ink: "#0B5C7A" } }}
      >
        <Box
          id="layout"
          style={{
            width: 760,
            height: 180,
            padding: 12,
            backgroundColor: "#FFFFFF",
          }}
        >
          <Text
            id="headline"
            x={12}
            y={12}
            width={720}
            height={48}
            style={{ fontSize: 24 }}
          >
            <P>
              <Span style={{ fontWeight: "bold" }}>Same </Span>
              <Span style={{ color: "ink" }}>intent</Span>
            </P>
          </Text>
          <Shape
            id="panel"
            x={12}
            y={72}
            width={240}
            height={80}
            text="Native panel"
            fill="#2F6F8F"
          />
        </Box>
        <Image
          id="logo"
          src={png}
          x={800}
          y={20}
          width={100}
          height={80}
          fit="cover"
        />
        <Group id="badge">
          <Shape
            id="badge-shape"
            x={800}
            y={120}
            width={100}
            height={60}
            text="Badge"
            fill="accent1"
          />
        </Group>
        <Table
          id="metrics"
          x={20}
          y={220}
          width={420}
          columns={[140, "1fr"]}
          style={{ fontSize: 14 }}
        >
          <Row style={{ height: 28 }}>
            <Cell>Metric</Cell>
            <Cell>Value</Cell>
          </Row>
          <Row style={{ height: 28 }}>
            <Cell>North</Cell>
            <Cell>12</Cell>
          </Row>
        </Table>
        <Chart
          id="revenue"
          x={480}
          y={220}
          width={400}
          height={220}
          groups={[
            {
              kind: "column",
              data: { kind: "shared-categories", categories: ["Q1", "Q2"] },
              series: [
                { name: "North", values: [12, 18] },
                { name: "South", values: [9, 15] },
              ],
            },
          ]}
        />
        <Step n={1}>
          <Shape
            id="fade-in"
            x={20}
            y={460}
            width={120}
            height={40}
            enter={fade({ duration: 200 })}
          />
          <MotionScope scope={lead} publishEnd={lead.cue("ready")}>
            <Shape
              id="scope-lead"
              x={160}
              y={460}
              width={120}
              height={40}
              enter={fade({ duration: 100 })}
            />
          </MotionScope>
          <MotionScope scope={follower} startAt={after(lead.cue("ready"), 20)}>
            <Shape
              id="scope-follow"
              x={300}
              y={460}
              width={120}
              height={40}
              enter={fade({ duration: 80 })}
            />
          </MotionScope>
        </Step>
        <Shape
          id="next"
          x={440}
          y={460}
          width={120}
          height={40}
          text="Details"
          fill="accent1"
          link={{ id: "details" }}
        />
      </Slide>
      <Slide id="details" transition={transitions.fade()}>
        <Text
          id="details-title"
          x={20}
          y={20}
          width={400}
          height={48}
          style={{ fontSize: 24, color: "ink" }}
        >
          Details
        </Text>
      </Slide>
    </Deck>
  );
}

function slideXml(result: { readonly pptx: Uint8Array }, index = 1): string {
  return strFromU8(unzipSync(result.pptx)[`ppt/slides/slide${index}.xml`]!);
}

function motionProjection(result: Awaited<ReturnType<typeof build>>) {
  return result.report.slides.map(({ capabilities }) => {
    const motion = capabilities?.["motion"] as MotionObservation | undefined;
    return motion === undefined
      ? undefined
      : {
          steps: motion.steps.map(({ trigger, number, effects }) => ({
            trigger,
            number,
            effects: effects.map(
              ({ target, preset, startMs, endMs, durationMs }) => ({
                target: target.authorId,
                preset,
                startMs,
                endMs,
                durationMs,
              }),
            ),
          })),
          scopes: motion.scopes?.map(
            ({ mode, delayMs, startMs, endMs, completion }) => ({
              mode,
              delayMs,
              startMs,
              endMs,
              completion,
            }),
          ),
        };
  });
}

function readerProjection(result: Awaited<ReturnType<typeof validatePptx>>) {
  return result.slides.map(({ id, objects, animations, transition }) => ({
    id,
    objects: objects.map(({ kind, boxEmu, text, link, table, members }) => ({
      kind,
      boxEmu,
      text,
      link,
      table,
      members,
    })),
    animations,
    transition,
  }));
}

describe("generation closure", () => {
  it("builds the same authoring intent through the default plain and React entries", async () => {
    const [plainResult, reactResult] = await Promise.all([
      build(plainDeck()),
      build(<ReactDeck />),
    ]);

    assert.deepEqual(plainResult.report.slides, reactResult.report.slides);
    assert.deepEqual(
      motionProjection(plainResult),
      motionProjection(reactResult),
    );
    assert.deepEqual(
      plainResult.report.diagnostics,
      reactResult.report.diagnostics,
    );
    assert.equal(
      plainResult.report.diagnostics.some(
        ({ severity }) => severity === "error",
      ),
      false,
    );

    const [plainFile, reactFile] = await Promise.all([
      validatePptx(plainResult.pptx),
      validatePptx(reactResult.pptx),
    ]);
    assert.deepEqual(plainFile.errors, []);
    assert.deepEqual(reactFile.errors, []);
    assert.deepEqual(readerProjection(plainFile), readerProjection(reactFile));

    const plainSlideXml = slideXml(plainResult);
    const reactSlideXml = slideXml(reactResult);
    assert.match(plainSlideXml, /0B5C7A/u);
    assert.match(reactSlideXml, /0B5C7A/u);
    assert.match(plainSlideXml, /<p:timing/u);
    assert.match(reactSlideXml, /<p:timing/u);
    const plainParts = unzipSync(plainResult.pptx);
    const reactParts = unzipSync(reactResult.pptx);
    assert.equal(
      strFromU8(plainParts["ppt/theme/theme1.xml"]!),
      strFromU8(reactParts["ppt/theme/theme1.xml"]!),
    );
    assert.deepEqual(
      plainResult.report.slides[0]?.objects.find(
        ({ authorId }) => authorId === "next",
      )?.link,
      "slide:details",
    );
    assert.deepEqual(
      reactResult.report.slides[0]?.objects.find(
        ({ authorId }) => authorId === "next",
      )?.link,
      "slide:details",
    );
    const plainOverview = plainFile.slides.find(({ id }) => id === "overview");
    const plainObjects = plainOverview?.objects ?? [];
    const kinds = plainObjects.map(({ kind }) => kind);
    const expectedKinds = [
      "shape",
      "image",
      "group",
      "table",
      "chart",
    ] as const;
    for (const kind of expectedKinds) {
      assert.ok(kinds.includes(kind), `missing ${kind} in plain output`);
    }
    const plainText = plainObjects
      .flatMap((object) => object.text?.paragraphs ?? [])
      .map(({ text: content }) => content)
      .join("\n");
    assert.match(plainText, /Same intent/u);
    assert.match(plainText, /Native panel/u);
    assert.equal(
      plainObjects.find(({ kind }) => kind === "table")?.table?.cells[0]?.[0],
      "Metric",
    );
    const plainChart = plainObjects.find(({ kind }) => kind === "chart")?.chart;
    assert.equal(plainChart?.types[0], "barChart");
    assert.deepEqual(plainChart?.series[0]?.values, [12, 18]);
    const plainNext = plainObjects.find(({ link }) => link !== undefined)?.link;
    assert.ok(plainNext !== undefined && "slideId" in plainNext);
    assert.equal(plainNext.slideId, "details");
  });

  it("reports the same capability boundary when either entry is trimmed", async () => {
    const resolveShapeOnly = createResolver([shapeCapability]);
    const plainFailure = await Promise.resolve()
      .then(() =>
        emit(
          resolveShapeOnly(
            deck({ width: 240, height: 160 }).append(
              slide({ id: "trimmed" }).append(
                text("not registered", {
                  id: "plain-text",
                  x: 0,
                  y: 0,
                  width: 160,
                  height: 32,
                }),
              ),
            ),
          ),
        ),
      )
      .then(
        () => undefined,
        (error: unknown) => error,
      );
    const reactFailure = await expandDocument(
      <Deck width={240} height={160}>
        <Slide id="trimmed">
          <Text id="plain-text" x={0} y={0} width={160} height={32}>
            not registered
          </Text>
        </Slide>
      </Deck>,
    )
      .then((expanded) => emit(resolveShapeOnly(expanded)))
      .then(
        () => undefined,
        (error: unknown) => error,
      );

    assert.ok(plainFailure instanceof Error);
    assert.ok(reactFailure instanceof Error);
    const plainDiagnostics = (
      plainFailure as unknown as {
        readonly diagnostics: readonly {
          readonly code: string;
          readonly path: string;
          readonly message: string;
        }[];
      }
    ).diagnostics;
    assert.deepEqual(
      plainDiagnostics,
      (
        reactFailure as unknown as {
          readonly diagnostics: readonly unknown[];
        }
      ).diagnostics,
    );
    assert.equal(plainDiagnostics[0]?.code, "missing-capability");
  });
});
