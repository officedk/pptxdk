import assert from "node:assert/strict";

import {
  emit,
  motionSchedule,
  resolve,
  type ResolvedObject,
  type TextMetrics,
} from "@officedk/pptxdk";
import { paint } from "@officedk/pptxdk/preview";
import {
  Box,
  Deck,
  Slide,
  Text,
  expandDocument,
  fromVanilla,
} from "@officedk/pptxdk/react";
import { deck, slide } from "@officedk/pptxdk/vanilla";
import { validatePptx } from "@pptxdk/validation";
import { describe, it } from "vitest";

import {
  createMetricTile,
  createTakeawayStrip,
  createTelemetryPanel,
} from "../src/components/cards.ts";
import { createBeamDivergenceDiagram } from "../src/components/diagrams.ts";
import { AetherLinkDeck } from "../src/index.ts";
import { PALETTE, accent2Muted, aetherTheme } from "../src/theme.ts";

function assertValidPng1280x720(png: Uint8Array, label: string): void {
  const pngSignature = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a];
  assert.ok(png.byteLength > 24, `${label}: PNG too small`);
  assert.deepEqual(
    [...png.subarray(0, 8)],
    pngSignature,
    `${label}: invalid PNG signature`,
  );
  const view = new DataView(png.buffer, png.byteOffset, png.byteLength);
  assert.equal(view.getUint32(16, false), 1280, `${label}: PNG width`);
  assert.equal(view.getUint32(20, false), 720, `${label}: PNG height`);
}

function flattenObjects(objects: readonly ResolvedObject[]): ResolvedObject[] {
  return objects.flatMap((object) => [
    object,
    ...(object.members === undefined ? [] : flattenObjects(object.members)),
  ]);
}

function assertMeasuredTextFits(
  metrics: TextMetrics | undefined,
  label: string,
): void {
  assert.ok(metrics !== undefined, `${label}: expected textMetrics`);
  assert.equal(
    metrics.status,
    "measured",
    `${label}: expected measured textMetrics, got ${metrics.status}`,
  );
  assert.equal(
    metrics.fits,
    true,
    `${label}: expected text to fit (available=${JSON.stringify(metrics.available)}, content=${JSON.stringify(metrics.content)})`,
  );
}

describe("@pptxdk/aether-link full deck acceptance", () => {
  it("resolves, embeds all 6 English static fonts, paints 10 1280x720 PNGs, and validates native charts, placeholders, transitions, and motion", async () => {
    const resolution = resolve(await expandDocument(<AetherLinkDeck />));
    assert.deepEqual(resolution.report.dimensions, {
      width: 1280,
      height: 720,
    });
    assert.equal(resolution.slides.length, 10);
    assert.equal(resolution.report.masters?.length, 1);
    assert.equal(resolution.report.masters?.[0]?.layouts.length, 7);

    // Every font face must resolve with embedded data and zero missing glyphs.
    assert.ok(resolution.report.fonts.length >= 5);
    for (const fontUsage of resolution.report.fonts) {
      assert.equal(
        fontUsage.hasResource,
        true,
        `font ${fontUsage.family} (${fontUsage.weight}, ${fontUsage.style}) must have resource`,
      );
      assert.equal(
        fontUsage.missingGlyphs,
        undefined,
        `font ${fontUsage.family} (${fontUsage.weight}) has missing glyphs: ${fontUsage.missingGlyphs?.join(", ")}`,
      );
    }
    const fontFamilies = new Set(
      resolution.report.fonts.map((entry) => entry.family),
    );
    assert.ok(fontFamilies.has("Jost"));
    assert.ok(fontFamilies.has("Jost Medium"));
    assert.ok(fontFamilies.has("Darker Grotesque SemiBold"));

    // All resolved objects stay within 1280x720 bounds and every non-empty text/cell fits.
    for (const resolvedSlide of resolution.slides) {
      assert.deepEqual(resolvedSlide.dimensions, { width: 1280, height: 720 });
      for (const object of flattenObjects(resolvedSlide.objects)) {
        const bounds = object.bounds ?? object.box;
        if (bounds !== undefined) {
          assert.ok(
            bounds.x >= 0,
            `${resolvedSlide.id}/${object.path} overflows left (${bounds.x})`,
          );
          assert.ok(
            bounds.y >= 0,
            `${resolvedSlide.id}/${object.path} overflows top (${bounds.y})`,
          );
          assert.ok(
            bounds.x + bounds.width <= 1280,
            `${resolvedSlide.id}/${object.path} overflows right (${bounds.x + bounds.width})`,
          );
          assert.ok(
            bounds.y + bounds.height <= 720,
            `${resolvedSlide.id}/${object.path} overflows bottom (${bounds.y + bounds.height})`,
          );
        }
        if (
          (object.kind === "text" || object.kind === "shape") &&
          object.text !== undefined &&
          object.text.trim() !== ""
        ) {
          assertMeasuredTextFits(
            object.textMetrics,
            `${resolvedSlide.id}/${object.path} (${JSON.stringify(object.text)})`,
          );
        }
        if (object.table !== undefined) {
          for (const tableCell of object.table.cells) {
            if (tableCell.text.trim() !== "") {
              assertMeasuredTextFits(
                tableCell.textMetrics,
                `${resolvedSlide.id}/${tableCell.path} (${JSON.stringify(tableCell.text)})`,
              );
            }
          }
        }
      }
    }

    const result = emit(resolution, { embedFonts: "full" });
    assert.equal(result.report.slides.length, 10);

    const painted = paint(resolution);
    assert.equal(painted.pages.length, 10);
    for (const page of painted.pages) {
      assert.deepEqual(page.errors, [], `preview page ${page.page} errors`);
      assert.deepEqual(
        page.placeholders,
        [],
        `preview page ${page.page} placeholders`,
      );
      assert.ok(page.png !== undefined, `preview page ${page.page} png`);
      assertValidPng1280x720(page.png, `preview page ${page.page}`);
    }

    const file = await validatePptx(result.pptx, {
      allowedExtensions: ["p14", "p159"],
    });
    assert.deepEqual(file.errors, []);
    assert.equal(file.slides.length, 10);
    assert.deepEqual(
      file.slides.map((s) => s.id),
      [
        "cover",
        "exec-summary",
        "link-budget",
        "architecture",
        "pointing-acq",
        "telemetry-campaign",
        "ground-network",
        "subsystem-matrix",
        "anomaly-postmortem",
        "roadmap-close",
      ],
    );
    assert.deepEqual(
      [...new Set(file.slides.map((s) => s.section))],
      [
        "01 · Mission & Physics",
        "02 · System Architecture",
        "03 · Flight Campaign & Roadmap",
      ],
    );

    // Every emitted top-level object stays within 1280x720 EMU bounds.
    for (const emittedSlide of file.slides) {
      for (const object of emittedSlide.objects) {
        if (object.boxEmu === undefined) continue;
        assert.ok(
          object.boxEmu.x >= 0,
          `${emittedSlide.id}/${object.name} left`,
        );
        assert.ok(
          object.boxEmu.y >= 0,
          `${emittedSlide.id}/${object.name} top`,
        );
        assert.ok(
          object.boxEmu.x + object.boxEmu.width <= 1280 * 9525,
          `${emittedSlide.id}/${object.name} right`,
        );
        assert.ok(
          object.boxEmu.y + object.boxEmu.height <= 720 * 9525,
          `${emittedSlide.id}/${object.name} bottom`,
        );
      }
    }

    // Native charts & synchronized embedded Excel workbooks (including pie & error bars).
    const charts = file.slides.flatMap((s) =>
      s.objects.flatMap((o) =>
        o.chart === undefined ? [] : [{ slideId: s.id, chart: o.chart }],
      ),
    );
    assert.equal(charts.length, 5);
    for (const { slideId, chart } of charts) {
      assert.ok(chart.series.length > 0, `Chart on ${slideId} needs series`);
      const workbookSeries = chart.embeddedWorkbook?.series ?? [];
      assert.equal(
        workbookSeries.length,
        chart.series.length,
        `Embedded workbook on ${slideId} series count`,
      );
      for (let i = 0; i < chart.series.length; i += 1) {
        assert.equal(workbookSeries[i]?.name, chart.series[i]?.name);
        assert.deepEqual(
          workbookSeries[i]?.categories,
          chart.series[i]?.categories,
        );
        assert.deepEqual(workbookSeries[i]?.values, chart.series[i]?.values);
        assert.deepEqual(workbookSeries[i]?.x, chart.series[i]?.x);
        assert.deepEqual(workbookSeries[i]?.y, chart.series[i]?.y);
      }
    }
    const groundPieChart = charts.find(
      ({ slideId }) => slideId === "ground-network",
    )?.chart;
    assert.equal(groundPieChart?.groups[0]?.type, "pieChart");

    const patScatterChart = charts.find(
      ({ slideId }) => slideId === "pointing-acq",
    )?.chart;
    assert.ok((patScatterChart?.series[1]?.errorBars?.length ?? 0) >= 1);

    // Native placeholder bindings: chart (S6), pic (S7), tbl (S8), unpopulated obj (S10).
    const campaignSlide = file.slides.find(
      ({ id }) => id === "telemetry-campaign",
    )!;
    assert.ok(
      campaignSlide.objects.some(
        (o) =>
          o.chart !== undefined &&
          o.placeholder?.type === "chart" &&
          o.placeholder.idx === 12,
      ),
    );

    const groundSlide = file.slides.find(({ id }) => id === "ground-network")!;
    assert.ok(
      groundSlide.objects.some(
        (o) =>
          o.kind === "image" &&
          o.placeholder?.type === "pic" &&
          o.placeholder.idx === 10,
      ),
    );

    const matrixSlide = file.slides.find(
      ({ id }) => id === "subsystem-matrix",
    )!;
    assert.ok(
      matrixSlide.objects.some(
        (o) =>
          o.kind === "table" &&
          o.placeholder?.type === "tbl" &&
          o.placeholder.idx === 11,
      ),
    );
    const resolvedTable = resolution
      .slide("subsystem-matrix")
      .get("verification-table").objects[0]?.table;
    assert.equal(resolvedTable?.rows.length, 9);
    assert.equal(resolvedTable?.columns.length, 7);

    const roadmapSlide = file.slides.find(({ id }) => id === "roadmap-close")!;
    assert.equal(roadmapSlide.background?.type, "ref");
    assert.ok(
      roadmapSlide.objects.some(
        (o) =>
          o.kind === "shape" &&
          o.placeholder?.type === "obj" &&
          o.placeholder.idx === 9 &&
          o.text === undefined,
      ),
    );

    // Slide 1 image background, Slide 5 flip transition, Slide 9 colorMapOverride & themeOverride.
    const coverSlide = file.slides.find(({ id }) => id === "cover")!;
    assert.equal(coverSlide.background?.type, "image");

    const pointingSlide = file.slides.find(({ id }) => id === "pointing-acq")!;
    assert.equal(pointingSlide.transition?.kind, "flip");

    const anomalyLayout = file.masters?.[0]?.layouts.find(
      (layout) => layout.name === "Anomaly Brief",
    );
    assert.ok(
      anomalyLayout?.colorMapOverride !== undefined &&
        typeof anomalyLayout.colorMapOverride === "object",
    );
    assert.equal(anomalyLayout.colorMapOverride["accent6"], "accent2");
    assert.equal(anomalyLayout.themeOverride?.colors?.accent1, "#C9401C");

    // Morph continuity between Slide 2 and Slide 3.
    const linkBudgetSlide = file.slides.find(({ id }) => id === "link-budget")!;
    assert.equal(linkBudgetSlide.transition?.kind, "morph");
    assert.ok(
      linkBudgetSlide.objects.some((o) => o.name.startsWith("!!")),
      "Link budget slide should carry Morph !! object",
    );

    // Click steps on ExecSummary (2 click steps) & TelemetryCampaign (1 click step).
    const execSchedule = motionSchedule(resolution.slide("exec-summary"));
    assert.equal(
      execSchedule?.steps.filter((s) => s.trigger === "click").length,
      2,
    );
    const campaignSchedule = motionSchedule(
      resolution.slide("telemetry-campaign"),
    );
    assert.equal(
      campaignSchedule?.steps.filter((s) => s.trigger === "click").length,
      1,
    );
  }, 60_000);
});

describe("@pptxdk/aether-link Vanilla component library", () => {
  it("constructs closed and slotted components without React and adopts slotted children via fromVanilla", async () => {
    const panel = createTelemetryPanel({
      id: "vanilla-panel",
      kicker: "STANDALONE",
      title: "Pure Vanilla Telemetry",
      badge: "VERIFIED",
      accentColor: PALETTE.cobalt,
      badgeBackground: accent2Muted,
      width: 560,
    });
    panel.slot.append(
      createMetricTile({
        id: "vanilla-tile",
        label: "LINK MARGIN",
        value: "4.2",
        unit: "dB",
        detail: "Standalone Vanilla resolution without React.",
      }),
    );

    const standalone = deck({
      width: 1280,
      height: 720,
      theme: aetherTheme,
    }).append(
      slide({ id: "vanilla-check" }).append(
        panel.node,
        createBeamDivergenceDiagram({
          id: "vanilla-diagram",
          rfLabel: "Ka-Band",
          rfDetail: "3.8 mrad",
          opticalLabel: "1550 nm",
          opticalDetail: "11.4 µrad",
          caption: "Standalone group",
        }),
        createTakeawayStrip({
          id: "vanilla-strip",
          tag: "STANDALONE",
          text: "Zero-React Vanilla components resolve cleanly.",
        }),
      ),
    );

    const resolved = resolve(standalone);
    const objs = resolved.slide("vanilla-check").objects;
    assert.ok(objs.some((o) => o.authorId === "vanilla-panel"));
    assert.ok(objs.some((o) => o.authorId === "vanilla-tile"));
    assert.ok(objs.some((o) => o.authorId === "vanilla-diagram"));
    assert.ok(objs.some((o) => o.authorId === "vanilla-strip"));
    const standaloneBuilt = emit(resolved);
    assert.deepEqual((await validatePptx(standaloneBuilt.pptx)).errors, []);

    // Closed Vanilla factory bridged with fromVanilla statically rejects children.
    const MetricTile = fromVanilla(createMetricTile);
    const TelemetryPanel = fromVanilla(createTelemetryPanel);
    const rejectChildrenCheck = () => (
      <MetricTile id="bad" label="L" value="1" detail="D">
        {/* @ts-expect-error closed MetricTile component must reject JSX children */}
        <Text>disallowed</Text>
      </MetricTile>
    );
    void rejectChildrenCheck;

    const bridgedResolution = resolve(
      await expandDocument(
        <Deck width={800} height={450} theme={aetherTheme}>
          <Slide id="bridge-check">
            <Box
              style={{
                position: "absolute",
                left: 24,
                top: 24,
                width: 520,
                height: 220,
                display: "flex",
                flexDirection: "column",
              }}
            >
              <TelemetryPanel id="bridged-panel" title="Bridged Panel">
                <MetricTile
                  id="bridged-tile"
                  label="SIGNAL"
                  value="78.2"
                  unit="%"
                  accentColor="#0E8A74"
                  detail="Resolved inside slotted panel with explicit accentColor."
                />
              </TelemetryPanel>
            </Box>
          </Slide>
        </Deck>,
      ),
    );
    const tileObj = bridgedResolution
      .slide("bridge-check")
      .get("bridged-panel")
      .get("bridged-tile").objects[0];
    assert.ok(tileObj !== undefined);
    const built = emit(bridgedResolution);
    assert.deepEqual((await validatePptx(built.pptx)).errors, []);
  });
});
