import {
  CapabilityError,
  parseCapabilityLayoutContract,
  type BoundObjectCapability,
  type BoundObjectNativeCollaboration,
  type CapabilityFontUsageClaim,
  type CapabilityInput,
  type CapabilityLayoutMeasurement,
  type CapabilityLayoutRequest,
  type CapabilityParseContext,
  type ColorValue,
  type EmitContext,
  type EmitOutput,
  type Rect,
  type LayoutStyle,
  type NativePlaceholderRef,
  type ObjectCapabilityModule,
  type ObjectResolveContext,
  type ObjectResolution,
  type PaintContext,
  type PaintOutput,
  type SealedCanvasMechanism,
  type ThemeContext,
} from "@pptxdk/core/capability";
import { defineCanvasMechanism } from "@pptxdk/import";

import {
  fontPresentationContributions,
  type FontEmbeddingUse,
} from "../font/font-embedding.ts";
import { motionAnnotations, pickMotionProps } from "../motion/motion-api.ts";
import {
  parseLinkTarget,
  parseMorphKey,
  parseOptionalName,
} from "../shared/identity.ts";
import { resolvePlaceholderGeometry } from "../shared/style-inheritance.ts";
import {
  writeChartTextProperties,
  writeChartTitle,
  type MissingGlyphIssue,
  type TextResourceClaim,
} from "../shared/text-service.ts";
import type { TextStyle } from "../shared/text-service.ts";
import { chartCodegen } from "./chart-codegen.ts";
import {
  chartImportHandler,
  chartPlaceholderImportHandler,
} from "./chart-import.ts";
import { mergeDataLabels } from "./chart-labels.ts";
import {
  textOnlyChartStyle,
  type ParsedAxisCrossing,
  type ParsedCategoryAxis,
  type ParsedChart,
  type ParsedChartAxis,
  type ParsedChartGroup,
  type ParsedChartSeries,
  type ParsedScatterSeries,
  type ParsedValueAxis,
  type ResolvedCartesianAxes,
  type ResolvedScatterAxes,
} from "./chart-model.ts";
import {
  chartAllowedPropKeys,
  chartStyleKeys,
  parseChartAxisBySchema,
  parseChartCoordinateBySchema,
  parseChartDisplayBlanksAsBySchema,
  parseChartExtentBySchema,
  parseChartGroupBySchema,
  parseChartLegendBySchema,
  parseChartPlotAreaBySchema,
  parseChartTitleBySchema,
  parseChartXmlTextBySchema,
} from "./chart-schemas.ts";
import type {
  ChartAxis,
  ChartAxisPosition,
  ChartCategoryAxisReferences,
  ChartDataLabelSettings,
  ChartDisplayBlanksAs,
  ChartGroup,
  ChartGroupData,
  ChartKind,
  ChartLine,
  ChartMarker,
  ChartSeries,
  ChartTitle,
  ChartXYAxisReferences,
  IndependentCategoryData,
  IndependentCategorySeries,
  ScatterChartSeries,
  ScatterData,
  ScatterLine,
  SharedCategoryData,
} from "./chart-types.ts";
import {
  chartFrameXml,
  chartPartXml,
  chartUsesCategoryLabelRange,
  chartWorkbook,
  type ChartTextStyleXml,
} from "./chart-xml.ts";
import { drawChartPreview } from "./preview/index.ts";

const chartRelationship =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
const packageRelationship =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package";
const chartContentType =
  "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
const workbookContentType =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const horizontalAxisPositions = ["bottom", "top"] as const;
const verticalAxisPositions = ["left", "right"] as const;
const nativeAxisIds = [111111, 222222, 333333, 444444] as const;

function fail(code: string, message: string, path?: string): never {
  throw new CapabilityError(code, message, path);
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

type ParsedDraftSeries = ChartSeries & {
  readonly categories: readonly string[];
};

function hasVisibleDataLabels(
  value: false | ChartDataLabelSettings | undefined,
): value is ChartDataLabelSettings {
  return (
    value !== undefined &&
    value !== false &&
    (value.value === true ||
      value.percent === true ||
      value.category === true ||
      value.xValue === true ||
      value.seriesName === true)
  );
}

function hasLeaderLinesWithoutVisibleLabels(
  groupLabels: false | ChartDataLabelSettings | undefined,
  series: readonly ChartSeries[],
): boolean {
  // Only the effective labels written to each series dLbls matter. A group declaration that
  // every series overrides is not written and must not create an extra validation target.
  return series.some((item) => {
    const labels = mergeDataLabels(groupLabels, item.dataLabels);
    return (
      labels !== undefined &&
      labels !== false &&
      labels.leaderLines !== undefined &&
      labels.leaderLines !== false &&
      !hasVisibleDataLabels(labels)
    );
  });
}

type ParsedGroupDraft =
  | {
      readonly kind: Exclude<ChartKind, "scatter">;
      readonly data: ChartGroupData | IndependentCategoryData;
      readonly series: readonly ParsedDraftSeries[];
      readonly axes?: ChartCategoryAxisReferences;
      readonly dataLabels?: false | ChartDataLabelSettings;
      readonly stacking?: "stacked" | "percent";
      readonly gapWidth?: number;
      readonly overlap?: number;
      readonly explosion?: number;
      readonly firstSliceAngle?: number;
      readonly holeSize?: number;
    }
  | {
      readonly kind: "scatter";
      readonly data: ScatterData;
      readonly series: readonly ScatterChartSeries[];
      readonly axes?: ChartXYAxisReferences;
      readonly dataLabels?: false | ChartDataLabelSettings;
      readonly line?: ScatterLine;
      readonly markers: boolean;
    };

function chartGroupDrafts(value: unknown): ParsedGroupDraft[] {
  if (!Array.isArray(value) || value.length === 0) {
    fail(
      "invalid-chart-groups",
      "Chart.groups must contain at least one group",
    );
  }
  return value.map((rawItem, groupIndex) => {
    const label = `Chart.groups[${groupIndex}]`;
    const group: ChartGroup = parseChartGroupBySchema(rawItem, label);
    if (group.kind === "scatter") {
      return {
        kind: "scatter",
        data: group.data,
        series: group.series,
        markers: group.markers ?? true,
        ...(group.axes === undefined ? {} : { axes: group.axes }),
        ...(group.dataLabels === undefined
          ? {}
          : { dataLabels: group.dataLabels }),
        ...(group.line === undefined ? {} : { line: group.line }),
      };
    }

    const categoriesForSeries = (
      item: ChartSeries | IndependentCategorySeries,
    ): readonly string[] =>
      group.data.kind === "independent-categories"
        ? (item as IndependentCategorySeries).categories
        : group.data.categories;

    const series: ParsedDraftSeries[] = group.series.map((item) => ({
      ...item,
      categories: categoriesForSeries(item),
    }));

    if (
      (group.kind === "pie" || group.kind === "doughnut") &&
      hasLeaderLinesWithoutVisibleLabels(group.dataLabels, series)
    ) {
      fail(
        "invalid-data-labels",
        "leaderLines requires at least one visible data label value, percent, or category",
      );
    }

    if (group.kind === "column" || group.kind === "bar") {
      return {
        kind: group.kind,
        data: group.data,
        series,
        ...(group.axes === undefined ? {} : { axes: group.axes }),
        ...(group.dataLabels === undefined
          ? {}
          : { dataLabels: group.dataLabels }),
        ...(group.stacking === undefined ? {} : { stacking: group.stacking }),
        ...(group.gapWidth === undefined ? {} : { gapWidth: group.gapWidth }),
        ...(group.overlap === undefined ? {} : { overlap: group.overlap }),
      };
    }

    if (group.kind === "line" || group.kind === "area") {
      return {
        kind: group.kind,
        data: group.data,
        series,
        ...(group.axes === undefined ? {} : { axes: group.axes }),
        ...(group.dataLabels === undefined
          ? {}
          : { dataLabels: group.dataLabels }),
        ...(group.stacking === undefined ? {} : { stacking: group.stacking }),
      };
    }

    if (group.kind === "pie") {
      return {
        kind: "pie",
        data: group.data,
        series,
        ...(group.dataLabels === undefined
          ? {}
          : { dataLabels: group.dataLabels }),
        ...(group.explosion === undefined
          ? {}
          : { explosion: group.explosion }),
        ...(group.firstSliceAngle === undefined
          ? {}
          : { firstSliceAngle: group.firstSliceAngle }),
      };
    }

    // doughnut
    if (group.kind !== "doughnut") {
      throw new Error(`Unexpected chart group kind: ${String(group.kind)}`);
    }
    return {
      kind: "doughnut",
      data: group.data,
      series,
      ...(group.dataLabels === undefined
        ? {}
        : { dataLabels: group.dataLabels }),
      ...(group.explosion === undefined ? {} : { explosion: group.explosion }),
      ...(group.firstSliceAngle === undefined
        ? {}
        : { firstSliceAngle: group.firstSliceAngle }),
      ...(group.holeSize === undefined ? {} : { holeSize: group.holeSize }),
    };
  });
}

type ParsedAxisDraft = {
  readonly name: string;
  readonly spec: ChartAxis;
  direction?: "horizontal" | "vertical";
  position?: ChartAxisPosition;
  crossing?: ParsedAxisCrossing;
  id?: number;
  crossAxisId?: number;
};

type ResolvedGroupAxes =
  | {
      readonly kind: "category";
      readonly category: ParsedAxisDraft;
      readonly value: ParsedAxisDraft;
    }
  | {
      readonly kind: "scatter";
      readonly x: ParsedAxisDraft;
      readonly y: ParsedAxisDraft;
    };

const nativeAxisIdLimit = nativeAxisIds.length;

function axisEntries(
  value: unknown,
  groups: readonly ParsedGroupDraft[],
): {
  readonly entries: { name: string; spec: ChartAxis }[];
  readonly axisless: boolean;
} {
  const scatterOnly = groups.every((group) => group.kind === "scatter");
  if (
    groups.every((group) => group.kind === "pie" || group.kind === "doughnut")
  ) {
    if (groups.length !== 1) {
      fail(
        "unsupported-chart-group-combination",
        "Pie and doughnut charts support one chart group",
      );
    }
    if (value !== undefined) {
      fail(
        "invalid-chart-axes",
        "Pie and doughnut charts do not use category or value axes",
      );
    }
    return { entries: [], axisless: true };
  }
  if (
    groups.some((group) => group.kind === "pie" || group.kind === "doughnut")
  ) {
    fail(
      "unsupported-chart-group-combination",
      "Pie and doughnut charts support one chart group",
    );
  }
  if (
    value === undefined &&
    !scatterOnly &&
    new Set(groups.map((group) => group.kind)).size > 1
  ) {
    fail(
      "invalid-chart-axes",
      "Charts that combine different group kinds must declare chart-level axes and explicit group references",
    );
  }
  let entries: { name: string; spec: ChartAxis }[];
  if (value === undefined) {
    entries = scatterOnly
      ? [
          { name: "x", spec: { kind: "value" } },
          { name: "y", spec: { kind: "value" } },
        ]
      : [
          { name: "category", spec: { kind: "category" } },
          { name: "value", spec: { kind: "value" } },
        ];
  } else {
    if (!isRecord(value)) fail("invalid-axis", "Chart.axes must be an object");
    entries = Object.entries(value).map(([name, spec]) => {
      if (name.trim() === "") {
        fail("invalid-axis", "Chart.axes names must be non-empty strings");
      }
      if (!isRecord(spec)) {
        fail("invalid-axis", `Chart.axes.${name} must be an object`);
      }
      const kind = spec["kind"];
      if (kind !== "category" && kind !== "value") {
        fail(
          "invalid-chart-axis-role",
          `Chart.axes.${name}.kind must be "category" or "value"`,
        );
      }
      return {
        name,
        spec: parseChartAxisBySchema(spec, `Chart.axes.${name}`, kind),
      };
    });
  }
  if (entries.length === 0) {
    fail("invalid-axis", "Chart.axes must define at least one axis");
  }
  if (entries.length > nativeAxisIdLimit) {
    fail(
      "unsupported-chart-axis",
      "Office supports at most four chart axes per chart",
    );
  }
  return { entries, axisless: false };
}

type ResolvedParsedGroupAxes =
  | ({ readonly kind: "category" } & ResolvedCartesianAxes)
  | ({ readonly kind: "scatter" } & ResolvedScatterAxes)
  | undefined;

function parsedAxes(
  value: unknown,
  groups: readonly ParsedGroupDraft[],
): {
  axes: ParsedChartAxis[];
  groupAxes: ResolvedParsedGroupAxes[];
} {
  const { entries, axisless } = axisEntries(value, groups);
  if (axisless) {
    return { axes: [], groupAxes: groups.map(() => undefined) };
  }

  const entryByName = new Map(entries.map((entry) => [entry.name, entry]));
  const draftByName = new Map<string, ParsedAxisDraft>();
  const firstUse: ParsedAxisDraft[] = [];
  const resolvedGroups: ResolvedGroupAxes[] = [];

  const resolve = (
    name: string,
    expected: "category" | "value",
    label: string,
  ): ParsedAxisDraft => {
    const entry = entryByName.get(name);
    if (entry === undefined) {
      fail("unknown-chart-axis", `${label} references unknown axis "${name}"`);
    }
    if (entry.spec.kind !== expected) {
      fail(
        "invalid-chart-axis-role",
        `${label} references "${name}", which is not a ${expected} axis`,
      );
    }
    let draft = draftByName.get(name);
    if (draft === undefined) {
      draft = { name, spec: entry.spec };
      draftByName.set(name, draft);
    }
    if (!firstUse.includes(draft)) firstUse.push(draft);
    return draft;
  };

  const assignDirection = (
    axis: ParsedAxisDraft,
    direction: "horizontal" | "vertical",
  ): void => {
    if (axis.direction !== undefined && axis.direction !== direction) {
      fail(
        "invalid-chart-axes",
        `Chart axis "${axis.name}" cannot be both horizontal and vertical`,
      );
    }
    axis.direction = direction;
  };

  const defaultPair = (
    first: string,
    second: string,
    kind: "category" | "value",
  ) =>
    entries.length === 2 &&
    entryByName.get(first)?.spec.kind === kind &&
    entryByName.get(second)?.spec.kind === "value";

  for (const [groupIndex, group] of groups.entries()) {
    const label = `Chart.groups[${groupIndex}]`;
    if (group.kind === "scatter") {
      let references = group.axes;
      if (references === undefined) {
        if (!defaultPair("x", "y", "value")) {
          fail(
            "invalid-chart-axes",
            `${label} must reference chart-level X and Y axes`,
          );
        }
        references = { x: "x", y: "y" };
      }
      const x = resolve(references.x, "value", `${label} X`);
      const y = resolve(references.y, "value", `${label} Y`);
      if (x === y) {
        fail(
          "invalid-chart-axes",
          `${label} must reference distinct X and Y value axes`,
        );
      }
      assignDirection(x, "horizontal");
      assignDirection(y, "vertical");
      resolvedGroups.push({ kind: "scatter", x, y });
      continue;
    }

    let references = group.axes;
    if (references === undefined) {
      if (!defaultPair("category", "value", "category")) {
        fail(
          "invalid-chart-axes",
          `${label} must reference chart-level category and value axes`,
        );
      }
      references = { category: "category", value: "value" };
    }
    const category = resolve(
      references.category,
      "category",
      `${label} category axis`,
    );
    const valueAxis = resolve(references.value, "value", `${label} value axis`);
    if (category === valueAxis) {
      fail(
        "invalid-chart-axes",
        `${label} must reference distinct category and value axes`,
      );
    }
    if (group.kind === "bar") {
      assignDirection(category, "vertical");
      assignDirection(valueAxis, "horizontal");
    } else {
      assignDirection(category, "horizontal");
      assignDirection(valueAxis, "vertical");
    }
    resolvedGroups.push({ kind: "category", category, value: valueAxis });
  }

  for (const entry of entries) {
    if (!firstUse.some((axis) => axis.name === entry.name)) {
      fail(
        "invalid-chart-axes",
        `Chart axis "${entry.name}" is not referenced by any chart group`,
      );
    }
  }

  const partnersByAxis = new Map<ParsedAxisDraft, ParsedAxisDraft[]>();
  const addPartner = (left: ParsedAxisDraft, right: ParsedAxisDraft) => {
    const leftPartners = partnersByAxis.get(left);
    if (leftPartners === undefined) partnersByAxis.set(left, [right]);
    else if (!leftPartners.includes(right)) leftPartners.push(right);
  };
  for (const group of resolvedGroups) {
    const first = group.kind === "category" ? group.category : group.x;
    const second = group.kind === "category" ? group.value : group.y;
    addPartner(first, second);
    addPartner(second, first);
  }
  const partnersOf = (axis: ParsedAxisDraft): readonly ParsedAxisDraft[] =>
    partnersByAxis.get(axis) ?? [];
  const sharesPartner = (left: ParsedAxisDraft, right: ParsedAxisDraft) =>
    partnersOf(left).some((partner) => partnersOf(right).includes(partner));

  const explicitCrossing = (
    axis: ParsedAxisDraft,
  ): ParsedAxisCrossing | undefined => {
    if (axis.spec.crossesAt !== undefined) {
      return { kind: "crossesAt", value: axis.spec.crossesAt };
    }
    if (axis.spec.crosses !== undefined) {
      return { kind: "crosses", value: axis.spec.crosses };
    }
    return undefined;
  };

  const positionAxes = new Map<ChartAxisPosition, ParsedAxisDraft[]>();
  const claimPosition = (
    axis: ParsedAxisDraft,
    position: ChartAxisPosition,
  ) => {
    const occupants = positionAxes.get(position) ?? [];
    axis.position = position;
    if (occupants.length === 0) positionAxes.set(position, [axis]);
    else occupants.push(axis);
  };
  for (const axis of firstUse) {
    const position = axis.spec.position;
    if (position === undefined) continue;
    const allowed =
      axis.direction === "horizontal"
        ? horizontalAxisPositions
        : verticalAxisPositions;
    if (!(allowed as readonly string[]).includes(position)) {
      fail(
        "invalid-chart-axes",
        `Chart axis "${axis.name}" position "${position}" does not match its ${axis.direction} direction`,
      );
    }
    claimPosition(axis, position);
  }
  for (const axis of firstUse) {
    if (axis.position !== undefined) continue;
    const allowed =
      axis.direction === "horizontal"
        ? horizontalAxisPositions
        : verticalAxisPositions;
    // Defaults may share a side when the axes share a perpendicular partner; explicit
    // requests are not forced to the same side.
    const shared = firstUse.find(
      (other) =>
        other !== axis &&
        other.position !== undefined &&
        other.direction === axis.direction &&
        sharesPartner(axis, other),
    );
    if (shared !== undefined) {
      claimPosition(axis, shared.position!);
      continue;
    }
    const position = allowed.find((item) => !positionAxes.has(item));
    if (position === undefined) {
      fail(
        "unsupported-chart-axis",
        `Office supports at most two independent ${axis.direction} axes per chart`,
      );
    }
    claimPosition(axis, position);
  }

  const maxSide = (direction: "horizontal" | "vertical") =>
    direction === "vertical" ? "right" : "top";
  const rawCrossingForPosition = (
    axis: ParsedAxisDraft,
    partner: ParsedAxisDraft,
    position: ChartAxisPosition,
  ): ParsedAxisCrossing => {
    const edgeRequested =
      axis.spec.position !== undefined || position === maxSide(axis.direction!);
    if (!edgeRequested) return { kind: "crosses", value: "autoZero" };
    const partnerReversed = partner.spec.reversed === true;
    if (axis.direction === "vertical") {
      const rightIsMax = !partnerReversed;
      return {
        kind: "crosses",
        value:
          position === "right"
            ? rightIsMax
              ? "max"
              : "min"
            : rightIsMax
              ? "min"
              : "max",
      };
    }
    const topIsMax = !partnerReversed;
    return {
      kind: "crosses",
      value:
        position === "top"
          ? topIsMax
            ? "max"
            : "min"
          : topIsMax
            ? "min"
            : "max",
    };
  };
  const positionForRawCrossing = (
    crossing: ParsedAxisCrossing,
    partner: ParsedAxisDraft,
    direction: "horizontal" | "vertical",
  ): ChartAxisPosition | undefined => {
    if (crossing.kind !== "crosses" || crossing.value === "autoZero") {
      return undefined;
    }
    const maxAtFarEnd = partner.spec.reversed !== true;
    if (direction === "vertical") {
      return (crossing.value === "max") === maxAtFarEnd ? "right" : "left";
    }
    return (crossing.value === "max") === maxAtFarEnd ? "top" : "bottom";
  };
  // A known finite value domain proves an interior crossing cannot realize an
  // explicit edge position. Unknown auto-scale domains stay undecided.
  const knownValueEndpoint = (
    crossing: ParsedAxisCrossing,
    partner: ParsedAxisDraft,
  ): "min" | "max" | undefined => {
    if (partner.spec.kind !== "value") return undefined;
    const min = partner.spec.min;
    const max = partner.spec.max;
    if (min === undefined || max === undefined || min === max) return undefined;
    if (crossing.kind === "crossesAt") {
      if (crossing.value === min) return "min";
      if (crossing.value === max) return "max";
      return undefined;
    }
    if (crossing.kind !== "crosses" || crossing.value !== "autoZero") {
      return undefined;
    }
    if (min === 0) return "min";
    if (max === 0) return "max";
    return undefined;
  };
  const crossesKnownValueInterior = (
    crossing: ParsedAxisCrossing,
    partner: ParsedAxisDraft,
  ): boolean => {
    if (partner.spec.kind !== "value") return false;
    const min = partner.spec.min;
    const max = partner.spec.max;
    if (min === undefined || max === undefined) return false;
    const low = Math.min(min, max);
    const high = Math.max(min, max);
    if (crossing.kind === "crossesAt") {
      return crossing.value > low && crossing.value < high;
    }
    return (
      crossing.kind === "crosses" &&
      crossing.value === "autoZero" &&
      low < 0 &&
      high > 0
    );
  };
  const sameCrossing = (left: ParsedAxisCrossing, right: ParsedAxisCrossing) =>
    left.kind === right.kind &&
    (left.kind === "crossesAt"
      ? right.kind === "crossesAt" && left.value === right.value
      : right.kind === "crosses" && left.value === right.value);

  // An author crossing on axis B positions B's perpendicular partner A. The raw OOXML
  // field lives on A and positions A in B's domain, so adapt authors to the raw owner.
  for (const axis of firstUse) {
    const desired = new Set<ChartAxisPosition>();
    for (const partner of partnersOf(axis)) {
      const crossing = explicitCrossing(partner);
      if (crossing === undefined) continue;
      if (axis.spec.position !== undefined) {
        if (crossesKnownValueInterior(crossing, partner)) {
          fail(
            "invalid-chart-axes",
            `Chart axis "${axis.name}" position "${axis.spec.position}" conflicts with the known interior crossing on "${partner.name}"`,
          );
        }
        const endpoint = knownValueEndpoint(crossing, partner);
        if (endpoint !== undefined) {
          const edge = positionForRawCrossing(
            { kind: "crosses", value: endpoint },
            partner,
            axis.direction!,
          );
          if (edge !== undefined && edge !== axis.spec.position) {
            fail(
              "invalid-chart-axes",
              `Chart axis "${axis.name}" position "${axis.spec.position}" conflicts with the known crossing endpoint on "${partner.name}"`,
            );
          }
        }
      }
      const position = positionForRawCrossing(
        crossing,
        partner,
        axis.direction!,
      );
      if (position !== undefined) desired.add(position);
    }
    if (desired.size > 1) {
      fail(
        "invalid-chart-axes",
        `Explicit crossings on axes crossing "${axis.name}" disagree on its physical side`,
      );
    }
    const wanted = [...desired][0];
    if (wanted === undefined) continue;
    if (axis.spec.position !== undefined && axis.spec.position !== wanted) {
      fail(
        "invalid-chart-axes",
        `Chart axis "${axis.name}" position "${axis.spec.position}" conflicts with its explicit crossing controls`,
      );
    }
    axis.position = wanted;
  }

  for (const axis of firstUse) {
    let crossing: ParsedAxisCrossing | undefined;
    for (const partner of partnersOf(axis)) {
      const candidate =
        explicitCrossing(partner) ??
        rawCrossingForPosition(axis, partner, axis.position!);
      if (crossing === undefined) {
        crossing = candidate;
        continue;
      }
      if (!sameCrossing(crossing, candidate)) {
        fail(
          "invalid-chart-axes",
          `Shared crossing controls for axis "${axis.name}" disagree; use an explicit position or separate axes`,
        );
      }
    }
    if (crossing === undefined) {
      fail(
        "invalid-chart-axes",
        `Chart axis "${axis.name}" is not crossed by another axis`,
      );
    }
    axis.crossing = crossing;
  }

  for (const [index, axis] of firstUse.entries()) {
    axis.id = nativeAxisIds[index]!;
  }
  for (const group of resolvedGroups) {
    const first = group.kind === "category" ? group.category : group.x;
    const second = group.kind === "category" ? group.value : group.y;
    first.crossAxisId ??= second.id!;
    second.crossAxisId ??= first.id!;
  }

  const parsed: ParsedChartAxis[] = firstUse.map((axis) =>
    axis.spec.kind === "category"
      ? {
          kind: "category",
          name: axis.name,
          id: axis.id!,
          crossAxisId: axis.crossAxisId!,
          direction: axis.direction!,
          position: axis.position!,
          crossing: axis.crossing!,
          spec: axis.spec,
        }
      : {
          kind: "value",
          name: axis.name,
          id: axis.id!,
          crossAxisId: axis.crossAxisId!,
          direction: axis.direction!,
          position: axis.position!,
          crossing: axis.crossing!,
          spec: axis.spec,
        },
  );

  const parsedByDraft = new Map(
    firstUse.map((axis, index) => [axis, parsed[index]!]),
  );
  const groupAxes: ResolvedParsedGroupAxes[] = resolvedGroups.map((group) =>
    group.kind === "category"
      ? {
          kind: "category",
          category: parsedByDraft.get(group.category) as ParsedCategoryAxis,
          value: parsedByDraft.get(group.value) as ParsedValueAxis,
        }
      : {
          kind: "scatter",
          x: parsedByDraft.get(group.x) as ParsedValueAxis,
          y: parsedByDraft.get(group.y) as ParsedValueAxis,
        },
  );
  return { axes: parsed, groupAxes };
}

function displayBlanksAs(
  value: unknown,
  groups: readonly ParsedGroupDraft[],
): ChartDisplayBlanksAs | undefined {
  const requested = parseChartDisplayBlanksAsBySchema(value);
  if (requested === undefined) return undefined;
  const unsupported = groups.find((group) => {
    if (group.kind === "scatter") {
      // Span needs a connecting line; markers-only groups cannot span the missing point.
      return (
        requested === "span" &&
        (group.line === undefined || group.line === false)
      );
    }
    if (group.kind === "line") {
      return group.stacking === undefined
        ? !["gap", "zero", "span"].includes(requested)
        : requested !== "zero";
    }
    if (group.kind === "area") {
      return group.stacking === undefined
        ? !["gap", "zero"].includes(requested)
        : requested !== "zero";
    }
    return !["gap", "zero"].includes(requested);
  });
  if (unsupported !== undefined) {
    const detail =
      unsupported.kind === "scatter"
        ? `scatter chart group without a connecting line does not support displayBlanksAs "${requested}"`
        : `${unsupported.kind}${unsupported.stacking === undefined ? "" : ` ${unsupported.stacking}`} chart does not support displayBlanksAs "${requested}"`;
    fail("invalid-chart-missing-values", detail);
  }
  return requested;
}

type ResolvedChartModel = {
  readonly parsed: ParsedChart;
  readonly box: Rect;
  readonly name: string;
  readonly textStyles: ChartTextStyleXml;
  readonly fontUses: readonly FontEmbeddingUse[];
  readonly missingGlyphErrors: readonly MissingGlyphIssue[];
  readonly placeholder?: NativePlaceholderRef;
};

function emitChartModel(
  model: ResolvedChartModel,
  emitContext: EmitContext,
): EmitOutput {
  const embedFonts = emitContext.options.embedFonts;
  if (
    (embedFonts === "full" || embedFonts === "subset") &&
    model.missingGlyphErrors.length > 0
  ) {
    const issue = model.missingGlyphErrors[0]!;
    throw new CapabilityError("font-missing-glyph", issue.message, undefined, {
      font: issue.font,
      characters: issue.characters,
    });
  }
  const workbook = chartWorkbook(model.parsed);
  const relationshipId = emitContext.registerPackagePart({
    part: {
      pathTemplate: "ppt/charts/chart{n}.xml",
      contentType: chartContentType,
    },
    relationshipType: chartRelationship,
    content(relatedRelationshipIds) {
      const [workbookRelationshipId] = relatedRelationshipIds;
      if (workbookRelationshipId === undefined) {
        fail(
          "invalid-part",
          "Chart part requires embedded workbook relationship",
        );
      }
      return chartPartXml(
        model.parsed,
        model.box,
        workbookRelationshipId,
        model.textStyles,
        emitContext.theme,
      );
    },
    relatedParts: [
      {
        pathTemplate: "ppt/embeddings/Microsoft_Excel_Worksheet{n}.xlsx",
        contentType: workbookContentType,
        relationshipType: packageRelationship,
        bytes: workbook,
      },
    ],
  });
  const presentationContributions = fontPresentationContributions(
    model.fontUses,
    embedFonts,
  );
  return {
    xml: chartFrameXml(
      emitContext.id,
      model.name,
      model.box,
      relationshipId,
      emitContext.nativeLink,
      model.placeholder,
    ),
    ...(presentationContributions.length === 0
      ? {}
      : { presentationContributions }),
  };
}

function paintChartModel(
  model: ResolvedChartModel,
  paintContext: PaintContext,
): PaintOutput {
  const paintBox = paintContext.box ?? model.box;
  return drawChartPreview(model.parsed, paintBox, paintContext);
}

function defineChartCanvasMechanism<const Key extends string>(key: Key) {
  return defineCanvasMechanism({
    key,
    role: "canvas",
    emit: emitChartModel,
    paint: paintChartModel,
    import: chartImportHandler,
  });
}

export const chartColumnMechanism = defineChartCanvasMechanism("chart-column");
export const chartBarMechanism = defineChartCanvasMechanism("chart-bar");
export const chartLineMechanism = defineChartCanvasMechanism("chart-line");
export const chartAreaMechanism = defineChartCanvasMechanism("chart-area");
export const chartPieMechanism = defineChartCanvasMechanism("chart-pie");
export const chartDoughnutMechanism =
  defineChartCanvasMechanism("chart-doughnut");
export const chartScatterMechanism =
  defineChartCanvasMechanism("chart-scatter");
/** Typed native canvas mechanism connecting emit, paint, and import for `chart-placeholder`. */
export const chartPlaceholderMechanism = defineCanvasMechanism({
  key: "chart-placeholder",
  role: "canvas",
  emit: emitChartModel,
  paint: paintChartModel,
  import: chartPlaceholderImportHandler,
});

const chartCompatiblePlaceholderRoles = new Set<string>(["chart", "obj"]);

export const chartDataLabelsMechanism =
  defineChartCanvasMechanism("chart-data-labels");
export const chartTrendlineMechanism =
  defineChartCanvasMechanism("chart-trendline");
export const chartErrorBarsMechanism =
  defineChartCanvasMechanism("chart-error-bars");
export const chartMarkersAndPointsMechanism = defineChartCanvasMechanism(
  "chart-markers-and-points",
);

const chartMechanismsByKind = {
  column: chartColumnMechanism,
  bar: chartBarMechanism,
  line: chartLineMechanism,
  area: chartAreaMechanism,
  pie: chartPieMechanism,
  doughnut: chartDoughnutMechanism,
  scatter: chartScatterMechanism,
} as const;

/** Chart module validates input and resolves chart data/configuration for native emit and static paint. */
export const chartCapability: ObjectCapabilityModule<
  BoundObjectCapability,
  readonly [
    SealedCanvasMechanism<"chart-column">,
    SealedCanvasMechanism<"chart-bar">,
    SealedCanvasMechanism<"chart-line">,
    SealedCanvasMechanism<"chart-area">,
    SealedCanvasMechanism<"chart-pie">,
    SealedCanvasMechanism<"chart-doughnut">,
    SealedCanvasMechanism<"chart-scatter">,
    SealedCanvasMechanism<"chart-placeholder">,
    SealedCanvasMechanism<"chart-data-labels">,
    SealedCanvasMechanism<"chart-trendline">,
    SealedCanvasMechanism<"chart-error-bars">,
    SealedCanvasMechanism<"chart-markers-and-points">,
  ]
> = {
  declaration: "chart",
  codegen: chartCodegen,
  mechanisms: [
    chartColumnMechanism,
    chartBarMechanism,
    chartLineMechanism,
    chartAreaMechanism,
    chartPieMechanism,
    chartDoughnutMechanism,
    chartScatterMechanism,
    chartPlaceholderMechanism,
    chartDataLabelsMechanism,
    chartTrendlineMechanism,
    chartErrorBarsMechanism,
    chartMarkersAndPointsMechanism,
  ] as const,

  bind(
    input: CapabilityInput,
    context: CapabilityParseContext,
  ): BoundObjectCapability {
    const placeholder = context.placeholder;
    if (placeholder !== undefined) {
      const role =
        placeholder.layoutPlaceholder?.type ?? placeholder.nativeRef.type;
      if (!chartCompatiblePlaceholderRoles.has(role)) {
        fail(
          "incompatible-placeholder-content",
          `Chart cannot be bound to ${role} placeholder`,
        );
      }
    }
    if (input.children.length > 0) {
      fail(
        "unsupported-children",
        "Chart does not accept children; chart text and data should be passed in groups",
      );
    }
    const unknown = Object.keys(input.props).find(
      (key) => !chartAllowedPropKeys.includes(key),
    );
    if (unknown !== undefined) {
      fail(
        "unknown-parameter",
        `Chart contains unimplemented parameter ${unknown}`,
      );
    }

    const allowedStyleKeys = new Set<string>(chartStyleKeys);
    const contract = parseCapabilityLayoutContract({
      nativeKind: "graphicFrame",
      props: input.props,
      allowedStyleKeys,
      allowPadding: false,
      path: context.path,
    });

    const requestedName = parseOptionalName(input.props["name"]);
    if (requestedName !== undefined) {
      parseChartXmlTextBySchema(requestedName, "Chart.name");
    }
    const morphKey = parseMorphKey(input.props["morphKey"]);
    const link = parseLinkTarget(input.props["link"]);

    const drafts = chartGroupDrafts(input.props["groups"]);
    const { axes, groupAxes } = parsedAxes(input.props["axes"], drafts);
    const categoryDomains = new Map<string, readonly string[]>();
    for (const [groupIndex, draft] of drafts.entries()) {
      if (draft.kind === "scatter") continue;
      if (draft.data.kind !== "shared-categories") continue;
      const resolved = groupAxes[groupIndex];
      if (resolved === undefined || resolved.kind !== "category") continue;
      const domain = draft.data.categories;
      const existing = categoryDomains.get(resolved.category.name);
      if (existing === undefined) {
        categoryDomains.set(resolved.category.name, domain);
        continue;
      }
      if (
        existing.length !== domain.length ||
        existing.some((category, index) => category !== domain[index])
      ) {
        fail(
          "chart-category-domain-mismatch",
          `Chart groups sharing category axis "${resolved.category.name}" must have the same ordered category domain`,
        );
      }
    }

    let seriesIndex = 0;
    const groups: ParsedChartGroup[] = drafts.map((draft, groupIndex) => {
      const resolvedAxes = groupAxes[groupIndex];
      if (draft.kind === "scatter") {
        if (resolvedAxes === undefined || resolvedAxes.kind !== "scatter") {
          fail(
            "invalid-chart-axes",
            `Chart.groups[${groupIndex}] is missing resolved X and Y axes`,
          );
        }
        const series = draft.series.map((item, groupSeriesIndex) => ({
          ...item,
          groupIndex,
          groupSeriesIndex,
          seriesIndex: seriesIndex++,
        }));
        return {
          kind: "scatter",
          data: draft.data,
          series,
          axes: { x: resolvedAxes.x, y: resolvedAxes.y },
          line: draft.line ?? false,
          markers: draft.markers,
          ...(draft.dataLabels === undefined
            ? {}
            : { dataLabels: draft.dataLabels }),
        };
      }
      const series = draft.series.map((item, groupSeriesIndex) => ({
        ...item,
        groupIndex,
        groupSeriesIndex,
        seriesIndex: seriesIndex++,
      }));
      if (draft.kind === "pie") {
        const firstSeries = series[0];
        if (firstSeries === undefined || series.length !== 1) {
          fail(
            "invalid-data",
            "pie chart currently only supports a single data series",
          );
        }
        return {
          kind: "pie",
          data: draft.data as SharedCategoryData,
          series: [firstSeries],
          firstSliceAngle: draft.firstSliceAngle ?? 0,
          ...(draft.explosion === undefined
            ? {}
            : { explosion: draft.explosion }),
          ...(draft.dataLabels === undefined
            ? {}
            : { dataLabels: draft.dataLabels }),
        };
      }
      if (draft.kind === "doughnut") {
        return {
          kind: "doughnut",
          data: draft.data,
          series,
          firstSliceAngle: draft.firstSliceAngle ?? 0,
          holeSize: draft.holeSize ?? 50,
          ...(draft.explosion === undefined
            ? {}
            : { explosion: draft.explosion }),
          ...(draft.dataLabels === undefined
            ? {}
            : { dataLabels: draft.dataLabels }),
        };
      }
      if (resolvedAxes === undefined || resolvedAxes.kind !== "category") {
        fail(
          "invalid-chart-axes",
          `Chart.groups[${groupIndex}] is missing resolved category and value axes`,
        );
      }
      const cartesianAxes: ResolvedCartesianAxes = {
        category: resolvedAxes.category,
        value: resolvedAxes.value,
      };
      if (draft.kind === "column" || draft.kind === "bar") {
        return {
          kind: draft.kind,
          data: draft.data as SharedCategoryData,
          series,
          axes: cartesianAxes,
          stacking: draft.stacking ?? "clustered",
          gapWidth: draft.gapWidth ?? 150,
          ...(draft.overlap === undefined ? {} : { overlap: draft.overlap }),
          ...(draft.dataLabels === undefined
            ? {}
            : { dataLabels: draft.dataLabels }),
        };
      }
      return {
        kind: draft.kind,
        data: draft.data as SharedCategoryData,
        series,
        axes: cartesianAxes,
        stacking: draft.stacking ?? "standard",
        ...(draft.dataLabels === undefined
          ? {}
          : { dataLabels: draft.dataLabels }),
      };
    });
    const series = groups.flatMap(
      (group): readonly (ParsedChartSeries | ParsedScatterSeries)[] =>
        group.series,
    );
    const categorySeries = series.filter(
      (item): item is ParsedChartSeries => "categories" in item,
    );
    const categories = categorySeries[0]?.categories ?? [];
    const categoryCount = Math.max(
      0,
      ...categorySeries.map((item) => item.categories.length),
    );

    const placeholderGeometry =
      placeholder === undefined
        ? undefined
        : resolvePlaceholderGeometry(placeholder, input.props);
    const placeholderWidth =
      typeof placeholderGeometry?.width === "number"
        ? placeholderGeometry.width
        : undefined;
    const placeholderHeight =
      typeof placeholderGeometry?.height === "number"
        ? placeholderGeometry.height
        : undefined;
    const declaredWidth = placeholderWidth ?? contract.explicitPixelSize?.width;
    const declaredHeight =
      placeholderHeight ?? contract.explicitPixelSize?.height;
    const rawX = placeholderGeometry?.x;
    const rawY = placeholderGeometry?.y;
    const declaredX =
      rawX !== undefined ? parseChartCoordinateBySchema(rawX, "x") : undefined;
    const declaredY =
      rawY !== undefined ? parseChartCoordinateBySchema(rawY, "y") : undefined;

    const legend = parseChartLegendBySchema(input.props["legend"]);
    const checkedPlotArea = parseChartPlotAreaBySchema(input.props["plotArea"]);
    const title = parseChartTitleBySchema(input.props["title"], "Chart.title");
    const blanks = displayBlanksAs(input.props["displayBlanksAs"], drafts);

    const parsed: ParsedChart = {
      axes,
      hasExplicitAxes: input.props["axes"] !== undefined,
      groups,
      series,
      categories,
      categoryCount,
      contract,
      ...(requestedName === undefined ? {} : { name: requestedName }),
      ...(title === undefined ? {} : { title }),
      ...(blanks === undefined ? {} : { displayBlanksAs: blanks }),
      ...(morphKey === undefined ? {} : { morphKey }),
      ...(link === undefined ? {} : { link }),
      ...(declaredX === undefined ? {} : { declaredX }),
      ...(declaredY === undefined ? {} : { declaredY }),
      ...(declaredWidth === undefined ? {} : { declaredWidth }),
      ...(declaredHeight === undefined ? {} : { declaredHeight }),
      ...(legend === undefined ? {} : { legend: legend }),
      ...(checkedPlotArea === undefined ? {} : { plotArea: checkedPlotArea }),
      ...(input.props["style"] === undefined
        ? {}
        : { style: input.props["style"] as TextStyle & LayoutStyle }),
    };

    const authorId =
      typeof input.props["id"] === "string" ? input.props["id"] : undefined;
    const motionProps: Readonly<Record<string, unknown>> = pickMotionProps(
      input.props,
    );

    const native: BoundObjectNativeCollaboration = {
      ...(morphKey === undefined ? {} : { morphKey }),
      ...(requestedName === undefined ? {} : { hasExplicitName: true }),
      ...(link === undefined ? {} : { link }),
    };

    return {
      contract,
      ...(morphKey !== undefined ||
      requestedName !== undefined ||
      link !== undefined
        ? { native }
        : {}),
      measureLayout(
        request: CapabilityLayoutRequest,
      ): CapabilityLayoutMeasurement {
        const width =
          request.availableWidth ??
          declaredWidth ??
          (declaredHeight === undefined ? 320 : declaredHeight * (16 / 9));
        const height = declaredHeight ?? width * (9 / 16);
        return {
          width,
          height,
          minContentWidth: declaredWidth ?? 0,
          minContentHeight: declaredHeight ?? 0,
        };
      },
      resolve(resolveContext: ObjectResolveContext): ObjectResolution {
        const x = parseChartCoordinateBySchema(
          resolveContext.layoutBox?.x ?? declaredX ?? 0,
          "x",
        );
        const y = parseChartCoordinateBySchema(
          resolveContext.layoutBox?.y ?? declaredY ?? 0,
          "y",
        );
        const rawBoxWidth =
          resolveContext.layoutBox?.width ??
          declaredWidth ??
          (declaredHeight === undefined ? 320 : declaredHeight * (16 / 9));
        const rawBoxHeight =
          resolveContext.layoutBox?.height ??
          declaredHeight ??
          rawBoxWidth * (9 / 16);
        const width = parseChartExtentBySchema(rawBoxWidth, "width");
        const height = parseChartExtentBySchema(rawBoxHeight, "height");
        const box: Rect = { x, y, width, height };

        if (parsed.plotArea !== undefined) {
          if (
            parsed.plotArea.left < 0 ||
            parsed.plotArea.top < 0 ||
            parsed.plotArea.width <= 0 ||
            parsed.plotArea.height <= 0 ||
            parsed.plotArea.left + parsed.plotArea.width > box.width ||
            parsed.plotArea.top + parsed.plotArea.height > box.height
          ) {
            fail(
              "invalid-plot-area",
              "Chart.plotArea must be inside chart box with positive width and height",
            );
          }
        }

        const name =
          resolveContext.morphName ??
          requestedName ??
          placeholder?.layoutPlaceholder?.name ??
          `Chart ${resolveContext.id - 1}`;
        parseChartXmlTextBySchema(name, "Chart.name");

        const annotations = motionAnnotations(motionProps, {
          kind: "chart",
          name,
          ...(authorId === undefined ? {} : { authorId }),
          textBody: false,
          // Chart ranges own their gridLegend step; a chart has no whole-text fill/background flag.
          filled: false,
          box,
          rotationDegrees: 0,
          chartSeriesPointCounts: parsed.series.map((item) =>
            "x" in item ? item.x.length : item.values.length,
          ),
        });
        const {
          textStyles,
          resourceClaims,
          fontUses,
          fontUsages,
          missingGlyphErrors,
        } = chartResourceClaims(parsed, resolveContext.theme);
        validateChartColors(parsed, resolveContext.theme);
        const usesCategoryLabelRange = chartUsesCategoryLabelRange(parsed);

        const primaryMechanism =
          placeholder !== undefined
            ? chartPlaceholderMechanism
            : chartMechanismsByKind[parsed.groups[0]?.kind ?? "column"];
        const resolvedModel: ResolvedChartModel = {
          parsed,
          box,
          name,
          textStyles,
          fontUses,
          missingGlyphErrors,
          ...(placeholder === undefined
            ? {}
            : { placeholder: placeholder.nativeRef }),
        };

        return {
          observation: {
            id: resolveContext.id,
            kind: "chart",
            name,
            box,
          },
          extensions: usesCategoryLabelRange ? ["c15"] : [],
          ...(annotations.length === 0 ? {} : { annotations }),
          resourceClaims,
          ...(fontUsages.length === 0 ? {} : { fontUsages }),
          emit(emitContext: EmitContext): EmitOutput {
            return primaryMechanism.emit(resolvedModel, emitContext);
          },
          paint(paintContext: PaintContext): PaintOutput {
            return primaryMechanism.paint(resolvedModel, paintContext);
          },
        };
      },
    };
  },
} satisfies ObjectCapabilityModule;

const generalNumberCharacters = "0123456789.,-+eE";
const dateNameCharacters =
  "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz";
// When simplified Chinese LCID is explicitly specified, date names use localized glyphs instead of English characters.
const dateNameCharactersByLocale: Readonly<Record<string, string>> = {
  "0804": "星期日一二三四五六七八九十年月",
};
const numberFormatLiteralCharacters = new Set("$+():^'{}<=>-/!&~ ");

function dateNameCharactersForLocale(
  locale: string | undefined,
): string | undefined {
  if (locale === undefined) return dateNameCharacters;
  const languageId = /^(?:0x)?([\da-f]{1,4})(?:,.*)?$/iu.exec(locale)?.[1];
  if (languageId === undefined) return undefined;
  const normalizedLanguageId = languageId.padStart(4, "0").toUpperCase();
  if (normalizedLanguageId === "0409") return dateNameCharacters;
  // Locales explicitly specified without registered name glyphs do not masquerade as English.
  return dateNameCharactersByLocale[normalizedLanguageId];
}

function hasDigitPlaceholderAfter(
  format: readonly string[],
  start: number,
): boolean {
  for (let index = start; index < format.length; index += 1) {
    const character = format[index]!;
    if (character === ";") return false;
    if (character === '"') {
      index += 1;
      while (index < format.length && format[index] !== '"') index += 1;
      continue;
    }
    if (character === "\\" || character === "_") {
      index += 1;
      continue;
    }
    if (character === "[") {
      while (index < format.length && format[index] !== "]") index += 1;
      continue;
    }
    if (character === "0" || character === "#" || character === "?") {
      return true;
    }
  }
  return false;
}

/** Claims glyphs that chart number formatting might display, without simulating application-level number formatting at build time. */
function numberLabelCharacters(
  numberFormat: string | undefined,
  percent = false,
): string {
  const characters = new Set<string>();
  const add = (value: string) => {
    for (const character of value) characters.add(character);
  };
  const format = Array.from(numberFormat ?? "");
  const defaultFormat =
    numberFormat === undefined ||
    numberFormat.trim().toLocaleLowerCase("en-US") === "general";
  let hasNumericValue = defaultFormat;
  let hasGeneralValue = defaultFormat;
  let sectionLocale: string | undefined;
  let sectionCount = 1;

  if (defaultFormat) add(generalNumberCharacters);
  for (let index = 0; index < format.length; index += 1) {
    const character = format[index]!;
    if (character === '"') {
      for (index += 1; index < format.length; index += 1) {
        if (format[index] === '"') break;
        characters.add(format[index]!);
      }
      continue;
    }
    if (character === "\\" || character === "*") {
      const following = format[index + 1];
      if (following !== undefined) {
        characters.add(following);
        index += 1;
      }
      continue;
    }
    if (character === "_") {
      index += 1;
      continue;
    }
    if (character === "[") {
      const end = format.indexOf("]", index + 1);
      if (end < 0) continue;
      const directive = format.slice(index + 1, end).join("");
      if (directive.startsWith("$")) {
        const currencyAndLocale = directive.slice(1);
        const separator = currencyAndLocale.indexOf("-");
        const currency =
          separator < 0
            ? currencyAndLocale
            : currencyAndLocale.slice(0, separator);
        add(currency);
        sectionLocale =
          separator < 0 ? undefined : currencyAndLocale.slice(separator + 1);
      } else if (/^[hms]+$/iu.test(directive)) {
        hasNumericValue = true;
      }
      index = end;
      continue;
    }
    if (character === ";") {
      sectionCount += 1;
      if (sectionCount > 3) break;
      sectionLocale = undefined;
      continue;
    }
    if (character === "0" || character === "#" || character === "?") {
      if (character === "?") characters.add(" ");
      hasNumericValue = true;
      continue;
    }
    if (character === ",") {
      if (hasDigitPlaceholderAfter(format, index + 1))
        characters.add(character);
      continue;
    }
    if (character === "." || character === "%") {
      characters.add(character);
      continue;
    }
    if (character === "A" || character === "a") {
      const amPm = format.slice(index, index + 5).join("");
      const amPmHalf = format.slice(index, index + 3).join("");
      if (/^am\/pm$/iu.test(amPm)) {
        add(amPm.replace("/", ""));
        index += 4;
        continue;
      }
      if (/^a\/p$/iu.test(amPmHalf)) {
        add(amPmHalf.replace("/", ""));
        index += 2;
        continue;
      }
    }
    if (/[yYmdDhHsS]/u.test(character)) {
      let end = index + 1;
      while (
        end < format.length &&
        format[end]!.toLocaleLowerCase("en-US") ===
          character.toLocaleLowerCase("en-US")
      ) {
        end += 1;
      }
      const width = end - index;
      hasNumericValue = true;
      if (
        (character.toLocaleLowerCase("en-US") === "m" ||
          character.toLocaleLowerCase("en-US") === "d") &&
        width >= 3
      ) {
        const dateNames = dateNameCharactersForLocale(sectionLocale);
        if (dateNames !== undefined) add(dateNames);
      }
      index = end - 1;
      continue;
    }
    if (character === "G" || character === "g") {
      const general = format.slice(index, index + 7).join("");
      if (
        general.toLocaleLowerCase("en-US") === "general" &&
        (format[index + 7] === undefined || !/[a-z]/iu.test(format[index + 7]!))
      ) {
        add(generalNumberCharacters);
        hasNumericValue = true;
        hasGeneralValue = true;
        index += 6;
        continue;
      }
    }
    if (character === "E" || character === "e") {
      const sign = format[index + 1];
      if (
        (sign === "+" || sign === "-") &&
        hasDigitPlaceholderAfter(format, index + 2)
      ) {
        add(character + sign);
        hasNumericValue = true;
        index += 1;
        continue;
      }
    }
    if (numberFormatLiteralCharacters.has(character)) {
      characters.add(character);
      continue;
    }
    if (/[1-9]/u.test(character) || character.codePointAt(0)! > 0x7f) {
      characters.add(character);
    }
  }

  if (hasNumericValue) {
    add("0123456789");
    if (sectionCount === 1) characters.add("-");
  }
  if (hasGeneralValue) add("+eE");
  if (percent && numberFormat === undefined) characters.add("%");
  return [...characters].join("");
}

function effectiveSeriesLabels(
  group: ParsedChartGroup,
  series: ParsedChartSeries | ParsedScatterSeries,
): false | ChartDataLabelSettings | undefined {
  return mergeDataLabels(group.dataLabels, series.dataLabels);
}

function effectivePointLabels(
  group: ParsedChartGroup,
  series: ParsedChartSeries | ParsedScatterSeries,
  index: number,
): false | ChartDataLabelSettings | undefined {
  return mergeDataLabels(
    effectiveSeriesLabels(group, series),
    series.points?.[index]?.dataLabels,
  );
}

function pointDataLabelTexts(
  group: ParsedChartGroup,
  series: ParsedChartSeries | ParsedScatterSeries,
  labels: ChartDataLabelSettings,
  index: number,
): string[] {
  if (group.kind === "scatter") {
    const scatter = series as ParsedScatterSeries;
    const texts: string[] = [];
    const numericText = () =>
      numberLabelCharacters(labels.numberFormat, labels.percent === true);
    if (labels.seriesName === true) {
      texts.push(scatter.name);
    }
    if (
      (labels.value === true || labels.percent === true) &&
      scatter.y[index] !== null &&
      scatter.y[index] !== undefined
    ) {
      texts.push(numericText());
    }
    if (
      labels.xValue === true &&
      scatter.x[index] !== null &&
      scatter.x[index] !== undefined
    ) {
      texts.push(numericText());
    }
    if (labels.separator !== undefined) {
      texts.push(labels.separator);
    }
    return texts;
  }
  const value = (series as ParsedChartSeries).values[index];
  if (value === null || value === undefined) {
    return [];
  }
  const texts: string[] = [];
  if (labels.seriesName === true) {
    texts.push(series.name);
  }
  if (labels.value === true || labels.percent === true) {
    texts.push(
      numberLabelCharacters(labels.numberFormat, labels.percent === true),
    );
  }
  if (labels.category === true) {
    texts.push((series as ParsedChartSeries).categories[index]!);
  }
  if (labels.separator !== undefined) {
    texts.push(labels.separator);
  }
  return texts;
}

function groupDataLabelTexts(group: ParsedChartGroup): string[] {
  const texts: string[] = [];
  for (const series of group.series) {
    const count =
      group.kind === "scatter"
        ? (series as ParsedScatterSeries).x.length
        : (series as ParsedChartSeries).values.length;
    for (let index = 0; index < count; index += 1) {
      const labels = effectivePointLabels(group, series, index);
      if (labels === undefined || labels === false) continue;
      texts.push(...pointDataLabelTexts(group, series, labels, index));
    }
  }
  return texts;
}

function mergeTextStyle(
  base: TextStyle | undefined,
  override: TextStyle,
): TextStyle {
  return { ...base, ...override };
}

function titleParts(
  value: ChartTitle | undefined,
): { readonly text: string; readonly style?: TextStyle } | undefined {
  if (value === undefined) return undefined;
  if (typeof value === "string") return { text: value };
  return {
    text: value.text,
    ...(value.style === undefined ? {} : { style: value.style }),
  };
}

function validateChartColors(parsed: ParsedChart, theme: ThemeContext): void {
  const checkColor = (color: ColorValue | undefined) => {
    if (color !== undefined) theme.resolveColor(color);
  };
  const checkLine = (line: boolean | ChartLine | undefined) => {
    if (typeof line === "object") checkColor(line.color);
  };
  const checkMarker = (marker: boolean | ChartMarker | undefined) => {
    if (typeof marker === "object") {
      checkColor(marker.fill);
      checkLine(marker.line);
    }
  };
  const checkLabels = (
    dataLabels: false | ChartDataLabelSettings | undefined,
  ) => {
    if (typeof dataLabels === "object") checkLine(dataLabels.leaderLines);
  };

  const checkCommonSeries = (
    series: ParsedChartSeries | ParsedScatterSeries,
  ) => {
    checkColor(series.color);
    checkLine(series.line);
    checkMarker(series.marker);
    for (const point of Object.values(series.points ?? {})) {
      checkColor(point.color);
      checkLine(point.line);
      checkMarker(point.marker);
      checkColor(point.labelColor);
    }
    for (const trendline of series.trendlines ?? []) {
      checkLine(trendline.line);
    }
  };

  for (const { spec } of parsed.axes) {
    checkLine(spec.line);
    checkLine(spec.gridlines);
    checkLine(spec.minorGridlines);
  }
  for (const group of parsed.groups) {
    checkLabels(group.dataLabels);
    if (group.kind === "scatter") {
      for (const series of group.series) {
        checkCommonSeries(series);
        checkLine(series.errorBars?.x?.line);
        checkLine(series.errorBars?.y?.line);
      }
    } else {
      for (const series of group.series) {
        checkCommonSeries(series);
        checkLabels(series.dataLabels);
        checkLine(series.errorBars?.line);
      }
    }
  }
}

function chartResourceClaims(
  parsed: ParsedChart,
  theme: ThemeContext,
): {
  readonly textStyles: ChartTextStyleXml;
  readonly resourceClaims: readonly TextResourceClaim[];
  readonly fontUses: readonly FontEmbeddingUse[];
  readonly fontUsages: readonly CapabilityFontUsageClaim[];
  readonly missingGlyphErrors: readonly MissingGlyphIssue[];
} {
  const categoryLabels = new Map<string, readonly string[]>();
  for (const group of parsed.groups) {
    if (
      group.kind === "scatter" ||
      group.kind === "pie" ||
      group.kind === "doughnut"
    ) {
      continue;
    }
    categoryLabels.set(group.axes.category.name, group.data.categories);
  }
  const dataLabelText = parsed.groups.flatMap((group) =>
    group.dataLabels !== undefined &&
    group.dataLabels !== false &&
    group.dataLabels.style !== undefined
      ? []
      : groupDataLabelTexts(group),
  );
  const chartText = [
    ...parsed.series.flatMap((item) =>
      "categories" in item ? item.categories : [],
    ),
    ...parsed.series.map(({ name }) => name),
    ...dataLabelText,
  ];
  const percentStacked = parsed.groups.some(
    (group) =>
      (group.kind === "column" ||
        group.kind === "bar" ||
        group.kind === "line" ||
        group.kind === "area") &&
      group.stacking === "percent",
  );
  const axisCharacters = (spec: ChartAxis): string[] =>
    spec.visible === false
      ? []
      : [
          numberLabelCharacters(
            spec.kind === "value" ? spec.numberFormat : undefined,
            percentStacked,
          ),
        ];
  for (const { spec } of parsed.axes) {
    if (
      spec.kind === "value" &&
      spec.visible !== false &&
      spec.style === undefined
    ) {
      chartText.push(...axisCharacters(spec));
    }
  }
  const chartStyle = textOnlyChartStyle(parsed.style);
  const chart = writeChartTextProperties(
    chartStyle,
    chartText,
    "Chart.style",
    theme,
  );
  const resourceClaims: TextResourceClaim[] = [...chart.resourceClaims];
  const fontUses: FontEmbeddingUse[] = chart.fontUses.map((use) => ({
    ...use,
    path: "Chart.style",
  }));
  const fontUsages: CapabilityFontUsageClaim[] = [...chart.fontUsages];
  const missingGlyphErrors: MissingGlyphIssue[] = [...chart.missingGlyphErrors];
  const textStyles: {
    chart: string;
    title?: string;
    axisTitles?: Map<string, string>;
    axisStyles?: Map<string, string>;
    groupDataLabels?: Map<number, string>;
    pointLabels?: Map<string, string>;
  } = { chart: chart.xml };
  const chartTitle = titleParts(parsed.title);
  if (chartTitle !== undefined) {
    const output = writeChartTitle(
      chartTitle.text,
      mergeTextStyle(chartStyle, chartTitle.style ?? {}),
      "Chart.title",
      theme,
    );
    textStyles.title = output.xml;
    resourceClaims.push(...output.resourceClaims);
    fontUses.push(
      ...output.fontUses.map((use) => ({ ...use, path: "Chart.title" })),
    );
    fontUsages.push(...output.fontUsages);
    missingGlyphErrors.push(...output.missingGlyphErrors);
  }
  for (const { name, spec } of parsed.axes) {
    if (spec.style !== undefined) {
      const output = writeChartTextProperties(
        mergeTextStyle(chartStyle, spec.style),
        spec.kind === "category"
          ? (categoryLabels.get(name) ?? [])
          : axisCharacters(spec),
        `Chart.axes.${name}.style`,
        theme,
      );
      textStyles.axisStyles ??= new Map();
      textStyles.axisStyles.set(name, output.xml);
      resourceClaims.push(...output.resourceClaims);
      fontUses.push(
        ...output.fontUses.map((use) => ({
          ...use,
          path: `Chart.axes.${name}.style`,
        })),
      );
      fontUsages.push(...output.fontUsages);
      missingGlyphErrors.push(...output.missingGlyphErrors);
    }
    const axisTitle = titleParts(spec.title);
    if (axisTitle !== undefined) {
      const output = writeChartTitle(
        axisTitle.text,
        mergeTextStyle(
          mergeTextStyle(chartStyle, spec.style ?? {}),
          axisTitle.style ?? {},
        ),
        `Chart.axes.${name}.title`,
        theme,
      );
      textStyles.axisTitles ??= new Map();
      textStyles.axisTitles.set(name, output.xml);
      resourceClaims.push(...output.resourceClaims);
      fontUses.push(
        ...output.fontUses.map((use) => ({
          ...use,
          path: `Chart.axes.${name}.title`,
        })),
      );
      fontUsages.push(...output.fontUsages);
      missingGlyphErrors.push(...output.missingGlyphErrors);
    }
  }
  for (const [groupIndex, group] of parsed.groups.entries()) {
    const groupLabels = group.dataLabels;
    if (
      groupLabels !== undefined &&
      groupLabels !== false &&
      groupLabels.style !== undefined
    ) {
      const output = writeChartTextProperties(
        mergeTextStyle(chartStyle, groupLabels.style),
        groupDataLabelTexts(group),
        `Chart.groups[${groupIndex}].dataLabels.style`,
        theme,
      );
      textStyles.groupDataLabels ??= new Map();
      textStyles.groupDataLabels.set(groupIndex, output.xml);
      resourceClaims.push(...output.resourceClaims);
      fontUses.push(
        ...output.fontUses.map((use) => ({
          ...use,
          path: `Chart.groups[${groupIndex}].dataLabels.style`,
        })),
      );
      fontUsages.push(...output.fontUsages);
      missingGlyphErrors.push(...output.missingGlyphErrors);
    }
    const groupLabelStyle =
      groupLabels !== undefined && groupLabels !== false
        ? groupLabels.style
        : undefined;
    for (const series of group.series) {
      for (const [pointIndex, point] of Object.entries(series.points ?? {})) {
        if (point.labelColor === undefined) continue;
        const index = Number(pointIndex);
        const pointLabels = effectivePointLabels(group, series, index);
        if (pointLabels === undefined || pointLabels === false) continue;
        const pointText = pointDataLabelTexts(
          group,
          series,
          pointLabels,
          index,
        );
        if (pointText.length === 0) continue;
        const output = writeChartTextProperties(
          mergeTextStyle(chartStyle, {
            ...groupLabelStyle,
            color: point.labelColor,
          }),
          pointText,
          "Chart.dataLabels.point.style",
          theme,
        );
        textStyles.pointLabels ??= new Map();
        textStyles.pointLabels.set(
          `${series.seriesIndex}:${pointIndex}`,
          output.xml,
        );
        resourceClaims.push(...output.resourceClaims);
        fontUses.push(
          ...output.fontUses.map((use) => ({
            ...use,
            path: "Chart.dataLabels.point.style",
          })),
        );
        fontUsages.push(...output.fontUsages);
        missingGlyphErrors.push(...output.missingGlyphErrors);
      }
    }
  }
  return {
    textStyles,
    resourceClaims,
    fontUses,
    fontUsages,
    missingGlyphErrors,
  };
}

export {
  createChartNode,
  type ChartAxis,
  type ChartAxisCrossing,
  type ChartAxisPosition,
  type ChartAxisReferences,
  type ChartCategoryAxis,
  type ChartCategoryAxisReferences,
  type ChartCustomErrorBars,
  type ChartDataLabelOverride,
  type ChartDataLabels,
  type ChartDisplayBlanksAs,
  type ChartErrorBarInclude,
  type ChartErrorBars,
  type ChartErrorBarValues,
  type ChartGroup,
  type ChartGroupData,
  type ChartKind,
  type ChartLine,
  type ChartLineDash,
  type ChartMarker,
  type ChartMarkerSymbol,
  type ChartNode,
  type ChartPoint,
  type ChartProps,
  type ChartSeries,
  type ChartSeriesDataLabelOverride,
  type ChartTextStyle,
  type ChartTickLabelPosition,
  type ChartTickMark,
  type ChartTitle,
  type ChartTrendline,
  type ChartTrendlineType,
  type ChartValueAxis,
  type ChartXYAxisReferences,
  type ColumnOrBarChartGroup,
  type DoughnutChartGroup,
  type IndependentCategoryData,
  type IndependentCategorySeries,
  type LineOrAreaChartGroup,
  type PieChartGroup,
  type ScatterChartGroup,
  type ScatterChartPoint,
  type ScatterChartSeries,
  type ScatterData,
  type ScatterDataLabelOverride,
  type ScatterDataLabels,
  type ScatterErrorBars,
  type ScatterLine,
  type SharedCategoryData,
} from "./chart-declaration.ts";
