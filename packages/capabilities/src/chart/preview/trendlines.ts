import {
  CapabilityError,
  type ColorValue,
  type Rect,
  type ThemeContext,
} from "@pptxdk/core/capability";

import type { TextStyle } from "../../shared/text-service.ts";
import type { ParsedChart } from "../chart-model.ts";
import type { ChartTrendline } from "../chart-types.ts";
import { resolveChartLineStroke, strokeAttrs } from "./style.ts";
import {
  measureChartText,
  mergeChartTextStyle,
  renderChartTextSvg,
  resolveChartFontSizePixels,
  type ChartTextClipAllocator,
} from "./text-render.ts";
import {
  assertNever,
  type AxisScale,
  type CategoryAxisScale,
} from "./types.ts";

export type FitPoint = {
  readonly x: number;
  readonly y: number;
};

export type FittedTrendlineModel =
  | {
      readonly kind: "curve";
      readonly xMin: number;
      readonly xMax: number;
      readonly evalY: (x: number) => number;
      readonly rSquared: number;
      readonly equationText: string;
    }
  | {
      readonly kind: "movingAverage";
      readonly points: readonly FitPoint[];
    };

function solveGaussianPartialPivoting(
  matrix: number[][],
  rhs: number[],
): number[] | undefined {
  const n = rhs.length;
  const a = matrix.map((row) => [...row]);
  const b = [...rhs];

  for (let col = 0; col < n; col += 1) {
    let pivotRow = col;
    let pivotAbs = Math.abs(a[col]![col]!);
    for (let row = col + 1; row < n; row += 1) {
      const candidate = Math.abs(a[row]![col]!);
      if (candidate > pivotAbs) {
        pivotAbs = candidate;
        pivotRow = row;
      }
    }
    if (!Number.isFinite(pivotAbs) || pivotAbs < 1e-12) {
      return undefined;
    }
    if (pivotRow !== col) {
      const tmpRow = a[col]!;
      a[col] = a[pivotRow]!;
      a[pivotRow] = tmpRow;
      const tmpB = b[col]!;
      b[col] = b[pivotRow]!;
      b[pivotRow] = tmpB;
    }

    const diag = a[col]![col]!;
    for (let row = col + 1; row < n; row += 1) {
      const factor = a[row]![col]! / diag;
      if (factor === 0) continue;
      for (let k = col; k < n; k += 1) {
        a[row]![k] = a[row]![k]! - factor * a[col]![k]!;
      }
      b[row] = b[row]! - factor * b[col]!;
    }
  }

  const sol = Array.from({ length: n }, () => 0);
  for (let row = n - 1; row >= 0; row -= 1) {
    let sum = b[row]!;
    for (let col = row + 1; col < n; col += 1) {
      sum -= a[row]![col]! * sol[col]!;
    }
    const val = sum / a[row]![row]!;
    if (!Number.isFinite(val)) return undefined;
    sol[row] = val;
  }
  return sol;
}

function fitLinearOls(
  points: readonly FitPoint[],
  fixedIntercept?: number,
): { readonly slope: number; readonly intercept: number } | undefined {
  const n = points.length;
  if (fixedIntercept !== undefined) {
    if (n < 1) return undefined;
    let sumXX = 0;
    let sumXY = 0;
    for (const pt of points) {
      sumXX += pt.x * pt.x;
      sumXY += pt.x * (pt.y - fixedIntercept);
    }
    if (sumXX <= 1e-15) return undefined;
    return { slope: sumXY / sumXX, intercept: fixedIntercept };
  }
  if (n < 2) return undefined;
  let sumX = 0;
  let sumY = 0;
  for (const pt of points) {
    sumX += pt.x;
    sumY += pt.y;
  }
  const meanX = sumX / n;
  const meanY = sumY / n;
  let sxx = 0;
  let sxy = 0;
  for (const pt of points) {
    const dx = pt.x - meanX;
    sxx += dx * dx;
    sxy += dx * (pt.y - meanY);
  }
  if (sxx <= 1e-15) return undefined;
  const slope = sxy / sxx;
  return { slope, intercept: meanY - slope * meanX };
}

function computeRSquared(
  points: readonly FitPoint[],
  evalY: (x: number) => number,
): number {
  const n = points.length;
  if (n === 0) return 1;
  let sumY = 0;
  for (const pt of points) sumY += pt.y;
  const meanY = sumY / n;
  let ssTot = 0;
  let ssRes = 0;
  for (const pt of points) {
    const dy = pt.y - meanY;
    ssTot += dy * dy;
    const err = pt.y - evalY(pt.x);
    ssRes += err * err;
  }
  if (ssTot <= 1e-15) return 1;
  return Math.max(0, 1 - ssRes / ssTot);
}

function formatNum(value: number): string {
  const clean = Math.abs(value) < 1e-10 ? 0 : Number(value.toPrecision(4));
  return String(clean);
}

function formatLinearTerm(slope: number, variable: string): string {
  return `${formatNum(slope)}${variable}`;
}

function formatSignedOffset(value: number): string {
  const clean = Math.abs(value) < 1e-10 ? 0 : value;
  if (clean === 0) return "";
  return clean > 0 ? ` + ${formatNum(clean)}` : ` - ${formatNum(-clean)}`;
}

function formatPolynomialEquation(coeffsX: readonly number[]): string {
  const terms: string[] = [];
  for (let p = coeffsX.length - 1; p >= 0; p -= 1) {
    const c = coeffsX[p]!;
    if (Math.abs(c) < 1e-10 && p > 0) continue;
    if (p === 0 && Math.abs(c) < 1e-10 && terms.length > 0) continue;
    const absStr = formatNum(Math.abs(c));
    const varPart = p === 0 ? "" : p === 1 ? "x" : `x^${p}`;
    if (terms.length === 0) {
      terms.push(`${c < 0 ? "-" : ""}${absStr}${varPart}`);
    } else {
      terms.push(`${c < 0 ? " - " : " + "}${absStr}${varPart}`);
    }
  }
  return `y = ${terms.length > 0 ? terms.join("") : "0"}`;
}

export function collectCategoryFitPoints(
  values: readonly (number | null)[],
): FitPoint[] {
  const pts: FitPoint[] = [];
  for (let i = 0; i < values.length; i += 1) {
    const y = values[i];
    if (y !== null && y !== undefined && Number.isFinite(y)) {
      pts.push({ x: i + 1, y });
    }
  }
  return pts;
}

export function collectScatterFitPoints(
  xs: readonly (number | null)[],
  ys: readonly (number | null)[],
): FitPoint[] {
  const pts: FitPoint[] = [];
  const len = Math.min(xs.length, ys.length);
  for (let i = 0; i < len; i += 1) {
    const x = xs[i];
    const y = ys[i];
    if (
      x !== null &&
      x !== undefined &&
      Number.isFinite(x) &&
      y !== null &&
      y !== undefined &&
      Number.isFinite(y)
    ) {
      pts.push({ x, y });
    }
  }
  pts.sort((a, b) => a.x - b.x);
  return pts;
}

export function fitTrendline(
  points: readonly FitPoint[],
  trendline: ChartTrendline,
): FittedTrendlineModel | undefined {
  if (points.length === 0) return undefined;

  if (trendline.type === "movingAverage") {
    const p = trendline.period;
    if (points.length < p || p < 2) return undefined;
    const sorted = [...points].sort((a, b) => a.x - b.x);
    const out: FitPoint[] = [];
    let windowSum = 0;
    for (let i = 0; i < sorted.length; i += 1) {
      windowSum += sorted[i]!.y;
      if (i >= p) {
        windowSum -= sorted[i - p]!.y;
      }
      if (i >= p - 1) {
        out.push({ x: sorted[i]!.x, y: windowSum / p });
      }
    }
    return { kind: "movingAverage", points: out };
  }

  let rawMinX = points[0]!.x;
  let rawMaxX = points[0]!.x;
  for (const pt of points) {
    if (pt.x < rawMinX) rawMinX = pt.x;
    if (pt.x > rawMaxX) rawMaxX = pt.x;
  }
  let xMin = rawMinX - (trendline.backward ?? 0);
  const xMax = rawMaxX + (trendline.forward ?? 0);

  switch (trendline.type) {
    case "linear": {
      const fit = fitLinearOls(points, trendline.intercept);
      if (fit === undefined) return undefined;
      const evalY = (x: number) => fit.slope * x + fit.intercept;
      return {
        kind: "curve",
        xMin,
        xMax,
        evalY,
        rSquared: computeRSquared(points, evalY),
        equationText: `y = ${formatLinearTerm(fit.slope, "x")}${formatSignedOffset(fit.intercept)}`,
      };
    }
    case "exponential": {
      if (points.some((pt) => pt.y <= 0)) return undefined;
      if (trendline.intercept !== undefined && trendline.intercept <= 0) {
        return undefined;
      }
      const logPts = points.map((pt) => ({ x: pt.x, y: Math.log(pt.y) }));
      const fixedLnA =
        trendline.intercept !== undefined
          ? Math.log(trendline.intercept)
          : undefined;
      const fit = fitLinearOls(logPts, fixedLnA);
      if (fit === undefined) return undefined;
      const a = Math.exp(fit.intercept);
      const b = fit.slope;
      if (!Number.isFinite(a) || !Number.isFinite(b)) return undefined;
      const evalY = (x: number) => a * Math.exp(b * x);
      return {
        kind: "curve",
        xMin,
        xMax,
        evalY,
        rSquared: computeRSquared(points, evalY),
        equationText: `y = ${formatNum(a)}e^(${formatNum(b)}x)`,
      };
    }
    case "logarithmic": {
      if (points.some((pt) => pt.x <= 0)) return undefined;
      const logPts = points.map((pt) => ({ x: Math.log(pt.x), y: pt.y }));
      const fit = fitLinearOls(logPts);
      if (fit === undefined) return undefined;
      xMin = Math.max(1e-6, xMin);
      if (xMax <= xMin) return undefined;
      const evalY = (x: number) =>
        fit.slope * Math.log(Math.max(1e-12, x)) + fit.intercept;
      return {
        kind: "curve",
        xMin,
        xMax,
        evalY,
        rSquared: computeRSquared(points, evalY),
        equationText: `y = ${formatNum(fit.slope)}ln(x)${formatSignedOffset(fit.intercept)}`,
      };
    }
    case "power": {
      if (points.some((pt) => pt.x <= 0 || pt.y <= 0)) return undefined;
      const logPts = points.map((pt) => ({
        x: Math.log(pt.x),
        y: Math.log(pt.y),
      }));
      const fit = fitLinearOls(logPts);
      if (fit === undefined) return undefined;
      const a = Math.exp(fit.intercept);
      const b = fit.slope;
      if (!Number.isFinite(a) || !Number.isFinite(b)) return undefined;
      xMin = Math.max(1e-6, xMin);
      if (xMax <= xMin) return undefined;
      const evalY = (x: number) => a * Math.max(1e-12, x) ** b;
      return {
        kind: "curve",
        xMin,
        xMax,
        evalY,
        rSquared: computeRSquared(points, evalY),
        equationText: `y = ${formatNum(a)}x^${formatNum(b)}`,
      };
    }
    case "polynomial": {
      const k = trendline.order;
      if (trendline.intercept !== undefined) {
        const c0 = trendline.intercept;
        if (points.length < k) return undefined;
        let maxAbsX = 1;
        for (const pt of points) {
          maxAbsX = Math.max(maxAbsX, Math.abs(pt.x));
        }
        const mat = Array.from({ length: k }, () =>
          Array.from({ length: k }, () => 0),
        );
        const rhs = Array.from({ length: k }, () => 0);
        for (const pt of points) {
          const u = pt.x / maxAbsX;
          const resY = pt.y - c0;
          const uPows = Array.from({ length: 2 * k + 1 }, () => 1);
          for (let p = 1; p <= 2 * k; p += 1) {
            uPows[p] = uPows[p - 1]! * u;
          }
          for (let r = 0; r < k; r += 1) {
            rhs[r] = rhs[r]! + resY * uPows[r + 1]!;
            for (let c = 0; c < k; c += 1) {
              mat[r]![c] = mat[r]![c]! + uPows[r + c + 2]!;
            }
          }
        }
        const sol = solveGaussianPartialPivoting(mat, rhs);
        if (sol === undefined) return undefined;
        const coeffsX: number[] = [c0];
        for (let p = 1; p <= k; p += 1) {
          coeffsX.push(sol[p - 1]! / maxAbsX ** p);
        }
        const evalY = (x: number) => {
          const u = x / maxAbsX;
          let sum = c0;
          let uPow = u;
          for (let p = 0; p < k; p += 1) {
            sum += sol[p]! * uPow;
            uPow *= u;
          }
          return sum;
        };
        return {
          kind: "curve",
          xMin,
          xMax,
          evalY,
          rSquared: computeRSquared(points, evalY),
          equationText: formatPolynomialEquation(coeffsX),
        };
      }

      if (points.length < k + 1) return undefined;
      let sumX = 0;
      for (const pt of points) sumX += pt.x;
      const meanX = sumX / points.length;
      let scaleX = 1;
      for (const pt of points) {
        scaleX = Math.max(scaleX, Math.abs(pt.x - meanX));
      }
      const dim = k + 1;
      const mat = Array.from({ length: dim }, () =>
        Array.from({ length: dim }, () => 0),
      );
      const rhs = Array.from({ length: dim }, () => 0);
      for (const pt of points) {
        const u = (pt.x - meanX) / scaleX;
        const uPows = Array.from({ length: 2 * k + 1 }, () => 1);
        for (let p = 1; p <= 2 * k; p += 1) {
          uPows[p] = uPows[p - 1]! * u;
        }
        for (let r = 0; r < dim; r += 1) {
          rhs[r] = rhs[r]! + pt.y * uPows[r]!;
          for (let c = 0; c < dim; c += 1) {
            mat[r]![c] = mat[r]![c]! + uPows[r + c]!;
          }
        }
      }
      const sol = solveGaussianPartialPivoting(mat, rhs);
      if (sol === undefined) return undefined;

      const coeffsX = Array.from({ length: dim }, () => 0);
      for (let p = 0; p <= k; p += 1) {
        const factor = sol[p]! / scaleX ** p;
        let binom = 1;
        for (let j = 0; j <= p; j += 1) {
          coeffsX[j] = coeffsX[j]! + factor * binom * (-meanX) ** (p - j);
          binom = (binom * (p - j)) / (j + 1);
        }
      }
      const evalY = (x: number) => {
        const u = (x - meanX) / scaleX;
        let sum = 0;
        let uPow = 1;
        for (let p = 0; p <= k; p += 1) {
          sum += sol[p]! * uPow;
          uPow *= u;
        }
        return sum;
      };
      return {
        kind: "curve",
        xMin,
        xMax,
        evalY,
        rSquared: computeRSquared(points, evalY),
        equationText: formatPolynomialEquation(coeffsX),
      };
    }
    default:
      return assertNever(trendline);
  }
}

function categoryFitPositionToNormalized(
  catScale: CategoryAxisScale,
  oneBasedPos: number,
): number {
  const index = oneBasedPos - 1;
  const rawT =
    catScale.crossBetween === "midCat" && catScale.count > 1
      ? index / (catScale.count - 1)
      : (index + 0.5) / catScale.count;
  return catScale.reversed ? 1 - rawT : rawT;
}

function renderSafeTrendlineText(options: {
  readonly candidates: readonly string[];
  readonly box: Rect;
  readonly style: TextStyle | undefined;
  readonly theme: ThemeContext;
  readonly slideNumber: number;
  readonly labelPath: string;
  readonly clipBox: Rect;
  readonly clipAllocator: ChartTextClipAllocator;
}): string {
  for (const text of options.candidates) {
    if (text === "") continue;
    try {
      return renderChartTextSvg({
        text,
        box: options.box,
        style: options.style,
        theme: options.theme,
        slideNumber: options.slideNumber,
        labelPath: options.labelPath,
        align: "right",
        verticalAlign: "middle",
        wrap: false,
        clipBox: options.clipBox,
        clipAllocator: options.clipAllocator,
      });
    } catch (error) {
      // Never allow synthetic trendline annotation characters (e.g. superscript 2 or equation letters)
      // to trigger missing-glyph failures on restricted subset fonts; still propagate unsupported-font.
      if (
        error instanceof CapabilityError &&
        error.code === "preview-missing-glyph"
      ) {
        continue;
      }
      throw error;
    }
  }
  return "";
}

function safeMeasureCandidate(options: {
  readonly candidates: readonly string[];
  readonly style: TextStyle | undefined;
  readonly theme: ThemeContext;
  readonly slideNumber: number;
  readonly labelPath: string;
}):
  | { readonly text: string; readonly width: number; readonly height: number }
  | undefined {
  for (const text of options.candidates) {
    if (text === "") continue;
    try {
      const m = measureChartText({
        text,
        style: options.style,
        theme: options.theme,
        slideNumber: options.slideNumber,
        labelPath: options.labelPath,
      });
      return { text, width: m.width, height: m.height };
    } catch (error) {
      if (
        error instanceof CapabilityError &&
        error.code === "preview-missing-glyph"
      ) {
        continue;
      }
      throw error;
    }
  }
  return undefined;
}

export function renderTrendlinesSvg(options: {
  readonly parsed: ParsedChart;
  readonly scales: ReadonlyMap<string, AxisScale>;
  readonly plotRect: Rect;
  readonly chartBox: Rect;
  readonly baseChartTextStyle: TextStyle | undefined;
  readonly theme: ThemeContext;
  readonly slideNumber: number;
  readonly clipAllocator: ChartTextClipAllocator;
}): { readonly linesSvg: string; readonly labelsSvg: string } {
  const {
    parsed,
    scales,
    plotRect,
    chartBox,
    baseChartTextStyle,
    theme,
    slideNumber,
    clipAllocator,
  } = options;
  const lineParts: string[] = [];
  const labelParts: string[] = [];

  for (let groupIndex = 0; groupIndex < parsed.groups.length; groupIndex += 1) {
    const group = parsed.groups[groupIndex]!;
    if (group.kind === "pie" || group.kind === "doughnut") continue;

    const seriesTargets: {
      readonly seriesIndex: number;
      readonly seriesColor: ColorValue | undefined;
      readonly trendlines: readonly ChartTrendline[];
      readonly fitPoints: FitPoint[];
      readonly mapDataToScreen: (
        xVal: number,
        yVal: number,
      ) => { readonly x: number; readonly y: number } | undefined;
      readonly isLinearCartesian: boolean;
    }[] = [];

    if (group.kind === "scatter") {
      const xScale = scales.get(group.axes.x.name);
      const yScale = scales.get(group.axes.y.name);
      if (xScale?.kind !== "value" || yScale?.kind !== "value") continue;
      for (const series of group.series) {
        if (series.trendlines === undefined || series.trendlines.length === 0) {
          continue;
        }
        seriesTargets.push({
          seriesIndex: series.seriesIndex,
          seriesColor: series.color,
          trendlines: series.trendlines,
          fitPoints: collectScatterFitPoints(series.x, series.y),
          isLinearCartesian: false,
          mapDataToScreen: (xVal, yVal) => {
            if (
              !Number.isFinite(xVal) ||
              !Number.isFinite(yVal) ||
              (xScale.logBase !== undefined && xVal <= 0) ||
              (yScale.logBase !== undefined && yVal <= 0)
            ) {
              return undefined;
            }
            return {
              x: xScale.valueToScreen(xVal, plotRect),
              y: yScale.valueToScreen(yVal, plotRect),
            };
          },
        });
      }
    } else {
      const catScale = scales.get(group.axes.category.name);
      const valScale = scales.get(group.axes.value.name);
      if (catScale?.kind !== "category" || valScale?.kind !== "value") {
        continue;
      }
      const isHorizontalBar = group.kind === "bar";
      const isLinearCartesian = valScale.logBase === undefined;
      for (const series of group.series) {
        if (series.trendlines === undefined || series.trendlines.length === 0) {
          continue;
        }
        seriesTargets.push({
          seriesIndex: series.seriesIndex,
          seriesColor: series.color,
          trendlines: series.trendlines,
          fitPoints: collectCategoryFitPoints(series.values),
          isLinearCartesian,
          mapDataToScreen: (xVal, yVal) => {
            if (
              !Number.isFinite(xVal) ||
              !Number.isFinite(yVal) ||
              (valScale.logBase !== undefined && yVal <= 0)
            ) {
              return undefined;
            }
            const catScreen = catScale.mapScreen(
              categoryFitPositionToNormalized(catScale, xVal),
              plotRect,
            );
            const valScreen = valScale.valueToScreen(yVal, plotRect);
            return isHorizontalBar
              ? { x: valScreen, y: catScreen }
              : { x: catScreen, y: valScreen };
          },
        });
      }
    }

    for (const target of seriesTargets) {
      for (const trendline of target.trendlines) {
        const model = fitTrendline(target.fitPoints, trendline);
        if (model === undefined) continue;

        const stroke = resolveChartLineStroke(trendline.line, theme, {
          color: target.seriesColor ?? `accent${(target.seriesIndex % 6) + 1}`,
          width: 1.5,
        });
        if (stroke === undefined) continue;

        const screenPts: { readonly x: number; readonly y: number }[] = [];
        if (model.kind === "movingAverage") {
          for (const pt of model.points) {
            const mapped = target.mapDataToScreen(pt.x, pt.y);
            if (mapped !== undefined) screenPts.push(mapped);
          }
        } else {
          const sampleCount =
            trendline.type === "linear" && target.isLinearCartesian ? 1 : 64;
          for (let k = 0; k <= sampleCount; k += 1) {
            const xVal =
              model.xMin + ((model.xMax - model.xMin) * k) / sampleCount;
            const yVal = model.evalY(xVal);
            const mapped = target.mapDataToScreen(xVal, yVal);
            if (mapped !== undefined) screenPts.push(mapped);
          }
        }

        if (screenPts.length >= 2) {
          const d =
            `M${screenPts[0]!.x} ${screenPts[0]!.y}` +
            screenPts
              .slice(1)
              .map((p) => `L${p.x} ${p.y}`)
              .join("");
          lineParts.push(
            `<path d="${d}" fill="none" ${strokeAttrs(stroke)} stroke-linecap="round" stroke-linejoin="round"/>`,
          );
        }

        if (
          model.kind === "curve" &&
          (trendline.displayEquation === true ||
            trendline.displayRSquared === true)
        ) {
          const anchorPt = screenPts[screenPts.length - 1];
          if (anchorPt === undefined) continue;

          const labelStyle = mergeChartTextStyle(baseChartTextStyle, {
            fontSize: Math.max(
              9,
              resolveChartFontSizePixels(baseChartTextStyle?.fontSize) * 0.85,
            ),
          });
          const labelPath = `Chart.groups[${groupIndex}].trendline`;
          const r2Val = formatNum(model.rSquared);

          const lines: { readonly candidates: readonly string[] }[] = [];
          if (trendline.displayEquation === true) {
            lines.push({ candidates: [model.equationText] });
          }
          if (trendline.displayRSquared === true) {
            lines.push({
              candidates: [`R\u00B2 = ${r2Val}`, `R^2 = ${r2Val}`],
            });
          }

          const lineSpacing =
            resolveChartFontSizePixels(labelStyle?.fontSize) * 1.25;
          let cursorY = Math.max(
            plotRect.y + 2,
            Math.min(
              plotRect.y + plotRect.height - lines.length * lineSpacing - 2,
              anchorPt.y - lines.length * lineSpacing - 4,
            ),
          );

          for (const lineSpec of lines) {
            const measured = safeMeasureCandidate({
              candidates: lineSpec.candidates,
              style: labelStyle,
              theme,
              slideNumber,
              labelPath,
            });
            if (measured === undefined) continue;
            const boxW = Math.max(20, measured.width + 4);
            const boxX = Math.max(
              plotRect.x + 2,
              Math.min(
                plotRect.x + plotRect.width - boxW - 2,
                anchorPt.x - boxW,
              ),
            );
            const rendered = renderSafeTrendlineText({
              candidates: [measured.text],
              box: {
                x: boxX,
                y: cursorY,
                width: boxW,
                height: Math.max(measured.height, lineSpacing),
              },
              style: labelStyle,
              theme,
              slideNumber,
              labelPath,
              clipBox: chartBox,
              clipAllocator,
            });
            if (rendered !== "") {
              labelParts.push(rendered);
              cursorY += lineSpacing;
            }
          }
        }
      }
    }
  }

  return {
    linesSvg: lineParts.join(""),
    labelsSvg: labelParts.join(""),
  };
}
