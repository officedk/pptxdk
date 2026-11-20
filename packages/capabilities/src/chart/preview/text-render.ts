import type { GeometryBox, ThemeContext } from "@react-pptx/core/capability";

import {
  measureTextLayoutFromParagraphs,
  parseTextContent,
  renderTextOutlinesSvg,
  withPreviewFontBoundary,
  type FontSize,
  type Paragraph,
  type ResolvedTextStyle,
  type TextStyle,
} from "../../shared/text-service.ts";

export type ChartTextClipAllocator = {
  readonly defs: string[];
  nextClipId(): string;
};

export function createChartClipAllocator(
  slideNumber: number,
  objectId: number,
): ChartTextClipAllocator {
  let seq = 0;
  return {
    defs: [],
    nextClipId() {
      seq += 1;
      return `rp-chart-clip-${slideNumber}-${objectId}-${seq}`;
    },
  };
}

/** Converts a FontSize token into CSS/SVG pixels, falling back to 14px when omitted. */
export function resolveChartFontSizePixels(
  fontSize: FontSize | undefined,
): number {
  if (fontSize === undefined) return 14;
  if (typeof fontSize === "number") return fontSize;
  const parsed = Number.parseFloat(fontSize);
  if (!Number.isFinite(parsed) || parsed <= 0) return 14;
  return fontSize.endsWith("pt") ? (parsed * 4) / 3 : parsed;
}

export function mergeChartTextStyle(
  base: TextStyle | undefined,
  override: TextStyle | undefined,
): TextStyle | undefined {
  if (base === undefined) return override;
  if (override === undefined) return base;
  return { ...base, ...override };
}

export type MeasuredChartText = {
  readonly paragraphs: readonly Paragraph[];
  readonly resolvedStyle: ResolvedTextStyle;
  readonly width: number;
  readonly height: number;
  readonly minContentWidth: number;
};

/** Measures text using the shared text-service without mutating or inventing glyphs. */
export function measureChartText(options: {
  readonly text: string;
  readonly style: TextStyle | undefined;
  readonly theme: ThemeContext;
  readonly slideNumber: number;
  readonly labelPath: string;
  readonly availableWidth?: number;
}): MeasuredChartText {
  const { paragraphs, style: resolvedStyle } = parseTextContent(
    [{ type: "#text", text: options.text }],
    options.style,
    options.theme,
    { slideNumber: options.slideNumber },
    options.labelPath,
  );
  const measured = withPreviewFontBoundary(() =>
    measureTextLayoutFromParagraphs(
      paragraphs,
      resolvedStyle,
      options.availableWidth,
    ),
  );
  return {
    paragraphs,
    resolvedStyle,
    width: measured.width,
    height: measured.height,
    minContentWidth: measured.minContentWidth,
  };
}

/**
 * Renders vector glyph outlines inside `box`, adapting cleanly via optional line wrapping
 * and SVG clipPath bounds without ever emitting overflow errors or synthetic ellipsis glyphs.
 */
export function renderChartTextSvg(options: {
  readonly text: string;
  readonly box: GeometryBox;
  readonly style: TextStyle | undefined;
  readonly theme: ThemeContext;
  readonly slideNumber: number;
  readonly labelPath: string;
  readonly align?: "left" | "center" | "right";
  readonly verticalAlign?: "top" | "middle" | "bottom";
  readonly wrap?: boolean;
  readonly clipBox?: GeometryBox;
  readonly clipAllocator?: ChartTextClipAllocator;
  readonly rotationDegrees?: -90 | 90;
}): string {
  if (options.text === "") return "";

  const align = options.align ?? options.style?.align ?? "center";
  const verticalAlign =
    options.verticalAlign ?? options.style?.verticalAlign ?? "middle";
  const effectiveStyle: TextStyle = {
    ...options.style,
    align,
    verticalAlign,
  };

  const { paragraphs, style: resolvedStyle } = parseTextContent(
    [{ type: "#text", text: options.text }],
    effectiveStyle,
    options.theme,
    { slideNumber: options.slideNumber },
    options.labelPath,
  );

  const rawWidth = Math.max(1, options.box.width);
  const rawHeight = Math.max(1, options.box.height);

  let layoutBox: GeometryBox;
  if (options.rotationDegrees !== undefined) {
    const cx = options.box.x + options.box.width / 2;
    const cy = options.box.y + options.box.height / 2;
    layoutBox = {
      x: cx - rawHeight / 2,
      y: cy - rawWidth / 2,
      width: rawHeight,
      height: rawWidth,
    };
  } else {
    layoutBox = {
      x: options.box.x,
      y: options.box.y,
      width: rawWidth,
      height: rawHeight,
    };
  }

  const singleLineMeasure = withPreviewFontBoundary(() =>
    measureTextLayoutFromParagraphs(paragraphs, resolvedStyle),
  );
  const allowWrap = options.wrap ?? false;
  let effectiveLayoutBox = layoutBox;

  if (!allowWrap && singleLineMeasure.width > layoutBox.width) {
    const neededWidth = singleLineMeasure.width + 1;
    const delta = neededWidth - layoutBox.width;
    const shiftedX =
      align === "right"
        ? layoutBox.x - delta
        : align === "center"
          ? layoutBox.x - delta / 2
          : layoutBox.x;
    effectiveLayoutBox = {
      ...layoutBox,
      x: shiftedX,
      width: neededWidth,
    };
  }

  const wrappedMeasure = allowWrap
    ? withPreviewFontBoundary(() =>
        measureTextLayoutFromParagraphs(
          paragraphs,
          resolvedStyle,
          effectiveLayoutBox.width,
        ),
      )
    : singleLineMeasure;

  const idPrefix =
    options.clipAllocator !== undefined &&
    resolvedStyle.textShadow !== undefined
      ? options.clipAllocator.nextClipId()
      : `rp-chart-txt-${options.slideNumber}`;
  let content = withPreviewFontBoundary(() =>
    renderTextOutlinesSvg({
      paragraphs,
      box: effectiveLayoutBox,
      verticalAlign,
      theme: options.theme,
      idPrefix,
      rotate: options.rotationDegrees ?? 0,
    }),
  );
  if (content === "") return "";
  if (options.rotationDegrees !== undefined) {
    const cx = options.box.x + options.box.width / 2;
    const cy = options.box.y + options.box.height / 2;
    content = `<g transform="rotate(${options.rotationDegrees} ${cx} ${cy})">${content}</g>`;
  }

  const clipTarget = options.clipBox ?? options.box;
  const overflowsBounds =
    options.rotationDegrees === undefined
      ? wrappedMeasure.width > clipTarget.width + 0.5 ||
        wrappedMeasure.height > clipTarget.height + 0.5
      : wrappedMeasure.width > clipTarget.height + 0.5 ||
        wrappedMeasure.height > clipTarget.width + 0.5;

  if (
    overflowsBounds &&
    options.clipAllocator !== undefined &&
    clipTarget.width > 0 &&
    clipTarget.height > 0
  ) {
    const clipId = options.clipAllocator.nextClipId();
    options.clipAllocator.defs.push(
      `<clipPath id="${clipId}" clipPathUnits="userSpaceOnUse"><rect x="${clipTarget.x}" y="${clipTarget.y}" width="${clipTarget.width}" height="${clipTarget.height}"/></clipPath>`,
    );
    return `<g clip-path="url(#${clipId})">${content}</g>`;
  }

  return content;
}
