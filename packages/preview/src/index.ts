import { createRequire } from "node:module";

import {
  CapabilityError,
  type BuildQuery,
  type Resolution,
  type ResolutionReport,
  type SlideLocator,
} from "@pptxdk/core/build";
import type {
  GeometryBox,
  PreviewTransform,
  PreviewWarning,
  Rect,
  ResolvedLayout,
  ResolvedObject,
} from "@pptxdk/core/capability";
import {
  getInternalResolution,
  type ResolvedObjectLayer,
} from "@pptxdk/core/resolution";
import type { Resvg as ResvgClass } from "@resvg/resvg-js";

import {
  PreviewPageRangeError,
  paintOptionsSchema,
  previewDiagnosticSchema,
  previewOptionsSchema,
  previewPageNumberSchema,
  previewPagesSchema,
  previewSeveritySchema,
  screenGeometrySchema,
  screenPointSchema,
  validatePaintOptions,
  type PaintOptions,
  type PreviewDiagnostic,
  type PreviewOptions,
  type PreviewSeverity,
  type ScreenGeometry,
  type ScreenPoint,
} from "#/options.ts";

export type { PreviewTransform, PreviewWarning };
export {
  PreviewPageRangeError,
  paintOptionsSchema,
  previewDiagnosticSchema,
  previewOptionsSchema,
  previewPageNumberSchema,
  previewPagesSchema,
  previewSeveritySchema,
  screenGeometrySchema,
  screenPointSchema,
  validatePaintOptions,
};
export type {
  PaintOptions,
  PreviewDiagnostic,
  PreviewOptions,
  PreviewSeverity,
  ScreenGeometry,
  ScreenPoint,
};

export type PreviewNode = Omit<ResolvedObject, "corners"> & {
  readonly path: string;
  readonly bounds?: Rect;
  readonly corners?: readonly ScreenPoint[];
  /** Transitional alias retained until Ticket 11. */
  readonly screenBox?: Rect;
};

/** Transitional alias retained until Ticket 11. */
export type PreviewNodeObservation = PreviewNode;

export type PreviewPageResult = {
  readonly id: string;
  readonly page: number;
  readonly svg?: string;
  readonly png?: Uint8Array;
  readonly nodes: readonly PreviewNode[];
  readonly layouts: readonly ResolvedLayout[];
  readonly diagnostics: readonly PreviewDiagnostic[];
};

export type PaintResult = {
  readonly pages: readonly PreviewPageResult[];
};

export type PreviewResult = {
  readonly report: ResolutionReport;
  readonly pages: readonly PreviewPageResult[];
  readonly slides: readonly SlideLocator[];
  slide(id: string): SlideLocator;
  /** Transitional property retained until Ticket 11. */
  readonly query: BuildQuery;
};

const defaultCanvasBackground =
  '<rect width="100%" height="100%" fill="white"/>';

function roundScreenMetric(value: number): number {
  const rounded = Math.round(value * 1e9) / 1e9;
  return Object.is(rounded, -0) ? 0 : rounded;
}

/**
 * Computes transformed screen corners and axis-aligned screen bounds for a layout box.
 * Applies the local object transform first, followed by enclosing group transforms outward.
 */
export function screenGeometry(
  box: GeometryBox,
  transforms: readonly PreviewTransform[],
): ScreenGeometry {
  const corners: { x: number; y: number }[] = [
    { x: box.x, y: box.y },
    { x: box.x + box.width, y: box.y },
    { x: box.x + box.width, y: box.y + box.height },
    { x: box.x, y: box.y + box.height },
  ];
  for (const transform of transforms) {
    const cx = transform.box.x + transform.box.width / 2;
    const cy = transform.box.y + transform.box.height / 2;
    const angle = (transform.angle * Math.PI) / 180;
    for (const point of corners) {
      const x = (point.x - cx) * (transform.flipH ? -1 : 1);
      const y = (point.y - cy) * (transform.flipV ? -1 : 1);
      point.x = cx + x * Math.cos(angle) - y * Math.sin(angle);
      point.y = cy + x * Math.sin(angle) + y * Math.cos(angle);
    }
  }
  const rounded = corners.map((point) =>
    Object.freeze({
      x: roundScreenMetric(point.x),
      y: roundScreenMetric(point.y),
    }),
  );
  const left = Math.min(...rounded.map((point) => point.x));
  const top = Math.min(...rounded.map((point) => point.y));
  const right = Math.max(...rounded.map((point) => point.x));
  const bottom = Math.max(...rounded.map((point) => point.y));
  return {
    corners: Object.freeze(rounded),
    screenBox: Object.freeze({
      x: left,
      y: top,
      width: roundScreenMetric(right - left),
      height: roundScreenMetric(bottom - top),
    }),
  };
}

let cachedResvg: typeof ResvgClass | undefined;

function getResvg(): typeof ResvgClass {
  if (cachedResvg === undefined) {
    const require = createRequire(import.meta.url);
    const loaded: unknown = require("@resvg/resvg-js");
    if (
      typeof loaded !== "object" ||
      loaded === null ||
      !("Resvg" in loaded) ||
      typeof loaded.Resvg !== "function"
    ) {
      throw new Error("Failed to load @resvg/resvg-js Resvg constructor");
    }
    cachedResvg = loaded.Resvg as typeof ResvgClass;
  }
  return cachedResvg;
}

function maxPaintObjectId(layers: readonly ResolvedObjectLayer[]): number {
  let maxId = 0;
  for (const layer of layers) {
    if (layer.paintObjectId > maxId) maxId = layer.paintObjectId;
    const memberMax = maxPaintObjectId(layer.memberLayers);
    if (memberMax > maxId) maxId = memberMax;
    const siblingMax = maxPaintObjectId(layer.siblingLayers);
    if (siblingMax > maxId) maxId = siblingMax;
  }
  return maxId;
}

function paintLayer(
  layer: ResolvedObjectLayer,
  pageNumber: number,
  diagnostics: PreviewDiagnostic[],
  nodePaths: Map<number, string>,
  transforms: Map<number, PreviewTransform>,
  paintIdOffset = 0,
  recordNodes = true,
): string {
  const childrenSvg =
    layer.childComposition === "members"
      ? layer.memberLayers
          .map((child) =>
            paintLayer(
              child,
              pageNumber,
              diagnostics,
              nodePaths,
              transforms,
              paintIdOffset,
              recordNodes,
            ),
          )
          .join("")
      : layer.childComposition === "siblings"
        ? layer.siblingLayers
            .map((child) =>
              paintLayer(
                child,
                pageNumber,
                diagnostics,
                nodePaths,
                transforms,
                paintIdOffset,
                recordNodes,
              ),
            )
            .join("")
        : undefined;

  if (
    layer.childComposition !== "leaf" &&
    layer.emitsObject &&
    !layer.hasObservation
  ) {
    return "";
  }

  const recordNodeIdentity = (): void => {
    if (!recordNodes || !layer.hasObservation) return;
    nodePaths.set(layer.objectId, layer.path);
    if (layer.transform !== undefined) {
      transforms.set(layer.objectId, layer.transform);
    }
  };

  if (layer.paint !== undefined) {
    try {
      const output = layer.paint({
        path: layer.path,
        slideNumber: pageNumber,
        objectId: layer.paintObjectId + paintIdOffset,
        ...(layer.box === undefined ? {} : { box: layer.box }),
        theme: layer.theme,
        ...(childrenSvg === undefined ? {} : { childrenSvg }),
      });
      recordNodeIdentity();
      for (const warning of output.warnings ?? []) {
        diagnostics.push(
          Object.freeze({
            severity: "warning",
            page: pageNumber,
            path:
              warning.path === undefined
                ? layer.path
                : `${layer.path}/${warning.path}`,
            message: warning.message,
          }),
        );
      }
      return output.svg;
    } catch (error) {
      recordNodeIdentity();
      diagnostics.push(
        Object.freeze({
          severity: "error",
          page: pageNumber,
          path:
            error instanceof CapabilityError && error.path !== undefined
              ? `${layer.path}/${error.path}`
              : layer.path,
          message: error instanceof Error ? error.message : String(error),
        }),
      );
      return "";
    }
  }

  recordNodeIdentity();
  diagnostics.push(
    Object.freeze({
      severity: "error",
      page: pageNumber,
      path: layer.path,
      message: `Preview does not support ${layer.declarationType}`,
    }),
  );
  return "";
}

/** Paints static preview pages directly from a Resolution without running emit(). */
export function paint(
  resolution: Resolution,
  options: PaintOptions = {},
): PaintResult {
  const { pages: requestedPages, signal } = validatePaintOptions(options);
  if (signal?.aborted) throw signal.reason;

  const internal = getInternalResolution(resolution);
  if (requestedPages?.some((page) => page > internal.slides.length)) {
    throw new PreviewPageRangeError(requestedPages, internal.slides.length);
  }
  const selected =
    requestedPages === undefined ? undefined : new Set(requestedPages);

  const pages: PreviewPageResult[] = [];
  for (const [index, slide] of internal.slides.entries()) {
    const pageNumber = slide.index;
    if (selected !== undefined && !selected.has(pageNumber)) continue;

    const slideObservation = resolution.report.slides[index];
    if (slideObservation === undefined) {
      throw new Error("Missing resolved slide observation for paint");
    }

    const diagnostics: PreviewDiagnostic[] = [];
    const nodePaths = new Map<number, string>();
    const transforms = new Map<number, PreviewTransform>();

    const masterOffset = maxPaintObjectId(slide.layers) + 1000;
    const wrapSharedObjectSvg = (svg: string): string =>
      svg !== "" && internal.mastersCanvas !== undefined
        ? internal.mastersCanvas.paintSharedObject(svg).svg
        : svg;
    const masterLayerFragments = (slide.masterLayers ?? []).map((layer) =>
      wrapSharedObjectSvg(
        paintLayer(
          layer,
          pageNumber,
          diagnostics,
          nodePaths,
          transforms,
          masterOffset,
          false,
        ),
      ),
    );
    const masterLayersSvg =
      internal.mastersCanvas !== undefined
        ? internal.mastersCanvas.paintMasterLayers(masterLayerFragments).svg
        : masterLayerFragments.join("");
    const layoutOffset =
      masterOffset + maxPaintObjectId(slide.masterLayers ?? []) + 1000;
    const layoutLayerFragments = (slide.layoutLayers ?? []).map((layer) =>
      wrapSharedObjectSvg(
        paintLayer(
          layer,
          pageNumber,
          diagnostics,
          nodePaths,
          transforms,
          layoutOffset,
          false,
        ),
      ),
    );
    const layoutLayersSvg =
      internal.mastersCanvas !== undefined
        ? internal.mastersCanvas.paintLayoutLayers(layoutLayerFragments).svg
        : layoutLayerFragments.join("");

    const layers = slide.layers
      .map((layer) =>
        paintLayer(layer, pageNumber, diagnostics, nodePaths, transforms),
      )
      .join("");

    let background: string | undefined;
    let shapeTreeCapSvg = "";
    let hasExplicitSlideBackground = false;
    for (const slideCap of slide.slideCapabilities) {
      if (slideCap.slot === "background" && slideCap.hasSlideProps) {
        hasExplicitSlideBackground = true;
      }
      const runSlidePaint = slideCap.paint;
      if (runSlidePaint !== undefined) {
        try {
          const previewOutput = runSlidePaint({
            path: slideCap.path,
            dimensions: internal.dimensions,
            theme: slideCap.theme,
            slideNumber: pageNumber,
          });
          if (slideCap.slot === "background" && previewOutput.svg !== "") {
            background = previewOutput.svg;
          } else if (
            slideCap.slot === "shape-tree" &&
            previewOutput.svg !== ""
          ) {
            shapeTreeCapSvg += previewOutput.svg;
          }
          for (const warning of previewOutput.warnings ?? []) {
            diagnostics.push(
              Object.freeze({
                severity: "warning",
                page: pageNumber,
                path:
                  warning.path === undefined
                    ? slideCap.path
                    : `${slideCap.path}/${warning.path}`,
                message: warning.message,
              }),
            );
          }
        } catch (error) {
          diagnostics.push(
            Object.freeze({
              severity: "error",
              page: pageNumber,
              path:
                error instanceof CapabilityError && error.path !== undefined
                  ? error.path
                  : slideCap.path,
              message: error instanceof Error ? error.message : String(error),
            }),
          );
        }
      } else if (slideCap.slot === "background" && slideCap.hasSlideProps) {
        diagnostics.push(
          Object.freeze({
            severity: "error",
            page: pageNumber,
            path: slideCap.path,
            message: `Preview does not support ${slideCap.declaration}`,
          }),
        );
      }
    }
    if (
      !hasExplicitSlideBackground &&
      slide.inheritedBackground?.paint !== undefined
    ) {
      try {
        const inheritedOutput = slide.inheritedBackground.paint({
          path: slide.inheritedBackground.path,
          dimensions: internal.dimensions,
          theme: slide.inheritedBackground.theme,
          slideNumber: pageNumber,
        });
        if (inheritedOutput.svg !== "") {
          background = inheritedOutput.svg;
        }
      } catch (error) {
        diagnostics.push(
          Object.freeze({
            severity: "error",
            page: pageNumber,
            path:
              error instanceof CapabilityError && error.path !== undefined
                ? error.path
                : slide.inheritedBackground.path,
            message: error instanceof Error ? error.message : String(error),
          }),
        );
      }
    }

    const nodes: PreviewNode[] = [];
    const collect = (
      objects: readonly ResolvedObject[],
      ancestors: readonly PreviewTransform[] = [],
    ): void => {
      for (const object of objects) {
        const own = transforms.get(object.id);
        const chain = own === undefined ? ancestors : [own, ...ancestors];
        const screen =
          object.box === undefined
            ? undefined
            : screenGeometry(object.box, chain);
        nodes.push(
          Object.freeze({
            ...object,
            ...(screen === undefined
              ? {}
              : {
                  corners: screen.corners,
                  bounds: screen.screenBox,
                  screenBox: screen.screenBox,
                }),
            path: nodePaths.get(object.id) ?? object.path,
          }),
        );
        if (object.members !== undefined) collect(object.members, chain);
      }
    };
    collect(slideObservation.objects);

    const pageDiagnostics = Object.freeze([...diagnostics]);
    if (pageDiagnostics.some((item) => item.severity === "error")) {
      pages.push(
        Object.freeze({
          id: slide.id,
          page: pageNumber,
          nodes: Object.freeze(nodes),
          layouts: slideObservation.layouts ?? Object.freeze([]),
          diagnostics: pageDiagnostics,
        }),
      );
      continue;
    }

    const fallbackBackground =
      internal.mastersCanvas !== undefined
        ? internal.mastersCanvas.paintBackdrop({
            theme: slide.theme,
            ...(slide.hasCustomColorMap ? { hasCustomColorMap: true } : {}),
            ...(slide.effectiveColorMapOverride === undefined
              ? {}
              : { colorMapOverride: slide.effectiveColorMapOverride }),
            ...(slide.effectiveThemeOverride === undefined
              ? {}
              : { themeOverride: slide.effectiveThemeOverride }),
          }).svg
        : defaultCanvasBackground;
    const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${internal.dimensions.width}" height="${internal.dimensions.height}" viewBox="0 0 ${internal.dimensions.width} ${internal.dimensions.height}">${background ?? fallbackBackground}${masterLayersSvg}${layoutLayersSvg}${shapeTreeCapSvg}${layers}</svg>`;
    const Resvg = getResvg();
    const png = new Uint8Array(
      new Resvg(svg, { font: { loadSystemFonts: false } }).render().asPng(),
    );
    pages.push(
      Object.freeze({
        id: slide.id,
        page: pageNumber,
        svg,
        png,
        nodes: Object.freeze(nodes),
        layouts: slideObservation.layouts ?? Object.freeze([]),
        diagnostics: pageDiagnostics,
      }),
    );
  }

  return Object.freeze({ pages: Object.freeze(pages) });
}
