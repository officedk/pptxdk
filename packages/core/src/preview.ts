import { Resvg } from "@resvg/resvg-js";

import { CapabilityError } from "#/diagnostics.ts";
import {
  getInternalResolution,
  type ResolvedObjectLayer,
} from "#/resolution.ts";
import type {
  GeometryBox,
  ObjectObservation,
  PaintOptions,
  PaintResult,
  PreviewDiagnostic,
  PreviewNodeObservation,
  PreviewPageResult,
  PreviewTransform,
  Resolution,
  ScreenGeometry,
} from "#/types.ts";

export type {
  PaintOptions,
  PaintResult,
  PreviewDiagnostic,
  PreviewNodeObservation,
  PreviewOptions,
  PreviewPageResult,
  PreviewResult,
  PreviewSeverity,
  PreviewTransform,
  PreviewWarning,
  ScreenGeometry,
  ScreenPoint,
} from "#/types.ts";

const defaultCanvasBackground =
  '<rect width="100%" height="100%" fill="white"/>';

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
  const left = Math.min(...corners.map((point) => point.x));
  const top = Math.min(...corners.map((point) => point.y));
  return {
    corners: Object.freeze(corners.map((point) => Object.freeze({ ...point }))),
    screenBox: Object.freeze({
      x: left,
      y: top,
      width: Math.max(...corners.map((point) => point.x)) - left,
      height: Math.max(...corners.map((point) => point.y)) - top,
    }),
  };
}

function validateRequestedPages(pages: readonly number[] | undefined): void {
  if (pages?.some((page) => !Number.isInteger(page) || page < 1)) {
    throw new RangeError("Preview page numbers must be positive integers");
  }
}

function paintLayer(
  layer: ResolvedObjectLayer,
  pageNumber: number,
  diagnostics: PreviewDiagnostic[],
  nodePaths: Map<number, string>,
  transforms: Map<number, PreviewTransform>,
): string {
  const childrenSvg =
    layer.childComposition === "members"
      ? layer.memberLayers
          .map((child) =>
            paintLayer(child, pageNumber, diagnostics, nodePaths, transforms),
          )
          .join("")
      : layer.childComposition === "siblings"
        ? layer.siblingLayers
            .map((child) =>
              paintLayer(child, pageNumber, diagnostics, nodePaths, transforms),
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
    if (!layer.hasObservation) return;
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
        objectId: layer.paintObjectId,
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
  const { pages: requestedPages, signal } = options;
  validateRequestedPages(requestedPages);
  if (signal?.aborted) throw signal.reason;

  const internal = getInternalResolution(resolution);
  if (requestedPages?.some((page) => page > internal.slides.length)) {
    throw new RangeError("Preview page number is outside the document");
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

    const layers = slide.layers
      .map((layer) =>
        paintLayer(layer, pageNumber, diagnostics, nodePaths, transforms),
      )
      .join("");

    let background: string | undefined;
    for (const slideCap of slide.slideCapabilities) {
      const runSlidePaint = slideCap.paint;
      if (runSlidePaint !== undefined) {
        try {
          const previewOutput = runSlidePaint({
            path: slideCap.path,
            dimensions: internal.dimensions,
            theme: slideCap.theme,
            slideNumber: pageNumber,
          });
          if (slideCap.slot === "background") {
            background = previewOutput.svg;
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

    const nodes: PreviewNodeObservation[] = [];
    const collect = (
      objects: readonly ObjectObservation[],
      ancestors: readonly PreviewTransform[] = [],
    ): void => {
      for (const object of objects) {
        const own = transforms.get(object.id);
        const chain = own === undefined ? ancestors : [own, ...ancestors];
        nodes.push(
          Object.freeze({
            ...object,
            ...(object.box === undefined
              ? {}
              : screenGeometry(object.box, chain)),
            path: nodePaths.get(object.id),
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

    const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${internal.dimensions.width}" height="${internal.dimensions.height}" viewBox="0 0 ${internal.dimensions.width} ${internal.dimensions.height}">${background ?? defaultCanvasBackground}${layers}</svg>`;
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
