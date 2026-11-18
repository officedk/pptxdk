import { Resvg } from "@resvg/resvg-js";

import type { DeclarationContainer } from "#/declaration.ts";
import type { CommonNode } from "#/document.ts";
import {
  createDeclarationPreviewBuild,
  type PreviewSession,
} from "#/formal-build.ts";
import type {
  CapabilityModule,
  GeometryBox,
  ObjectObservation,
  PreviewDiagnostic,
  PreviewNodeObservation,
  PreviewOptions,
  PreviewPageResult,
  PreviewResult,
  PreviewTransform,
  ScreenGeometry,
} from "#/types.ts";

export type {
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

function runPreviewPipeline(
  buildForPreview: ReturnType<typeof createDeclarationPreviewBuild>,
  declaration: DeclarationContainer,
  options: PreviewOptions,
  checkSignal: boolean,
): PreviewResult {
  const { pages: requestedPages, ...buildOptions } = options;
  validateRequestedPages(requestedPages);
  const selected =
    requestedPages === undefined ? undefined : new Set(requestedPages);

  const dimensionsByPage = new Map<
    number,
    { readonly width: number; readonly height: number }
  >();
  const backgroundByPage = new Map<number, string>();
  const layersByPage = new Map<number, string>();
  const diagnosticsByPage = new Map<number, PreviewDiagnostic[]>();
  const nodePaths = new Map<string, string>();
  const transforms = new Map<string, PreviewTransform>();

  const session: PreviewSession = {
    shouldPreviewPage(page) {
      return selected === undefined || selected.has(page);
    },
    setSlideDimensions(page, dimensions) {
      dimensionsByPage.set(page, dimensions);
    },
    setSlideBackground(page, svg) {
      backgroundByPage.set(page, svg);
    },
    recordDiagnostic(page, entry) {
      const list = diagnosticsByPage.get(page) ?? [];
      list.push(Object.freeze({ ...entry }));
      diagnosticsByPage.set(page, list);
    },
    recordNode(page, objectId, path, transform) {
      const key = `${page}:${objectId}`;
      nodePaths.set(key, path);
      if (transform !== undefined) transforms.set(key, transform);
    },
    setPageLayers(page, svg) {
      layersByPage.set(page, svg);
    },
  };

  const build = buildForPreview(
    declaration,
    buildOptions,
    session,
    checkSignal,
  );
  if (requestedPages?.some((page) => page > build.report.slides.length)) {
    throw new RangeError("Preview page number is outside the document");
  }

  const pages: PreviewPageResult[] = [];
  for (const [index, slide] of build.report.slides.entries()) {
    const pageNumber = index + 1;
    if (selected !== undefined && !selected.has(pageNumber)) continue;

    const nodes: PreviewNodeObservation[] = [];
    const collect = (
      objects: readonly ObjectObservation[],
      ancestors: readonly PreviewTransform[] = [],
    ): void => {
      for (const object of objects) {
        const own = transforms.get(`${pageNumber}:${object.id}`);
        const chain = own === undefined ? ancestors : [own, ...ancestors];
        nodes.push(
          Object.freeze({
            ...object,
            ...(object.box === undefined
              ? {}
              : screenGeometry(object.box, chain)),
            path: nodePaths.get(`${pageNumber}:${object.id}`),
          }),
        );
        if (object.members !== undefined) collect(object.members, chain);
      }
    };
    collect(slide.objects);

    const pageDiagnostics = Object.freeze([
      ...(diagnosticsByPage.get(pageNumber) ?? []),
    ]);
    if (pageDiagnostics.some((item) => item.severity === "error")) {
      pages.push(
        Object.freeze({
          id: slide.id,
          page: pageNumber,
          nodes: Object.freeze(nodes),
          layouts: slide.layouts ?? Object.freeze([]),
          diagnostics: pageDiagnostics,
        }),
      );
      continue;
    }

    const dimensions = dimensionsByPage.get(pageNumber);
    if (dimensions === undefined) {
      throw new Error("Missing formal slide dimensions for preview");
    }
    const background =
      backgroundByPage.get(pageNumber) ?? defaultCanvasBackground;
    const layers = layersByPage.get(pageNumber) ?? "";
    const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${dimensions.width}" height="${dimensions.height}" viewBox="0 0 ${dimensions.width} ${dimensions.height}">${background}${layers}</svg>`;
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
        layouts: slide.layouts ?? Object.freeze([]),
        diagnostics: pageDiagnostics,
      }),
    );
  }

  return Object.freeze({
    build,
    pages: Object.freeze(pages),
  });
}

/**
 * Creates a static preview runner over a plain editable document without loading React.
 * Reuses a single formal build pass per invocation for PPTX, layout, diagnostics, and PNGs.
 */
export function createDocumentPreview(modules: readonly CapabilityModule[]) {
  const buildForPreview = createDeclarationPreviewBuild(modules);

  return async function previewDocument(
    document: CommonNode,
    options: PreviewOptions = {},
  ): Promise<PreviewResult> {
    return runPreviewPipeline(
      buildForPreview,
      { children: [document] },
      options,
      true,
    );
  };
}

/**
 * Internal preview entry for already-expanded declaration trees (e.g. JSX handoff).
 */
export function createCompletedDeclarationPreview(
  modules: readonly CapabilityModule[],
) {
  const buildForPreview = createDeclarationPreviewBuild(modules);

  return function previewCompletedDeclaration(
    declaration: DeclarationContainer,
    options: PreviewOptions = {},
  ): PreviewResult {
    return runPreviewPipeline(buildForPreview, declaration, options, false);
  };
}
