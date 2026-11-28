import {
  CapabilityError,
  parseCapabilityLayoutContract,
  type BoundObjectCapability,
  type BoundObjectNativeCollaboration,
  type CapabilityInput,
  type CapabilityLayoutMeasurement,
  type CapabilityLayoutRequest,
  type CapabilityParseContext,
  type EmitContext,
  type EmitOutput,
  type ObjectCapabilityModule,
  type ObjectResolution,
  type ObjectResolveContext,
  type PaintContext,
  type PaintOutput,
  type PreviewTransform,
  type ScopeTargetObject,
  type ScreenPoint,
  type SealedCanvasMechanism,
} from "@pptxdk/core/capability";
import { pxToEmu, xmlEscape } from "@pptxdk/emit";
import { defineCanvasMechanism } from "@pptxdk/import";

import { motionAnnotations } from "../motion/motion-api.ts";
import { motionPropsShape } from "../motion/motion-schemas.ts";
import type { Geometry, PresetGeometry } from "../shared/geometry.ts";
import {
  nativeLinkXml,
  parseLinkTarget,
  parseMorphKey,
  parseOptionalName,
} from "../shared/identity.ts";
import {
  assertVisualBoxExtent,
  emitVisualModel,
  paintVisualModel,
  resolveLineStroke,
  resolveVisualModel,
  type ResolvedVisualModel,
} from "../shared/visual.ts";
import { lineCodegen } from "./line-codegen.ts";
import {
  lineAdjustSchema,
  lineAllowedPropKeyList,
  lineEndpointSchema,
  lineRoutingSchema,
  lineStyleKeyList,
  type LineAdjust,
  type LineEndpoint,
  type LinePort,
  type LineRouting,
} from "./line-declaration.ts";
import {
  connectorLineImportHandler,
  presetLineShapeImportHandler,
} from "./line-import.ts";
import { lineXml, type LineConnectionXml } from "./line-xml.ts";

export {
  createLineNode,
  lineAdjustSchema,
  lineAllowedPropKeyList,
  lineDeclarationPropsSchema,
  lineEndpointSchema,
  linePortSchema,
  lineRoutingSchema,
  lineStrokeSchema,
  lineStrokeValueSchema,
  lineStyleKeyList,
  lineStyleSchema,
  type ArrowType,
  type LineAdjust,
  type LineDeclarationProps,
  type LineEndpoint,
  type LineLayoutStyle,
  type LinePort,
  type LineRouting,
  type LineStroke,
  type LineStrokeValue,
  type LineStyle,
  type LineVisualStyle,
} from "./line-declaration.ts";

const allowedParameters: readonly string[] = lineAllowedPropKeyList;
export const lineStyleKeys = lineStyleKeyList;
const allowedStyleKeySet = new Set<string>(lineStyleKeyList);
const endpointModeAllowedStyleKeys = new Set<string>([
  "display",
  "zIndex",
  "opacity",
  "boxShadow",
]);

function record(input: unknown): Record<string, unknown> {
  return typeof input === "object" && input !== null && !Array.isArray(input)
    ? (input as Record<string, unknown>)
    : {};
}

function validateCoordinate(value: unknown, name: "x" | "y"): number {
  if (
    typeof value !== "number" ||
    !Number.isFinite(value) ||
    value < -1_000_000 ||
    value > 1_000_000
  ) {
    throw new CapabilityError(
      "invalid-coordinate",
      `${name} must be a finite pixel number between -1000000 and 1000000`,
    );
  }
  return value;
}

function validateNonNegativeSize(
  value: unknown,
  name: "width" | "height",
): number {
  if (typeof value !== "number" || !Number.isFinite(value) || value < 0) {
    throw new CapabilityError(
      "invalid-coordinate",
      `${name} must be a non-negative finite pixel number`,
    );
  }
  const extentEmu = pxToEmu(value);
  if (
    !Number.isSafeInteger(extentEmu) ||
    extentEmu < 0 ||
    extentEmu > 2_147_483_647
  ) {
    throw new CapabilityError(
      "invalid-coordinate",
      `${name} must convert to an EMU value between 0 and 2147483647 inclusive`,
    );
  }
  return value;
}

function resolveConnectorGeometry(
  rawRouting: unknown,
  rawAdjust: unknown,
): { readonly routing: LineRouting; readonly geometry: Geometry } {
  const routingResult =
    rawRouting === undefined
      ? { success: true as const, data: "straight" as const }
      : lineRoutingSchema.safeParse(rawRouting);
  if (!routingResult.success) {
    throw new CapabilityError(
      "invalid-geometry",
      `Unsupported line routing ${String(rawRouting)}`,
    );
  }
  const routing = routingResult.data;

  let parsedAdjust: LineAdjust | undefined;
  if (rawAdjust !== undefined) {
    const adjustResult = lineAdjustSchema.safeParse(rawAdjust);
    if (!adjustResult.success) {
      const issue = adjustResult.error.issues[0];
      const msg =
        issue?.code === "unrecognized_keys" && issue.keys[0] !== undefined
          ? `Line adjust contains unsupported key ${issue.keys[0]} (only adj1, adj2, adj3 are supported)`
          : (issue?.message ?? "Invalid line adjust parameters");
      throw new CapabilityError("invalid-geometry", msg);
    }
    parsedAdjust = adjustResult.data;
  }

  if (routing === "straight") {
    if (parsedAdjust !== undefined && Object.keys(parsedAdjust).length > 0) {
      throw new CapabilityError(
        "invalid-geometry",
        "Straight line routing does not support adjust parameters",
      );
    }
    return { routing, geometry: "straightConnector1" };
  }

  const prefix = routing === "elbow" ? "bentConnector" : "curvedConnector";
  if (parsedAdjust === undefined) {
    return { routing, geometry: `${prefix}3` as const };
  }
  const definedEntries = Object.entries(parsedAdjust).filter(
    ([, val]) => val !== undefined,
  );
  if (definedEntries.length === 0) {
    return { routing, geometry: `${prefix}2` as const };
  }
  const cleanAdjust = Object.fromEntries(definedEntries) as Readonly<
    Record<string, number>
  >;
  const name =
    parsedAdjust.adj3 !== undefined
      ? (`${prefix}5` as const)
      : parsedAdjust.adj2 !== undefined
        ? (`${prefix}4` as const)
        : (`${prefix}3` as const);
  return {
    routing,
    geometry: {
      type: "preset",
      name,
      adjust: cleanAdjust,
    } satisfies PresetGeometry,
  };
}

function portToSiteIndex(port: LinePort): number {
  switch (port) {
    case "top":
      return 0;
    case "left":
      return 1;
    case "bottom":
      return 2;
    case "right":
      return 3;
    default:
      return port;
  }
}

function computeTargetConnectionSites(
  target: ScopeTargetObject,
): readonly ScreenPoint[] {
  if (target.connectionSites !== undefined) {
    return target.connectionSites;
  }
  const { box, transform } = target;
  const hw = box.width / 2;
  const hh = box.height / 2;
  const cx = box.x + hw;
  const cy = box.y + hh;
  const flipH = transform?.flipH ?? false;
  const flipV = transform?.flipV ?? false;
  const angleRad = ((transform?.angle ?? 0) * Math.PI) / 180;
  const cos = Math.cos(angleRad);
  const sin = Math.sin(angleRad);
  // ECMA-376 standard 4 cardinal connection sites: 0=top, 1=left, 2=bottom, 3=right.
  const localOffsets: readonly ScreenPoint[] = [
    { x: 0, y: -hh },
    { x: -hw, y: 0 },
    { x: 0, y: hh },
    { x: hw, y: 0 },
  ];
  return localOffsets.map((offset) => {
    const fx = flipH ? -offset.x : offset.x;
    const fy = flipV ? -offset.y : offset.y;
    return {
      x: cx + fx * cos - fy * sin,
      y: cy + fx * sin + fy * cos,
    };
  });
}

function resolveEndpointPosition(
  endpoint: LineEndpoint,
  resolveTargetObject: ObjectResolveContext["resolveTargetObject"],
): {
  readonly point: ScreenPoint;
  readonly connection?: { readonly id: number; readonly idx: number };
} {
  if ("x" in endpoint) {
    return { point: { x: endpoint.x, y: endpoint.y } };
  }
  const target = resolveTargetObject?.(endpoint.id);
  if (target === undefined) {
    throw new CapabilityError(
      "broken-reference",
      `Line endpoint references missing, hidden, non-native, or out-of-scope target "${endpoint.id}"`,
    );
  }
  const sites = computeTargetConnectionSites(target);
  const idx = portToSiteIndex(endpoint.port);
  if (!Number.isSafeInteger(idx) || idx < 0 || idx >= sites.length) {
    throw new CapabilityError(
      "invalid-connection-port",
      `Connection port ${String(endpoint.port)} is out of bounds for target "${endpoint.id}" (valid indices: 0..${sites.length - 1})`,
    );
  }
  const site = sites[idx]!;
  return {
    point: site,
    connection: { id: target.spid, idx },
  };
}

/** Closed resolved line model produced in resolve() and consumed by emit and paint. */
type ResolvedLineModel = {
  readonly escapedName: string;
  readonly visual: ResolvedVisualModel;
  readonly connection?: LineConnectionXml;
};

function emitConnectorLineModel(
  model: ResolvedLineModel,
  emitContext: EmitContext,
): EmitOutput {
  const properties = emitVisualModel(
    model.visual,
    (resource) => emitContext.registerResource(resource),
    { omitFill: true, allowSingleZeroExtent: true },
  );
  return {
    xml: lineXml(
      emitContext.id,
      model.escapedName,
      properties.xml,
      nativeLinkXml(emitContext.nativeLink),
      model.connection,
    ),
  };
}

function paintConnectorLineModel(
  model: ResolvedLineModel,
  paintContext: PaintContext,
): PaintOutput {
  return paintVisualModel({
    model: model.visual,
    slideNumber: paintContext.slideNumber,
    objectId: paintContext.objectId,
  });
}

/** Native canvas mechanism for DrawingML `<p:cxnSp>` connectors and lines. */
export const connectorLineMechanism = defineCanvasMechanism({
  key: "connector-line",
  role: "canvas",
  emit: emitConnectorLineModel,
  paint: paintConnectorLineModel,
  import: connectorLineImportHandler,
});

/** Native canvas mechanism normalizing empty no-fill line `<p:sp>` elements into `<Line>`. */
export const normalizedSpLineMechanism = defineCanvasMechanism({
  key: "line-normalized-shape",
  role: "canvas",
  emit: emitConnectorLineModel,
  paint: paintConnectorLineModel,
  import: presetLineShapeImportHandler,
});

/** Line module independently validates 1D geometry, endpoint anchors, strokes, and emits native `<p:cxnSp>`. */
export const lineCapability: ObjectCapabilityModule<
  BoundObjectCapability,
  readonly [
    SealedCanvasMechanism<"connector-line">,
    SealedCanvasMechanism<"line-normalized-shape">,
  ]
> = {
  declaration: "line",
  codegen: lineCodegen,
  mechanisms: [connectorLineMechanism, normalizedSpLineMechanism] as const,

  bind(
    input: CapabilityInput,
    context: CapabilityParseContext,
  ): BoundObjectCapability {
    if (input.children.length > 0) {
      throw new CapabilityError(
        "unsupported-children",
        "Line does not support children or text content",
      );
    }

    const unknown = Object.keys(input.props).find(
      (key) => key !== "children" && !allowedParameters.includes(key),
    );
    if (unknown !== undefined) {
      throw new CapabilityError(
        "unknown-parameter",
        `Line contains unsupported parameter ${unknown}`,
        undefined,
        { prop: unknown, scope: "props" },
      );
    }

    const requestedName = parseOptionalName(input.props["name"]);
    const morphKey = parseMorphKey(input.props["morphKey"]);
    const link = parseLinkTarget(input.props["link"]);

    const rawFrom = input.props["from"];
    const rawTo = input.props["to"];
    const hasFrom = rawFrom !== undefined;
    const hasTo = rawTo !== undefined;
    if (hasFrom !== hasTo) {
      throw new CapabilityError(
        "invalid-line-endpoints",
        "Line requires both from and to when declaring anchored endpoints",
      );
    }
    const hasEndpoints = hasFrom && hasTo;

    let fromEndpoint: LineEndpoint | undefined;
    let toEndpoint: LineEndpoint | undefined;
    if (hasEndpoints) {
      const parsedFrom = lineEndpointSchema.safeParse(rawFrom);
      if (!parsedFrom.success) {
        throw new CapabilityError(
          "invalid-line-endpoints",
          parsedFrom.error.issues[0]?.message ?? "Invalid Line.from endpoint",
        );
      }
      const parsedTo = lineEndpointSchema.safeParse(rawTo);
      if (!parsedTo.success) {
        throw new CapabilityError(
          "invalid-line-endpoints",
          parsedTo.error.issues[0]?.message ?? "Invalid Line.to endpoint",
        );
      }
      fromEndpoint = parsedFrom.data;
      toEndpoint = parsedTo.data;
    }

    const rawStyle = record(input.props["style"]);
    if (hasEndpoints) {
      for (const [key, value] of Object.entries(rawStyle)) {
        if (value === undefined) continue;
        if (!allowedStyleKeySet.has(key)) {
          throw new CapabilityError(
            "invalid-style",
            `Line.style does not support ${key}`,
          );
        }
        if (!endpointModeAllowedStyleKeys.has(key)) {
          throw new CapabilityError(
            "invalid-style",
            `Line with from/to endpoints cannot declare geometric or layout style.${key}`,
            undefined,
            { field: `style.${key}` },
          );
        }
        if (key === "display" && value !== "none") {
          throw new CapabilityError(
            "invalid-style",
            `Line with from/to endpoints only allows display: "none", got ${JSON.stringify(value)}`,
            undefined,
            { field: "style.display" },
          );
        }
      }
    }

    const contract = parseCapabilityLayoutContract({
      nativeKind: "cxnSp",
      props: input.props,
      allowedStyleKeys: allowedStyleKeySet,
      allowPadding: false,
      excludedFromLayout: hasEndpoints,
      path: context.path,
    });

    const { geometry } = resolveConnectorGeometry(
      input.props["routing"],
      input.props["adjust"],
    );

    const visualStyle = { ...rawStyle };
    const opacity =
      typeof visualStyle["opacity"] === "number" &&
      Number.isFinite(visualStyle["opacity"])
        ? visualStyle["opacity"]
        : undefined;
    const rawStroke = input.props["stroke"];

    // Validate stroke, opacity, boxShadow, rotate, flip during bind.
    const initialStroke = resolveLineStroke(context.theme, rawStroke, opacity);
    resolveVisualModel({
      theme: context.theme,
      style: visualStyle,
      box: { x: 0, y: 0, width: 100, height: 0 },
      geometry,
      borderOverride: initialStroke ?? null,
    });

    const authorId =
      typeof input.props["id"] === "string" ? input.props["id"] : undefined;
    const motionProps: Record<string, unknown> = {};
    for (const key of Object.keys(motionPropsShape)) {
      if (input.props[key] !== undefined) {
        motionProps[key] = input.props[key];
      }
    }

    const declaredWidth = contract.explicitPixelSize?.width;
    const declaredHeight = contract.explicitPixelSize?.height;

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
        _request: CapabilityLayoutRequest,
      ): CapabilityLayoutMeasurement | undefined {
        return { minContentWidth: 0, minContentHeight: 0 };
      },
      resolve(resolveContext: ObjectResolveContext): ObjectResolution {
        let box: { x: number; y: number; width: number; height: number };
        let resolvedStyle = { ...visualStyle };
        let connection: LineConnectionXml | undefined;

        if (fromEndpoint !== undefined && toEndpoint !== undefined) {
          const resolveTarget: ObjectResolveContext["resolveTargetObject"] = (
            id,
          ) => resolveContext.resolveTargetObject?.(id);
          const fromResolved = resolveEndpointPosition(
            fromEndpoint,
            resolveTarget,
          );
          const toResolved = resolveEndpointPosition(toEndpoint, resolveTarget);
          const x1 = validateCoordinate(fromResolved.point.x, "x");
          const y1 = validateCoordinate(fromResolved.point.y, "y");
          const x2 = validateCoordinate(toResolved.point.x, "x");
          const y2 = validateCoordinate(toResolved.point.y, "y");
          box = {
            x: Math.min(x1, x2),
            y: Math.min(y1, y2),
            width: Math.abs(x2 - x1),
            height: Math.abs(y2 - y1),
          };
          assertVisualBoxExtent(box, { allowSingleZeroExtent: true });
          const flipH = x2 < x1;
          const flipV = y2 < y1;
          resolvedStyle = {
            ...resolvedStyle,
            ...(flipH ? { flipH: true } : {}),
            ...(flipV ? { flipV: true } : {}),
          };
          if (
            fromResolved.connection !== undefined ||
            toResolved.connection !== undefined
          ) {
            connection = {
              ...(fromResolved.connection === undefined
                ? {}
                : { start: fromResolved.connection }),
              ...(toResolved.connection === undefined
                ? {}
                : { end: toResolved.connection }),
            };
          }
        } else {
          box =
            resolveContext.layoutBox === undefined
              ? {
                  x: validateCoordinate(0, "x"),
                  y: validateCoordinate(0, "y"),
                  width: validateNonNegativeSize(declaredWidth ?? 0, "width"),
                  height: validateNonNegativeSize(
                    declaredHeight ?? 0,
                    "height",
                  ),
                }
              : {
                  x: validateCoordinate(resolveContext.layoutBox.x, "x"),
                  y: validateCoordinate(resolveContext.layoutBox.y, "y"),
                  width: validateNonNegativeSize(
                    resolveContext.layoutBox.width,
                    "width",
                  ),
                  height: validateNonNegativeSize(
                    resolveContext.layoutBox.height,
                    "height",
                  ),
                };
          assertVisualBoxExtent(box, { allowSingleZeroExtent: true });
        }

        const resolvedStroke = resolveLineStroke(
          resolveContext.theme,
          rawStroke,
          opacity,
        );
        const visual = resolveVisualModel({
          theme: resolveContext.theme,
          style: resolvedStyle,
          box,
          geometry,
          borderOverride: resolvedStroke ?? null,
        });
        const geometryName =
          visual.geometry.kind === "custom" ? "custom" : visual.geometry.name;

        const name =
          resolveContext.morphName ??
          requestedName ??
          `Line ${resolveContext.id - 1}`;
        const annotations = motionAnnotations(motionProps, {
          kind: "line",
          name,
          ...(authorId === undefined ? {} : { authorId }),
          textBody: false,
          filled: false,
          solidFilled: false,
          box,
          rotationDegrees: visual.rotate,
          hasBackground: visual.border !== undefined,
        });
        let escapedName: string;
        try {
          escapedName = xmlEscape(name);
        } catch {
          throw new CapabilityError(
            "invalid-xml-character",
            "Line.name contains characters not allowed in XML 1.0",
          );
        }
        const resolvedModel: ResolvedLineModel = {
          escapedName,
          visual,
          ...(connection === undefined ? {} : { connection }),
        };
        const transform: PreviewTransform = {
          box,
          angle: visual.rotate,
          flipH: visual.flipH,
          flipV: visual.flipV,
        };
        return {
          observation: {
            id: resolveContext.id,
            kind: "line",
            name,
            geometry: geometryName,
            box,
          },
          transform,
          extensions: [],
          ...(annotations.length === 0 ? {} : { annotations }),
          emit(emitContext: EmitContext): EmitOutput {
            return connectorLineMechanism.emit(resolvedModel, emitContext);
          },
          paint(paintContext: PaintContext): PaintOutput {
            return connectorLineMechanism.paint(resolvedModel, paintContext);
          },
        };
      },
    };
  },
} satisfies ObjectCapabilityModule;
