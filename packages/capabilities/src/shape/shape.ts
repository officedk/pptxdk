import {
  assertValidLang,
  CapabilityError,
  paddingStyleKeys,
  parseCapabilityLayoutContract,
  textStyleKeys,
  type BoundObjectCapability,
  type BoundObjectNativeCollaboration,
  type CapabilityInput,
  type CapabilityLayoutMeasurement,
  type CapabilityLayoutRequest,
  type CapabilityParseContext,
  type EmitContext,
  type EmitOutput,
  type ObjectCapabilityModule,
  type ObjectResolveContext,
  type PaintContext,
  type PaintOutput,
  type PreviewTransform,
  type ObjectResolution,
  type SealedCanvasMechanism,
} from "@pptxdk/core/capability";
import { pxToEmu, xmlEscape } from "@pptxdk/emit";
import { defineCanvasMechanism } from "@pptxdk/import";

import { evaluateGeometry, type Geometry } from "../shared/geometry.ts";
export {
  blockArc,
  path,
  preset,
  presetGeometryNames,
  presetGeometryTuple,
} from "../shared/geometry.ts";
export type {
  CustomGeometry,
  Geometry,
  GeometryPath,
  PathCommand,
  PathOptions,
  PresetGeometry,
  PresetGeometryName,
} from "../shared/geometry.ts";
import {
  nativeLinkXml,
  parseLinkTarget,
  parseMorphKey,
  parseOptionalName,
} from "../shared/identity.ts";
import {
  customGeometryShapeImportHandler,
  presetShapeImportHandler,
} from "./shape-import.ts";
import { shapeXml } from "./shape-xml.ts";
export { linearGradient, radialGradient } from "../shared/visual.ts";
export type { PaddingStyle, VisualStyle } from "@pptxdk/core/capability";
import {
  motionAnnotations,
  motionParagraphAnnotations,
  pickMotionProps,
} from "../motion/motion-api.ts";
import {
  computeTextMetrics,
  emitTextBody,
  measureTextLayoutFromParagraphs,
  parseTextContent,
  renderTextOutlinesSvg,
  resolveTextBody,
  textInsets,
  type Paragraph,
  type ParsedParagraphMotion,
  type ResolvedTextBody,
  type TextStyle,
} from "../shared/text-service.ts";
import {
  emitVisualModel,
  gradientFillMechanism,
  imageFillMechanism,
  paintVisualModel,
  resolveVisualModel,
  solidFillMechanism,
  type ResolvedVisualModel,
} from "../shared/visual.ts";
import { shapeCodegen } from "./shape-codegen.ts";
import {
  shapeCoordinateSchema,
  shapeDeclarationPropsKeys,
  shapeSizeSchema,
  shapeStyleKeyList,
} from "./shape-declaration.ts";

export {
  createShapeNode,
  shapeDeclarationPropsSchema,
  shapeStyleSchema,
  type ShapeDeclarationProps,
  type ShapeStyle,
} from "./shape-declaration.ts";

const allowedParameters: readonly string[] = shapeDeclarationPropsKeys;

export const shapeStyleKeys = shapeStyleKeyList;

function record(input: unknown): Record<string, unknown> {
  return typeof input === "object" && input !== null && !Array.isArray(input)
    ? (input as Record<string, unknown>)
    : {};
}

function shapeTextStyle(
  props: Readonly<Record<string, unknown>>,
): Record<string, unknown> {
  const layoutStyle = record(props["style"]);
  if (layoutStyle["lang"] !== undefined) {
    assertValidLang(layoutStyle["lang"]);
  }
  return Object.fromEntries(
    Object.entries(layoutStyle).filter(([key]) => textStyleKeys.has(key)),
  );
}

function validateCoordinate(
  value: unknown,
  name: "x" | "y",
  path?: string,
): number {
  const parsed = shapeCoordinateSchema.safeParse(value);
  if (!parsed.success) {
    throw new CapabilityError(
      "invalid-coordinate",
      `${name} must be a finite pixel number, width and height must be greater than 0, and coordinates must be between -1000000 and 1000000`,
      path,
    );
  }
  return parsed.data;
}

function validateSize(
  value: unknown,
  name: "width" | "height",
  path?: string,
): number {
  if (typeof value === "number" && Number.isFinite(value)) {
    if (value === 0 || pxToEmu(value) === 0) {
      throw new CapabilityError(
        "invalid-coordinate",
        `Computed ${name} is 0px; ${name} must be greater than 0px`,
        path,
      );
    }
    if (value < 0 || pxToEmu(value) < 1) {
      throw new CapabilityError(
        "invalid-coordinate",
        `Computed ${name} is ${value}px; ${name} must be greater than 0px`,
        path,
      );
    }
  } else {
    throw new CapabilityError(
      "invalid-coordinate",
      `${name} must be a finite pixel number greater than 0px`,
      path,
    );
  }
  const parsed = shapeSizeSchema.safeParse(value);
  if (!parsed.success) {
    throw new CapabilityError(
      "invalid-coordinate",
      `${name} must convert to an EMU value between 1 and 2147483647 inclusive`,
      path,
    );
  }
  return parsed.data;
}

type TextInsets = ReturnType<typeof textInsets>;

/** Closed resolved shape text body consumed exhaustively by emitShapeModel and paintShapeModel. */
type ResolvedShapeText =
  | { readonly kind: "none" }
  | {
      readonly kind: "body";
      readonly body: ResolvedTextBody;
      readonly insets: TextInsets;
      readonly verticalAlign: NonNullable<TextStyle["verticalAlign"]>;
    };

/** Closed resolved shape model produced in resolve() and exhaustively consumed by emit and paint. */
type ResolvedShapeModel = {
  readonly escapedName: string;
  readonly visual: ResolvedVisualModel;
  readonly text: ResolvedShapeText;
};

function roundSvgNumber(value: number): string {
  const rounded = Math.round(value * 1_000_000) / 1_000_000;
  return Object.is(rounded, -0) ? "0" : String(rounded);
}

function emitShapeModel(
  model: ResolvedShapeModel,
  emitContext: EmitContext,
): EmitOutput {
  const properties = emitVisualModel(model.visual, (resource) =>
    emitContext.registerResource(resource),
  );
  const embedFonts = emitContext.options.embedFonts;
  switch (model.text.kind) {
    case "none":
      return {
        xml: shapeXml(
          emitContext.id,
          model.escapedName,
          properties.xml,
          "",
          nativeLinkXml(emitContext.nativeLink),
        ),
      };
    case "body": {
      const emittedBody = emitTextBody(model.text.body, emitContext.theme, {
        createFieldId: emitContext.createTextFieldId,
        ...(embedFonts === undefined ? {} : { embedFonts }),
      });
      return {
        xml: shapeXml(
          emitContext.id,
          model.escapedName,
          properties.xml,
          emittedBody.xml,
          nativeLinkXml(emitContext.nativeLink),
        ),
        presentationContributions: emittedBody.presentationContributions,
      };
    }
    default: {
      const exhaustive: never = model.text;
      return exhaustive;
    }
  }
}

function paintShapeModel(
  model: ResolvedShapeModel,
  paintContext: PaintContext,
): PaintOutput {
  const visualPaint = paintVisualModel({
    model: model.visual,
    slideNumber: paintContext.slideNumber,
    objectId: paintContext.objectId,
  });
  switch (model.text.kind) {
    case "none":
      return visualPaint;
    case "body": {
      const { box, flipH, flipV, rotate, geometry } = model.visual;
      const paintBox = paintContext.box ?? box;
      const evaluated = evaluateGeometry(geometry, paintBox);
      const { textRect } = evaluated;
      const cx = paintBox.x + paintBox.width / 2;
      const cy = paintBox.y + paintBox.height / 2;
      // Flip mirrors the geometry and its embedded textRect around the shape center,
      // while inline text glyphs stay upright and only follow shape rotation.
      const effectiveTextRect = {
        x: flipH ? 2 * cx - (textRect.x + textRect.width) : textRect.x,
        y: flipV ? 2 * cy - (textRect.y + textRect.height) : textRect.y,
        width: textRect.width,
        height: textRect.height,
      };
      const glyphsSvg = renderTextOutlinesSvg({
        paragraphs: model.text.body.paragraphs,
        box: effectiveTextRect,
        insets: model.text.insets,
        verticalAlign: model.text.verticalAlign,
        theme: paintContext.theme,
        measuredLayout:
          paintBox.width === box.width && paintBox.height === box.height
            ? model.text.body.measuredLayout
            : undefined,
        idPrefix: `obj-s${paintContext.slideNumber}-o${paintContext.objectId}`,
        rotate,
      });
      const textLayer =
        glyphsSvg === ""
          ? ""
          : rotate === 0
            ? glyphsSvg
            : `<g transform="rotate(${roundSvgNumber(rotate)} ${roundSvgNumber(cx)} ${roundSvgNumber(cy)})">${glyphsSvg}</g>`;
      return {
        svg: `${visualPaint.svg}${textLayer}`,
        ...(visualPaint.placeholders === undefined
          ? {}
          : { placeholders: visualPaint.placeholders }),
      };
    }
    default: {
      const exhaustive: never = model.text;
      return exhaustive;
    }
  }
}

/** Native canvas mechanism for DrawingML `<a:prstGeom>` shapes. */
export const presetShapeMechanism = defineCanvasMechanism({
  key: "preset-shape",
  role: "canvas",
  emit: emitShapeModel,
  paint: paintShapeModel,
  import: presetShapeImportHandler,
});

/** Native canvas mechanism for DrawingML `<a:custGeom>` shapes. */
export const customGeometryShapeMechanism = defineCanvasMechanism({
  key: "custom-geometry-shape",
  role: "canvas",
  emit: emitShapeModel,
  paint: paintShapeModel,
  import: customGeometryShapeImportHandler,
});

/** Shape module independently validates input, encapsulates bound state, and emits native PresentationML objects. */
export const shapeCapability: ObjectCapabilityModule<
  BoundObjectCapability,
  readonly [
    SealedCanvasMechanism<"preset-shape">,
    SealedCanvasMechanism<"custom-geometry-shape">,
    SealedCanvasMechanism<"solid-fill">,
    SealedCanvasMechanism<"gradient-fill">,
    SealedCanvasMechanism<"image-fill">,
  ]
> = {
  declaration: "shape",
  codegen: shapeCodegen,
  mechanisms: [
    presetShapeMechanism,
    customGeometryShapeMechanism,
    solidFillMechanism,
    gradientFillMechanism,
    imageFillMechanism,
  ] as const,

  bind(
    input: CapabilityInput,
    context: CapabilityParseContext,
  ): BoundObjectCapability {
    const unknown = Object.keys(input.props).find(
      (key) => key !== "children" && !allowedParameters.includes(key),
    );
    if (unknown !== undefined) {
      throw new CapabilityError(
        "unknown-parameter",
        `Shape contains unsupported parameter ${unknown}`,
        undefined,
        { prop: unknown, scope: "props" },
      );
    }

    const requestedName = parseOptionalName(input.props["name"]);
    const morphKey = parseMorphKey(input.props["morphKey"]);
    const link = parseLinkTarget(input.props["link"]);

    const layoutStyleRecord = record(input.props["style"]);
    const hasDeclaredTextOrPaddingStyle = Object.keys(layoutStyleRecord).some(
      (key) =>
        textStyleKeys.has(key) ||
        (paddingStyleKeys as readonly string[]).includes(key),
    );
    const hasNonEmptyText = input.children.some(
      (child) =>
        !("text" in child && typeof child.text === "string") ||
        child.text !== "",
    );
    const hasExplicitEmptyText = input.children.length > 0 && !hasNonEmptyText;
    const hasTextBody = hasNonEmptyText || hasDeclaredTextOrPaddingStyle;

    const allowedStyleKeys = new Set<string>(shapeStyleKeys);

    const contract = parseCapabilityLayoutContract({
      nativeKind: "sp",
      props: input.props,
      allowedStyleKeys,
      allowPadding: true,
      path: context.path,
    });

    const insets = hasTextBody
      ? textInsets(input.props["style"])
      : { top: 0, right: 0, bottom: 0, left: 0 };

    const rawTextStyle = shapeTextStyle(input.props);
    const effectiveTextStyle = {
      ...rawTextStyle,
      verticalAlign: rawTextStyle["verticalAlign"] ?? "middle",
    };
    const {
      paragraphs,
      paragraphMotions,
      style: parsedTextStyle,
    } = hasTextBody
      ? parseTextContent(
          input.children,
          effectiveTextStyle,
          context.theme,
          { slideNumber: context.slideNumber },
          "Shape.style",
          undefined,
          undefined,
          { allowParagraphMotion: true },
        )
      : {
          paragraphs: [] as readonly Paragraph[],
          paragraphMotions: [] as readonly ParsedParagraphMotion[],
          style: parseTextContent(
            [],
            {},
            context.theme,
            { slideNumber: context.slideNumber },
            "Shape.style",
          ).style,
        };

    const visualStyle = { ...record(input.props["style"]) };
    const geometry = input.props["geometry"] as Geometry | undefined;
    const authorId =
      typeof input.props["id"] === "string" ? input.props["id"] : undefined;
    const motionProps: Readonly<Record<string, unknown>> = pickMotionProps(
      input.props,
    );

    // Validate geometry and visual properties during bind before layout runs.
    resolveVisualModel({
      theme: context.theme,
      style: visualStyle,
      box: { x: 0, y: 0, width: 100, height: 100 },
      ...(geometry !== undefined ? { geometry } : {}),
    });

    const declaredWidth = contract.explicitPixelSize?.width;
    const declaredHeight = contract.explicitPixelSize?.height;
    const hasFixedHeight = contract.fixedAxes?.height ?? false;

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
      ): CapabilityLayoutMeasurement | undefined {
        if (!hasTextBody || paragraphs.length === 0) {
          return { minContentWidth: 0, minContentHeight: 0 };
        }

        const width = request.availableWidth ?? declaredWidth;
        const innerWidth =
          request.availableWidth !== undefined
            ? request.availableWidth
            : typeof width === "number"
              ? Math.max(0, width - insets.left - insets.right)
              : undefined;

        try {
          const measured = measureTextLayoutFromParagraphs(
            paragraphs,
            parsedTextStyle,
            innerWidth,
          );
          const baseline =
            parsedTextStyle.verticalAlign === "top"
              ? (measured.baseline ?? "unavailable")
              : "unavailable";

          return {
            ...measured,
            baseline,
          };
        } catch (error) {
          if (
            hasFixedHeight &&
            request.purpose !== "baseline" &&
            error instanceof CapabilityError &&
            (error.code === "font-measurement-unavailable" ||
              error.code === "font-missing-glyph")
          ) {
            return {
              baseline: "unavailable",
            };
          }
          throw error;
        }
      },
      resolve(resolveContext: ObjectResolveContext): ObjectResolution {
        const box =
          resolveContext.layoutBox === undefined
            ? {
                x: validateCoordinate(undefined, "x"),
                y: validateCoordinate(undefined, "y"),
                width: validateSize(declaredWidth, "width"),
                height: validateSize(declaredHeight, "height"),
              }
            : {
                x: validateCoordinate(resolveContext.layoutBox.x, "x"),
                y: validateCoordinate(resolveContext.layoutBox.y, "y"),
                width: validateSize(resolveContext.layoutBox.width, "width"),
                height: validateSize(resolveContext.layoutBox.height, "height"),
              };
        const visual = resolveVisualModel({
          theme: resolveContext.theme,
          style: visualStyle,
          box,
          ...(geometry === undefined ? {} : { geometry }),
        });
        const geometryName =
          visual.geometry.kind === "custom" ? "custom" : visual.geometry.name;

        let resolvedText: ResolvedTextBody | undefined;
        if (hasTextBody && paragraphs.length > 0) {
          const { textRect } = evaluateGeometry(visual.geometry, box);
          resolvedText = resolveTextBody(
            paragraphs,
            parsedTextStyle,
            resolveContext.theme,
            parsedTextStyle.verticalAlign ?? "middle",
            {
              width: Math.max(0, textRect.width - insets.left - insets.right),
              height: Math.max(0, textRect.height - insets.top - insets.bottom),
              contentSized: !hasFixedHeight,
              insets,
            },
            "text-box",
          );
        }
        const name =
          resolveContext.morphName ??
          requestedName ??
          `Shape ${resolveContext.id - 1}`;
        const target = {
          kind: "shape",
          name,
          ...(authorId === undefined ? {} : { authorId }),
          textBody: resolvedText !== undefined,
          filled: visual.fill.kind !== "none",
          solidFilled: visual.fill.kind === "solid",
          box,
          rotationDegrees: visual.rotate,
          hasBackground:
            visual.fill.kind !== "none" || visual.border !== undefined,
          ...(resolvedText === undefined
            ? {}
            : { paragraphCount: resolvedText.paragraphCount }),
        } as const;
        const annotations = [
          ...motionAnnotations(motionProps, target),
          ...motionParagraphAnnotations(paragraphMotions, target, motionProps),
        ];
        let escapedName: string;
        try {
          escapedName = xmlEscape(name);
        } catch {
          throw new CapabilityError(
            "invalid-xml-character",
            "Shape.name contains characters not allowed in XML 1.0",
          );
        }
        const resolvedModel: ResolvedShapeModel = {
          escapedName,
          visual,
          text:
            resolvedText === undefined
              ? { kind: "none" }
              : {
                  kind: "body",
                  body: resolvedText,
                  insets,
                  verticalAlign: parsedTextStyle.verticalAlign ?? "middle",
                },
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
            kind: "shape",
            name,
            geometry: geometryName,
            box,
            ...(resolvedText === undefined
              ? hasExplicitEmptyText
                ? { text: "" }
                : {}
              : {
                  text: resolvedText.text,
                  textMetrics:
                    resolvedText.textMetrics ??
                    computeTextMetrics(resolvedText, 0, 0),
                }),
          },
          transform,
          extensions: [],
          ...(annotations.length === 0 ? {} : { annotations }),
          ...(resolvedText === undefined
            ? {}
            : {
                resourceClaims: resolvedText.resourceClaims,
                fontUsages: resolvedText.fontUsages,
              }),
          emit(emitContext: EmitContext): EmitOutput {
            return resolvedModel.visual.geometry.kind === "custom"
              ? customGeometryShapeMechanism.emit(resolvedModel, emitContext)
              : presetShapeMechanism.emit(resolvedModel, emitContext);
          },
          paint(paintContext: PaintContext): PaintOutput {
            return resolvedModel.visual.geometry.kind === "custom"
              ? customGeometryShapeMechanism.paint(resolvedModel, paintContext)
              : presetShapeMechanism.paint(resolvedModel, paintContext);
          },
        };
      },
    };
  },
} satisfies ObjectCapabilityModule;
