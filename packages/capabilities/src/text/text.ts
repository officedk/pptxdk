import {
  CapabilityError,
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
  type Rect,
  type NativePlaceholderRef,
  type ObjectCapabilityModule,
  type ObjectResolveContext,
  type ObjectResolution,
  type PaintContext,
  type PaintOutput,
  type PreviewTransform,
  type SealedCanvasMechanism,
  type SealedShowMechanism,
  type SlideEmitOutput,
} from "@pptxdk/core/capability";
import { assertXml10Characters, xmlEscape } from "@pptxdk/emit";
import { defineCanvasMechanism, defineShowMechanism } from "@pptxdk/import";

import {
  fontPresentationContributions,
  type FontEmbeddingMode,
  type FontEmbeddingUse,
} from "../font/font-embedding.ts";
import { embeddedFontImportHandler } from "../font/font-import.ts";
import {
  motionAnnotations,
  motionParagraphAnnotations,
  pickMotionProps,
} from "../motion/motion-api.ts";
import {
  nativeLinkXml,
  parseLinkTarget,
  parseMorphKey,
  parseOptionalName,
} from "../shared/identity.ts";
import {
  hasLocalDeclaration,
  resolvePlaceholderGeometry,
  resolvePlaceholderVisualStyle,
} from "../shared/style-inheritance.ts";
import {
  parseTextCoordinate,
  parseTextExtent,
  textBoxAllowedParamKeys,
  textBoxStyleKeys,
} from "../shared/text-schemas.ts";
import {
  computeTextMetrics,
  emitSlideNumberRunXml,
  emitTextBody,
  measureTextLayoutFromParagraphs,
  parseTextContent,
  renderTextOutlinesSvg,
  resolveTextBody,
  textInsets,
  type ResolvedTextBody,
  type Run,
} from "../shared/text-service.ts";
import {
  emitPlaceholderVisualSpPr,
  emitVisualModel,
  paintVisualModel,
  resolveVisualModel,
  visualSvgTransform,
  type PlaceholderSpPrFlags,
  type ResolvedVisualModel,
} from "../shared/visual.ts";
import { textCodegen } from "./text-codegen.ts";
import {
  textBoxImportHandler,
  textFieldSlideNumImportHandler,
  textPlaceholderImportHandler,
} from "./text-import.ts";
import { textBoxXml } from "./text-xml.ts";

export type {
  ParagraphDeclarationProps,
  SlideNumberDeclarationProps,
  ShapeTextStyle,
  SpanDeclarationProps,
  SubDeclarationProps,
  SupDeclarationProps,
  TextDeclarationProps,
} from "./text-declaration.ts";
export { font, fromFontCollection } from "../font/font.ts";
export type {
  Font,
  FontFace,
  FontFaceStyle,
  FontOptions,
} from "../font/font.ts";
export type {
  FontReference,
  FontSize,
  LineHeight,
  ParagraphStyle,
  RunTextStyle,
  TextAlignment,
  TextList,
  TextStyle,
} from "../shared/text-service.ts";

function record(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function coordinate(value: unknown, label: "x" | "y", path?: string): number {
  return parseTextCoordinate(value, label, path);
}

function size(
  value: unknown,
  label: "width" | "height",
  path?: string,
): number {
  return parseTextExtent(value, label, path);
}

function textStyle(value: unknown): unknown {
  if (!record(value)) return value;
  return Object.fromEntries(
    Object.entries(value).filter(([key]) => textStyleKeys.has(key)),
  );
}

type TextInsets = ReturnType<typeof textInsets>;

/** Closed resolved text model produced in resolve() and exhaustively consumed by emitTextModel and paintTextModel. */
type ResolvedTextModel = {
  readonly escapedName: string;
  readonly box: Rect;
  readonly visual: ResolvedVisualModel;
  readonly text: ResolvedTextBody;
  readonly insets: TextInsets;
  readonly placeholder?: {
    readonly nativeRef: NativePlaceholderRef;
    readonly spPrFlags: PlaceholderSpPrFlags;
  };
};

const textCompatiblePlaceholderRoles = new Set<string>([
  "title",
  "ctrTitle",
  "subTitle",
  "body",
  "dt",
  "ftr",
  "hdr",
  "obj",
  "sldNum",
]);

function emitTextModel(
  model: ResolvedTextModel,
  emitContext: EmitContext,
): EmitOutput {
  const embedFonts = emitContext.options.embedFonts;
  const emittedBody = emitTextBody(model.text, emitContext.theme, {
    createFieldId: emitContext.createTextFieldId,
    ...(embedFonts === undefined ? {} : { embedFonts }),
  });
  let placeholderEmit:
    | { readonly nativeRef: NativePlaceholderRef; readonly spPrXml?: string }
    | undefined;
  let propertiesXml = "";
  if (model.placeholder !== undefined) {
    placeholderEmit = {
      nativeRef: model.placeholder.nativeRef,
      spPrXml: emitPlaceholderVisualSpPr(
        model.visual,
        model.placeholder.spPrFlags,
        (resource) => emitContext.registerResource(resource),
      ),
    };
  } else {
    propertiesXml = emitVisualModel(model.visual, (resource) =>
      emitContext.registerResource(resource),
    ).xml;
  }
  return {
    xml: textBoxXml(
      emitContext.id,
      model.escapedName,
      model.box,
      emittedBody.xml,
      propertiesXml,
      nativeLinkXml(emitContext.nativeLink),
      placeholderEmit,
    ),
    ...(emittedBody.presentationContributions.length === 0
      ? {}
      : {
          presentationContributions: emittedBody.presentationContributions,
        }),
  };
}

function paintTextModel(
  model: ResolvedTextModel,
  paintContext: PaintContext,
): PaintOutput {
  const paintBox = paintContext.box ?? model.box;
  const visualPaint = model.visual.hasVisibleStyle
    ? paintVisualModel({
        model: model.visual,
        slideNumber: paintContext.slideNumber,
        objectId: paintContext.objectId,
      })
    : undefined;
  const backgroundSvg = visualPaint?.svg ?? "";
  const glyphs = renderTextOutlinesSvg({
    paragraphs: model.text.paragraphs,
    box: paintBox,
    insets: model.insets,
    verticalAlign: model.text.verticalAlign,
    theme: paintContext.theme,
    measuredLayout:
      model.box.width === paintBox.width
        ? model.text.measuredLayout
        : undefined,
    idPrefix: `obj-s${paintContext.slideNumber}-o${paintContext.objectId}`,
    rotate: model.visual.rotate,
  });
  return {
    svg: `${backgroundSvg}<g transform="${visualSvgTransform(model.visual)}">${glyphs}</g>`,
    ...(visualPaint?.placeholders === undefined
      ? {}
      : { placeholders: visualPaint.placeholders }),
  };
}

/** Typed native canvas mechanism connecting emit, paint, and import for `text-box`. */
export const textBoxMechanism = defineCanvasMechanism({
  key: "text-box",
  role: "canvas",
  emit: emitTextModel,
  paint: paintTextModel,
  import: textBoxImportHandler,
});

/** Typed native canvas mechanism connecting emit, paint, and import for `text-placeholder`. */
export const textPlaceholderMechanism = defineCanvasMechanism({
  key: "text-placeholder",
  role: "canvas",
  emit: emitTextModel,
  paint: paintTextModel,
  import: textPlaceholderImportHandler,
});

type ResolvedSlideNumberFieldModel = {
  readonly run: Run;
  readonly box: Rect;
  readonly idPrefix: string;
};

/** Typed native canvas mechanism connecting emit, paint, and import for `text-field-slidenum`. */
export const textFieldSlideNumMechanism = defineCanvasMechanism({
  key: "text-field-slidenum",
  role: "canvas",
  emit(
    model: ResolvedSlideNumberFieldModel,
    emitContext: EmitContext,
  ): EmitOutput {
    return {
      xml: emitSlideNumberRunXml(
        model.run,
        emitContext.theme,
        emitContext.createTextFieldId,
      ),
    };
  },
  paint(
    model: ResolvedSlideNumberFieldModel,
    paintContext: PaintContext,
  ): PaintOutput {
    const svg = renderTextOutlinesSvg({
      paragraphs: [
        {
          runs: [model.run],
          style: {},
          endStyle: model.run.style,
          wrap: "normal",
        },
      ],
      box: paintContext.box ?? model.box,
      theme: paintContext.theme,
      idPrefix: model.idPrefix,
    });
    return { svg };
  },
  import: textFieldSlideNumImportHandler,
});

type ResolvedEmbeddedFontModel = {
  readonly uses: readonly FontEmbeddingUse[];
  readonly mode?: FontEmbeddingMode;
};

/** Typed native show mechanism connecting emit and import for presentation `embedded-font`. */
export const embeddedFontMechanism = defineShowMechanism({
  key: "embedded-font",
  role: "show",
  emit(model: ResolvedEmbeddedFontModel): SlideEmitOutput {
    return {
      xml: "",
      presentationContributions: fontPresentationContributions(
        model.uses,
        model.mode,
      ),
    };
  },
  import: embeddedFontImportHandler,
});

/** Fixed-region and content-sized text box capability bound per formal build. */
export const textCapability: ObjectCapabilityModule<
  BoundObjectCapability,
  readonly [
    SealedCanvasMechanism<"text-box">,
    SealedCanvasMechanism<"text-placeholder">,
    SealedCanvasMechanism<"text-field-slidenum">,
    SealedShowMechanism<"embedded-font">,
  ]
> = {
  declaration: "text",
  codegen: textCodegen,
  mechanisms: [
    textBoxMechanism,
    textPlaceholderMechanism,
    textFieldSlideNumMechanism,
    embeddedFontMechanism,
  ] as const,

  bind(
    input: CapabilityInput,
    context: CapabilityParseContext,
  ): BoundObjectCapability {
    const unknown = Object.keys(input.props).find(
      (key) => !textBoxAllowedParamKeys.has(key),
    );
    if (unknown !== undefined) {
      throw new CapabilityError(
        "unknown-parameter",
        `Text contains unsupported parameter ${unknown}`,
      );
    }
    const placeholder = context.placeholder;
    if (placeholder !== undefined) {
      const role =
        placeholder.layoutPlaceholder?.type ?? placeholder.nativeRef.type;
      if (!textCompatiblePlaceholderRoles.has(role)) {
        throw new CapabilityError(
          "incompatible-placeholder-content",
          `Text cannot be bound to ${role} placeholder`,
        );
      }
    }

    const requestedName = parseOptionalName(input.props["name"]);
    const morphKey = parseMorphKey(input.props["morphKey"]);
    const link = parseLinkTarget(input.props["link"]);

    const placeholderGeometry =
      placeholder === undefined
        ? undefined
        : resolvePlaceholderGeometry(placeholder, input.props);
    const declaredX = placeholderGeometry?.x;
    const declaredY = placeholderGeometry?.y;
    const declaredWidth = placeholderGeometry?.width;
    if (declaredX !== undefined) coordinate(declaredX, "x");
    if (declaredY !== undefined) coordinate(declaredY, "y");
    if (declaredWidth !== undefined && typeof declaredWidth === "number") {
      size(declaredWidth, "width");
    }

    const allowedStyleKeys = new Set<string>(textBoxStyleKeys);

    const contract = parseCapabilityLayoutContract({
      nativeKind: "sp",
      props: input.props,
      allowedStyleKeys,
      allowPadding: true,
      path: context.path,
    });

    const insets = textInsets(input.props["style"], placeholder);
    const placeholderHeight = placeholderGeometry?.height;
    const placeholderFixedHeight =
      placeholderHeight !== undefined && placeholderHeight !== "auto";
    const isContentSized =
      !placeholderFixedHeight && !(contract.fixedAxes?.height ?? false);
    const declaredPixelHeight =
      typeof placeholderHeight === "number"
        ? size(placeholderHeight, "height")
        : contract.explicitPixelSize?.height;

    const {
      paragraphs,
      paragraphMotions,
      style: parsedTextStyle,
      localBodyDeclared,
    } = parseTextContent(
      input.children,
      textStyle(input.props["style"]),
      placeholder?.masterTheme ?? context.theme,
      {
        slideNumber: context.slideNumber,
      },
      "Text.style",
      placeholder,
      input.props["style"],
      { allowParagraphMotion: true },
    );

    const rawSlideStyle = record(input.props["style"])
      ? input.props["style"]
      : undefined;
    const inheritedPlaceholderGeometry =
      placeholder?.layoutPlaceholder?.geometry ??
      placeholder?.masterPlaceholder?.geometry;
    const visualStyle =
      placeholder !== undefined
        ? resolvePlaceholderVisualStyle(placeholder, rawSlideStyle)
        : rawSlideStyle !== undefined
          ? { ...rawSlideStyle }
          : {};
    resolveVisualModel({
      theme: placeholder?.masterTheme ?? context.theme,
      style: visualStyle,
      box: { x: 0, y: 0, width: 100, height: 100 },
      ...(inheritedPlaceholderGeometry === undefined
        ? {}
        : { geometry: inheritedPlaceholderGeometry }),
    });

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
      ): CapabilityLayoutMeasurement | undefined {
        const width =
          request.availableWidth ??
          contract.explicitPixelSize?.width ??
          declaredWidth;
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
          return {
            ...measured,
            baseline: measured.baseline ?? "unavailable",
          };
        } catch (error) {
          if (
            !isContentSized &&
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
        const dimensions = {
          x: coordinate(resolveContext.layoutBox?.x ?? declaredX, "x"),
          y: coordinate(resolveContext.layoutBox?.y ?? declaredY, "y"),
          width: size(
            resolveContext.layoutBox?.width ??
              declaredWidth ??
              contract.explicitPixelSize?.width,
            "width",
          ),
        };
        const measuredHeight = isContentSized ? undefined : declaredPixelHeight;
        const outputHeight = resolveContext.layoutBox?.height ?? measuredHeight;
        const innerWidth = Math.max(
          0,
          dimensions.width - insets.left - insets.right,
        );
        const innerHeight =
          outputHeight === undefined
            ? undefined
            : Math.max(0, outputHeight - insets.top - insets.bottom);
        const resolvedText = resolveTextBody(
          paragraphs,
          parsedTextStyle,
          placeholder?.masterTheme ?? resolveContext.theme,
          parsedTextStyle.verticalAlign ?? "top",
          {
            width: innerWidth,
            ...(innerHeight === undefined ? {} : { height: innerHeight }),
            contentSized: isContentSized,
            insets,
          },
          "text-box",
          placeholder === undefined
            ? undefined
            : {
                isPlaceholder: true,
                ...(localBodyDeclared === undefined
                  ? {}
                  : { localBodyDeclared }),
              },
        );
        const box: Rect = {
          ...dimensions,
          height: size(
            resolveContext.layoutBox?.height ??
              (isContentSized
                ? resolvedText.measurement === undefined
                  ? undefined
                  : resolvedText.measurement.height + insets.top + insets.bottom
                : measuredHeight),
            "height",
          ),
        };
        const textMetrics = computeTextMetrics(
          resolvedText,
          Math.max(0, box.width - insets.left - insets.right),
          Math.max(0, box.height - insets.top - insets.bottom),
        );
        const visual = resolveVisualModel({
          theme: placeholder?.masterTheme ?? resolveContext.theme,
          style: visualStyle,
          box,
          ...(inheritedPlaceholderGeometry === undefined
            ? {}
            : { geometry: inheritedPlaceholderGeometry }),
        });
        const name =
          resolveContext.morphName ??
          requestedName ??
          placeholder?.layoutPlaceholder?.name ??
          `Text ${resolveContext.id - 1}`;
        let escapedName: string;
        try {
          assertXml10Characters(name);
          escapedName = xmlEscape(name);
        } catch {
          throw new CapabilityError(
            "invalid-xml-character",
            "Text.name contains characters not allowed in XML 1.0",
          );
        }
        const target = {
          kind: "text",
          name,
          ...(authorId === undefined ? {} : { authorId }),
          textBody: true,
          filled: visual.fill.kind !== "none",
          box,
          rotationDegrees: visual.rotate,
          hasBackground:
            visual.fill.kind !== "none" || visual.border !== undefined,
          paragraphCount: resolvedText.paragraphCount,
        } as const;
        const annotations = [
          ...motionAnnotations(motionProps, target),
          ...motionParagraphAnnotations(paragraphMotions, target, motionProps),
        ];
        const transform: PreviewTransform = {
          box,
          angle: visual.rotate,
          flipH: visual.flipH,
          flipV: visual.flipV,
        };
        const model: ResolvedTextModel = {
          escapedName,
          box,
          visual,
          text: resolvedText,
          insets,
          ...(placeholder === undefined
            ? {}
            : {
                placeholder: {
                  nativeRef: placeholder.nativeRef,
                  spPrFlags: {
                    hasBox: placeholderGeometry?.hasLocalBoxOverride === true,
                    hasRotate: hasLocalDeclaration(rawSlideStyle, "rotate"),
                    hasFlipH: hasLocalDeclaration(rawSlideStyle, "flipH"),
                    hasFlipV: hasLocalDeclaration(rawSlideStyle, "flipV"),
                    hasGeometryOrRadius: hasLocalDeclaration(
                      rawSlideStyle,
                      "borderRadius",
                    ),
                    hasFill: hasLocalDeclaration(rawSlideStyle, "background"),
                    hasBorder: hasLocalDeclaration(rawSlideStyle, "border"),
                    hasEffects:
                      hasLocalDeclaration(rawSlideStyle, "boxShadow") ||
                      hasLocalDeclaration(rawSlideStyle, "softEdge"),
                    includeDefaultRectWhenHasBox: false,
                  },
                },
              }),
        };
        return {
          observation: {
            id: resolveContext.id,
            kind: "text",
            name,
            box,
            text: resolvedText.text,
            textMetrics,
          },
          transform,
          extensions: [],
          ...(annotations.length === 0 ? {} : { annotations }),
          resourceClaims: resolvedText.resourceClaims,
          fontUsages: resolvedText.fontUsages,
          emit(emitContext: EmitContext): EmitOutput {
            return placeholder !== undefined
              ? textPlaceholderMechanism.emit(model, emitContext)
              : textBoxMechanism.emit(model, emitContext);
          },
          paint(paintContext: PaintContext): PaintOutput {
            return placeholder !== undefined
              ? textPlaceholderMechanism.paint(model, paintContext)
              : textBoxMechanism.paint(model, paintContext);
          },
        };
      },
    };
  },
} satisfies ObjectCapabilityModule;
