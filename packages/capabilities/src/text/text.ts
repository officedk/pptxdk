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
  type GeometryBox,
  type NativePlaceholderRef,
  type ObjectCapabilityModule,
  type ObjectResolveContext,
  type PaintContext,
  type PaintOutput,
  type PreviewTransform,
  type ResolvedObject,
  type SlideEmitOutput,
} from "@pptxdk/core/capability";
import { assertXml10Characters, pxToEmu, xmlEscape } from "@pptxdk/emit";
import { defineCanvasMechanism, defineShowMechanism } from "@pptxdk/import";

import {
  fontPresentationContributions,
  type FontEmbeddingMode,
  type FontEmbeddingUse,
} from "../font/font-embedding.ts";
import { embeddedFontImportHandler } from "../font/font-import.ts";
import { motionAnnotations } from "../motion/motion-api.ts";
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
import { layoutStyleValue } from "../shared/style-parse.ts";
import {
  parseTextCoordinate,
  parseTextExtent,
  textBoxAllowedParamKeys,
  textBoxStyleSchema,
} from "../shared/text-schemas.ts";
import {
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
  emitVisualModel,
  paintVisualModel,
  resolveVisualModel,
  visualStyleKeys,
  visualSvgTransform,
  type ResolvedVisualModel,
} from "../shared/visual.ts";
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
  TextDeclarationProps,
} from "./text-declaration.ts";
export { font } from "../font/font.ts";
export type {
  Font,
  FontFace,
  FontFaceStyle,
  FontOptions,
} from "../font/font.ts";
export type {
  FontReference,
  FontSize,
  ParagraphStyle,
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

function height(
  value: unknown,
  path?: string,
):
  | { readonly contentSized: true }
  | { readonly contentSized: false; readonly value: number | undefined } {
  if (value === "auto" || value === undefined) return { contentSized: true };
  if (
    typeof value === "string" &&
    /^-?(?:\d+(?:\.\d*)?|\.\d+)%$/u.test(value)
  ) {
    return { contentSized: false, value: undefined };
  }
  return { contentSized: false, value: size(value, "height", path) };
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
  readonly box: GeometryBox;
  readonly visual: ResolvedVisualModel;
  readonly text: ResolvedTextBody;
  readonly insets: TextInsets;
  readonly placeholder?: {
    readonly nativeRef: NativePlaceholderRef;
    readonly hasLocalBoxOverride: boolean;
    readonly hasLocalVisualOverride: boolean;
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
    if (
      !model.placeholder.hasLocalBoxOverride &&
      !model.placeholder.hasLocalVisualOverride &&
      !model.visual.hasVisibleStyle &&
      (model.visual.rotationUnits === undefined ||
        model.visual.rotationUnits === 0)
    ) {
      placeholderEmit = {
        nativeRef: model.placeholder.nativeRef,
        spPrXml: "<p:spPr/>",
      };
    } else if (
      !model.placeholder.hasLocalVisualOverride &&
      !model.visual.hasVisibleStyle
    ) {
      const rotAttr =
        model.visual.rotationUnits === undefined ||
        model.visual.rotationUnits === 0
          ? ""
          : ` rot="${model.visual.rotationUnits}"`;
      placeholderEmit = {
        nativeRef: model.placeholder.nativeRef,
        spPrXml: `<p:spPr><a:xfrm${rotAttr}><a:off x="${pxToEmu(model.box.x)}" y="${pxToEmu(model.box.y)}"/><a:ext cx="${pxToEmu(model.box.width)}" cy="${pxToEmu(model.box.height)}"/></a:xfrm></p:spPr>`,
      };
    } else {
      propertiesXml = emitVisualModel(model.visual, (resource) =>
        emitContext.registerResource(resource),
      ).xml;
      placeholderEmit = {
        nativeRef: model.placeholder.nativeRef,
      };
    }
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
  const backgroundSvg = model.visual.hasVisibleStyle
    ? paintVisualModel({
        model: model.visual,
        slideNumber: paintContext.slideNumber,
        objectId: paintContext.objectId,
      }).svg
    : "";
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
  readonly box: GeometryBox;
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
export const textCapability = {
  declaration: "text",
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
    const declaredX = placeholderGeometry?.x ?? input.props["x"];
    const declaredY = placeholderGeometry?.y ?? input.props["y"];
    const declaredWidth = placeholderGeometry?.width ?? input.props["width"];
    if (declaredX !== undefined) coordinate(declaredX, "x");
    if (declaredY !== undefined) coordinate(declaredY, "y");
    if (input.props["width"] !== undefined) {
      size(input.props["width"], "width");
    }
    if (declaredWidth !== undefined && typeof declaredWidth === "number") {
      size(declaredWidth, "width");
    }

    const allowedStyleKeys = new Set<string>(
      Object.keys(textBoxStyleSchema.unwrap().shape),
    );

    const contract = parseCapabilityLayoutContract({
      nativeKind: "sp",
      props: input.props,
      allowedStyleKeys,
      allowPadding: true,
      path: context.path,
    });

    const insets = textInsets(input.props["style"], placeholder);
    const styleHeight = layoutStyleValue(input, "height");
    const requestedHeight = height(
      styleHeight ?? input.props["height"] ?? placeholderGeometry?.height,
    );

    const {
      paragraphs,
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
    );

    const rawSlideStyle = record(input.props["style"])
      ? input.props["style"]
      : undefined;
    const hasLocalVisualOverride =
      rawSlideStyle !== undefined &&
      visualStyleKeys.some((key) => hasLocalDeclaration(rawSlideStyle, key));
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
    });

    const authorId =
      typeof input.props["id"] === "string" ? input.props["id"] : undefined;
    const motionProps: Readonly<Record<string, unknown>> = {
      ...(input.props["enter"] !== undefined
        ? { enter: input.props["enter"] }
        : {}),
      ...(input.props["emphasis"] !== undefined
        ? { emphasis: input.props["emphasis"] }
        : {}),
      ...(input.props["motion"] !== undefined
        ? { motion: input.props["motion"] }
        : {}),
    };

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
        const isContentSized = requestedHeight.contentSized;
        const width = request.availableWidth ?? contract.width ?? declaredWidth;
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
            error instanceof CapabilityError &&
            error.code === "font-measurement-unavailable"
          ) {
            return {
              baseline: "unavailable",
            };
          }
          throw error;
        }
      },
      resolve(resolveContext: ObjectResolveContext): ResolvedObject {
        const dimensions = {
          x: coordinate(
            resolveContext.layoutBox?.x ?? contract.x ?? declaredX,
            "x",
          ),
          y: coordinate(
            resolveContext.layoutBox?.y ?? contract.y ?? declaredY,
            "y",
          ),
          width: size(
            resolveContext.layoutBox?.width ?? contract.width ?? declaredWidth,
            "width",
          ),
        };
        const measuredHeight = requestedHeight.contentSized
          ? undefined
          : requestedHeight.value;
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
            contentSized: requestedHeight.contentSized,
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
        const box: GeometryBox = {
          ...dimensions,
          height: size(
            resolveContext.layoutBox?.height ??
              (requestedHeight.contentSized
                ? resolvedText.measurement === undefined
                  ? undefined
                  : resolvedText.measurement.height + insets.top + insets.bottom
                : requestedHeight.value),
            "height",
          ),
        };
        const visual = resolveVisualModel({
          theme: placeholder?.masterTheme ?? resolveContext.theme,
          style: visualStyle,
          box,
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
        const annotations = motionAnnotations(motionProps, {
          kind: "text",
          name,
          ...(authorId === undefined ? {} : { authorId }),
          textBody: true,
          filled:
            visualStyle["backgroundColor"] !== undefined ||
            visualStyle["backgroundImage"] !== undefined,
          box,
          rotationDegrees: visual.rotate,
          hasBackground:
            visualStyle["backgroundColor"] !== undefined ||
            visualStyle["backgroundImage"] !== undefined ||
            visualStyle["borderColor"] !== undefined ||
            (typeof visualStyle["borderWidth"] === "number" &&
              visualStyle["borderWidth"] > 0),
          paragraphCount: resolvedText.paragraphCount,
        });
        const transform: PreviewTransform = {
          box,
          angle: visual.rotate,
          flipH: false,
          flipV: false,
        };
        const warnings = [...resolvedText.warnings, ...visual.warnings];
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
                  hasLocalBoxOverride:
                    placeholderGeometry?.hasLocalBoxOverride === true,
                  hasLocalVisualOverride,
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
          },
          transform,
          extensions: [],
          ...(warnings.length === 0 ? {} : { warnings }),
          ...(annotations.length === 0 ? {} : { annotations }),
          resourceClaims: resolvedText.resourceClaims,
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
