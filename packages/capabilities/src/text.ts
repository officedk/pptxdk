import {
  CapabilityError,
  layoutStyleKeys,
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
  type GeometryBox,
  type ObjectCapabilityModule,
  type ObjectResolveContext,
  type PaintContext,
  type PaintOutput,
  type PreviewTransform,
  type ResolvedObject,
} from "@react-pptx/core/capability";
import { assertXml10Characters, xmlEscape } from "@react-pptx/emit";

import {
  nativeLinkXml,
  parseLinkTarget,
  parseMorphKey,
  parseOptionalName,
} from "./identity.ts";
import { motionAnnotations } from "./motion-api.ts";
import { layoutStyleValue } from "./style-parse.ts";
import {
  emitTextBody,
  measureTextLayoutFromParagraphs,
  parseTextContent,
  renderTextOutlinesSvg,
  resolveTextBody,
  textExtent,
  textInsets,
} from "./text-service.ts";
import { textBoxXml } from "./text-xml.ts";
import {
  resolveBoxPreviewVisuals,
  shapePropertiesXml,
  visualStyleKeys,
} from "./visual.ts";

export type {
  ParagraphDeclarationProps,
  SlideNumberDeclarationProps,
  ShapeTextStyle,
  SpanDeclarationProps,
  TextDeclarationProps,
} from "./text-declaration.ts";
export { font } from "./font.ts";
export type { Font, FontFace, FontFaceStyle, FontOptions } from "./font.ts";
export type {
  FontReference,
  FontSize,
  ParagraphStyle,
  TextAlignment,
  TextList,
  TextStyle,
} from "./text-service.ts";

const allowedParameters = [
  "id",
  "name",
  "morphKey",
  "link",
  "x",
  "y",
  "width",
  "height",
  "style",
  "children",
  "enter",
  "emphasis",
  "motion",
] as const;

function record(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function coordinate(value: unknown, label: "x" | "y", path?: string): number {
  if (
    typeof value !== "number" ||
    !Number.isFinite(value) ||
    value < -1_000_000 ||
    value > 1_000_000
  ) {
    throw new CapabilityError(
      "invalid-text-box",
      `${label} must be a finite pixel number between -1000000 and 1000000`,
      path,
    );
  }
  return value;
}

function size(
  value: unknown,
  label: "width" | "height",
  path?: string,
): number {
  if (typeof value !== "number") {
    throw new CapabilityError(
      "invalid-text-box",
      `${label} must be a finite positive pixel number`,
      path,
    );
  }
  textExtent(value, label);
  return value;
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

/** Fixed-region and content-sized text box capability bound per formal build. */
export const textCapability: ObjectCapabilityModule = {
  declaration: "rp:text",

  bind(
    input: CapabilityInput,
    context: CapabilityParseContext,
  ): BoundObjectCapability {
    const unknown = Object.keys(input.props).find(
      (key) => !(allowedParameters as readonly string[]).includes(key),
    );
    if (unknown !== undefined) {
      throw new CapabilityError(
        "unknown-parameter",
        `Text contains unsupported parameter ${unknown}`,
      );
    }
    const requestedName = parseOptionalName(input.props["name"]);
    const morphKey = parseMorphKey(input.props["morphKey"]);
    const link = parseLinkTarget(input.props["link"]);

    const declaredX = input.props["x"];
    const declaredY = input.props["y"];
    const declaredWidth = input.props["width"];
    if (declaredX !== undefined) coordinate(declaredX, "x");
    if (declaredY !== undefined) coordinate(declaredY, "y");
    if (declaredWidth !== undefined && typeof declaredWidth === "number") {
      size(declaredWidth, "width");
    }

    const allowedStyleKeys = new Set<string>([
      ...layoutStyleKeys,
      ...paddingStyleKeys,
      ...textStyleKeys,
      ...visualStyleKeys,
    ]);

    const contract = parseCapabilityLayoutContract({
      nativeKind: "sp",
      props: input.props,
      allowedStyleKeys,
      allowPadding: true,
      path: context.path,
    });

    const insets = textInsets(input.props["style"]);
    const styleHeight = layoutStyleValue(input, "height");
    const requestedHeight = height(styleHeight ?? input.props["height"]);

    const { paragraphs, style: parsedTextStyle } = parseTextContent(
      input.children,
      textStyle(input.props["style"]),
      context.theme,
      {
        slideNumber: context.slideNumber,
      },
    );

    const visualStyle = record(input.props["style"])
      ? { ...input.props["style"] }
      : {};
    shapePropertiesXml({
      theme: context.theme,
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
          resolveContext.theme,
          "top",
          {
            width: innerWidth,
            ...(innerHeight === undefined ? {} : { height: innerHeight }),
            contentSized: requestedHeight.contentSized,
            insets,
          },
          "text-box",
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
        const properties = shapePropertiesXml({
          theme: resolveContext.theme,
          style: visualStyle,
          box,
        });
        const name =
          resolveContext.morphName ??
          requestedName ??
          `Text ${resolveContext.id - 1}`;
        try {
          assertXml10Characters(name);
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
          rotationDegrees:
            typeof visualStyle["rotate"] === "number"
              ? visualStyle["rotate"]
              : 0,
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
          angle:
            typeof visualStyle["rotate"] === "number"
              ? visualStyle["rotate"]
              : 0,
          flipH: false,
          flipV: false,
        };
        const warnings = [...resolvedText.warnings, ...properties.warnings];
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
            const embedFonts = emitContext.options.embedFonts;
            const emittedBody = emitTextBody(resolvedText, emitContext.theme, {
              createFieldId: emitContext.createTextFieldId,
              ...(embedFonts === undefined ? {} : { embedFonts }),
            });
            const emittedProperties = shapePropertiesXml({
              theme: emitContext.theme,
              style: visualStyle,
              box,
            });
            return {
              xml: textBoxXml(
                emitContext.id,
                xmlEscape(name),
                box,
                emittedBody.xml,
                emittedProperties.xml,
                nativeLinkXml(emitContext.nativeLink),
              ),
              ...(emittedBody.presentationContributions.length === 0
                ? {}
                : {
                    presentationContributions:
                      emittedBody.presentationContributions,
                  }),
            };
          },
          paint(paintContext: PaintContext): PaintOutput {
            const paintBox = paintContext.box ?? box;
            const visuals = resolveBoxPreviewVisuals({
              label: "text",
              visualStyle,
              box: paintBox,
              theme: paintContext.theme,
            });
            const backgroundRect = visuals.hasFillOrStroke
              ? `<rect x="${paintBox.x}" y="${paintBox.y}" width="${paintBox.width}" height="${paintBox.height}" ${visuals.attributes}/>`
              : "";
            const glyphs = renderTextOutlinesSvg({
              paragraphs: resolvedText.paragraphs,
              box: paintBox,
              insets,
              verticalAlign: parsedTextStyle.verticalAlign ?? "top",
              theme: paintContext.theme,
              measuredLayout:
                box.width === paintBox.width
                  ? resolvedText.measuredLayout
                  : undefined,
            });
            return {
              svg: `<g transform="${visuals.svgTransform}">${backgroundRect}${glyphs}</g>`,
            };
          },
        };
      },
    };
  },
};
