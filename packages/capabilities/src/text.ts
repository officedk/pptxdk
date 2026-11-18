import {
  CapabilityError,
  xmlEscape,
  layoutStyleKeys,
  paddingStyleKeys,
  parseCapabilityLayoutContract,
  textStyleKeys,
  type BoundObjectCapability,
  type BoundObjectCapabilityModule,
  type BoundObjectNativeCollaboration,
  type CapabilityContext,
  type CapabilityInput,
  type CapabilityLayoutMeasurement,
  type CapabilityLayoutRequest,
  type CapabilityParseContext,
  type GeometryBox,
  type ObjectPreviewContext,
  type ObjectPreviewOutput,
} from "@react-pptx/core/capability";

import { fontPresentationContributions } from "./font-embedding.ts";
import {
  nativeLinkXml,
  parseLinkTarget,
  parseMorphKey,
  parseOptionalName,
} from "./identity.ts";
import { motionAnnotations } from "./motion-api.ts";
import { layoutStyleValue } from "./style-parse.ts";
import {
  measureTextLayoutFromParagraphs,
  parseTextContent,
  renderTextOutlinesSvg,
  textInsets,
  textExtent,
  writeTextBodyFromParagraphs,
  type MeasuredTextLayout,
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
export const textCapability: BoundObjectCapabilityModule = {
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
        context.path,
      );
    }
    const requestedName = parseOptionalName(input.props["name"], context.path);
    const morphKey = parseMorphKey(input.props["morphKey"], context.path);
    const link = parseLinkTarget(input.props["link"], context.path);

    const declaredX = input.props["x"];
    const declaredY = input.props["y"];
    const declaredWidth = input.props["width"];
    if (declaredX !== undefined) coordinate(declaredX, "x", context.path);
    if (declaredY !== undefined) coordinate(declaredY, "y", context.path);
    if (declaredWidth !== undefined && typeof declaredWidth === "number") {
      size(declaredWidth, "width", context.path);
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
    const requestedHeight = height(
      styleHeight ?? input.props["height"],
      context.path,
    );

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

    let compiledState:
      | {
          readonly box: GeometryBox;
          readonly measuredLayout?: MeasuredTextLayout | undefined;
        }
      | undefined;

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
      compile(compileContext: CapabilityContext) {
        const dimensions = {
          x: coordinate(
            compileContext.layoutBox?.x ?? contract.x ?? declaredX,
            "x",
            compileContext.path,
          ),
          y: coordinate(
            compileContext.layoutBox?.y ?? contract.y ?? declaredY,
            "y",
            compileContext.path,
          ),
          width: size(
            compileContext.layoutBox?.width ?? contract.width ?? declaredWidth,
            "width",
            compileContext.path,
          ),
        };
        const measuredHeight = requestedHeight.contentSized
          ? undefined
          : requestedHeight.value;
        const outputHeight = compileContext.layoutBox?.height ?? measuredHeight;
        const innerWidth = Math.max(
          0,
          dimensions.width - insets.left - insets.right,
        );
        const innerHeight =
          outputHeight === undefined
            ? undefined
            : Math.max(0, outputHeight - insets.top - insets.bottom);
        const body = writeTextBodyFromParagraphs(
          paragraphs,
          parsedTextStyle,
          compileContext.theme,
          "top",
          {
            width: innerWidth,
            ...(innerHeight === undefined ? {} : { height: innerHeight }),
            contentSized: requestedHeight.contentSized,
            insets,
          },
          "text-box",
          compileContext.options.embedFonts,
          compileContext.createTextFieldId,
        );
        const box = {
          ...dimensions,
          height: size(
            compileContext.layoutBox?.height ??
              (requestedHeight.contentSized
                ? body.measurement === undefined
                  ? undefined
                  : body.measurement.height + insets.top + insets.bottom
                : requestedHeight.value),
            "height",
            compileContext.path,
          ),
        };
        compiledState = { box, measuredLayout: body.measuredLayout };

        const properties = shapePropertiesXml({
          theme: compileContext.theme,
          style: visualStyle,
          box,
        });
        const name =
          compileContext.morphName ??
          requestedName ??
          `Text ${compileContext.id - 1}`;
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
          paragraphCount: body.paragraphCount,
        });
        let escapedName: string;
        try {
          escapedName = xmlEscape(name);
        } catch {
          throw new CapabilityError(
            "invalid-xml-character",
            "Text.name contains characters not allowed in XML 1.0",
            compileContext.path,
          );
        }
        return {
          xml: textBoxXml(
            compileContext.id,
            escapedName,
            box,
            body.xml,
            properties.xml,
            nativeLinkXml(compileContext.nativeLink),
          ),
          observation: {
            id: compileContext.id,
            kind: "text",
            name,
            box,
            text: body.text,
          },
          extensions: [],
          warnings: [...body.warnings, ...properties.warnings],
          ...(annotations.length === 0 ? {} : { annotations }),
          resourceClaims: body.resourceClaims,
          presentationContributions: fontPresentationContributions(
            body.fontUses,
            compileContext.options.embedFonts,
          ),
        };
      },
      preview(previewContext: ObjectPreviewContext): ObjectPreviewOutput {
        const box = previewContext.box ?? compiledState?.box;
        if (box === undefined) {
          throw new CapabilityError(
            "preview-missing-geometry",
            "Preview requires compiled text geometry",
          );
        }
        const visuals = resolveBoxPreviewVisuals({
          label: "text",
          visualStyle,
          box,
          theme: previewContext.theme,
        });
        const backgroundRect = visuals.hasFillOrStroke
          ? `<rect x="${box.x}" y="${box.y}" width="${box.width}" height="${box.height}" ${visuals.attributes}/>`
          : "";
        const glyphs = renderTextOutlinesSvg({
          paragraphs,
          box,
          insets,
          verticalAlign: parsedTextStyle.verticalAlign ?? "top",
          theme: previewContext.theme,
          measuredLayout:
            compiledState !== undefined && compiledState.box.width === box.width
              ? compiledState.measuredLayout
              : undefined,
        });
        return {
          svg: `<g transform="${visuals.svgTransform}">${backgroundRect}${glyphs}</g>`,
          transform: visuals.transform,
        };
      },
    };
  },
};
