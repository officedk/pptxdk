import {
  assertValidLang,
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
  type ColorValue,
  type EmitContext,
  type EmitOutput,
  type LayoutStyle,
  type LinkTarget,
  type ObjectCapabilityModule,
  type ObjectResolveContext,
  type PaddingStyle,
  type PaintContext,
  type PaintOutput,
  type PreviewTransform,
  type ResolvedObject,
} from "@react-pptx/core/capability";
import { RichTextContainer } from "@react-pptx/core/document";
import { pxToEmu, xmlEscape } from "@react-pptx/emit";

import type { Geometry } from "../shared/geometry.ts";
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
import { shapeXml } from "./shape-xml.ts";
export { linearGradient, radialGradient } from "../shared/visual.ts";
export type { PaddingStyle, VisualStyle } from "@react-pptx/core/capability";
import { motionAnnotations, type MotionProps } from "../motion/motion-api.ts";
import type { MorphProps } from "../shared/morph-api.ts";
import {
  emitTextBody,
  measureTextLayoutFromParagraphs,
  parseTextContent,
  resolveTextBody,
  textInsets,
  type Paragraph,
  type ResolvedTextBody,
  type TextStyle,
} from "../shared/text-service.ts";
import {
  emitVisualModel,
  lineEndStyleKeys,
  paintVisualModel,
  resolveVisualModel,
  visualStyleKeys,
  type ShapeVisualStyle,
} from "../shared/visual.ts";

export type ShapeDeclarationProps = MotionProps &
  MorphProps & {
    readonly link?: LinkTarget;
    readonly id?: string;
    readonly name?: string;
    readonly geometry?: Geometry;
    readonly flipH?: boolean;
    readonly flipV?: boolean;
    readonly x?: number;
    readonly y?: number;
    readonly width?: number;
    readonly height?: number;
    readonly style?: LayoutStyle & ShapeVisualStyle & TextStyle & PaddingStyle;
    readonly fill?: ColorValue;
    /** Directly writes native text into this shape without creating an additional text box; mutually exclusive with children. */
    readonly text?: string;
    /** Shape text uses shared text services and can still be emitted when text box capability is omitted. */
    readonly textStyle?: TextStyle;
  };

/** Creates the editable native shape shared by the command factory and the Shape JSX element. */
export function createShapeNode(
  props: ShapeDeclarationProps,
): RichTextContainer<
  "shape",
  ShapeDeclarationProps,
  NonNullable<ShapeDeclarationProps["style"]>
> {
  return new RichTextContainer<
    "shape",
    ShapeDeclarationProps,
    NonNullable<ShapeDeclarationProps["style"]>
  >("shape", "rp:shape", props);
}

const maximumShapeExtentEmu = 2_147_483_647;

const allowedParameters = [
  "id",
  "name",
  "morphKey",
  "link",
  "geometry",
  "flipH",
  "flipV",
  "x",
  "y",
  "width",
  "height",
  "style",
  "fill",
  "text",
  "textStyle",
  "children",
  "enter",
  "emphasis",
  "motion",
] as const;

export const shapeStyleKeys = [
  ...layoutStyleKeys,
  ...visualStyleKeys,
  ...lineEndStyleKeys,
  ...textStyleKeys,
  ...paddingStyleKeys,
] as const;

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
  const inherited = Object.fromEntries(
    Object.entries(layoutStyle).filter(([key]) => textStyleKeys.has(key)),
  );
  const explicit = props["textStyle"];
  if (explicit === undefined) return inherited;
  if (
    typeof explicit !== "object" ||
    explicit === null ||
    Array.isArray(explicit)
  ) {
    return inherited;
  }
  const explicitObj = explicit as Record<string, unknown>;
  if (explicitObj["lang"] !== undefined) {
    assertValidLang(explicitObj["lang"]);
  }
  return { ...inherited, ...explicitObj };
}

function shapeTextContent(
  input: CapabilityInput,
  path?: string,
): readonly unknown[] {
  const text = input.props["text"];
  if (text !== undefined) {
    if (typeof text !== "string") {
      throw new CapabilityError(
        "invalid-text",
        "Shape.text must be a string",
        path,
      );
    }
    if (input.children.length > 0) {
      throw new CapabilityError(
        "conflicting-text-content",
        "Shape.text and Shape children text cannot be declared simultaneously",
        path,
      );
    }
    return text === "" ? [] : [{ type: "#text", text }];
  }
  return input.children;
}

function validateCoordinate(
  value: unknown,
  name: "x" | "y",
  path?: string,
): number {
  if (
    typeof value !== "number" ||
    !Number.isFinite(value) ||
    value < -1_000_000 ||
    value > 1_000_000
  ) {
    throw new CapabilityError(
      "invalid-coordinate",
      `${name} must be a finite pixel number, width and height must be greater than 0, and coordinates must be between -1000000 and 1000000`,
      path,
    );
  }
  return value;
}

function validateSize(
  value: unknown,
  name: "width" | "height",
  path?: string,
): number {
  if (typeof value !== "number" || !Number.isFinite(value) || value <= 0) {
    throw new CapabilityError(
      "invalid-coordinate",
      `${name} must be a finite pixel number, width and height must be greater than 0, and coordinates must be between -1000000 and 1000000`,
      path,
    );
  }
  const extentEmu = pxToEmu(value);
  if (extentEmu < 1 || extentEmu > maximumShapeExtentEmu) {
    throw new CapabilityError(
      "invalid-coordinate",
      `${name} must convert to an EMU value between 1 and ${maximumShapeExtentEmu} inclusive`,
      path,
    );
  }
  return value;
}

/** Shape module independently validates input, encapsulates bound state, and emits native PresentationML objects. */
export const shapeCapability: ObjectCapabilityModule = {
  declaration: "rp:shape",

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
        `Shape contains unsupported parameter ${unknown}`,
      );
    }

    const requestedName = parseOptionalName(input.props["name"]);
    const morphKey = parseMorphKey(input.props["morphKey"]);
    const link = parseLinkTarget(input.props["link"]);

    const declaredX = input.props["x"];
    const declaredY = input.props["y"];
    if (declaredX !== undefined) {
      validateCoordinate(declaredX, "x");
    }
    if (declaredY !== undefined) {
      validateCoordinate(declaredY, "y");
    }
    if (
      input.props["width"] !== undefined &&
      typeof input.props["width"] === "number"
    ) {
      validateSize(input.props["width"], "width");
    }
    if (
      input.props["height"] !== undefined &&
      typeof input.props["height"] === "number"
    ) {
      validateSize(input.props["height"], "height");
    }

    const textProp = input.props["text"];
    if (textProp !== undefined && typeof textProp !== "string") {
      throw new CapabilityError("invalid-text", "Shape.text must be a string");
    }
    if (textProp !== undefined && input.children.length > 0) {
      throw new CapabilityError(
        "conflicting-text-content",
        "Shape.text and Shape children text cannot be declared simultaneously",
      );
    }

    const hasText =
      (typeof textProp === "string" && textProp !== "") ||
      (textProp === undefined && input.children.length > 0);

    const allowedStyleKeys = new Set<string>([
      ...layoutStyleKeys,
      ...visualStyleKeys,
      ...lineEndStyleKeys,
      ...(hasText ? [...textStyleKeys, ...paddingStyleKeys] : []),
    ]);

    const contract = parseCapabilityLayoutContract({
      nativeKind: "sp",
      props: input.props,
      allowedStyleKeys,
      allowPadding: hasText,
      path: context.path,
    });

    const insets = hasText
      ? textInsets(input.props["style"])
      : { top: 0, right: 0, bottom: 0, left: 0 };

    const textChildren = shapeTextContent(input);
    const rawTextStyle = shapeTextStyle(input.props);
    const effectiveTextStyle = {
      ...rawTextStyle,
      verticalAlign: rawTextStyle["verticalAlign"] ?? "middle",
    };
    const { paragraphs, style: parsedTextStyle } = hasText
      ? parseTextContent(
          textChildren,
          effectiveTextStyle,
          context.theme,
          { slideNumber: context.slideNumber },
          "Shape.style",
        )
      : {
          paragraphs: [] as readonly Paragraph[],
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
    const flipH = input.props["flipH"] as boolean | undefined;
    const flipV = input.props["flipV"] as boolean | undefined;
    const fill = input.props["fill"] as ColorValue | undefined;
    const explicitText = typeof textProp === "string" ? textProp : undefined;
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

    // Validate geometry and visual properties during bind before layout runs.
    resolveVisualModel({
      theme: context.theme,
      style: visualStyle,
      box: { x: 0, y: 0, width: 100, height: 100 },
      ...(geometry !== undefined ? { geometry } : {}),
      ...(flipH !== undefined ? { flipH } : {}),
      ...(flipV !== undefined ? { flipV } : {}),
      ...(fill !== undefined ? { fill } : {}),
    });

    const declaredWidth = contract.width;
    const declaredHeight = contract.height;
    const hasFixedSize =
      typeof declaredWidth === "number" && typeof declaredHeight === "number";

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
        if (!hasText || paragraphs.length === 0) {
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
            hasFixedSize &&
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
        const box =
          resolveContext.layoutBox === undefined
            ? {
                x: validateCoordinate(declaredX, "x"),
                y: validateCoordinate(declaredY, "y"),
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
          ...(flipH === undefined ? {} : { flipH }),
          ...(flipV === undefined ? {} : { flipV }),
          ...(fill === undefined ? {} : { fill }),
        });
        const geometryName =
          visual.geometry.kind === "custom" ? "custom" : visual.geometry.name;

        let resolvedText: ResolvedTextBody | undefined;
        if (hasText && paragraphs.length > 0) {
          resolvedText = resolveTextBody(
            paragraphs,
            parsedTextStyle,
            resolveContext.theme,
            parsedTextStyle.verticalAlign ?? "middle",
            {
              width: Math.max(0, box.width - insets.left - insets.right),
              height: Math.max(0, box.height - insets.top - insets.bottom),
              insets,
            },
            "text-box",
          );
        }
        const name =
          resolveContext.morphName ??
          requestedName ??
          `Shape ${resolveContext.id - 1}`;
        const annotations = motionAnnotations(motionProps, {
          kind: "shape",
          name,
          ...(authorId === undefined ? {} : { authorId }),
          textBody: resolvedText !== undefined,
          filled:
            fill !== undefined ||
            visualStyle["backgroundColor"] !== undefined ||
            visualStyle["backgroundImage"] !== undefined,
          solidFilled:
            (fill ?? visualStyle["backgroundColor"]) !== undefined &&
            (fill ?? visualStyle["backgroundColor"]) !== "none" &&
            visualStyle["backgroundImage"] === undefined,
          box,
          rotationDegrees: visual.rotate,
          hasBackground:
            fill !== undefined ||
            visualStyle["backgroundColor"] !== undefined ||
            visualStyle["backgroundImage"] !== undefined ||
            visualStyle["borderColor"] !== undefined ||
            (typeof visualStyle["borderWidth"] === "number" &&
              visualStyle["borderWidth"] > 0),
          ...(resolvedText === undefined
            ? {}
            : { paragraphCount: resolvedText.paragraphCount }),
        });
        let escapedName: string;
        try {
          escapedName = xmlEscape(name);
        } catch {
          throw new CapabilityError(
            "invalid-xml-character",
            "Shape.name contains characters not allowed in XML 1.0",
          );
        }
        const warnings = [
          ...visual.warnings,
          ...(resolvedText?.warnings ?? []),
        ];
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
              ? explicitText === undefined
                ? {}
                : { text: explicitText }
              : { text: resolvedText.text }),
          },
          transform,
          extensions: [],
          ...(warnings.length === 0 ? {} : { warnings }),
          ...(annotations.length === 0 ? {} : { annotations }),
          ...(resolvedText === undefined
            ? {}
            : { resourceClaims: resolvedText.resourceClaims }),
          emit(emitContext: EmitContext): EmitOutput {
            const embedFonts = emitContext.options.embedFonts;
            const emittedBody =
              resolvedText === undefined
                ? undefined
                : emitTextBody(resolvedText, emitContext.theme, {
                    createFieldId: emitContext.createTextFieldId,
                    ...(embedFonts === undefined ? {} : { embedFonts }),
                  });
            const properties = emitVisualModel(visual, (resource) =>
              emitContext.registerResource(resource),
            );
            return {
              xml: shapeXml(
                emitContext.id,
                escapedName,
                properties.xml,
                emittedBody?.xml ?? "",
                nativeLinkXml(emitContext.nativeLink),
              ),
              ...(emittedBody === undefined
                ? {}
                : {
                    presentationContributions:
                      emittedBody.presentationContributions,
                  }),
            };
          },
          paint(paintContext: PaintContext): PaintOutput {
            if (hasText) {
              throw new CapabilityError(
                "preview-unsupported-text",
                "Preview does not yet support text inside shapes",
              );
            }
            return paintVisualModel({
              model: visual,
              slideNumber: paintContext.slideNumber,
              objectId: paintContext.objectId,
            });
          },
        };
      },
    };
  },
};
