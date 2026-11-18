import {
  assertValidLang,
  CapabilityError,
  isMorphKey,
  layoutStyleKeys,
  paddingStyleKeys,
  parseCapabilityLayoutContract,
  pxToEmu,
  textStyleKeys,
  xmlEscape,
  type BoundObjectCapability,
  type BoundObjectCapabilityModule,
  type BoundObjectNativeCollaboration,
  type CapabilityContext,
  type CapabilityInput,
  type CapabilityLayoutMeasurement,
  type CapabilityLayoutRequest,
  type CapabilityParseContext,
  type ColorValue,
  type LayoutStyle,
  type LinkTarget,
  type MorphKey,
  type PaddingStyle,
} from "@react-pptx/core/capability";
import { RichTextContainer } from "@react-pptx/core/document";

import type { Geometry } from "./geometry.ts";
export { blockArc, path, preset } from "./geometry.ts";
export type {
  CustomGeometry,
  Geometry,
  GeometryPath,
  PathCommand,
  PathOptions,
  PresetGeometry,
} from "./geometry.ts";
import { nativeLinkXml } from "./identity.ts";
import { shapeXml } from "./shape-xml.ts";
export { linearGradient, radialGradient } from "./visual.ts";
export type { PaddingStyle, VisualStyle } from "@react-pptx/core/capability";
import { fontPresentationContributions } from "./font-embedding.ts";
import type { MorphProps } from "./morph-api.ts";
import { motionAnnotations, type MotionProps } from "./motion-api.ts";
import {
  measureTextLayoutFromParagraphs,
  parseTextContent,
  textInsets,
  writeTextBodyFromParagraphs,
  type Paragraph,
  type TextServiceResult,
  type TextStyle,
} from "./text-service.ts";
import {
  lineEndStyleKeys,
  shapePropertiesXml,
  visualStyleKeys,
  type ShapeVisualStyle,
} from "./visual.ts";

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

function parseLinkTarget(
  value: unknown,
  path?: string,
): LinkTarget | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    throw new CapabilityError(
      "invalid-link",
      "link must specify id, url, or home action",
      path,
    );
  }
  const entries = Object.entries(value).filter(
    ([, target]) => target !== undefined,
  );
  if (entries.length !== 1) {
    throw new CapabilityError(
      "invalid-link",
      "link must specify exactly one target",
      path,
    );
  }
  const [kind, target] = entries[0] ?? [];
  if (kind === "id" && typeof target === "string") {
    return { id: target };
  }
  if (kind === "url" && typeof target === "string") {
    return { url: target };
  }
  if (kind === "action" && target === "home") {
    return { action: "home" };
  }
  throw new CapabilityError(
    "invalid-link",
    "link supports { id }, { url }, or { action: 'home' }",
    path,
  );
}

function optionalText(
  props: Record<string, unknown>,
  name: "name",
  path?: string,
): string | undefined {
  const value = props[name];
  if (value === undefined) return undefined;
  if (typeof value !== "string" || value.trim() === "") {
    throw new CapabilityError(
      "invalid-parameter",
      `${name} must be a non-empty string`,
      path,
    );
  }
  return value;
}

/** Shape module independently validates input, encapsulates bound state, and emits native PresentationML objects. */
export const shapeCapability: BoundObjectCapabilityModule = {
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
        context.path,
      );
    }

    const requestedName = optionalText(input.props, "name", context.path);
    const rawMorphKey = input.props["morphKey"];
    let morphKey: MorphKey | undefined;
    if (rawMorphKey !== undefined) {
      if (!isMorphKey(rawMorphKey)) {
        throw new CapabilityError(
          "invalid-morph-key",
          "morphKey must be created by createMorphScope().handle(name)",
          context.path,
        );
      }
      morphKey = rawMorphKey;
    }
    const link = parseLinkTarget(input.props["link"], context.path);

    const declaredX = input.props["x"];
    const declaredY = input.props["y"];
    if (declaredX !== undefined) {
      validateCoordinate(declaredX, "x", context.path);
    }
    if (declaredY !== undefined) {
      validateCoordinate(declaredY, "y", context.path);
    }
    if (
      input.props["width"] !== undefined &&
      typeof input.props["width"] === "number"
    ) {
      validateSize(input.props["width"], "width", context.path);
    }
    if (
      input.props["height"] !== undefined &&
      typeof input.props["height"] === "number"
    ) {
      validateSize(input.props["height"], "height", context.path);
    }

    const textProp = input.props["text"];
    if (textProp !== undefined && typeof textProp !== "string") {
      throw new CapabilityError(
        "invalid-text",
        "Shape.text must be a string",
        context.path,
      );
    }
    if (textProp !== undefined && input.children.length > 0) {
      throw new CapabilityError(
        "conflicting-text-content",
        "Shape.text and Shape children text cannot be declared simultaneously",
        context.path,
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

    const textChildren = shapeTextContent(input, context.path);
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
    shapePropertiesXml({
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
      compile(compileContext: CapabilityContext) {
        const box =
          compileContext.layoutBox === undefined
            ? {
                x: validateCoordinate(declaredX, "x", compileContext.path),
                y: validateCoordinate(declaredY, "y", compileContext.path),
                width: validateSize(
                  declaredWidth,
                  "width",
                  compileContext.path,
                ),
                height: validateSize(
                  declaredHeight,
                  "height",
                  compileContext.path,
                ),
              }
            : {
                x: validateCoordinate(
                  compileContext.layoutBox.x,
                  "x",
                  compileContext.path,
                ),
                y: validateCoordinate(
                  compileContext.layoutBox.y,
                  "y",
                  compileContext.path,
                ),
                width: validateSize(
                  compileContext.layoutBox.width,
                  "width",
                  compileContext.path,
                ),
                height: validateSize(
                  compileContext.layoutBox.height,
                  "height",
                  compileContext.path,
                ),
              };
        const properties = shapePropertiesXml({
          theme: compileContext.theme,
          style: visualStyle,
          box,
          ...(geometry === undefined ? {} : { geometry }),
          ...(flipH === undefined ? {} : { flipH }),
          ...(flipV === undefined ? {} : { flipV }),
          ...(fill === undefined ? {} : { fill }),
        });

        let textOutput: TextServiceResult | undefined;
        if (hasText && paragraphs.length > 0) {
          textOutput = writeTextBodyFromParagraphs(
            paragraphs,
            parsedTextStyle,
            compileContext.theme,
            parsedTextStyle.verticalAlign ?? "middle",
            {
              width: Math.max(0, box.width - insets.left - insets.right),
              height: Math.max(0, box.height - insets.top - insets.bottom),
              insets,
            },
            "text-box",
            compileContext.options.embedFonts,
            compileContext.createTextFieldId,
          );
        }
        const name =
          compileContext.morphName ??
          requestedName ??
          `Shape ${compileContext.id - 1}`;
        const annotations = motionAnnotations(motionProps, {
          kind: "shape",
          name,
          ...(authorId === undefined ? {} : { authorId }),
          textBody: textOutput !== undefined,
          filled:
            fill !== undefined ||
            visualStyle["backgroundColor"] !== undefined ||
            visualStyle["backgroundImage"] !== undefined,
          solidFilled:
            (fill ?? visualStyle["backgroundColor"]) !== undefined &&
            (fill ?? visualStyle["backgroundColor"]) !== "none" &&
            visualStyle["backgroundImage"] === undefined,
          box,
          rotationDegrees:
            typeof visualStyle["rotate"] === "number"
              ? visualStyle["rotate"]
              : 0,
          hasBackground:
            fill !== undefined ||
            visualStyle["backgroundColor"] !== undefined ||
            visualStyle["backgroundImage"] !== undefined ||
            visualStyle["borderColor"] !== undefined ||
            (typeof visualStyle["borderWidth"] === "number" &&
              visualStyle["borderWidth"] > 0),
          ...(textOutput === undefined
            ? {}
            : { paragraphCount: textOutput.paragraphCount }),
        });
        let escapedName: string;
        try {
          escapedName = xmlEscape(name);
        } catch {
          throw new CapabilityError(
            "invalid-xml-character",
            "Shape.name contains characters not allowed in XML 1.0",
            compileContext.path,
          );
        }
        return {
          xml: shapeXml(
            compileContext.id,
            escapedName,
            properties.xml,
            textOutput?.xml ?? "",
            nativeLinkXml(compileContext.nativeLink),
          ),
          observation: {
            id: compileContext.id,
            kind: "shape",
            name,
            geometry: properties.geometryName,
            box,
            ...(textOutput === undefined
              ? explicitText === undefined
                ? {}
                : { text: explicitText }
              : { text: textOutput.text }),
          },
          extensions: [],
          ...(properties.warnings.length === 0 && textOutput === undefined
            ? {}
            : {
                warnings: [
                  ...properties.warnings,
                  ...(textOutput?.warnings ?? []),
                ],
              }),
          ...(annotations.length === 0 ? {} : { annotations }),
          ...(textOutput === undefined
            ? {}
            : {
                resourceClaims: textOutput.resourceClaims,
                presentationContributions: fontPresentationContributions(
                  textOutput.fontUses,
                  compileContext.options.embedFonts,
                ),
              }),
        };
      },
    };
  },
};
