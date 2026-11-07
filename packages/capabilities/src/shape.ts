import {
  assertValidLang,
  CapabilityError,
  layoutStyleKeys,
  paddingStyleKeys,
  parseCapabilityLayoutContract,
  pxToEmu,
  textStyleKeys,
  xmlEscape,
  type CapabilityContext,
  type CapabilityInput,
  type CapabilityLayoutContract,
  type CapabilityLayoutMeasurement,
  type CapabilityLayoutRequest,
  type CapabilityParseContext,
  type CapabilityParseResult,
  type ColorValue,
  type LayoutStyle,
  type LinkTarget,
  type ObjectCapabilityModule,
  type PaddingStyle,
  type VisualStyle,
} from "@react-pptx/core";
import type { ReactNode } from "react";

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
export { linearGradient, radialGradient } from "./visual.ts";
export type { PaddingStyle, VisualStyle } from "@react-pptx/core";
import { fontPresentationContributions } from "./font-embedding.ts";
import type { MorphProps } from "./morph-api.ts";
import {
  motionAnnotations,
  motionElement,
  type MotionProps,
} from "./motion-api.ts";
import {
  measureTextLayoutFromParagraphs,
  parseTextContent,
  textInsets,
  writeTextBodyFromParagraphs,
  type Paragraph,
  type ResolvedTextStyle,
  type TextServiceResult,
  type TextStyle,
} from "./text-service.ts";
import {
  shapePropertiesXml,
  visualStyleKeys,
  type VisualStyleRecord,
} from "./visual.ts";

export type ShapeProps = MotionProps &
  MorphProps & {
    readonly link?: LinkTarget;
    readonly testId?: string;
    readonly part?: string;
    readonly name?: string;
    readonly geometry?: Geometry;
    readonly flipH?: boolean;
    readonly flipV?: boolean;
    readonly x?: number;
    readonly y?: number;
    readonly width?: number;
    readonly height?: number;
    readonly style?: LayoutStyle & VisualStyle & TextStyle & PaddingStyle;
    readonly fill?: ColorValue;
    /** Directly writes native text into this shape without creating an additional text box; mutually exclusive with children. */
    readonly text?: string;
    /** Shape text uses shared text services and can still be emitted when text box capability is omitted. */
    readonly textStyle?: TextStyle;
    /** Native text nodes shared with Text; may contain P, Span, and SlideNumber. */
    readonly children?: ReactNode;
  };

export type ParsedShape = {
  readonly testId?: string | undefined;
  readonly part?: string | undefined;
  readonly name?: string | undefined;
  readonly morphKey?: unknown;
  readonly link?: unknown;
  readonly geometry?: Geometry | undefined;
  readonly flipH?: boolean | undefined;
  readonly flipV?: boolean | undefined;
  readonly fill?: ColorValue | undefined;
  readonly visualStyle: VisualStyleRecord;
  readonly insets: {
    readonly top: number;
    readonly right: number;
    readonly bottom: number;
    readonly left: number;
  };
  readonly hasText: boolean;
  readonly text?: string | undefined;
  readonly paragraphs: readonly Paragraph[];
  readonly textStyle: ResolvedTextStyle;
  readonly requestedHeight:
    | { readonly contentSized: true }
    | { readonly contentSized: false; readonly value: number | undefined };
  readonly contract: CapabilityLayoutContract;
  readonly rawProps: Readonly<Record<string, unknown>>;
};

const maximumShapeExtentEmu = 2_147_483_647;

const allowedParameters = [
  "testId",
  "part",
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
  "__motionContext",
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

function requiredNumber(
  props: Record<string, unknown>,
  name: "x" | "y" | "width" | "height",
  path?: string,
): number {
  const value = props[name];
  const isExtent = name === "width" || name === "height";
  return isExtent
    ? validateSize(value, name, path)
    : validateCoordinate(value, name, path);
}

function optionalText(
  props: Record<string, unknown>,
  name: "testId" | "name" | "part",
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

function layoutStyleValue(
  input: CapabilityInput,
  name: "width" | "height",
): number | "auto" | `${number}%` | undefined {
  const value = input.props["style"];
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    return undefined;
  }
  const dimension = (value as Record<string, unknown>)[name];
  if (typeof dimension === "number" || dimension === "auto") return dimension;
  if (
    typeof dimension === "string" &&
    /^-?(?:\d+(?:\.\d*)?|\.\d+)%$/u.test(dimension)
  ) {
    return dimension as `${number}%`;
  }
  return undefined;
}

function height(
  value: unknown,
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
  return {
    contentSized: false,
    value: typeof value === "number" ? value : undefined,
  };
}

function shapeXml(
  id: number,
  escapedName: string,
  properties: string,
  textBody: string,
  linkXml: string,
): string {
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${escapedName}">${linkXml}</p:cNvPr><p:cNvSpPr/><p:nvPr/></p:nvSpPr>${properties}${textBody}</p:sp>`;
}

/** Shape module independently validates input, declares native geometry, and emits corresponding PresentationML objects. */
export const shapeCapability: ObjectCapabilityModule<ParsedShape> = {
  declaration: "rp:shape",

  parse(
    input: CapabilityInput,
    context: CapabilityParseContext,
  ): CapabilityParseResult<ParsedShape> {
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

    const testId = optionalText(input.props, "testId", context.path);
    const part = optionalText(input.props, "part", context.path);
    const requestedName = optionalText(input.props, "name", context.path);

    if (input.props["x"] !== undefined) {
      validateCoordinate(input.props["x"], "x", context.path);
    }
    if (input.props["y"] !== undefined) {
      validateCoordinate(input.props["y"], "y", context.path);
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
    const styleHeight = layoutStyleValue(input, "height");
    const requestedHeight = height(styleHeight ?? input.props["height"]);

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

    const visualStyle = record(input.props["style"]);

    // Validate geometry and visual properties
    shapePropertiesXml({
      theme: context.theme,
      style: visualStyle,
      box: { x: 0, y: 0, width: 100, height: 100 },
      ...(input.props["geometry"] !== undefined
        ? { geometry: input.props["geometry"] as Geometry }
        : {}),
      ...(input.props["flipH"] !== undefined
        ? { flipH: input.props["flipH"] as boolean }
        : {}),
      ...(input.props["flipV"] !== undefined
        ? { flipV: input.props["flipV"] as boolean }
        : {}),
      ...(input.props["fill"] !== undefined
        ? { fill: input.props["fill"] as ColorValue }
        : {}),
    });

    const parsed: ParsedShape = {
      ...(testId !== undefined ? { testId } : {}),
      ...(part !== undefined ? { part } : {}),
      ...(requestedName !== undefined ? { name: requestedName } : {}),
      ...(input.props["morphKey"] !== undefined
        ? { morphKey: input.props["morphKey"] }
        : {}),
      ...(input.props["link"] !== undefined
        ? { link: input.props["link"] }
        : {}),
      ...(input.props["geometry"] !== undefined
        ? { geometry: input.props["geometry"] as Geometry }
        : {}),
      ...(input.props["flipH"] !== undefined
        ? { flipH: input.props["flipH"] as boolean }
        : {}),
      ...(input.props["flipV"] !== undefined
        ? { flipV: input.props["flipV"] as boolean }
        : {}),
      ...(input.props["fill"] !== undefined
        ? { fill: input.props["fill"] as ColorValue }
        : {}),
      visualStyle,
      insets,
      hasText,
      ...(typeof textProp === "string" ? { text: textProp } : {}),
      paragraphs,
      textStyle: parsedTextStyle,
      requestedHeight,
      contract,
      rawProps: input.props,
    };

    return { contract, parsed };
  },

  measureLayout(
    inputOrParsed: ParsedShape | CapabilityInput,
    request: CapabilityLayoutRequest,
  ): CapabilityLayoutMeasurement | undefined {
    const parsed: ParsedShape =
      "paragraphs" in inputOrParsed && "contract" in inputOrParsed
        ? inputOrParsed
        : (
            this.parse as NonNullable<
              (
                input: CapabilityInput,
                context: CapabilityParseContext,
              ) => CapabilityParseResult<ParsedShape>
            >
          )(inputOrParsed, {
            path: "Shape",
            theme: request.theme,
            options: {},
            slideNumber: request.slideNumber,
          }).parsed;

    if (!parsed.hasText || parsed.paragraphs.length === 0) {
      return { minContentWidth: 0, minContentHeight: 0 };
    }

    const insets = parsed.insets;
    const styleWidth = layoutStyleValue(
      { type: "rp:shape", props: parsed.rawProps, children: [] },
      "width",
    );
    const width =
      request.availableWidth ?? styleWidth ?? parsed.rawProps["width"];
    const innerWidth =
      request.availableWidth !== undefined
        ? request.availableWidth
        : typeof width === "number"
          ? Math.max(0, width - insets.left - insets.right)
          : undefined;

    const hasFixedSize =
      typeof (styleWidth ?? parsed.rawProps["width"]) === "number" &&
      typeof (
        layoutStyleValue(
          { type: "rp:shape", props: parsed.rawProps, children: [] },
          "height",
        ) ?? parsed.rawProps["height"]
      ) === "number";

    try {
      const measured = measureTextLayoutFromParagraphs(
        parsed.paragraphs,
        parsed.textStyle,
        innerWidth,
      );
      const baseline =
        parsed.textStyle.verticalAlign === "top"
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

  compile(
    inputOrParsed: ParsedShape | CapabilityInput,
    context: CapabilityContext,
  ) {
    const parsed: ParsedShape =
      "paragraphs" in inputOrParsed && "contract" in inputOrParsed
        ? inputOrParsed
        : (
            this.parse as NonNullable<
              (
                input: CapabilityInput,
                context: CapabilityParseContext,
              ) => CapabilityParseResult<ParsedShape>
            >
          )(inputOrParsed, {
            path: context.path,
            theme: context.theme,
            options: context.options,
            slideNumber: context.slideNumber,
          }).parsed;

    const props = parsed.rawProps;
    const testId = parsed.testId;
    const part = parsed.part;
    const requestedName = parsed.name;
    const geometry = parsed.geometry;
    const layoutStyle = parsed.visualStyle;
    const insets = parsed.insets;
    const box =
      context.layoutBox === undefined
        ? {
            x: requiredNumber(props, "x", context.path),
            y: requiredNumber(props, "y", context.path),
            width: requiredNumber(
              { width: layoutStyle["width"] ?? props["width"] },
              "width",
              context.path,
            ),
            height: requiredNumber(
              { height: layoutStyle["height"] ?? props["height"] },
              "height",
              context.path,
            ),
          }
        : {
            x: requiredNumber({ x: context.layoutBox.x }, "x", context.path),
            y: requiredNumber({ y: context.layoutBox.y }, "y", context.path),
            width: requiredNumber(
              { width: context.layoutBox.width },
              "width",
              context.path,
            ),
            height: requiredNumber(
              { height: context.layoutBox.height },
              "height",
              context.path,
            ),
          };
    const properties = shapePropertiesXml({
      theme: context.theme,
      style: layoutStyle,
      box,
      ...(geometry === undefined ? {} : { geometry }),
      ...(parsed.flipH === undefined ? {} : { flipH: parsed.flipH }),
      ...(parsed.flipV === undefined ? {} : { flipV: parsed.flipV }),
      ...(parsed.fill === undefined ? {} : { fill: parsed.fill }),
    });

    let textOutput: TextServiceResult | undefined;
    if (parsed.hasText && parsed.paragraphs.length > 0) {
      textOutput = writeTextBodyFromParagraphs(
        parsed.paragraphs,
        parsed.textStyle,
        context.theme,
        parsed.textStyle.verticalAlign ?? "middle",
        {
          width: Math.max(0, box.width - insets.left - insets.right),
          height: Math.max(0, box.height - insets.top - insets.bottom),
          insets,
        },
        "text-box",
        context.options.embedFonts,
        context.createTextFieldId,
      );
    }
    const name =
      context.morphName ?? requestedName ?? testId ?? `Shape ${context.id - 1}`;
    const annotations = motionAnnotations(props, {
      kind: "shape",
      name,
      ...(testId === undefined ? {} : { testId }),
      textBody: textOutput !== undefined,
      filled:
        props["fill"] !== undefined ||
        layoutStyle["backgroundColor"] !== undefined ||
        layoutStyle["backgroundImage"] !== undefined,
      solidFilled:
        (props["fill"] ?? layoutStyle["backgroundColor"]) !== undefined &&
        (props["fill"] ?? layoutStyle["backgroundColor"]) !== "none" &&
        layoutStyle["backgroundImage"] === undefined,
      box,
      rotationDegrees:
        typeof layoutStyle["rotate"] === "number" ? layoutStyle["rotate"] : 0,
      hasBackground:
        props["fill"] !== undefined ||
        layoutStyle["backgroundColor"] !== undefined ||
        layoutStyle["backgroundImage"] !== undefined ||
        layoutStyle["borderColor"] !== undefined ||
        (typeof layoutStyle["borderWidth"] === "number" &&
          layoutStyle["borderWidth"] > 0),
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
        context.path,
      );
    }
    return {
      xml: shapeXml(
        context.id,
        escapedName,
        properties.xml,
        textOutput?.xml ?? "",
        nativeLinkXml(context.nativeLink),
      ),
      observation: {
        id: context.id,
        kind: "shape",
        name,
        ...(testId === undefined ? {} : { testId }),
        ...(part === undefined ? {} : { part }),
        geometry: properties.geometryName,
        box,
        ...(textOutput === undefined
          ? parsed.text === undefined
            ? {}
            : { text: parsed.text }
          : { text: textOutput.text }),
      },
      extensions: [],
      ...(properties.warnings.length === 0 && textOutput === undefined
        ? {}
        : {
            warnings: [...properties.warnings, ...(textOutput?.warnings ?? [])],
          }),
      ...(annotations.length === 0 ? {} : { annotations }),
      ...(textOutput === undefined
        ? {}
        : {
            resourceClaims: textOutput.resourceClaims,
            presentationContributions: fontPresentationContributions(
              textOutput.fontUses,
              context.options.embedFonts,
            ),
          }),
    };
  },
};

/** Declares a native object handled by the shape module within the React tree. */
export function Shape(props: ShapeProps): ReactNode {
  return motionElement("rp:shape", props);
}
