import {
  assertValidLang,
  CapabilityError,
  pxToEmu,
  xmlEscape,
  type CapabilityLayoutMeasurement,
  type CapabilityLayoutRequest,
  type ColorValue,
  type LayoutStyle,
  type LinkTarget,
  type TextStyle as CoreTextStyle,
  type VisualStyle,
  type CapabilityInput,
  type CapabilityModule,
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
export type { VisualStyle } from "@react-pptx/core";
import { fontPresentationContributions } from "./font-embedding.ts";
import type { MorphProps } from "./morph-api.ts";
import {
  motionAnnotations,
  motionElement,
  type MotionProps,
} from "./motion-api.ts";
import {
  measureTextLayout,
  textInsets,
  writeTextBody,
  type TextStyle,
} from "./text-service.ts";
import { shapePropertiesXml } from "./visual.ts";

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
    readonly style?: LayoutStyle & VisualStyle & CoreTextStyle;
    readonly fill?: ColorValue;
    /** Directly writes native text into this shape without creating an additional text box; mutually exclusive with children. */
    readonly text?: string;
    /** Shape text uses shared text services and can still be emitted when text box capability is omitted. */
    readonly textStyle?: TextStyle;
    /** Native text nodes shared with Text; may contain P, Span, and SlideNumber. */
    readonly children?: ReactNode;
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
  "enter",
  "emphasis",
  "motion",
  "__motionContext",
] as const;

function record(input: CapabilityInput): Record<string, unknown> {
  return input.props;
}

function shapeTextMeasurementStyle(value: unknown): unknown {
  if (value === undefined) return { verticalAlign: "middle" };
  if (typeof value !== "object" || value === null || Array.isArray(value))
    return value;
  const style = value as Record<string, unknown>;
  // Shape text defaults to middle anchor when emitted; measurement must use the same effective anchor.
  return { ...style, verticalAlign: style["verticalAlign"] ?? "middle" };
}

const shapeTextStyleKeys = new Set([
  "fontFamily",
  "fontSize",
  "fontWeight",
  "fontStyle",
  "color",
  "lineHeight",
  "verticalAlign",
  "align",
  "spacingAfter",
  "letterSpacing",
  "textTransform",
  "textDecoration",
  "textShadow",
  "textWrap",
  "lang",
]);

function shapeTextStyle(props: Readonly<Record<string, unknown>>): unknown {
  const layoutStyle =
    typeof props["style"] === "object" &&
    props["style"] !== null &&
    !Array.isArray(props["style"])
      ? (props["style"] as Record<string, unknown>)
      : {};
  if (layoutStyle["lang"] !== undefined) {
    assertValidLang(layoutStyle["lang"]);
  }
  const inherited = Object.fromEntries(
    Object.entries(layoutStyle).filter(([key]) => shapeTextStyleKeys.has(key)),
  );
  const explicit = props["textStyle"];
  if (explicit === undefined) return inherited;
  if (
    typeof explicit !== "object" ||
    explicit === null ||
    Array.isArray(explicit)
  )
    return explicit;
  const explicitObj = explicit as Record<string, unknown>;
  if (explicitObj["lang"] !== undefined) {
    assertValidLang(explicitObj["lang"]);
  }
  return { ...inherited, ...explicitObj };
}

function shapeTextContent(input: CapabilityInput): readonly unknown[] {
  const text = input.props["text"];
  if (typeof text === "string") {
    if (input.children.length > 0) {
      throw new CapabilityError(
        "conflicting-text-content",
        "Shape.text and Shape children text cannot be declared simultaneously",
      );
    }
    return text === "" ? [] : [{ type: "#text", text }];
  }
  return text === undefined ? input.children : [];
}

function requiredNumber(
  props: Record<string, unknown>,
  name: "x" | "y" | "width" | "height",
): number {
  const value = props[name];
  const isExtent = name === "width" || name === "height";
  const minimum = isExtent ? 0 : -1_000_000;
  if (
    typeof value !== "number" ||
    !Number.isFinite(value) ||
    (isExtent ? value <= minimum : value < minimum || value > 1_000_000)
  ) {
    throw new CapabilityError(
      "invalid-coordinate",
      `${name} must be a finite pixel number, width and height must be greater than 0, and coordinates must be between -1000000 and 1000000`,
    );
  }
  if (isExtent) {
    const extentEmu = pxToEmu(value);
    if (extentEmu < 1 || extentEmu > maximumShapeExtentEmu) {
      throw new CapabilityError(
        "invalid-coordinate",
        `${name} must convert to an EMU value between 1 and ${maximumShapeExtentEmu} inclusive`,
      );
    }
  }
  return value;
}

function optionalText(
  props: Record<string, unknown>,
  name: "testId" | "name" | "part",
): string | undefined {
  const value = props[name];
  if (value === undefined) return undefined;
  if (typeof value !== "string" || value.trim() === "") {
    throw new CapabilityError(
      "invalid-parameter",
      `${name} must be a non-empty string`,
    );
  }
  return value;
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
export const shapeCapability: CapabilityModule = {
  declaration: "rp:shape",
  measureLayout(
    input: CapabilityInput,
    request: CapabilityLayoutRequest,
  ): CapabilityLayoutMeasurement | undefined {
    const children = shapeTextContent(input);
    if (children.length === 0) {
      return { minContentWidth: 0, minContentHeight: 0 };
    }
    try {
      const measured = measureTextLayout(
        children,
        shapeTextMeasurementStyle(shapeTextStyle(input.props)),
        request.theme,
        request.availableWidth,
        request.slideNumber,
      );
      return measured;
    } catch (error) {
      const style = record(input);
      const layoutStyle =
        typeof style["style"] === "object" && style["style"] !== null
          ? (style["style"] as Record<string, unknown>)
          : {};
      const hasFixedSize =
        typeof (layoutStyle["width"] ?? input.props["width"]) === "number" &&
        typeof (layoutStyle["height"] ?? input.props["height"]) === "number";
      if (
        hasFixedSize &&
        error instanceof CapabilityError &&
        error.code === "font-measurement-unavailable"
      ) {
        return undefined;
      }
      throw error;
    }
  },
  compile(input, context) {
    const props = record(input);
    const unknown = Object.keys(props).find(
      (key) => !(allowedParameters as readonly string[]).includes(key),
    );
    if (unknown !== undefined) {
      throw new CapabilityError(
        "unknown-parameter",
        `Shape contains unsupported parameter ${unknown}`,
      );
    }
    const testId = optionalText(props, "testId");
    const part = optionalText(props, "part");
    const requestedName = optionalText(props, "name");
    const geometry = props["geometry"] as Geometry | undefined;
    const layoutStyle =
      typeof props["style"] === "object" && props["style"] !== null
        ? (props["style"] as Record<string, unknown>)
        : {};
    const insets = textInsets(layoutStyle);
    const box =
      context.layoutBox === undefined
        ? {
            x: requiredNumber(props, "x"),
            y: requiredNumber(props, "y"),
            width: requiredNumber(
              { width: layoutStyle["width"] ?? props["width"] },
              "width",
            ),
            height: requiredNumber(
              { height: layoutStyle["height"] ?? props["height"] },
              "height",
            ),
          }
        : {
            x: requiredNumber({ x: context.layoutBox.x }, "x"),
            y: requiredNumber({ y: context.layoutBox.y }, "y"),
            width: requiredNumber({ width: context.layoutBox.width }, "width"),
            height: requiredNumber(
              { height: context.layoutBox.height },
              "height",
            ),
          };
    const properties = shapePropertiesXml({
      theme: context.theme,
      style: layoutStyle,
      box,
      ...(geometry === undefined ? {} : { geometry }),
      ...(props["flipH"] === undefined
        ? {}
        : { flipH: props["flipH"] as boolean }),
      ...(props["flipV"] === undefined
        ? {}
        : { flipV: props["flipV"] as boolean }),
      ...(props["fill"] === undefined
        ? {}
        : { fill: props["fill"] as ColorValue }),
    });
    const text = props["text"];
    if (text !== undefined && typeof text !== "string") {
      throw new CapabilityError("invalid-text", "Shape.text must be a string");
    }
    const resolvedTextStyle = shapeTextStyle(props);
    const textChildren = shapeTextContent(input);
    let textOutput: ReturnType<typeof writeTextBody> | undefined;
    if (textChildren.length > 0) {
      textOutput = writeTextBody(
        textChildren,
        resolvedTextStyle,
        context.theme,
        "middle",
        {
          width: Math.max(0, box.width - insets.left - insets.right),
          height: Math.max(0, box.height - insets.top - insets.bottom),
          insets,
        },
        "text-box",
        {
          slideNumber: context.slideNumber,
          createId: context.createTextFieldId,
        },
        context.options.embedFonts,
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
          ? text === undefined
            ? {}
            : { text }
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
