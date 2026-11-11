import {
  CapabilityError,
  xmlEscape,
  layoutStyleKeys,
  paddingStyleKeys,
  parseCapabilityLayoutContract,
  textStyleKeys,
  type CapabilityContext,
  type CapabilityInput,
  type CapabilityLayoutContract,
  type CapabilityLayoutMeasurement,
  type CapabilityLayoutRequest,
  type ObjectCapabilityModule,
  type CapabilityParseContext,
  type CapabilityParseResult,
} from "@react-pptx/core/capability";

import { fontPresentationContributions } from "./font-embedding.ts";
import { nativeLinkXml } from "./identity.ts";
import { motionAnnotations } from "./motion-api.ts";
import { layoutStyleValue } from "./style-parse.ts";
import {
  measureTextLayoutFromParagraphs,
  parseTextContent,
  textInsets,
  textExtent,
  writeTextBodyFromParagraphs,
  type Paragraph,
  type ResolvedTextStyle,
} from "./text-service.ts";
import { textBoxXml } from "./text-xml.ts";
import { shapePropertiesXml, visualStyleKeys } from "./visual.ts";

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

export type ParsedText = {
  readonly name?: string | undefined;
  readonly morphKey?: unknown;
  readonly link?: unknown;
  readonly visualStyle: Readonly<Record<string, unknown>>;
  readonly insets: {
    readonly top: number;
    readonly right: number;
    readonly bottom: number;
    readonly left: number;
  };
  readonly paragraphs: readonly Paragraph[];
  readonly textStyle: ResolvedTextStyle;
  readonly requestedHeight:
    | { readonly contentSized: true }
    | { readonly contentSized: false; readonly value: number | undefined };
  readonly contract: CapabilityLayoutContract;
  readonly rawProps: Readonly<Record<string, unknown>>;
};

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

function optionalName(value: unknown, parameter: "name"): string | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== "string" || value.trim() === "") {
    throw new CapabilityError(
      "invalid-parameter",
      `${parameter} must be a non-empty string`,
    );
  }
  return value;
}

function coordinate(value: unknown, label: "x" | "y"): number {
  if (
    typeof value !== "number" ||
    !Number.isFinite(value) ||
    value < -1_000_000 ||
    value > 1_000_000
  ) {
    throw new CapabilityError(
      "invalid-text-box",
      `${label} must be a finite pixel number between -1000000 and 1000000`,
    );
  }
  return value;
}

function size(value: unknown, label: "width" | "height"): number {
  if (typeof value !== "number") {
    throw new CapabilityError(
      "invalid-text-box",
      `${label} must be a finite positive pixel number`,
    );
  }
  textExtent(value, label);
  return value;
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
  return { contentSized: false, value: size(value, "height") };
}

function textStyle(value: unknown): unknown {
  if (!record(value)) return value;
  return Object.fromEntries(
    Object.entries(value).filter(([key]) => textStyleKeys.has(key)),
  );
}

/** Fixed-region text box capability; specific text services are shared by text boxes and other native objects. */
export const textCapability: ObjectCapabilityModule<ParsedText> = {
  declaration: "rp:text",

  parse(
    input: CapabilityInput,
    context: CapabilityParseContext,
  ): CapabilityParseResult<ParsedText> {
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
    const requestedName = optionalName(input.props["name"], "name");

    if (input.props["x"] !== undefined) coordinate(input.props["x"], "x");
    if (input.props["y"] !== undefined) coordinate(input.props["y"], "y");
    if (
      input.props["width"] !== undefined &&
      typeof input.props["width"] === "number"
    ) {
      size(input.props["width"], "width");
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
      ? input.props["style"]
      : {};

    const parsed: ParsedText = {
      ...(requestedName !== undefined ? { name: requestedName } : {}),
      ...(input.props["morphKey"] !== undefined
        ? { morphKey: input.props["morphKey"] }
        : {}),
      ...(input.props["link"] !== undefined
        ? { link: input.props["link"] }
        : {}),
      visualStyle,
      insets,
      paragraphs,
      textStyle: parsedTextStyle,
      requestedHeight,
      contract,
      rawProps: input.props,
    };

    return { contract, parsed };
  },

  measureLayout(
    parsed: ParsedText,
    request: CapabilityLayoutRequest,
  ): CapabilityLayoutMeasurement | undefined {
    const insets = parsed.insets;
    const isContentSized = parsed.requestedHeight.contentSized;
    const styleWidth = layoutStyleValue(
      { type: "rp:text", props: parsed.rawProps, children: [] },
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

    try {
      const measured = measureTextLayoutFromParagraphs(
        parsed.paragraphs,
        parsed.textStyle,
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

  compile(parsed: ParsedText, context: CapabilityContext) {
    const insets = parsed.insets;
    const styleWidth = layoutStyleValue(
      { type: "rp:text", props: parsed.rawProps, children: [] },
      "width",
    );
    const dimensions = {
      x: coordinate(
        context.layoutBox?.x ?? parsed.contract.x ?? parsed.rawProps["x"],
        "x",
      ),
      y: coordinate(
        context.layoutBox?.y ?? parsed.contract.y ?? parsed.rawProps["y"],
        "y",
      ),
      width: size(
        context.layoutBox?.width ??
          parsed.contract.width ??
          styleWidth ??
          parsed.rawProps["width"],
        "width",
      ),
    };
    const requestedHeight = parsed.requestedHeight;
    const measuredHeight = requestedHeight.contentSized
      ? undefined
      : requestedHeight.value;
    const outputHeight = context.layoutBox?.height ?? measuredHeight;
    const innerWidth = Math.max(
      0,
      dimensions.width - insets.left - insets.right,
    );
    const innerHeight =
      outputHeight === undefined
        ? undefined
        : Math.max(0, outputHeight - insets.top - insets.bottom);
    const body = writeTextBodyFromParagraphs(
      parsed.paragraphs,
      parsed.textStyle,
      context.theme,
      "top",
      {
        width: innerWidth,
        ...(innerHeight === undefined ? {} : { height: innerHeight }),
        contentSized: requestedHeight.contentSized,
        insets,
      },
      "text-box",
      context.options.embedFonts,
      context.createTextFieldId,
    );
    const box = {
      ...dimensions,
      height: size(
        context.layoutBox?.height ??
          (requestedHeight.contentSized
            ? body.measurement === undefined
              ? undefined
              : body.measurement.height + insets.top + insets.bottom
            : requestedHeight.value),
        "height",
      ),
    };
    const visualStyle = parsed.visualStyle;
    const properties = shapePropertiesXml({
      theme: context.theme,
      style: visualStyle,
      box,
    });
    const name = context.morphName ?? parsed.name ?? `Text ${context.id - 1}`;
    const annotations = motionAnnotations(parsed.rawProps, {
      kind: "text",
      name,
      ...(typeof parsed.rawProps["id"] === "string"
        ? { authorId: parsed.rawProps["id"] }
        : {}),
      textBody: true,
      filled:
        visualStyle["backgroundColor"] !== undefined ||
        visualStyle["backgroundImage"] !== undefined,
      box,
      rotationDegrees:
        typeof visualStyle["rotate"] === "number" ? visualStyle["rotate"] : 0,
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
      );
    }
    const output = {
      xml: textBoxXml(
        context.id,
        escapedName,
        box,
        body.xml,
        properties.xml,
        nativeLinkXml(context.nativeLink),
      ),
      observation: {
        id: context.id,
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
        context.options.embedFonts,
      ),
    };
    return output;
  },
};
