import {
  CapabilityError,
  xmlEscape,
  type CapabilityLayoutMeasurement,
  type CapabilityLayoutRequest,
  type CapabilityInput,
  type CapabilityModule,
} from "@react-pptx/core";

import { fontPresentationContributions } from "./font-embedding.ts";
import { motionAnnotations } from "./motion-api.ts";
import {
  measureTextLayout,
  textInsets,
  textExtent,
  writeTextBody,
} from "./text-service.ts";
import { shapePropertiesXml } from "./visual.ts";

export { P, Span, Text } from "./text-elements.ts";
export type {
  ParagraphProps,
  ShapeTextStyle,
  SpanProps,
  TextProps,
} from "./text-elements.ts";
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
  "testId",
  "name",
  "morphKey",
  "x",
  "y",
  "width",
  "height",
  "style",
  "children",
  "enter",
  "emphasis",
  "motion",
  "__motionContext",
] as const;

function record(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function optionalName(
  value: unknown,
  parameter: "testId" | "name",
): string | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== "string" || value.trim() === "") {
    throw new CapabilityError(
      "invalid-parameter",
      `${parameter} 必须是非空字符串`,
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
      `${label} 必须是 -1000000 到 1000000 间的有限像素数值`,
    );
  }
  return value;
}

function size(value: unknown, label: "width" | "height"): number {
  if (typeof value !== "number") {
    throw new CapabilityError(
      "invalid-text-box",
      `${label} 必须是有限正像素数值`,
    );
  }
  textExtent(value, label);
  return value;
}

function height(
  value: unknown,
):
  | { readonly contentSized: true }
  | { readonly contentSized: false; readonly value: number } {
  if (value === "auto" || value === undefined) return { contentSized: true };
  return { contentSized: false, value: size(value, "height") };
}

const textStyleProperties = new Set([
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
]);

function textStyle(value: unknown): unknown {
  if (!record(value)) return value;
  return Object.fromEntries(
    Object.entries(value).filter(([key]) => textStyleProperties.has(key)),
  );
}

function layoutStyleValue(
  input: CapabilityInput,
  name: "width" | "height",
): number | "auto" | undefined {
  const value = input.props["style"];
  if (!record(value)) return undefined;
  const dimension = value[name];
  return typeof dimension === "number" || dimension === "auto"
    ? dimension
    : undefined;
}

function textBoxXml(
  id: number,
  escapedName: string,
  box: {
    readonly x: number;
    readonly y: number;
    readonly width: number;
    readonly height: number;
  },
  textBody: string,
  properties: string,
): string {
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${escapedName}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>${properties}${textBody}</p:sp>`;
}

/** 固定区域文本框能力；具体文字服务由文本框与其他原生对象共享。 */
export const textCapability: CapabilityModule = {
  declaration: "rp:text",
  measureLayout(
    input: CapabilityInput,
    request: CapabilityLayoutRequest,
  ): CapabilityLayoutMeasurement | undefined {
    const style = textStyle(input.props["style"]);
    const insets = textInsets(input.props["style"]);
    const heightStyle = layoutStyleValue(input, "height");
    const isContentSized =
      heightStyle === "auto" ||
      (heightStyle === undefined &&
        (input.props["height"] === "auto" ||
          input.props["height"] === undefined));
    const width =
      request.availableWidth ??
      layoutStyleValue(input, "width") ??
      input.props["width"];
    try {
      return measureTextLayout(
        input.children,
        style,
        request.theme,
        request.availableWidth !== undefined
          ? request.availableWidth
          : typeof width === "number"
            ? Math.max(0, width - insets.left - insets.right)
            : undefined,
      );
    } catch (error) {
      if (
        !isContentSized &&
        error instanceof CapabilityError &&
        (error.code === "font-measurement-unavailable" ||
          error.code === "font-missing-glyph")
      ) {
        return undefined;
      }
      throw error;
    }
  },
  compile(input: CapabilityInput, context) {
    const unknown = Object.keys(input.props).find(
      (key) => !(allowedParameters as readonly string[]).includes(key),
    );
    if (unknown !== undefined) {
      throw new CapabilityError(
        "unknown-parameter",
        `Text 含未实现参数 ${unknown}`,
      );
    }
    const testId = optionalName(input.props["testId"], "testId");
    const requestedName = optionalName(input.props["name"], "name");
    const insets = textInsets(input.props["style"]);
    const styleWidth = layoutStyleValue(input, "width");
    const styleHeight = layoutStyleValue(input, "height");
    const dimensions = {
      x: coordinate(context.layoutBox?.x ?? input.props["x"], "x"),
      y: coordinate(context.layoutBox?.y ?? input.props["y"], "y"),
      width: size(
        context.layoutBox?.width ?? styleWidth ?? input.props["width"],
        "width",
      ),
    };
    const requestedHeight = height(styleHeight ?? input.props["height"]);
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
    const body = writeTextBody(
      input.children,
      textStyle(input.props["style"]),
      context.theme,
      "top",
      {
        width: innerWidth,
        ...(innerHeight === undefined ? {} : { height: innerHeight }),
        contentSized: requestedHeight.contentSized,
        insets,
      },
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
    const visualStyle = record(input.props["style"])
      ? input.props["style"]
      : {};
    const properties = shapePropertiesXml({
      theme: context.theme,
      style: visualStyle,
      box,
    });
    const name =
      context.morphName ?? requestedName ?? testId ?? `Text ${context.id - 1}`;
    const annotations = motionAnnotations(input.props, {
      kind: "text",
      name,
      ...(testId === undefined ? {} : { testId }),
      textBody: true,
      filled:
        visualStyle["backgroundColor"] !== undefined ||
        visualStyle["backgroundImage"] !== undefined,
      paragraphCount: body.paragraphCount,
    });
    let escapedName: string;
    try {
      escapedName = xmlEscape(name);
    } catch {
      throw new CapabilityError(
        "invalid-xml-character",
        "Text.name 含 XML 1.0 不允许的字符",
      );
    }
    const output = {
      xml: textBoxXml(context.id, escapedName, box, body.xml, properties.xml),
      observation: {
        id: context.id,
        kind: "text",
        name,
        ...(testId === undefined ? {} : { testId }),
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
