import {
  CapabilityError,
  pxToEmu,
  xmlEscape,
  type CapabilityInput,
  type CapabilityModule,
} from "@react-pptx/core";

import { textExtent, writeTextBody } from "./text-service.ts";

export { P, Span, Text } from "./text-elements.ts";
export type {
  ParagraphProps,
  ShapeTextStyle,
  SpanProps,
  TextProps,
} from "./text-elements.ts";
export { font } from "./font.ts";
export type { Font, FontFace, FontFaceStyle } from "./font.ts";
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
  "x",
  "y",
  "width",
  "height",
  "style",
  "children",
] as const;

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
): string {
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${escapedName}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${pxToEmu(box.x)}" y="${pxToEmu(box.y)}"/><a:ext cx="${pxToEmu(box.width)}" cy="${pxToEmu(box.height)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr>${textBody}</p:sp>`;
}

/** 固定区域文本框能力；具体文字服务由文本框与其他原生对象共享。 */
export const textCapability: CapabilityModule = {
  declaration: "rp:text",
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
    const box = {
      x: coordinate(input.props["x"], "x"),
      y: coordinate(input.props["y"], "y"),
      width: size(input.props["width"], "width"),
      height: size(input.props["height"], "height"),
    };
    const body = writeTextBody(input.children, input.props["style"], "top");
    const name = requestedName ?? testId ?? `Text ${context.id - 1}`;
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
      xml: textBoxXml(context.id, escapedName, box, body.xml),
      observation: {
        id: context.id,
        kind: "text",
        name,
        ...(testId === undefined ? {} : { testId }),
        box,
        text: body.text,
      },
      extensions: [],
      resourceClaims: body.resourceClaims,
      warnings: body.warnings,
    };
    return output;
  },
};
