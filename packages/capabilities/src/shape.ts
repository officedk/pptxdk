import {
  CapabilityError,
  pxToEmu,
  xmlEscape,
  type CapabilityLayoutMeasurement,
  type CapabilityLayoutRequest,
  type LayoutStyle,
  type CapabilityInput,
  type CapabilityModule,
} from "@react-pptx/core";
import type { ReactNode } from "react";

import { fontPresentationContributions } from "./font-embedding.ts";
import {
  motionAnnotations,
  motionElement,
  type MotionProps,
} from "./motion-api.ts";
import {
  measureTextLayout,
  writeTextBody,
  type TextStyle,
} from "./text-service.ts";

export type ShapeProps = MotionProps & {
  readonly testId?: string;
  readonly name?: string;
  readonly geometry?: string;
  readonly x?: number;
  readonly y?: number;
  readonly width?: number;
  readonly height?: number;
  readonly style?: LayoutStyle;
  readonly fill?: string;
  /** 直接写入此形状的原生正文，不创建额外文本框。 */
  readonly text?: string;
  /** 形状正文使用共享文字服务，裁剪文本框能力后仍可写出。 */
  readonly textStyle?: TextStyle;
  readonly children?: ReactNode;
};

const geometryNames = new Set([
  "rect",
  "roundRect",
  "ellipse",
  "triangle",
  "rtTriangle",
  "diamond",
  "parallelogram",
  "trapezoid",
  "pentagon",
  "hexagon",
  "heptagon",
  "octagon",
  "star4",
  "star5",
  "star6",
  "star8",
  "cloud",
  "heart",
  "pie",
  "arc",
  "leftBrace",
  "rightBrace",
  "leftBracket",
  "rightBracket",
]);
const maximumShapeExtentEmu = 2_147_483_647;

const allowedParameters = [
  "testId",
  "name",
  "geometry",
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
      `${name} 必须是有限像素数值，宽高必须大于 0，坐标范围为 -1000000 到 1000000`,
    );
  }
  if (isExtent) {
    const extentEmu = pxToEmu(value);
    if (extentEmu < 1 || extentEmu > maximumShapeExtentEmu) {
      throw new CapabilityError(
        "invalid-coordinate",
        `${name} 换算后必须在 1 到 ${maximumShapeExtentEmu} EMU 的闭区间内`,
      );
    }
  }
  return value;
}

function optionalText(
  props: Record<string, unknown>,
  name: "testId" | "name",
): string | undefined {
  const value = props[name];
  if (value === undefined) return undefined;
  if (typeof value !== "string" || value.trim() === "") {
    throw new CapabilityError("invalid-parameter", `${name} 必须是非空字符串`);
  }
  return value;
}

function shapeXml(
  id: number,
  escapedName: string,
  geometry: string,
  box: { x: number; y: number; width: number; height: number },
  fill: string | undefined,
  textBody: string,
): string {
  const fillXml =
    fill === undefined
      ? "<a:noFill/>"
      : `<a:solidFill><a:srgbClr val="${fill}"/></a:solidFill>`;
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${escapedName}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${pxToEmu(box.x)}" y="${pxToEmu(box.y)}"/><a:ext cx="${pxToEmu(box.width)}" cy="${pxToEmu(box.height)}"/></a:xfrm><a:prstGeom prst="${geometry}"><a:avLst/></a:prstGeom>${fillXml}<a:ln><a:noFill/></a:ln></p:spPr>${textBody}</p:sp>`;
}

/** 形状模块独立校验输入、声明原生几何，并生成对应 PresentationML 对象。 */
export const shapeCapability: CapabilityModule = {
  declaration: "rp:shape",
  measureLayout(
    input: CapabilityInput,
    request: CapabilityLayoutRequest,
  ): CapabilityLayoutMeasurement | undefined {
    const text = input.props["text"];
    if (typeof text !== "string" || text === "") {
      return { minContentWidth: 0, minContentHeight: 0 };
    }
    try {
      const measured = measureTextLayout(
        [{ type: "#text", text }],
        input.props["textStyle"],
        request.availableWidth,
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
        (error.code === "font-measurement-unavailable" ||
          error.code === "font-missing-glyph")
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
        `Shape 含未实现参数 ${unknown}`,
      );
    }
    if (input.children.length > 0) {
      throw new CapabilityError(
        "unsupported-children",
        "本构建切片的 Shape 不接受子内容",
      );
    }

    const testId = optionalText(props, "testId");
    const requestedName = optionalText(props, "name");
    const geometry = props["geometry"] ?? "rect";
    if (typeof geometry !== "string" || !geometryNames.has(geometry)) {
      throw new CapabilityError(
        "unsupported-geometry",
        `不支持形状几何 ${typeof geometry === "string" ? geometry : "非字符串值"}`,
      );
    }
    const layoutStyle =
      typeof props["style"] === "object" && props["style"] !== null
        ? (props["style"] as Record<string, unknown>)
        : {};
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
    const fill = props["fill"];
    if (
      fill !== undefined &&
      (typeof fill !== "string" || !/^#[0-9a-f]{6}$/iu.test(fill))
    ) {
      throw new CapabilityError("invalid-fill", "fill 必须是 #RRGGBB 颜色");
    }
    const text = props["text"];
    if (text !== undefined && typeof text !== "string") {
      throw new CapabilityError("invalid-text", "Shape.text 必须是字符串");
    }
    let textOutput: ReturnType<typeof writeTextBody> | undefined;
    if (text !== undefined && text !== "") {
      textOutput = writeTextBody(
        [{ type: "#text", text }],
        props["textStyle"],
        "middle",
        { width: box.width, height: box.height },
      );
    }
    const name = requestedName ?? testId ?? `Shape ${context.id - 1}`;
    const annotations = motionAnnotations(props, {
      kind: "shape",
      name,
      ...(testId === undefined ? {} : { testId }),
      textBody: textOutput !== undefined,
      filled: fill !== undefined,
    });
    let escapedName: string;
    try {
      escapedName = xmlEscape(name);
    } catch {
      throw new CapabilityError(
        "invalid-xml-character",
        "Shape.name 含 XML 1.0 不允许的字符",
      );
    }
    return {
      xml: shapeXml(
        context.id,
        escapedName,
        geometry,
        box,
        fill?.slice(1).toUpperCase(),
        textOutput?.xml ?? "",
      ),
      observation: {
        id: context.id,
        kind: "shape",
        name,
        ...(testId === undefined ? {} : { testId }),
        geometry,
        box,
        ...(text === undefined ? {} : { text }),
      },
      extensions: [],
      ...(annotations.length === 0 ? {} : { annotations }),
      ...(textOutput === undefined
        ? {}
        : {
            resourceClaims: textOutput.resourceClaims,
            presentationContributions: fontPresentationContributions(
              textOutput.fontUses,
              context.options.embedFonts,
            ),
            warnings: textOutput.warnings,
          }),
    };
  },
};

/** 在 React root 中声明一个由官方形状模块处理的原生对象。 */
export function Shape(props: ShapeProps): ReactNode {
  return motionElement("rp:shape", props);
}
