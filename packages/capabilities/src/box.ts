import {
  CapabilityError,
  xmlEscape,
  type CapabilityInput,
  type CapabilityModule,
  type GeometryBox,
} from "@react-pptx/core";

import { shapePropertiesXml } from "./visual.ts";

function styleRecord(value: unknown): Record<string, unknown> {
  if (value === undefined) return {};
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    throw new CapabilityError("invalid-style", "Box.style 必须是对象");
  }
  return value as Record<string, unknown>;
}

function objectXml(
  id: number,
  name: string,
  box: GeometryBox,
  properties: string,
): string {
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${xmlEscape(name)}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>${properties}</p:sp>`;
}

/** 可见 Box 背景使用同一形状属性 writer；无填充/线/阴影时仅保留布局语义。 */
export const boxCapability: CapabilityModule = {
  declaration: "rp:box",
  compile(input: CapabilityInput, context) {
    const unknown = Object.keys(input.props).find(
      (key) => !["testId", "style", "theme"].includes(key),
    );
    if (unknown !== undefined) {
      throw new CapabilityError(
        "unknown-parameter",
        `Box 含未实现参数 ${unknown}`,
      );
    }
    const box = context.layoutBox;
    if (box === undefined)
      throw new CapabilityError("layout-unavailable", "Box 缺少布局结果");
    const testId = input.props["testId"];
    if (
      testId !== undefined &&
      (typeof testId !== "string" || testId.trim() === "")
    ) {
      throw new CapabilityError("invalid-parameter", "testId 必须是非空字符串");
    }
    const name = typeof testId === "string" ? testId : `Box ${context.id - 1}`;
    const style = styleRecord(input.props["style"]);
    const properties = shapePropertiesXml({ theme: context.theme, style, box });
    return {
      xml: properties.hasVisibleStyle
        ? objectXml(context.id, name, box, properties.xml)
        : "",
      observation: {
        id: context.id,
        kind: "shape",
        name,
        ...(typeof testId === "string" ? { testId } : {}),
        geometry: properties.geometryName,
        box,
      },
      extensions: [],
      ...(properties.warnings.length === 0
        ? {}
        : { warnings: properties.warnings }),
    };
  },
};
