import {
  CapabilityError,
  pxToEmu,
  xmlEscape,
  type CapabilityInput,
  type CapabilityModule,
  type TextStyle,
  type ThemeScope,
  type ObjectObservation,
} from "@react-pptx/core";
import type { ReactNode } from "react";

import {
  motionAnnotations,
  motionElement,
  type MotionProps,
} from "./motion-api.ts";

/** Group 只保留身份和原生成员，范围由成员几何求并集，不参与布局。 */
export type GroupProps = MotionProps & {
  readonly testId?: string;
  readonly name?: string;
  readonly style?: TextStyle;
  readonly theme?: ThemeScope;
  readonly children?: ReactNode;
};

const allowedParameters = [
  "testId",
  "name",
  "style",
  "theme",
  "enter",
  "emphasis",
  "motion",
  "__motionEnabled",
  "__motionStep",
] as const;
const emuPerPixel = 9525;
const maximumGroupExtentEmu = 2_147_483_647;

type EmuBox = {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
};

function optionalText(
  props: Readonly<Record<string, unknown>>,
  name: "testId" | "name",
): string | undefined {
  const value = props[name];
  if (value === undefined) return undefined;
  if (typeof value !== "string" || value.trim() === "") {
    throw new CapabilityError("invalid-parameter", `${name} 必须是非空字符串`);
  }
  return value;
}

function bounds(members: readonly ObjectObservation[]): {
  readonly pixels: { x: number; y: number; width: number; height: number };
  readonly emu: EmuBox;
} {
  if (members.length === 0) {
    throw new CapabilityError("empty-group", "Group 至少需要一个原生对象成员");
  }
  const memberBoxes = members.map((member) => {
    const box = member.box;
    if (
      box === undefined ||
      !Number.isFinite(box.x) ||
      !Number.isFinite(box.y) ||
      !Number.isFinite(box.width) ||
      !Number.isFinite(box.height) ||
      box.width <= 0 ||
      box.height <= 0
    ) {
      throw new CapabilityError(
        "member-without-geometry",
        `Group 成员 ${member.name} 缺少可组合的几何范围`,
      );
    }
    const x = pxToEmu(box.x);
    const y = pxToEmu(box.y);
    return {
      x,
      y,
      right: x + pxToEmu(box.width),
      bottom: y + pxToEmu(box.height),
    };
  });
  const x = Math.min(...memberBoxes.map((box) => box.x));
  const y = Math.min(...memberBoxes.map((box) => box.y));
  const right = Math.max(...memberBoxes.map((box) => box.right));
  const bottom = Math.max(...memberBoxes.map((box) => box.bottom));
  const width = right - x;
  const height = bottom - y;
  if (
    width < 1 ||
    height < 1 ||
    width > maximumGroupExtentEmu ||
    height > maximumGroupExtentEmu
  ) {
    throw new CapabilityError(
      "invalid-group-geometry",
      "Group 成员范围超出可写入的 PresentationML 尺寸",
    );
  }
  const emu = { x, y, width, height };
  return {
    emu,
    pixels: {
      x: x / emuPerPixel,
      y: y / emuPerPixel,
      width: width / emuPerPixel,
      height: height / emuPerPixel,
    },
  };
}

function groupXml(
  id: number,
  escapedName: string,
  box: EmuBox,
  members: string,
): string {
  // 组内坐标与页面坐标保持同一基准，避免成员因重复换算而发生缩放或位移。
  return `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${id}" name="${escapedName}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="${box.x}" y="${box.y}"/><a:ext cx="${box.width}" cy="${box.height}"/><a:chOff x="${box.x}" y="${box.y}"/><a:chExt cx="${box.width}" cy="${box.height}"/></a:xfrm></p:grpSpPr>${members}</p:grpSp>`;
}

/** 将成员保留为可编辑的原生对象；组合范围由成员几何求并集，不负责布局或编排作用域。 */
export const groupCapability: CapabilityModule = {
  declaration: "rp:group",
  compile(input: CapabilityInput, context) {
    const unknown = Object.keys(input.props).find(
      (key) => !(allowedParameters as readonly string[]).includes(key),
    );
    if (unknown !== undefined) {
      throw new CapabilityError(
        "unknown-parameter",
        `Group 含未实现参数 ${unknown}`,
      );
    }
    const testId = optionalText(input.props, "testId");
    const requestedName = optionalText(input.props, "name");
    const compiled = context.compileChildren(input.children);
    const geometry = bounds(compiled.objects);
    const name = requestedName ?? testId ?? `Group ${context.id - 1}`;
    const annotations = motionAnnotations(input.props, {
      kind: "group",
      name,
      ...(testId === undefined ? {} : { testId }),
      textBody: false,
      filled: false,
    });
    let escapedName: string;
    try {
      escapedName = xmlEscape(name);
    } catch {
      throw new CapabilityError(
        "invalid-xml-character",
        "Group.name 含 XML 1.0 不允许的字符",
      );
    }
    return {
      xml: groupXml(context.id, escapedName, geometry.emu, compiled.xml),
      observation: {
        id: context.id,
        kind: "group",
        name,
        ...(testId === undefined ? {} : { testId }),
        box: geometry.pixels,
        members: compiled.objects,
      },
      extensions: [],
      ...(annotations.length === 0 ? {} : { annotations }),
    };
  },
};

/** 在 React 根树中声明一个保留原生成员结构的可编辑组合。 */
export function Group(props: GroupProps): ReactNode {
  return motionElement("rp:group", props);
}
