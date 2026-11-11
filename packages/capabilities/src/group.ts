import {
  CapabilityError,
  layoutStyleKeys,
  paddingStyleKeys,
  parseCapabilityLayoutContract,
  pxToEmu,
  textStyleKeys,
  xmlEscape,
  type CapabilityChild,
  type CapabilityContext,
  type CapabilityInput,
  type CapabilityLayoutContract,
  type CapabilityParseContext,
  type CapabilityParseResult,
  type LayoutStyle,
  type LinkTarget,
  type MorphKey,
  type ObjectCapabilityModule,
  type ObjectObservation,
  type TextStyle,
  type ThemeScope,
  type VisualStyle,
} from "@react-pptx/core/capability";
import { CommonContainer } from "@react-pptx/core/document";

import { groupXml } from "./group-xml.ts";
import { nativeLinkXml } from "./identity.ts";
import type { MorphProps } from "./morph-api.ts";
import { motionAnnotations, type MotionProps } from "./motion-api.ts";

/** Native Group preserves member relationships; style allows it to arrange as a layout child, with output bounds computed from the union of member geometries. */
export type GroupDeclarationProps = MotionProps &
  MorphProps & {
    readonly id?: string;
    readonly testId?: string;
    readonly part?: string;
    readonly name?: string;
    readonly link?: LinkTarget;
    readonly style?: LayoutStyle & TextStyle & Pick<VisualStyle, "rotate">;
    readonly theme?: ThemeScope;
  };

/** Creates the editable native group shared by the command factory and the Group JSX element. */
export function createGroupNode(
  props: GroupDeclarationProps,
): CommonContainer<
  "group",
  GroupDeclarationProps,
  NonNullable<GroupDeclarationProps["style"]>
> {
  return new CommonContainer<
    "group",
    GroupDeclarationProps,
    NonNullable<GroupDeclarationProps["style"]>
  >("group", "rp:group", props);
}

const allowedParameters = [
  "id",
  "testId",
  "part",
  "name",
  "morphKey",
  "link",
  "style",
  "theme",
  "enter",
  "emphasis",
  "motion",
] as const;
const emuPerPixel = 9525;
const maximumGroupExtentEmu = 2_147_483_647;

type EmuBox = {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
};

type MotionNativeMember = {
  readonly id: number;
  readonly kind: string;
  readonly name: string;
  readonly testId?: string;
  readonly box?: NonNullable<ObjectObservation["box"]>;
  readonly members?: readonly MotionNativeMember[];
};

function motionMembers(
  objects: readonly ObjectObservation[],
): readonly MotionNativeMember[] {
  return objects.map((object) => ({
    id: object.id,
    kind: object.kind,
    name: object.name,
    ...(object.testId === undefined ? {} : { testId: object.testId }),
    ...(object.box === undefined ? {} : { box: object.box }),
    ...(object.members === undefined
      ? {}
      : { members: motionMembers(object.members) }),
  }));
}

function optionalText(
  props: Readonly<Record<string, unknown>>,
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

function bounds(members: readonly ObjectObservation[]): {
  readonly pixels: { x: number; y: number; width: number; height: number };
  readonly emu: EmuBox;
} {
  if (members.length === 0) {
    throw new CapabilityError(
      "empty-group",
      "Group requires at least one native object member",
    );
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
        `Group member ${member.name} is missing composable geometry bounds`,
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
      "Group member bounds exceed writable PresentationML dimensions",
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

const groupAllowedStyleKeys = new Set<string>([
  ...layoutStyleKeys,
  ...paddingStyleKeys,
  ...textStyleKeys,
  "rotate",
]);

export type ParsedGroup = {
  readonly testId?: string;
  readonly part?: string;
  readonly name?: string;
  readonly morphKey?: MorphKey;
  readonly link?: LinkTarget;
  readonly styleRecord: Readonly<Record<string, unknown>>;
  readonly rotation?: number;
  readonly contract: CapabilityLayoutContract;
  readonly rawProps: Readonly<Record<string, unknown>>;
  readonly children: readonly CapabilityChild[];
};

/** Preserves members as native objects and writes group bounds according to final member geometry; outer layout is resolved by core. */
export const groupCapability: ObjectCapabilityModule<ParsedGroup> = {
  declaration: "rp:group",

  parse(
    input: CapabilityInput,
    context: CapabilityParseContext,
  ): CapabilityParseResult<ParsedGroup> {
    const unknown = Object.keys(input.props).find(
      (key) => !(allowedParameters as readonly string[]).includes(key),
    );
    if (unknown !== undefined) {
      throw new CapabilityError(
        "unknown-parameter",
        `Group contains unsupported parameter ${unknown}`,
      );
    }
    const testId = optionalText(input.props, "testId");
    const part = optionalText(input.props, "part");
    const requestedName = optionalText(input.props, "name");
    const rawStyle = input.props["style"];
    const styleRecord =
      typeof rawStyle === "object" &&
      rawStyle !== null &&
      !Array.isArray(rawStyle)
        ? (rawStyle as Record<string, unknown>)
        : {};

    const contract = parseCapabilityLayoutContract({
      nativeKind: "grpSp",
      props: input.props,
      allowedStyleKeys: groupAllowedStyleKeys,
      allowPadding: true,
      isContainer: true,
      path: context.path,
    });

    const rotation = styleRecord["rotate"];
    if (typeof rotation === "number" && !Number.isFinite(rotation)) {
      throw new CapabilityError(
        "invalid-visual-style",
        "rotate must be a finite angle",
      );
    }
    const parsed: ParsedGroup = {
      ...(testId !== undefined ? { testId } : {}),
      ...(part !== undefined ? { part } : {}),
      ...(requestedName !== undefined ? { name: requestedName } : {}),
      ...(input.props["morphKey"] !== undefined
        ? { morphKey: input.props["morphKey"] as MorphKey }
        : {}),
      ...(input.props["link"] !== undefined
        ? { link: input.props["link"] as LinkTarget }
        : {}),
      styleRecord,
      ...(typeof rotation === "number" ? { rotation } : {}),
      contract,
      rawProps: input.props,
      children: input.children,
    };
    return { contract, parsed };
  },

  compile(parsed: ParsedGroup, context: CapabilityContext) {
    const testId = parsed.testId;
    const part = parsed.part;
    const requestedName = parsed.name;
    const compiled = context.compileChildren(parsed.children);
    const geometry = bounds(compiled.objects);
    const outputGeometry = context.layoutBox ?? geometry.pixels;
    if (outputGeometry.width <= 0 || outputGeometry.height <= 0) {
      throw new CapabilityError(
        "invalid-group-geometry",
        "Group layout width and height must be greater than 0",
      );
    }
    const outputEmu = {
      x: pxToEmu(outputGeometry.x),
      y: pxToEmu(outputGeometry.y),
      width: pxToEmu(outputGeometry.width),
      height: pxToEmu(outputGeometry.height),
    };
    const name =
      context.morphName ?? requestedName ?? testId ?? `Group ${context.id - 1}`;
    const rotation = parsed.rotation;
    const members = motionMembers(compiled.objects);
    const annotations = motionAnnotations(parsed.rawProps, {
      kind: "group",
      name,
      ...(testId === undefined ? {} : { testId }),
      textBody: false,
      filled: false,
      box: outputGeometry,
      rotationDegrees: typeof rotation === "number" ? rotation : 0,
      members,
    });
    let escapedName: string;
    try {
      escapedName = xmlEscape(name);
    } catch {
      throw new CapabilityError(
        "invalid-xml-character",
        "Group.name contains characters not allowed in XML 1.0",
      );
    }
    return {
      xml: groupXml(
        context.id,
        escapedName,
        outputEmu,
        compiled.xml,
        rotation === undefined ? undefined : rotation,
        nativeLinkXml(context.nativeLink),
      ),
      observation: {
        id: context.id,
        kind: "group",
        name,
        ...(testId === undefined ? {} : { testId }),
        ...(part === undefined ? {} : { part }),
        box: outputGeometry,
        members: compiled.objects,
      },
      extensions: [],
      ...(annotations.length === 0 ? {} : { annotations }),
    };
  },
};
