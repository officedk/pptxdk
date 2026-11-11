import {
  CapabilityError,
  layoutStyleKeys,
  paddingStyleKeys,
  parseCapabilityLayoutContract,
  textStyleKeys,
  xmlEscape,
  type CapabilityInput,
  type CapabilityLayoutContract,
  type CapabilityParseContext,
  type CapabilityParseResult,
  type GeometryBox,
  type LayoutStyle,
  type LinkTarget,
  type MorphKey,
  type ObjectCapabilityModule,
  type TextStyle,
  type ThemeScope,
  type VisualStyle,
} from "@react-pptx/core/capability";
import { CommonContainer } from "@react-pptx/core/document";

import { nativeLinkXml } from "./identity.ts";
import type { MorphProps } from "./morph-api.ts";
import { motionAnnotations, type MotionProps } from "./motion-api.ts";
import {
  hasVisibleBoxStyle,
  shapePropertiesXml,
  visualStyleKeys,
  type VisualStyleRecord,
} from "./visual.ts";

/** Layout container appearance with visual styling that optionally generates a native background shape. */
export type BoxDeclarationProps = MorphProps &
  Pick<MotionProps, "enter" | "motion"> & {
    readonly testId?: string;
    readonly part?: string;
    readonly link?: LinkTarget;
    readonly style?: LayoutStyle & TextStyle & VisualStyle;
    readonly theme?: ThemeScope;
  };

/** Creates the editable layout container shared by the command factory and the Box JSX element. */
export function createBoxNode(
  props: BoxDeclarationProps,
): CommonContainer<
  "box",
  BoxDeclarationProps,
  NonNullable<BoxDeclarationProps["style"]>
> {
  return new CommonContainer<
    "box",
    BoxDeclarationProps,
    NonNullable<BoxDeclarationProps["style"]>
  >("box", "rp:box", props);
}

const boxAllowedParameters = [
  "testId",
  "part",
  "style",
  "theme",
  "morphKey",
  "link",
  "enter",
  "motion",
  "__motionContext",
] as const;

const boxAllowedStyleKeys = new Set<string>([
  ...layoutStyleKeys,
  ...paddingStyleKeys,
  ...textStyleKeys,
  ...visualStyleKeys,
]);

export type ParsedBox = {
  readonly testId?: string;
  readonly part?: string;
  readonly morphKey?: MorphKey;
  readonly link?: LinkTarget;
  readonly styleRecord: VisualStyleRecord;
  readonly hasVisibleStyle: boolean;
  readonly contract: CapabilityLayoutContract;
  readonly rawProps: Readonly<Record<string, unknown>>;
};

function styleRecord(value: unknown): Record<string, unknown> {
  if (value === undefined) return {};
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    throw new CapabilityError("invalid-style", "Box.style must be an object");
  }
  return value as Record<string, unknown>;
}

function objectXml(
  id: number,
  name: string,
  box: GeometryBox,
  properties: string,
  linkXml: string,
): string {
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${xmlEscape(name)}">${linkXml}</p:cNvPr><p:cNvSpPr/><p:nvPr/></p:nvSpPr>${properties}</p:sp>`;
}

/** Visible Box backgrounds use the same shape property writer; boxes without fill, stroke, or shadow retain layout semantics only. */
export const boxCapability: ObjectCapabilityModule<ParsedBox> = {
  declaration: "rp:box",

  parse(
    input: CapabilityInput,
    context: CapabilityParseContext,
  ): CapabilityParseResult<ParsedBox> {
    const unknown = Object.keys(input.props).find(
      (key) => !(boxAllowedParameters as readonly string[]).includes(key),
    );
    if (unknown !== undefined) {
      throw new CapabilityError(
        "unknown-parameter",
        `Box contains unknown parameter ${unknown}`,
      );
    }
    const testId = input.props["testId"];
    if (
      testId !== undefined &&
      (typeof testId !== "string" || testId.trim() === "")
    ) {
      throw new CapabilityError(
        "invalid-parameter",
        "testId must be a non-empty string",
      );
    }
    const part = input.props["part"];
    if (
      part !== undefined &&
      (typeof part !== "string" || part.trim() === "")
    ) {
      throw new CapabilityError(
        "invalid-parameter",
        "part must be a non-empty string",
      );
    }

    const style = styleRecord(input.props["style"]);
    const hasVisibleStyle = hasVisibleBoxStyle(style);
    if (input.props["morphKey"] !== undefined && !hasVisibleStyle) {
      throw new CapabilityError(
        "morph-needs-object",
        "Only visible Box elements generating native background shapes can participate in Morph",
      );
    }

    const contract = parseCapabilityLayoutContract({
      nativeKind: "sp",
      props: input.props,
      allowedStyleKeys: boxAllowedStyleKeys,
      allowPadding: true,
      isContainer: true,
      path: context.path,
    });

    const parsed: ParsedBox = {
      ...(typeof testId === "string" ? { testId } : {}),
      ...(typeof part === "string" ? { part } : {}),
      ...(input.props["morphKey"] !== undefined
        ? { morphKey: input.props["morphKey"] as MorphKey }
        : {}),
      ...(input.props["link"] !== undefined
        ? { link: input.props["link"] as LinkTarget }
        : {}),
      styleRecord: style,
      hasVisibleStyle,
      contract,
      rawProps: input.props,
    };

    return { contract, parsed };
  },

  compile(parsed: ParsedBox, context) {
    const box = context.layoutBox;
    if (box === undefined) {
      throw new CapabilityError(
        "layout-unavailable",
        "Box is missing layout results",
      );
    }
    const testId = parsed.testId;
    const part = parsed.part;
    const name =
      context.morphName ??
      (typeof testId === "string" ? testId : `Box ${context.id - 1}`);
    const style = parsed.styleRecord;
    const properties = shapePropertiesXml({
      theme: context.theme,
      style,
      box,
      layoutOnlyWhenInvisible: true,
    });
    const motionTarget = {
      kind: "shape",
      name,
      ...(typeof testId === "string" ? { testId } : {}),
      textBody: false,
      filled:
        style["backgroundColor"] !== undefined ||
        style["backgroundImage"] !== undefined,
    } as const;
    const annotations = properties.hasVisibleStyle
      ? motionAnnotations(parsed.rawProps, motionTarget)
      : motionAnnotations(
          {
            motion: false,
            __motionContext: parsed.rawProps["__motionContext"],
          },
          motionTarget,
        );
    return {
      xml: properties.hasVisibleStyle
        ? objectXml(
            context.id,
            name,
            box,
            properties.xml,
            nativeLinkXml(context.nativeLink),
          )
        : "",
      observation: {
        id: context.id,
        kind: "shape",
        name,
        ...(typeof testId === "string" ? { testId } : {}),
        ...(typeof part === "string" ? { part } : {}),
        geometry: properties.geometryName,
        box,
      },
      extensions: [],
      ...(annotations.length === 0 ? {} : { annotations }),
      ...(properties.warnings.length === 0
        ? {}
        : { warnings: properties.warnings }),
    };
  },
};
