import {
  CommonContainer,
  CommonNode,
  RichTextContainer,
  type RichTextItem,
} from "@pptxdk/core/document";
import {
  A_NS,
  ooxmlNonNegativeIntSchema,
  ooxmlPositiveIntSchema,
  P_NS,
  type ImportDiagnostic,
  type ImportLocation,
  type SlideImportContext,
  type SlideImportOutcome,
  type XmlElementCursor,
} from "@pptxdk/import";

import { createParagraphNode } from "../text/text-declaration.ts";
import {
  after,
  appear,
  breathe,
  collapse,
  createMotionScopeHandle,
  createMotionScopeNode,
  createMotionStepIdentity,
  createMotionStepNode,
  disappear,
  fade,
  fadeOut,
  fillColor,
  floatDown,
  floatOut,
  floatUp,
  fly,
  flyOut,
  maxRepeatCount,
  maximumTimeNodeValue,
  onStep,
  relativeLuminance,
  repeat,
  spin,
  spinIn,
  spinOut,
  stretch,
  wipe,
  wipeOut,
  zoom,
  zoomOut,
  type Easing,
  type Effect,
  type EmphasisEffectDeclaration,
  type EntranceEffectDeclaration,
  type ExitEffectDeclaration,
  type MotionDirection,
  type MotionGranularity,
  type MotionScopeHandle,
  type MotionTargetRange,
  type MotionTiming,
  type OwnerMotion,
  type OwnerMotionEffect,
  type OwnerMotionStep,
  type OwnerMotionTarget,
  type RepeatCount,
} from "./motion-api.ts";

const supportedBehaviorNames = new Set([
  "set",
  "anim",
  "animEffect",
  "animRot",
  "animClr",
]);

const entrancePresetById: Readonly<Record<number, Effect["preset"]>> = {
  1: "appear",
  2: "fly",
  10: "fade",
  17: "stretch",
  22: "wipe",
  42: "floatUp",
  47: "floatDown",
  49: "spinIn",
  53: "zoom",
};

const emphasisPresetById: Readonly<Record<number, Effect["preset"]>> = {
  8: "spin",
  9: "fillColor",
};

const exitPresetById: Readonly<Record<number, Effect["preset"]>> = {
  1: "disappear",
  2: "flyOut",
  10: "fadeOut",
  17: "collapse",
  22: "wipeOut",
  42: "floatOut",
  47: "floatOut",
  49: "spinOut",
  53: "zoomOut",
};

function resolvePreset(
  presetClass: string | undefined,
  presetId: number | undefined,
  hasColorAnim: boolean,
  autoRev: boolean,
): Effect["preset"] | undefined {
  if (presetId === undefined) return undefined;
  if (presetClass === "emph") {
    return (
      emphasisPresetById[presetId] ??
      (presetId === 0 && hasColorAnim && autoRev
        ? ("breathe" as const)
        : undefined)
    );
  }
  if (presetClass === "exit") {
    return exitPresetById[presetId];
  }
  if (presetClass === "entr") {
    return entrancePresetById[presetId];
  }
  return undefined;
}

const directionBySubtype: Readonly<Record<number, MotionDirection>> = {
  1: "top",
  2: "right",
  4: "bottom",
  8: "left",
};

type ParsedTargetRef = {
  readonly spid: number;
  readonly isBackground: boolean;
  readonly range?: MotionTargetRange;
};

type ParsedBehavior = {
  readonly localName: string;
  readonly xpath: string;
  readonly targets: readonly ParsedTargetRef[];
  readonly durationMs?: number;
  readonly attrNames: readonly string[];
  readonly additive?: string;
  readonly fromColor?: string;
  readonly toColor?: string;
  readonly rotationFrom?: number;
  readonly rotationTo?: number;
  readonly rotationBy?: number;
};

type ParsedTimingCondition =
  | { readonly kind: "delay"; readonly delayMs: number }
  | {
      readonly kind: "end";
      readonly sourceTimingId: number;
      readonly offsetMs: number;
    }
  | { readonly kind: "onBegin"; readonly delayMs: number };

type ParsedBuildParagraph = {
  readonly build?: string;
  readonly animBg?: boolean;
};

type ParsedBuildGraphic = {
  readonly mode: "all" | "series" | "category" | "seriesEl" | "categoryEl";
};

type ParsedRawTimingNode =
  | {
      readonly kind: "effect";
      readonly timingId: number;
      readonly xpath: string;
      readonly parentTimingId?: number;
      readonly stepIndex: number;
      readonly presetId?: number;
      readonly presetClass?: string;
      readonly presetSubtype?: number;
      readonly buildGroupId?: number;
      readonly repeatCount?: RepeatCount;
      readonly autoRev: boolean;
      readonly accel?: number;
      readonly decel?: number;
      readonly delayMs: number;
      readonly dependsOn: readonly {
        readonly timingId: number;
        readonly offsetMs: number;
      }[];
      readonly iterateBy?: "word" | "letter";
      readonly iterateIntervalMs?: number;
      readonly hasSetBehavior: boolean;
      readonly behaviors: readonly ParsedBehavior[];
      readonly targetIds: readonly number[];
      readonly targetRange?: MotionTargetRange;
      readonly unsupportedBehavior?: {
        readonly namespaceUri: string;
        readonly localName: string;
        readonly xpath: string;
        readonly targetId?: number;
      };
      readonly cursor: XmlElementCursor;
    }
  | {
      readonly kind: "container";
      readonly timingId: number;
      readonly containerKind: "parallel" | "sequence";
      readonly xpath: string;
      readonly parentTimingId?: number;
      readonly stepIndex: number;
      readonly endSyncAll: boolean;
      readonly delayMs: number;
      readonly dependsOn: readonly {
        readonly timingId: number;
        readonly offsetMs: number;
      }[];
      readonly childIds: number[];
      readonly cursor: XmlElementCursor;
    };

type ImportedMotionEffect = {
  readonly effect: Effect;
  readonly declaration: Effect | ReturnType<typeof repeat>;
  readonly targetId: number;
  readonly targetRange?: MotionTargetRange;
  readonly buildGroupId?: number;
  readonly stepIndex: number;
  readonly parentTimingId?: number;
  readonly timingId: number;
  readonly effectiveDelayMs: number;
  readonly startAfterTimingId?: number;
  readonly startAfterOffsetMs?: number;
  readonly easing?: Easing;
  readonly repeatCount?: RepeatCount;
  readonly from?: MotionDirection | number;
  readonly to?: number;
  readonly by?: MotionGranularity | number;
  readonly intervalMs?: number;
  readonly fromColor?: string;
  readonly toColor?: string;
};

function locationFor(
  context: SlideImportContext,
  xpath: string,
  objectId?: number,
  objectName?: string,
): ImportLocation {
  return {
    part: context.part,
    xpath,
    ...(context.slideIndex === undefined
      ? {}
      : { slideIndex: context.slideIndex }),
    ...(objectId === undefined ? {} : { objectId }),
    ...(objectName === undefined ? {} : { objectName }),
  };
}

function parsePositiveInt(raw: string | undefined): number | undefined {
  if (raw === undefined) return undefined;
  const parsed = ooxmlPositiveIntSchema.safeParse(raw);
  return parsed.success ? parsed.data : undefined;
}

function parseNonNegativeInt(raw: string | undefined): number | undefined {
  if (raw === undefined) return undefined;
  const parsed = ooxmlNonNegativeIntSchema.safeParse(raw);
  return parsed.success ? parsed.data : undefined;
}

function parseSignedInt(raw: string | undefined): number | undefined {
  if (raw === undefined || !/^-?\d+$/u.test(raw)) return undefined;
  const n = Number(raw);
  return Number.isSafeInteger(n) ? n : undefined;
}

function consumeTargetElement(tgtEl: XmlElementCursor): ParsedTargetRef[] {
  const refs: ParsedTargetRef[] = [];
  tgtEl.children(P_NS, "sldTgt");
  for (const spTgt of tgtEl.children(P_NS, "spTgt")) {
    const spid = parsePositiveInt(spTgt.attr("spid"));
    let isBackground = false;
    for (const _bg of spTgt.children(P_NS, "bg")) {
      isBackground = true;
    }
    let range: MotionTargetRange | undefined;
    for (const txEl of spTgt.children(P_NS, "txEl")) {
      const pRg = txEl.child(P_NS, "pRg");
      if (pRg !== undefined) {
        const st = parseNonNegativeInt(pRg.attr("st"));
        const end = parseNonNegativeInt(pRg.attr("end"));
        if (st !== undefined && end !== undefined && st <= end) {
          range = { type: "paragraph", start: st, end };
        }
      }
      const charRg = txEl.child(P_NS, "charRg");
      if (charRg !== undefined) {
        charRg.consumeAttr("st");
        charRg.consumeAttr("end");
      }
    }
    for (const graphicEl of spTgt.children(P_NS, "graphicEl")) {
      const chart = graphicEl.child(A_NS, "chart");
      if (chart !== undefined) {
        const seriesIdx = parseSignedInt(chart.attr("seriesIdx"));
        const categoryIdx = parseSignedInt(chart.attr("categoryIdx"));
        const bldStep = chart.attr("bldStep");
        if (
          seriesIdx !== undefined &&
          categoryIdx !== undefined &&
          (bldStep === "gridLegend" ||
            bldStep === "series" ||
            bldStep === "category" ||
            bldStep === "ptInSeries" ||
            bldStep === "ptInCategory")
        ) {
          range = {
            type: "chart",
            step: bldStep,
            seriesIndex: seriesIdx,
            categoryIndex: categoryIdx,
          };
        }
      }
    }
    if (spid !== undefined) {
      refs.push({
        spid,
        isBackground,
        ...(range === undefined ? {} : { range }),
      });
    }
  }
  return refs;
}

function consumeConditionList(
  condList: XmlElementCursor | undefined,
): ParsedTimingCondition[] {
  if (condList === undefined) return [];
  const result: ParsedTimingCondition[] = [];
  for (const cond of condList.children(P_NS, "cond")) {
    const evt = cond.attr("evt");
    const rawDelay = cond.attr("delay");
    const delayMs =
      rawDelay === undefined || rawDelay === "indefinite"
        ? 0
        : (parseNonNegativeInt(rawDelay) ?? 0);
    const tn = cond.child(P_NS, "tn");
    const sourceId = parsePositiveInt(tn?.attr("val"));
    const tgtEl = cond.child(P_NS, "tgtEl");
    if (tgtEl !== undefined) {
      consumeTargetElement(tgtEl);
    }
    if (evt === "onBegin") {
      result.push({ kind: "onBegin", delayMs });
    } else if (evt === "end" && sourceId !== undefined) {
      result.push({ kind: "end", sourceTimingId: sourceId, offsetMs: delayMs });
    } else {
      result.push({ kind: "delay", delayMs });
    }
  }
  return result;
}

function consumeCommonTimeNodeAttrs(cTn: XmlElementCursor): {
  readonly id?: number;
  readonly dur?: string;
  readonly nodeType?: string;
  readonly presetId?: number;
  readonly presetClass?: string;
  readonly presetSubtype?: number;
  readonly grpId?: number;
  readonly repeatCount?: RepeatCount;
  readonly autoRev: boolean;
  readonly accel?: number;
  readonly decel?: number;
} {
  const id = parsePositiveInt(cTn.attr("id"));
  const dur = cTn.attr("dur");
  cTn.consumeAttr("fill");
  cTn.consumeAttr("restart");
  const nodeType = cTn.attr("nodeType");
  const presetId = parseNonNegativeInt(cTn.attr("presetID"));
  const presetClass = cTn.attr("presetClass");
  const presetSubtype = parseNonNegativeInt(cTn.attr("presetSubtype"));
  const grpId = parseNonNegativeInt(cTn.attr("grpId"));
  const rawRepeat = cTn.attr("repeatCount");
  let repeatCount: RepeatCount | undefined;
  if (rawRepeat === "indefinite") {
    repeatCount = "forever";
  } else if (rawRepeat !== undefined) {
    const milliCount = parsePositiveInt(rawRepeat);
    if (
      milliCount !== undefined &&
      milliCount >= 1000 &&
      milliCount % 1000 === 0
    ) {
      repeatCount = milliCount / 1000;
    }
  }
  const rawAutoRev = cTn.attr("autoRev");
  const autoRev = rawAutoRev === "1" || rawAutoRev === "true";
  const accel = parseNonNegativeInt(cTn.attr("accel"));
  const decel = parseNonNegativeInt(cTn.attr("decel"));
  return {
    ...(id === undefined ? {} : { id }),
    ...(dur === undefined ? {} : { dur }),
    ...(nodeType === undefined ? {} : { nodeType }),
    ...(presetId === undefined ? {} : { presetId }),
    ...(presetClass === undefined ? {} : { presetClass }),
    ...(presetSubtype === undefined ? {} : { presetSubtype }),
    ...(grpId === undefined ? {} : { grpId }),
    ...(repeatCount === undefined ? {} : { repeatCount }),
    autoRev,
    ...(accel === undefined ? {} : { accel }),
    ...(decel === undefined ? {} : { decel }),
  };
}

function consumeBehaviorValueWrapper(
  wrapper: XmlElementCursor | undefined,
): void {
  if (wrapper === undefined) return;
  for (const fltVal of wrapper.children(P_NS, "fltVal")) {
    fltVal.consumeAttr("val");
  }
  for (const strVal of wrapper.children(P_NS, "strVal")) {
    strVal.consumeAttr("val");
  }
  for (const srgb of wrapper.children(A_NS, "srgbClr")) {
    srgb.consumeAttr("val");
  }
}

function readColorFromWrapper(
  wrapper: XmlElementCursor | undefined,
): string | undefined {
  if (wrapper === undefined) return undefined;
  const srgb = wrapper.child(A_NS, "srgbClr");
  const hex = srgb?.attr("val");
  if (hex === undefined || !/^[0-9a-f]{6}$/iu.test(hex)) return undefined;
  return `#${hex.toUpperCase()}`;
}

function readBehaviorElement(behavior: XmlElementCursor): ParsedBehavior {
  const localName = behavior.localName;
  if (localName === "animEffect") {
    behavior.consumeAttr("transition");
    behavior.consumeAttr("filter");
  } else if (localName === "anim") {
    behavior.consumeAttr("calcmode");
    behavior.consumeAttr("valueType");
    behavior.consumeAttr("from");
    behavior.consumeAttr("to");
    behavior.consumeAttr("by");
  } else if (localName === "animClr") {
    behavior.consumeAttr("clrSpc");
  } else if (localName === "animRot") {
    // Attributes read below via attr("from") / attr("by") / attr("to").
  }

  const rawRotFrom =
    localName === "animRot" ? behavior.attr("from") : undefined;
  const rawRotTo = localName === "animRot" ? behavior.attr("to") : undefined;
  const rawRotBy = localName === "animRot" ? behavior.attr("by") : undefined;
  const rotationFrom = parseSignedInt(rawRotFrom);
  const rotationTo = parseSignedInt(rawRotTo);
  const rotationBy = parseSignedInt(rawRotBy);

  const cBhvr = behavior.child(P_NS, "cBhvr");
  const additive = cBhvr?.attr("additive");
  cBhvr?.consumeAttr("accumulate");
  cBhvr?.consumeAttr("override");
  const innerCTn = cBhvr?.child(P_NS, "cTn");
  const innerAttrs =
    innerCTn === undefined ? undefined : consumeCommonTimeNodeAttrs(innerCTn);
  if (innerCTn !== undefined) {
    consumeConditionList(innerCTn.child(P_NS, "stCondLst"));
  }
  const durationMs = parsePositiveInt(innerAttrs?.dur);

  const tgtEl = cBhvr?.child(P_NS, "tgtEl");
  const targets = tgtEl === undefined ? [] : consumeTargetElement(tgtEl);

  const attrNameLst = cBhvr?.child(P_NS, "attrNameLst");
  const attrNames: string[] = [];
  if (attrNameLst !== undefined) {
    for (const attrName of attrNameLst.children(P_NS, "attrName")) {
      attrNames.push(attrName.textContent.trim());
    }
  }

  const tavLst = behavior.child(P_NS, "tavLst");
  if (tavLst !== undefined) {
    for (const tav of tavLst.children(P_NS, "tav")) {
      tav.consumeAttr("tm");
      tav.consumeAttr("fmla");
      consumeBehaviorValueWrapper(tav.child(P_NS, "val"));
    }
  }

  const fromEl = behavior.child(P_NS, "from");
  const toEl = behavior.child(P_NS, "to");
  let fromColor: string | undefined;
  let toColor: string | undefined;
  if (localName === "animClr") {
    fromColor = readColorFromWrapper(fromEl);
    toColor = readColorFromWrapper(toEl);
  } else {
    consumeBehaviorValueWrapper(fromEl);
    consumeBehaviorValueWrapper(toEl);
  }

  return {
    localName,
    xpath: behavior.xpath,
    targets,
    ...(durationMs === undefined ? {} : { durationMs }),
    attrNames,
    ...(additive === undefined ? {} : { additive }),
    ...(fromColor === undefined ? {} : { fromColor }),
    ...(toColor === undefined ? {} : { toColor }),
    ...(rotationFrom === undefined ? {} : { rotationFrom }),
    ...(rotationTo === undefined ? {} : { rotationTo }),
    ...(rotationBy === undefined ? {} : { rotationBy }),
  };
}

function buildEffectDeclaration(args: {
  readonly preset: Effect["preset"];
  readonly kind: Effect["kind"];
  readonly durationMs: number;
  readonly delayMs: number;
  readonly from?: MotionDirection | number;
  readonly to?: number;
  readonly fromColor?: string;
  readonly toColor?: string;
  readonly easing?: Easing;
  readonly by?: MotionGranularity | number;
  readonly intervalMs?: number;
  readonly repeatCount?: RepeatCount;
}): {
  readonly effect: Effect;
  readonly declaration: Effect | ReturnType<typeof repeat>;
} {
  const baseOptions = {
    duration: args.durationMs,
    delay: args.delayMs,
    ...(args.easing === undefined ? {} : { easing: args.easing }),
  };
  const granularityOptions = {
    ...baseOptions,
    ...(typeof args.by === "string" ? { by: args.by } : {}),
    ...(args.intervalMs === undefined ? {} : { intervalMs: args.intervalMs }),
  };
  const instantOptions = {
    delay: args.delayMs,
    ...(typeof args.by === "string" ? { by: args.by } : {}),
    ...(args.intervalMs === undefined ? {} : { intervalMs: args.intervalMs }),
  };
  let effect: Effect;
  switch (args.preset) {
    case "appear":
      effect = appear(instantOptions);
      break;
    case "fade":
      effect = fade(granularityOptions);
      break;
    case "wipe":
      effect = wipe({
        ...granularityOptions,
        ...(typeof args.from === "string" ? { from: args.from } : {}),
      });
      break;
    case "zoom":
      effect = zoom(granularityOptions);
      break;
    case "fly":
      effect = fly({
        ...granularityOptions,
        ...(typeof args.from === "string" ? { from: args.from } : {}),
      });
      break;
    case "floatUp":
      effect = floatUp(granularityOptions);
      break;
    case "floatDown":
      effect = floatDown(granularityOptions);
      break;
    case "spinIn":
      effect = spinIn(granularityOptions);
      break;
    case "stretch":
      effect = stretch(granularityOptions);
      break;
    case "spin":
      effect =
        typeof args.by === "number"
          ? spin({
              ...baseOptions,
              by: args.by,
            })
          : spin({
              ...granularityOptions,
              ...(typeof args.from === "number" ? { from: args.from } : {}),
              ...(typeof args.to === "number" ? { to: args.to } : {}),
            });
      break;
    case "breathe":
      effect = breathe({
        period: args.durationMs,
        from: args.fromColor ?? "#000000",
        to: args.toColor ?? "#FFFFFF",
        delay: args.delayMs,
        ...(args.easing === undefined ? {} : { easing: args.easing }),
      });
      break;
    case "fillColor":
      effect = fillColor({
        ...baseOptions,
        from: args.fromColor ?? "#000000",
        to: args.toColor ?? "#FFFFFF",
      });
      break;
    case "disappear":
      effect = disappear(instantOptions);
      break;
    case "fadeOut":
      effect = fadeOut(granularityOptions);
      break;
    case "wipeOut":
      effect = wipeOut({
        ...granularityOptions,
        ...(typeof args.from === "string" ? { from: args.from } : {}),
      });
      break;
    case "flyOut":
      effect = flyOut({
        ...granularityOptions,
        ...(typeof args.from === "string" ? { from: args.from } : {}),
      });
      break;
    case "zoomOut":
      effect = zoomOut(granularityOptions);
      break;
    case "floatOut":
      effect = floatOut(granularityOptions);
      break;
    case "spinOut":
      effect = spinOut(granularityOptions);
      break;
    case "collapse":
      effect = collapse(granularityOptions);
      break;
  }
  const declaration =
    args.repeatCount === undefined ? effect : repeat(effect, args.repeatCount);
  return { effect, declaration };
}

function commonEffectPropsMatch(
  left: ImportedMotionEffect,
  right: ImportedMotionEffect,
): boolean {
  return (
    left.targetId === right.targetId &&
    left.effect.preset === right.effect.preset &&
    left.effect.kind === right.effect.kind &&
    left.effect.duration === right.effect.duration &&
    left.from === right.from &&
    left.to === right.to &&
    left.repeatCount === right.repeatCount &&
    JSON.stringify(left.easing ?? null) === JSON.stringify(right.easing ?? null)
  );
}

function targetParagraphCount(node: CommonNode): number {
  if (node.kind !== "text" && node.kind !== "shape") return 0;
  const explicitCount = node.children.filter(
    (c) => c instanceof CommonNode && c.type === "paragraph",
  ).length;
  if (explicitCount > 0) return explicitCount;
  return node.children.length > 0 || node.kind === "text" ? 1 : 0;
}

function shapeFillState(node: CommonNode): {
  readonly filled: boolean;
  readonly solidFilled: boolean;
} {
  if (node.kind !== "shape") {
    return { filled: false, solidFilled: false };
  }
  const style = node.props["style"];
  if (typeof style !== "object" || style === null) {
    return { filled: false, solidFilled: false };
  }
  const background = (style as Readonly<Record<string, unknown>>)["background"];
  if (background === undefined || background === "none") {
    return { filled: false, solidFilled: false };
  }
  const solidFilled =
    typeof background === "string" ||
    (typeof background === "object" &&
      background !== null &&
      "color" in background &&
      typeof (background as Readonly<Record<string, unknown>>)["color"] ===
        "string");
  return { filled: true, solidFilled };
}

function isValidRepeatModifier(args: {
  readonly preset: Effect["preset"];
  readonly repeatCount: RepeatCount;
  readonly by?: MotionGranularity | number;
  readonly durationMs: number;
  readonly delayMs: number;
}): boolean {
  const { preset, repeatCount, by, durationMs, delayMs } = args;
  if (preset !== "spin" && preset !== "breathe") return false;
  if (repeatCount === "forever") return preset === "breathe";
  if (by !== undefined && typeof by !== "number") return false;
  if (
    !Number.isSafeInteger(repeatCount) ||
    repeatCount < 1 ||
    repeatCount > maxRepeatCount ||
    repeatCount * 1000 > maximumTimeNodeValue
  ) {
    return false;
  }
  const unitDur =
    preset === "breathe" ? durationMs / 2 : Math.round(durationMs);
  return (
    Math.round(delayMs) <= maximumTimeNodeValue &&
    unitDur <= maximumTimeNodeValue &&
    Number.isSafeInteger(
      Math.round(delayMs) + Math.round(durationMs) * repeatCount,
    )
  );
}

/**
 * Reads slide, layout, or master `<p:timing>` back into live `enter`/`emphasis`,
 * `Step`/`MotionScope`, or `OwnerMotion` declarations, emitting localized diagnostics
 * with cascaded dependency tracking when targets or animation behaviors cannot be imported.
 */
export function readSlideMotionImport(
  context: SlideImportContext,
): SlideImportOutcome {
  const timingCursors = context.slideCursor.children(P_NS, "timing");
  if (timingCursors.length === 0) return {};

  const diagnostics: ImportDiagnostic[] = [];
  for (const extraTiming of timingCursors.slice(1)) {
    diagnostics.push({
      code: "invalid-xml-value",
      category: "corrupt-input",
      disposition: "recovered",
      location: locationFor(context, extraTiming.xpath),
      attributeOrElement: "p:timing",
    });
    extraTiming.suppressUnconsumed();
  }

  const timingCursor = timingCursors[0]!;

  // 1. Parse optional <p:bldLst>
  const paragraphBuilds = new Map<string, ParsedBuildParagraph>();
  const chartBuilds = new Map<string, ParsedBuildGraphic>();
  const bldLst = timingCursor.child(P_NS, "bldLst");
  if (bldLst !== undefined) {
    for (const bldP of bldLst.children(P_NS, "bldP")) {
      const spid = parsePositiveInt(bldP.attr("spid"));
      const grpId = parseNonNegativeInt(bldP.attr("grpId")) ?? 0;
      const build = bldP.attr("build");
      const rawAnimBg = bldP.attr("animBg");
      bldP.consumeAttr("rev");
      bldP.consumeAttr("advAuto");
      bldP.consumeAttr("uiExpand");
      bldP.consumeAttr("bld");
      const animBg =
        rawAnimBg === "1" || rawAnimBg === "true"
          ? true
          : rawAnimBg === "0" || rawAnimBg === "false"
            ? false
            : undefined;
      if (spid !== undefined) {
        paragraphBuilds.set(`${spid}:${grpId}`, {
          ...(build === undefined ? {} : { build }),
          ...(animBg === undefined ? {} : { animBg }),
        });
      }
    }
    for (const bldGraphic of bldLst.children(P_NS, "bldGraphic")) {
      const spid = parsePositiveInt(bldGraphic.attr("spid"));
      const grpId = parseNonNegativeInt(bldGraphic.attr("grpId")) ?? 0;
      bldGraphic.consumeAttr("uiExpand");
      let mode: ParsedBuildGraphic["mode"] | undefined;
      const bldAsOne = bldGraphic.child(P_NS, "bldAsOne");
      if (bldAsOne !== undefined) {
        mode = "all";
      }
      const bldSub = bldGraphic.child(P_NS, "bldSub");
      if (bldSub !== undefined) {
        const bldChart = bldSub.child(A_NS, "bldChart");
        const rawBld = bldChart?.attr("bld");
        bldChart?.consumeAttr("animBg");
        if (
          rawBld === "all" ||
          rawBld === "series" ||
          rawBld === "category" ||
          rawBld === "seriesEl" ||
          rawBld === "categoryEl"
        ) {
          mode = rawBld;
        }
      }
      if (spid !== undefined && mode !== undefined) {
        chartBuilds.set(`${spid}:${grpId}`, { mode });
      }
    }
  }

  // 2. Traverse <p:tnLst> -> tmRoot -> mainSeq -> steps
  const tnLst = timingCursor.child(P_NS, "tnLst");
  const rootPar = tnLst?.child(P_NS, "par");
  const rootCTn = rootPar?.child(P_NS, "cTn");
  if (rootCTn === undefined) {
    diagnostics.push({
      code: "corrupt-part",
      category: "corrupt-input",
      disposition: "omitted",
      location: locationFor(context, timingCursor.xpath),
    });
    timingCursor.suppressUnconsumed();
    return { diagnostics };
  }
  consumeCommonTimeNodeAttrs(rootCTn);
  consumeConditionList(rootCTn.child(P_NS, "stCondLst"));
  const rootChildTnLst = rootCTn.child(P_NS, "childTnLst");

  let mainSeq: XmlElementCursor | undefined;
  if (rootChildTnLst !== undefined) {
    for (const childSeq of rootChildTnLst.children(P_NS, "seq")) {
      childSeq.consumeAttr("concurrent");
      childSeq.consumeAttr("nextAc");
      const seqCTn = childSeq.child(P_NS, "cTn");
      const attrs =
        seqCTn === undefined ? undefined : consumeCommonTimeNodeAttrs(seqCTn);
      const prevCondLst = childSeq.child(P_NS, "prevCondLst");
      consumeConditionList(prevCondLst);
      const nextCondLst = childSeq.child(P_NS, "nextCondLst");
      consumeConditionList(nextCondLst);

      if (attrs?.nodeType === "mainSeq" && mainSeq === undefined) {
        mainSeq = childSeq;
      } else {
        diagnostics.push({
          code: "unsupported-mechanism",
          category: "unsupported",
          disposition: "omitted",
          location: locationFor(context, childSeq.xpath),
          namespaceUri: childSeq.namespaceUri,
          localName: childSeq.localName,
          mechanismHint: `timing-sequence:${attrs?.nodeType ?? "unknown"}`,
        });
        childSeq.suppressUnconsumed();
      }
    }
  }

  const mainSeqCTn = mainSeq?.child(P_NS, "cTn");
  const mainSeqChildren = mainSeqCTn?.child(P_NS, "childTnLst");
  if (mainSeqChildren === undefined) {
    return diagnostics.length === 0 ? {} : { diagnostics };
  }

  const rawNodesById = new Map<number, ParsedRawTimingNode>();
  const stepInfos: {
    readonly stepIndex: number;
    readonly trigger: "auto" | "click";
    readonly rootChildIds: number[];
    readonly xpath: string;
  }[] = [];

  const visitTimingNode = (
    cursor: XmlElementCursor,
    parentTimingId: number | undefined,
    stepIndex: number,
  ): number | undefined => {
    const cTn = cursor.child(P_NS, "cTn");
    if (cTn === undefined) return undefined;
    const attrs = consumeCommonTimeNodeAttrs(cTn);
    if (attrs.id === undefined) {
      diagnostics.push({
        code: "invalid-xml-value",
        category: "corrupt-input",
        disposition: "omitted",
        location: locationFor(context, cTn.xpath),
        attributeOrElement: "@id",
      });
      cursor.suppressUnconsumed();
      return undefined;
    }
    const timingId = attrs.id;
    const conds = consumeConditionList(cTn.child(P_NS, "stCondLst"));
    let delayMs = 0;
    const dependsOn: {
      readonly timingId: number;
      readonly offsetMs: number;
    }[] = [];
    for (const cond of conds) {
      if (cond.kind === "delay") {
        delayMs += cond.delayMs;
      } else if (cond.kind === "end") {
        dependsOn.push({
          timingId: cond.sourceTimingId,
          offsetMs: cond.offsetMs,
        });
      }
    }

    const endSyncEl = cTn.child(P_NS, "endSync");
    let endSyncAll = false;
    if (endSyncEl !== undefined) {
      const evt = endSyncEl.attr("evt");
      endSyncEl.consumeAttr("delay");
      const rtn = endSyncEl.child(P_NS, "rtn");
      const val = rtn?.attr("val");
      endSyncAll = evt === "end" && (val === "all" || val === undefined);
    }

    const iterateEl = cTn.child(P_NS, "iterate");
    let iterateBy: "word" | "letter" | undefined;
    let iterateIntervalMs: number | undefined;
    if (iterateEl !== undefined) {
      const rawType = iterateEl.attr("type");
      if (rawType === "wd") iterateBy = "word";
      else if (rawType === "lt") iterateBy = "letter";
      const tmAbs = iterateEl.child(P_NS, "tmAbs");
      iterateIntervalMs = parsePositiveInt(tmAbs?.attr("val"));
      iterateEl.consumeAttr("backwards");
    }

    const childTnLst = cTn.child(P_NS, "childTnLst");
    const subContainers: XmlElementCursor[] = [];
    const behaviorCursors: XmlElementCursor[] = [];
    const otherChildren: XmlElementCursor[] = [];

    if (childTnLst !== undefined) {
      for (const child of childTnLst.remainingElementChildren()) {
        if (
          child.namespaceUri === P_NS &&
          (child.localName === "par" || child.localName === "seq")
        ) {
          if (child.localName === "seq") {
            child.consumeAttr("concurrent");
            child.consumeAttr("nextAc");
          }
          subContainers.push(child);
        } else if (
          child.namespaceUri === P_NS &&
          supportedBehaviorNames.has(child.localName)
        ) {
          behaviorCursors.push(child);
        } else {
          otherChildren.push(child);
        }
      }
    }

    const isEffectNode =
      attrs.presetId !== undefined ||
      behaviorCursors.length > 0 ||
      otherChildren.length > 0;

    if (isEffectNode && subContainers.length === 0) {
      const behaviors = behaviorCursors.map(readBehaviorElement);
      const targetIds = [
        ...new Set(behaviors.flatMap((b) => b.targets.map((t) => t.spid))),
      ];
      let unsupportedBehavior:
        | {
            readonly namespaceUri: string;
            readonly localName: string;
            readonly xpath: string;
            readonly targetId?: number;
          }
        | undefined;
      if (otherChildren.length > 0) {
        const firstUnsupported = otherChildren[0]!;
        let unsupportedTargetId: number | undefined;
        const innerBhvr = firstUnsupported.child(P_NS, "cBhvr");
        const innerTgt = innerBhvr?.child(P_NS, "tgtEl");
        if (innerTgt !== undefined) {
          const refs = consumeTargetElement(innerTgt);
          unsupportedTargetId = refs[0]?.spid;
        }
        unsupportedBehavior = {
          namespaceUri: firstUnsupported.namespaceUri,
          localName: firstUnsupported.localName,
          xpath: firstUnsupported.xpath,
          ...(unsupportedTargetId === undefined
            ? {}
            : { targetId: unsupportedTargetId }),
        };
        cursor.suppressUnconsumed();
      }
      const hasSetBehavior = behaviors.some((b) => b.localName === "set");
      const firstRange = behaviors
        .flatMap((b) => b.targets)
        .find((t) => !t.isBackground && t.range !== undefined)?.range;

      rawNodesById.set(timingId, {
        kind: "effect",
        timingId,
        xpath: cursor.xpath,
        ...(parentTimingId === undefined ? {} : { parentTimingId }),
        stepIndex,
        ...(attrs.presetId === undefined ? {} : { presetId: attrs.presetId }),
        ...(attrs.presetClass === undefined
          ? {}
          : { presetClass: attrs.presetClass }),
        ...(attrs.presetSubtype === undefined
          ? {}
          : { presetSubtype: attrs.presetSubtype }),
        ...(attrs.grpId === undefined ? {} : { buildGroupId: attrs.grpId }),
        ...(attrs.repeatCount === undefined
          ? {}
          : { repeatCount: attrs.repeatCount }),
        autoRev: attrs.autoRev,
        ...(attrs.accel === undefined ? {} : { accel: attrs.accel }),
        ...(attrs.decel === undefined ? {} : { decel: attrs.decel }),
        delayMs,
        dependsOn,
        ...(iterateBy === undefined ? {} : { iterateBy }),
        ...(iterateIntervalMs === undefined ? {} : { iterateIntervalMs }),
        hasSetBehavior,
        behaviors,
        targetIds,
        ...(firstRange === undefined ? {} : { targetRange: firstRange }),
        ...(unsupportedBehavior === undefined ? {} : { unsupportedBehavior }),
        cursor,
      });
      return timingId;
    }

    const childIds: number[] = [];
    for (const sub of subContainers) {
      const childId = visitTimingNode(sub, timingId, stepIndex);
      if (childId !== undefined) childIds.push(childId);
    }
    rawNodesById.set(timingId, {
      kind: "container",
      timingId,
      containerKind: cursor.localName === "seq" ? "sequence" : "parallel",
      xpath: cursor.xpath,
      ...(parentTimingId === undefined ? {} : { parentTimingId }),
      stepIndex,
      endSyncAll,
      delayMs,
      dependsOn,
      childIds,
      cursor,
    });
    return timingId;
  };

  const stepPars = mainSeqChildren.children(P_NS, "par");
  for (let stepIndex = 0; stepIndex < stepPars.length; stepIndex += 1) {
    const stepPar = stepPars[stepIndex]!;
    const stepCTn = stepPar.child(P_NS, "cTn");
    if (stepCTn === undefined) continue;
    const stepAttrs = consumeCommonTimeNodeAttrs(stepCTn);
    const stepConds = consumeConditionList(stepCTn.child(P_NS, "stCondLst"));
    const trigger: "auto" | "click" = stepConds.some(
      (c) => c.kind === "onBegin",
    )
      ? "auto"
      : "click";
    const stepChildTnLst = stepCTn.child(P_NS, "childTnLst");
    const rootChildIds: number[] = [];
    if (stepChildTnLst !== undefined) {
      for (const childCursor of stepChildTnLst.children(P_NS, "par")) {
        const id = visitTimingNode(childCursor, stepAttrs.id, stepIndex);
        if (id !== undefined) rootChildIds.push(id);
      }
      for (const childCursor of stepChildTnLst.children(P_NS, "seq")) {
        childCursor.consumeAttr("concurrent");
        childCursor.consumeAttr("nextAc");
        const id = visitTimingNode(childCursor, stepAttrs.id, stepIndex);
        if (id !== undefined) rootChildIds.push(id);
      }
    }
    stepInfos.push({
      stepIndex,
      trigger,
      rootChildIds,
      xpath: stepPar.xpath,
    });
  }

  // 3. Validate targets / behaviors and resolve cascaded omission dependencies
  const rootDiagnosticByNodeId = new Map<
    number,
    Extract<ImportDiagnostic, { readonly code: "unsupported-mechanism" }>
  >();
  const omittedNodeIds = new Set<number>();

  const isKnownSlideTarget = (spid: number): boolean =>
    context.placeholderKeyBySpid?.has(spid) === true ||
    context.objectBySpid?.has(spid) === true;

  for (const [timingId, rawNode] of rawNodesById) {
    if (rawNode.kind !== "effect") continue;

    if (rawNode.unsupportedBehavior !== undefined) {
      const targetId =
        rawNode.unsupportedBehavior.targetId ?? rawNode.targetIds[0];
      const diag: Extract<
        ImportDiagnostic,
        { readonly code: "unsupported-mechanism" }
      > = {
        code: "unsupported-mechanism",
        category: "unsupported",
        disposition: "omitted",
        location: locationFor(
          context,
          rawNode.unsupportedBehavior.xpath,
          targetId,
        ),
        namespaceUri: rawNode.unsupportedBehavior.namespaceUri,
        localName: rawNode.unsupportedBehavior.localName,
        mechanismHint: `motion-behavior:${rawNode.unsupportedBehavior.localName}`,
      };
      rootDiagnosticByNodeId.set(timingId, diag);
      omittedNodeIds.add(timingId);
      rawNode.cursor.suppressUnconsumed();
      continue;
    }

    const preset = resolvePreset(
      rawNode.presetClass,
      rawNode.presetId,
      rawNode.behaviors.some((b) => b.localName === "animClr"),
      rawNode.autoRev,
    );

    if (preset === undefined) {
      const targetId = rawNode.targetIds[0];
      const diag: Extract<
        ImportDiagnostic,
        { readonly code: "unsupported-mechanism" }
      > = {
        code: "unsupported-mechanism",
        category: "unsupported",
        disposition: "omitted",
        location: locationFor(context, rawNode.xpath, targetId),
        namespaceUri: P_NS,
        localName: "cTn",
        mechanismHint: `motion-preset:${rawNode.presetClass ?? ""}#${String(rawNode.presetId ?? "")}`,
      };
      rootDiagnosticByNodeId.set(timingId, diag);
      omittedNodeIds.add(timingId);
      rawNode.cursor.suppressUnconsumed();
      continue;
    }

    const missingTargetId = rawNode.targetIds.find(
      (id) => !isKnownSlideTarget(id),
    );
    if (rawNode.targetIds.length !== 1 || missingTargetId !== undefined) {
      const omittedInfo =
        missingTargetId === undefined
          ? undefined
          : context.omittedSpids?.get(missingTargetId);
      const diag: Extract<
        ImportDiagnostic,
        { readonly code: "unsupported-mechanism" }
      > = {
        code: "unsupported-mechanism",
        category: "unsupported",
        disposition: "omitted",
        location: locationFor(
          context,
          rawNode.xpath,
          missingTargetId,
          omittedInfo?.objectName,
        ),
        namespaceUri: P_NS,
        localName: "spTgt",
        mechanismHint:
          omittedInfo === undefined
            ? "missing-animation-target"
            : `omitted-target:${omittedInfo.reason}`,
      };
      rootDiagnosticByNodeId.set(timingId, diag);
      omittedNodeIds.add(timingId);
      rawNode.cursor.suppressUnconsumed();
      continue;
    }

    const targetId = rawNode.targetIds[0]!;
    const targetNode = context.objectBySpid?.get(targetId);

    const omitInvalidEffect = (
      localName: string,
      mechanismHint: string,
    ): void => {
      rootDiagnosticByNodeId.set(timingId, {
        code: "unsupported-mechanism",
        category: "unsupported",
        disposition: "omitted",
        location: locationFor(context, rawNode.xpath, targetId),
        namespaceUri: P_NS,
        localName,
        mechanismHint,
      });
      omittedNodeIds.add(timingId);
      rawNode.cursor.suppressUnconsumed();
    };

    if (rawNode.targetRange?.type === "paragraph" && targetNode !== undefined) {
      const paraCount = targetParagraphCount(targetNode);
      if (paraCount === 0 || rawNode.targetRange.end >= paraCount) {
        omitInvalidEffect("pRg", "invalid-paragraph-range");
        continue;
      }
    }

    if (
      rawNode.targetRange?.type === "chart" &&
      targetNode !== undefined &&
      targetNode.kind !== "chart"
    ) {
      omitInvalidEffect("graphicEl", "invalid-chart-target");
      continue;
    }

    if (preset === "fillColor") {
      if (
        rawNode.targetRange !== undefined ||
        (targetNode !== undefined && !shapeFillState(targetNode).solidFilled)
      ) {
        omitInvalidEffect("animClr", "invalid-fill-color-target");
        continue;
      }
    }

    if (preset === "breathe") {
      const clrBehavior = rawNode.behaviors.find(
        (b) => b.localName === "animClr",
      );
      const fromColor = clrBehavior?.fromColor ?? "#000000";
      const toColor = clrBehavior?.toColor ?? "#FFFFFF";
      if (
        rawNode.presetClass !== "emph" ||
        rawNode.targetRange !== undefined ||
        (targetNode !== undefined && !shapeFillState(targetNode).filled)
      ) {
        omitInvalidEffect("animClr", "invalid-breathe-target");
        continue;
      }
      if (relativeLuminance(toColor) <= relativeLuminance(fromColor)) {
        omitInvalidEffect("animClr", "invalid-breathe-colors");
        continue;
      }
    }
  }

  // Cascade omission dependencies: if a downstream effect or container depends on an omitted
  // effect or an emptied container, mark it omitted and attach it to the root diagnostic's `cascaded`.
  const directDependentsOf = new Map<number, number[]>();
  for (const [timingId, rawNode] of rawNodesById) {
    for (const dep of rawNode.dependsOn) {
      const list = directDependentsOf.get(dep.timingId) ?? [];
      list.push(timingId);
      directDependentsOf.set(dep.timingId, list);
    }
  }

  const hasAnySurvivingEffect = (containerId: number): boolean => {
    const node = rawNodesById.get(containerId);
    if (node === undefined) return false;
    if (node.kind === "effect") return !omittedNodeIds.has(node.timingId);
    return node.childIds.some((cid) => hasAnySurvivingEffect(cid));
  };

  const cascadedByRootDiag = new Map<
    number,
    { readonly location: ImportLocation; readonly role: string }[]
  >();
  const addCascadedToRoot = (
    rootId: number,
    dependentNode: ParsedRawTimingNode,
  ): void => {
    const list = cascadedByRootDiag.get(rootId) ?? [];
    const targetId =
      dependentNode.kind === "effect" ? dependentNode.targetIds[0] : undefined;
    list.push({
      location: locationFor(context, dependentNode.xpath, targetId),
      role: "animation-dependency",
    });
    cascadedByRootDiag.set(rootId, list);
  };

  // Propagate omissions iteratively from each root diagnostic
  for (const [rootId] of rootDiagnosticByNodeId) {
    const queue = [rootId];
    const visited = new Set<number>([rootId]);
    while (queue.length > 0) {
      const currentId = queue.shift()!;
      const currentNode = rawNodesById.get(currentId);
      // If currentNode is a parent container of currentId, check if currentNode's parent becomes empty
      if (currentNode?.parentTimingId !== undefined) {
        const parentNode = rawNodesById.get(currentNode.parentTimingId);
        if (
          parentNode?.kind === "container" &&
          !omittedNodeIds.has(parentNode.timingId) &&
          !hasAnySurvivingEffect(parentNode.timingId)
        ) {
          omittedNodeIds.add(parentNode.timingId);
          if (!visited.has(parentNode.timingId)) {
            visited.add(parentNode.timingId);
            queue.push(parentNode.timingId);
          }
        }
      }
      for (const dependentId of directDependentsOf.get(currentId) ?? []) {
        if (visited.has(dependentId)) continue;
        visited.add(dependentId);
        const dependentNode = rawNodesById.get(dependentId);
        if (dependentNode === undefined) continue;
        omittedNodeIds.add(dependentId);
        dependentNode.cursor.suppressUnconsumed();
        if (dependentNode.kind === "effect") {
          addCascadedToRoot(rootId, dependentNode);
        }
        queue.push(dependentId);
      }
    }
  }

  for (const [rootId, diag] of rootDiagnosticByNodeId) {
    const cascaded = cascadedByRootDiag.get(rootId);
    if (cascaded !== undefined && cascaded.length > 0) {
      diagnostics.push({ ...diag, cascaded });
    } else {
      diagnostics.push(diag);
    }
  }

  // 4. Collect surviving effects in traversal order per step
  const survivingEffects: ImportedMotionEffect[] = [];
  const collectStepEffects = (
    nodeId: number,
    accumulatedDelayMs: number,
    inheritedStartAfter:
      | { readonly timingId: number; readonly offsetMs: number }
      | undefined,
  ): void => {
    if (omittedNodeIds.has(nodeId)) return;
    const node = rawNodesById.get(nodeId);
    if (node === undefined) return;

    const dep = node.dependsOn[0];
    const effectiveDelayMs =
      dep !== undefined
        ? dep.offsetMs
        : inheritedStartAfter !== undefined
          ? inheritedStartAfter.offsetMs + node.delayMs
          : accumulatedDelayMs + node.delayMs;
    const effectiveStartAfter =
      dep !== undefined && !omittedNodeIds.has(dep.timingId)
        ? dep
        : inheritedStartAfter;

    if (node.kind === "container") {
      for (const childId of node.childIds) {
        collectStepEffects(
          childId,
          effectiveStartAfter === undefined ? effectiveDelayMs : 0,
          effectiveStartAfter,
        );
      }
      return;
    }

    const preset = resolvePreset(
      node.presetClass,
      node.presetId,
      node.behaviors.some((b) => b.localName === "animClr"),
      node.autoRev,
    );
    if (preset === undefined || node.targetIds.length !== 1) return;
    const targetId = node.targetIds[0]!;
    const kind: Effect["kind"] =
      node.presetClass === "emph"
        ? "emphasis"
        : node.presetClass === "exit"
          ? "exit"
          : "entrance";
    const nonSetDurations = node.behaviors
      .filter((b) => b.localName !== "set")
      .map((b) => b.durationMs)
      .filter((d): d is number => d !== undefined);
    const traversalDurationMs =
      nonSetDurations.length > 0
        ? Math.max(...nonSetDurations)
        : preset === "appear" || preset === "disappear"
          ? 1
          : 500;
    const durationMs =
      preset === "breathe" ? traversalDurationMs * 2 : traversalDurationMs;

    let easing: Easing | undefined;
    if (
      (node.accel !== undefined && node.accel !== 0) ||
      (node.decel !== undefined && node.decel !== 0)
    ) {
      const accelRatio =
        node.accel !== undefined && node.accel !== 0 ? node.accel / 100_000 : 0;
      const decelRatio =
        node.decel !== undefined && node.decel !== 0 ? node.decel / 100_000 : 0;
      if (accelRatio < 0 || decelRatio < 0 || accelRatio + decelRatio > 1) {
        diagnostics.push({
          code: "invalid-xml-value",
          category: "corrupt-input",
          disposition: "recovered",
          location: locationFor(context, node.xpath, targetId),
          attributeOrElement: "p:cTn/@accel",
        });
      } else {
        easing = {
          ...(accelRatio > 0 ? { accelerate: accelRatio } : {}),
          ...(decelRatio > 0 ? { decelerate: decelRatio } : {}),
        };
      }
    }

    let from: MotionDirection | number | undefined;
    if (
      preset === "wipe" ||
      preset === "wipeOut" ||
      preset === "fly" ||
      preset === "flyOut"
    ) {
      from =
        node.presetSubtype === undefined
          ? undefined
          : directionBySubtype[node.presetSubtype];
    }

    let to: number | undefined;
    let by: MotionGranularity | number | undefined = node.iterateBy;
    let iterateIntervalMs = node.iterateIntervalMs;
    const targetNode = context.objectBySpid?.get(targetId);
    if (
      by !== undefined &&
      (preset === "fillColor" ||
        preset === "breathe" ||
        preset === "spin" ||
        (targetNode !== undefined && targetParagraphCount(targetNode) === 0))
    ) {
      diagnostics.push({
        code: "invalid-xml-value",
        category: "corrupt-input",
        disposition: "recovered",
        location: locationFor(context, node.xpath, targetId),
        attributeOrElement: "p:cTn/p:iterate",
        rawValue: by,
      });
      by = undefined;
      iterateIntervalMs = undefined;
    }

    if (preset === "spin") {
      const rotBehavior = node.behaviors.find((b) => b.localName === "animRot");
      if (rotBehavior?.rotationBy !== undefined) {
        by = rotBehavior.rotationBy / 60_000;
      } else {
        if (rotBehavior?.rotationFrom !== undefined) {
          from = rotBehavior.rotationFrom / 60_000;
        }
        if (
          rotBehavior?.rotationTo !== undefined &&
          rotBehavior.rotationTo !== 0
        ) {
          to = rotBehavior.rotationTo / 60_000;
        }
      }
    }

    const clrBehavior = node.behaviors.find((b) => b.localName === "animClr");
    const fromColor = clrBehavior?.fromColor;
    const toColor = clrBehavior?.toColor;

    // Exclude mixed-shape companion geometry nodes (paragraph target, grpId=0, NO <p:set>, entrance)
    // paired with a whole-shape entrance on the same target in the same step.
    if (
      !node.hasSetBehavior &&
      kind === "entrance" &&
      node.targetRange?.type === "paragraph" &&
      (node.buildGroupId ?? 0) === 0
    ) {
      return;
    }

    let effectiveRepeatCount = node.repeatCount;
    if (
      effectiveRepeatCount !== undefined &&
      !isValidRepeatModifier({
        preset,
        repeatCount: effectiveRepeatCount,
        ...(by === undefined ? {} : { by }),
        durationMs,
        delayMs: effectiveDelayMs,
      })
    ) {
      diagnostics.push({
        code: "invalid-xml-value",
        category: "corrupt-input",
        disposition: "recovered",
        location: locationFor(context, node.xpath, targetId),
        attributeOrElement: "p:cTn/@repeatCount",
        rawValue: String(effectiveRepeatCount),
      });
      effectiveRepeatCount = undefined;
    }

    const built = buildEffectDeclaration({
      preset,
      kind,
      durationMs,
      delayMs: effectiveDelayMs,
      ...(from === undefined ? {} : { from }),
      ...(to === undefined ? {} : { to }),
      ...(fromColor === undefined ? {} : { fromColor }),
      ...(toColor === undefined ? {} : { toColor }),
      ...(easing === undefined ? {} : { easing }),
      ...(by === undefined ? {} : { by }),
      ...(iterateIntervalMs === undefined
        ? {}
        : { intervalMs: iterateIntervalMs }),
      ...(effectiveRepeatCount === undefined
        ? {}
        : { repeatCount: effectiveRepeatCount }),
    });

    survivingEffects.push({
      effect: built.effect,
      declaration: built.declaration,
      targetId,
      ...(node.targetRange === undefined
        ? {}
        : { targetRange: node.targetRange }),
      ...(node.buildGroupId === undefined
        ? {}
        : { buildGroupId: node.buildGroupId }),
      stepIndex: node.stepIndex,
      ...(node.parentTimingId === undefined
        ? {}
        : { parentTimingId: node.parentTimingId }),
      timingId: node.timingId,
      effectiveDelayMs,
      ...(effectiveStartAfter === undefined
        ? {}
        : {
            startAfterTimingId: effectiveStartAfter.timingId,
            startAfterOffsetMs: effectiveStartAfter.offsetMs,
          }),
      ...(easing === undefined ? {} : { easing }),
      ...(effectiveRepeatCount === undefined
        ? {}
        : { repeatCount: effectiveRepeatCount }),
      ...(from === undefined ? {} : { from }),
      ...(to === undefined ? {} : { to }),
      ...(by === undefined ? {} : { by }),
      ...(iterateIntervalMs === undefined
        ? {}
        : { intervalMs: iterateIntervalMs }),
      ...(fromColor === undefined ? {} : { fromColor }),
      ...(toColor === undefined ? {} : { toColor }),
    });
  };

  for (const step of stepInfos) {
    for (const rootId of step.rootChildIds) {
      collectStepEffects(rootId, 0, undefined);
    }
  }

  // 5. SlideMaster / SlideLayout owner timing reconstruction
  if (context.placeholderKeyBySpid !== undefined) {
    // resolveSharedCanvasLayers() assigns emitted spids starting after declared
    // placeholders (2 + placeholderCount) in pre-order across object nodes.
    // Remap imported XML spids for surviving non-placeholder children to their
    // deterministic emitted spids so resolve -> emit -> validatePptx never writes
    // stale source spids.
    const placeholdersProp = context.slideNode.props["placeholders"];
    const placeholderCount =
      typeof placeholdersProp === "object" &&
      placeholdersProp !== null &&
      !Array.isArray(placeholdersProp)
        ? Object.keys(placeholdersProp).length
        : 0;
    const emittedSpidByNode = new Map<CommonNode, number>();
    let nextOwnerObjectId = 2 + placeholderCount;
    const assignEmittedOwnerSpids = (children: readonly unknown[]): void => {
      for (const child of children) {
        if (!(child instanceof CommonNode)) continue;
        emittedSpidByNode.set(child, nextOwnerObjectId++);
        if (child.kind === "group") {
          assignEmittedOwnerSpids(child.children);
        }
      }
    };
    assignEmittedOwnerSpids(context.slideNode.children);

    const ownerSteps: OwnerMotionStep[] = [];
    for (const stepInfo of stepInfos) {
      const stepEffects = survivingEffects.filter(
        (eff) => eff.stepIndex === stepInfo.stepIndex,
      );
      if (stepEffects.length === 0) continue;
      const mappedEffects: OwnerMotionEffect[] = [];
      for (const eff of stepEffects) {
        const placeholderKey = context.placeholderKeyBySpid?.get(eff.targetId);
        let target: OwnerMotionTarget | undefined;
        if (placeholderKey !== undefined) {
          target = { placeholder: placeholderKey };
        } else {
          const targetNode = context.objectBySpid?.get(eff.targetId);
          const emittedSpid =
            targetNode === undefined
              ? undefined
              : emittedSpidByNode.get(targetNode);
          if (emittedSpid !== undefined) {
            target = { spid: emittedSpid };
          } else {
            diagnostics.push({
              code: "unsupported-mechanism",
              category: "unsupported",
              disposition: "omitted",
              location: locationFor(context, timingCursor.xpath, eff.targetId),
              namespaceUri: P_NS,
              localName: "spTgt",
            });
            continue;
          }
        }
        mappedEffects.push({
          target,
          effect: eff.declaration as OwnerMotionEffect["effect"],
          ...(eff.targetRange === undefined ? {} : { range: eff.targetRange }),
        });
      }
      if (mappedEffects.length > 0) {
        ownerSteps.push({
          trigger: stepInfo.trigger,
          effects: mappedEffects,
        });
      }
    }
    if (ownerSteps.length > 0) {
      const motion: OwnerMotion = { steps: ownerSteps };
      context.slideNode.set({ motion });
    }
    return diagnostics.length === 0 ? {} : { diagnostics };
  }

  // 6. Collapse expanded paragraph (`by: "paragraph"`) and chart (`by: "series" | ...`) effects
  const collapsedEffects: ImportedMotionEffect[] = [];
  const consumedIndices = new Set<number>();

  for (let i = 0; i < survivingEffects.length; i += 1) {
    if (consumedIndices.has(i)) continue;
    const head = survivingEffects[i]!;

    // 6a. Whole-text paragraph projection (grpId=1, single effect already represents whole object)
    if (head.targetRange?.type === "paragraph" && head.buildGroupId === 1) {
      consumedIndices.add(i);
      const { targetRange: _omittedRange, ...rest } = head;
      collapsedEffects.push(rest);
      continue;
    }

    // 6b. Chart granular expansion (series / category / seriesElements / categoryElements)
    if (head.targetRange?.type === "chart") {
      const chartBuild = chartBuilds.get(
        `${head.targetId}:${head.buildGroupId ?? 0}`,
      );
      const groupIndices: number[] = [i];
      for (let j = i + 1; j < survivingEffects.length; j += 1) {
        if (consumedIndices.has(j)) continue;
        const cand = survivingEffects[j]!;
        if (
          cand.targetRange?.type === "chart" &&
          cand.stepIndex === head.stepIndex &&
          commonEffectPropsMatch(head, cand) &&
          cand.startAfterTimingId === head.startAfterTimingId &&
          cand.startAfterOffsetMs === head.startAfterOffsetMs
        ) {
          groupIndices.push(j);
        }
      }
      const mappedBy: MotionGranularity | undefined =
        chartBuild?.mode === "series"
          ? "series"
          : chartBuild?.mode === "category"
            ? "category"
            : chartBuild?.mode === "seriesEl"
              ? "seriesElements"
              : chartBuild?.mode === "categoryEl"
                ? "categoryElements"
                : head.targetRange.step === "series"
                  ? "series"
                  : head.targetRange.step === "category"
                    ? "category"
                    : head.targetRange.step === "ptInSeries"
                      ? "seriesElements"
                      : head.targetRange.step === "ptInCategory"
                        ? "categoryElements"
                        : undefined;
      for (const idx of groupIndices) consumedIndices.add(idx);
      const baseDelay = Math.min(
        ...groupIndices.map((idx) => survivingEffects[idx]!.effectiveDelayMs),
      );
      const effectiveRepeat =
        mappedBy !== undefined && head.repeatCount !== undefined
          ? undefined
          : head.repeatCount;
      if (mappedBy !== undefined && head.repeatCount !== undefined) {
        diagnostics.push({
          code: "invalid-xml-value",
          category: "corrupt-input",
          disposition: "recovered",
          location: locationFor(context, timingCursor.xpath, head.targetId),
          attributeOrElement: "p:cTn/@repeatCount",
          rawValue: String(head.repeatCount),
        });
      }
      const rebuilt = buildEffectDeclaration({
        preset: head.effect.preset,
        kind: head.effect.kind,
        durationMs: head.effect.duration,
        delayMs: baseDelay,
        ...(head.from === undefined ? {} : { from: head.from }),
        ...(head.to === undefined ? {} : { to: head.to }),
        ...(head.fromColor === undefined ? {} : { fromColor: head.fromColor }),
        ...(head.toColor === undefined ? {} : { toColor: head.toColor }),
        ...(head.easing === undefined ? {} : { easing: head.easing }),
        ...(mappedBy === undefined ? {} : { by: mappedBy }),
        ...(effectiveRepeat === undefined
          ? {}
          : { repeatCount: effectiveRepeat }),
      });
      const {
        targetRange: _omittedRange,
        repeatCount: _droppedRepeat,
        ...headRest
      } = head;
      collapsedEffects.push({
        ...headRest,
        effectiveDelayMs: baseDelay,
        ...(mappedBy === undefined ? {} : { by: mappedBy }),
        ...(effectiveRepeat === undefined
          ? {}
          : { repeatCount: effectiveRepeat }),
        effect: rebuilt.effect,
        declaration: rebuilt.declaration,
      });
      continue;
    }

    // 6c. Paragraph-by-paragraph expansion (grpId=0, build=p, full staggered pRg 0..N-1)
    if (
      head.targetRange?.type === "paragraph" &&
      head.targetRange.start === 0 &&
      head.targetRange.end === 0 &&
      (head.buildGroupId ?? 0) === 0
    ) {
      const pBuild = paragraphBuilds.get(`${head.targetId}:0`);
      const targetObj = context.objectBySpid?.get(head.targetId);
      const existingParaCount =
        targetObj === undefined
          ? 0
          : targetObj.children.filter(
              (c) => c instanceof CommonNode && c.type === "paragraph",
            ).length;
      const targetParaCount =
        existingParaCount > 0
          ? existingParaCount
          : targetObj?.kind === "shape" && targetObj.children.length > 0
            ? 1
            : 0;
      if (pBuild?.build === "p" && targetParaCount > 1) {
        const groupIndices: number[] = [i];
        let expectedNextStart = head.targetRange.end + 1;
        for (let j = i + 1; j < survivingEffects.length; j += 1) {
          if (consumedIndices.has(j)) continue;
          const cand = survivingEffects[j]!;
          if (
            cand.targetRange?.type === "paragraph" &&
            cand.targetRange.start === expectedNextStart &&
            cand.targetRange.end === expectedNextStart &&
            cand.stepIndex === head.stepIndex &&
            commonEffectPropsMatch(head, cand) &&
            cand.effectiveDelayMs ===
              head.effectiveDelayMs +
                expectedNextStart * head.effect.duration &&
            cand.startAfterTimingId === head.startAfterTimingId &&
            cand.startAfterOffsetMs === head.startAfterOffsetMs
          ) {
            groupIndices.push(j);
            expectedNextStart = cand.targetRange.end + 1;
          }
        }
        if (groupIndices.length === targetParaCount) {
          for (const idx of groupIndices) consumedIndices.add(idx);
          const baseDelay = Math.min(
            ...groupIndices.map(
              (idx) => survivingEffects[idx]!.effectiveDelayMs,
            ),
          );
          const effectiveRepeat =
            head.repeatCount !== undefined ? undefined : head.repeatCount;
          if (head.repeatCount !== undefined) {
            diagnostics.push({
              code: "invalid-xml-value",
              category: "corrupt-input",
              disposition: "recovered",
              location: locationFor(context, timingCursor.xpath, head.targetId),
              attributeOrElement: "p:cTn/@repeatCount",
              rawValue: String(head.repeatCount),
            });
          }
          const rebuilt = buildEffectDeclaration({
            preset: head.effect.preset,
            kind: head.effect.kind,
            durationMs: head.effect.duration,
            delayMs: baseDelay,
            ...(head.from === undefined ? {} : { from: head.from }),
            ...(head.to === undefined ? {} : { to: head.to }),
            ...(head.fromColor === undefined
              ? {}
              : { fromColor: head.fromColor }),
            ...(head.toColor === undefined ? {} : { toColor: head.toColor }),
            ...(head.easing === undefined ? {} : { easing: head.easing }),
            by: "paragraph",
            ...(effectiveRepeat === undefined
              ? {}
              : { repeatCount: effectiveRepeat }),
          });
          const {
            targetRange: _omittedRange,
            repeatCount: _droppedRepeat,
            ...headRest
          } = head;
          collapsedEffects.push({
            ...headRest,
            effectiveDelayMs: baseDelay,
            by: "paragraph",
            ...(effectiveRepeat === undefined
              ? {}
              : { repeatCount: effectiveRepeat }),
            effect: rebuilt.effect,
            declaration: rebuilt.declaration,
          });
          continue;
        }
      }
    }

    consumedIndices.add(i);
    collapsedEffects.push(head);
  }

  // 7. Resolve targetId -> live CommonNode on slideNode, handling Group spin lifting
  // 7. Resolve targetId -> live CommonNode on slideNode, handling Group spin lifting
  const objectBySpid = context.objectBySpid;
  const slideContainer = context.slideNode;
  if (
    objectBySpid === undefined ||
    !(slideContainer instanceof CommonContainer)
  ) {
    return diagnostics.length === 0 ? {} : { diagnostics };
  }

  // Build map from nested child CommonNode -> owning top-level slide child CommonNode,
  // and top-level/nested Group leaf memberships so Group spin projections lift cleanly back onto Group.
  const topLevelOwnerByNode = new Map<CommonNode, CommonNode>();
  const registerTopLevelOwner = (
    node: CommonNode,
    topLevel: CommonNode,
  ): void => {
    topLevelOwnerByNode.set(node, topLevel);
    for (const child of node.children) {
      if (child instanceof CommonNode) {
        registerTopLevelOwner(child, topLevel);
      }
    }
  };
  for (const child of context.slideNode.children) {
    if (child instanceof CommonNode) {
      registerTopLevelOwner(child, child);
    }
  }

  const spidsByNode = new Map<CommonNode, number>();
  for (const [spid, node] of objectBySpid) {
    spidsByNode.set(node, spid);
  }

  const collectGroupLeaves = (groupNode: CommonNode): CommonNode[] => {
    const leaves: CommonNode[] = [];
    for (const member of groupNode.children) {
      if (!(member instanceof CommonNode)) continue;
      if (member.kind === "group") {
        leaves.push(...collectGroupLeaves(member));
      } else {
        leaves.push(member);
      }
    }
    return leaves;
  };

  // Lift synchronized Group spin projections onto their enclosing GroupNode when all leaf members share the spin effect.
  const liftedCollapsed: ImportedMotionEffect[] = [];
  const spinConsumed = new Set<number>();
  const findAllGroups = (nodes: readonly unknown[]): CommonNode[] => {
    const groups: CommonNode[] = [];
    for (const n of nodes) {
      if (n instanceof CommonNode && n.kind === "group") {
        groups.push(n);
        groups.push(...findAllGroups(n.children));
      }
    }
    return groups;
  };
  for (const groupNode of findAllGroups(context.slideNode.children)) {
    const groupSpid = spidsByNode.get(groupNode);
    if (groupSpid === undefined) continue;
    const leaves = collectGroupLeaves(groupNode);
    if (leaves.length === 0) continue;
    const leafSpids = leaves
      .map((leaf) => spidsByNode.get(leaf))
      .filter((id): id is number => id !== undefined);
    if (leafSpids.length !== leaves.length) continue;

    // Look for synchronous spin effects in the same step matching every leaf spid
    for (let idx = 0; idx < collapsedEffects.length; idx += 1) {
      if (spinConsumed.has(idx)) continue;
      const candidate = collapsedEffects[idx]!;
      if (
        candidate.effect.preset !== "spin" ||
        candidate.targetId !== leafSpids[0]
      ) {
        continue;
      }
      const matchedLeafIndices: number[] = [idx];
      for (const leafSpid of leafSpids.slice(1)) {
        const matchIdx = collapsedEffects.findIndex(
          (other, otherIdx) =>
            !spinConsumed.has(otherIdx) &&
            other.effect.preset === "spin" &&
            other.targetId === leafSpid &&
            other.stepIndex === candidate.stepIndex &&
            other.effectiveDelayMs === candidate.effectiveDelayMs &&
            other.effect.duration === candidate.effect.duration &&
            other.by === candidate.by &&
            other.from === candidate.from &&
            other.to === candidate.to &&
            other.startAfterTimingId === candidate.startAfterTimingId &&
            other.startAfterOffsetMs === candidate.startAfterOffsetMs,
        );
        if (matchIdx >= 0) matchedLeafIndices.push(matchIdx);
      }
      if (matchedLeafIndices.length === leafSpids.length) {
        for (const m of matchedLeafIndices) spinConsumed.add(m);
        liftedCollapsed.push({
          ...candidate,
          targetId: groupSpid,
        });
      }
    }
  }
  for (let i = 0; i < collapsedEffects.length; i += 1) {
    if (!spinConsumed.has(i)) {
      liftedCollapsed.push(collapsedEffects[i]!);
    }
  }

  // 8. Compute click step numbering & attach enter/emphasis/exit (with slot-chain & onStep folding)
  const clickNumberByStepIndex = new Map<number, number>();
  let nextClickNumber = 1;
  for (const info of stepInfos) {
    if (info.trigger === "click") {
      const hasSurviving = liftedCollapsed.some(
        (e) => e.stepIndex === info.stepIndex,
      );
      if (hasSurviving) {
        clickNumberByStepIndex.set(info.stepIndex, nextClickNumber++);
      }
    }
  }

  const wholeEffectsByNode = new Map<CommonNode, ImportedMotionEffect[]>();
  const paragraphEffectsByNode = new Map<
    CommonNode,
    Map<number, ImportedMotionEffect[]>
  >();
  const topLevelStepIndices = new Map<CommonNode, Set<number>>();
  for (const eff of liftedCollapsed) {
    const targetNode = objectBySpid.get(eff.targetId);
    if (targetNode === undefined) continue;
    if (eff.targetRange?.type === "paragraph") {
      let byPara = paragraphEffectsByNode.get(targetNode);
      if (byPara === undefined) {
        byPara = new Map<number, ImportedMotionEffect[]>();
        paragraphEffectsByNode.set(targetNode, byPara);
      }
      const pIdx = eff.targetRange.start;
      const pList = byPara.get(pIdx) ?? [];
      pList.push(eff);
      byPara.set(pIdx, pList);
      continue;
    }
    const list = wholeEffectsByNode.get(targetNode) ?? [];
    list.push(eff);
    wholeEffectsByNode.set(targetNode, list);
    const topLevelNode = topLevelOwnerByNode.get(targetNode) ?? targetNode;
    const indices = topLevelStepIndices.get(topLevelNode) ?? new Set<number>();
    indices.add(eff.stepIndex);
    topLevelStepIndices.set(topLevelNode, indices);
  }

  const anchorStepIndexByTopLevel = new Map<CommonNode, number>();
  for (const [topLevelNode, indices] of topLevelStepIndices) {
    anchorStepIndexByTopLevel.set(topLevelNode, Math.min(...indices));
  }

  const internalChainTimingIds = new Set<number>();

  const finiteDurationOf = (eff: ImportedMotionEffect): number | undefined => {
    if (
      eff.repeatCount === "forever" ||
      eff.by === "word" ||
      eff.by === "letter" ||
      eff.by === "paragraph"
    ) {
      return undefined;
    }
    const rounds = typeof eff.repeatCount === "number" ? eff.repeatCount : 1;
    return eff.effect.duration * rounds;
  };

  const applySlotAdjustedPatch = (
    targetNode: CommonNode,
    list: readonly ImportedMotionEffect[],
    anchorStep: number | undefined,
  ): void => {
    const adjustedByEffect = new Map<
      ImportedMotionEffect,
      Effect | ReturnType<typeof repeat>
    >();
    const byStep = new Map<number, ImportedMotionEffect[]>();
    for (const eff of list) {
      const group = byStep.get(eff.stepIndex) ?? [];
      group.push(eff);
      byStep.set(eff.stepIndex, group);
    }

    for (const stepList of byStep.values()) {
      let prevSlotTimingIds = new Set<number>();
      let prevSlotEndMs: number | undefined = 0;
      let prevSlotEndKnown = true;
      let hasPrevSlot = false;
      for (const slotKind of ["entrance", "emphasis", "exit"] as const) {
        const slotItems = stepList.filter((e) => e.effect.kind === slotKind);
        if (slotItems.length === 0) continue;
        const slotTimingIds = new Set<number>(slotItems.map((e) => e.timingId));
        let slotMaxEnd: number | undefined = 0;
        let slotKnown: boolean = prevSlotEndKnown;
        for (const item of slotItems) {
          let adjustedDelay = item.effectiveDelayMs;
          if (hasPrevSlot) {
            if (
              item.startAfterTimingId !== undefined &&
              prevSlotTimingIds.has(item.startAfterTimingId)
            ) {
              internalChainTimingIds.add(item.timingId);
              adjustedDelay = item.startAfterOffsetMs ?? 0;
            } else if (prevSlotEndKnown && prevSlotEndMs !== undefined) {
              adjustedDelay = Math.max(
                0,
                item.effectiveDelayMs - prevSlotEndMs,
              );
            }
          }
          const rebuiltDecl =
            adjustedDelay === item.effectiveDelayMs
              ? item.declaration
              : buildEffectDeclaration({
                  preset: item.effect.preset,
                  kind: item.effect.kind,
                  durationMs: item.effect.duration,
                  delayMs: adjustedDelay,
                  ...(item.from === undefined ? {} : { from: item.from }),
                  ...(item.to === undefined ? {} : { to: item.to }),
                  ...(item.fromColor === undefined
                    ? {}
                    : { fromColor: item.fromColor }),
                  ...(item.toColor === undefined
                    ? {}
                    : { toColor: item.toColor }),
                  ...(item.easing === undefined ? {} : { easing: item.easing }),
                  ...(item.by === undefined ? {} : { by: item.by }),
                  ...(item.intervalMs === undefined
                    ? {}
                    : { intervalMs: item.intervalMs }),
                  ...(item.repeatCount === undefined
                    ? {}
                    : { repeatCount: item.repeatCount }),
                }).declaration;
          adjustedByEffect.set(item, rebuiltDecl);

          const finiteDur = finiteDurationOf(item);
          if (
            slotKnown &&
            finiteDur !== undefined &&
            prevSlotEndMs !== undefined
          ) {
            const itemEnd: number = prevSlotEndMs + adjustedDelay + finiteDur;
            slotMaxEnd =
              slotMaxEnd === undefined
                ? itemEnd
                : Math.max(slotMaxEnd, itemEnd);
          } else {
            slotKnown = false;
          }
        }
        prevSlotTimingIds = slotTimingIds;
        prevSlotEndMs = slotMaxEnd;
        prevSlotEndKnown = slotKnown;
        hasPrevSlot = true;
      }
    }

    const wrapWithStepIfNeeded = <
      D extends
        | EntranceEffectDeclaration
        | EmphasisEffectDeclaration
        | ExitEffectDeclaration,
    >(
      eff: ImportedMotionEffect,
      decl: Effect | ReturnType<typeof repeat>,
    ): D => {
      if (anchorStep !== undefined && eff.stepIndex !== anchorStep) {
        const clickNum = clickNumberByStepIndex.get(eff.stepIndex);
        if (clickNum !== undefined) {
          return onStep(
            clickNum,
            decl as Parameters<typeof onStep>[1],
          ) as unknown as D;
        }
      }
      return decl as D;
    };

    const enters = list
      .filter((e) => e.effect.kind === "entrance")
      .map((e) =>
        wrapWithStepIfNeeded<EntranceEffectDeclaration>(
          e,
          adjustedByEffect.get(e) ?? e.declaration,
        ),
      );
    const emphases = list
      .filter((e) => e.effect.kind === "emphasis")
      .map((e) =>
        wrapWithStepIfNeeded<EmphasisEffectDeclaration>(
          e,
          adjustedByEffect.get(e) ?? e.declaration,
        ),
      );
    const exits = list
      .filter((e) => e.effect.kind === "exit")
      .map((e) =>
        wrapWithStepIfNeeded<ExitEffectDeclaration>(
          e,
          adjustedByEffect.get(e) ?? e.declaration,
        ),
      );

    const patch: {
      enter?: EntranceEffectDeclaration | readonly EntranceEffectDeclaration[];
      emphasis?:
        | EmphasisEffectDeclaration
        | readonly EmphasisEffectDeclaration[];
      exit?: ExitEffectDeclaration | readonly ExitEffectDeclaration[];
    } = {};
    if (enters.length === 1) patch.enter = enters[0]!;
    else if (enters.length > 1) patch.enter = enters;
    if (emphases.length === 1) patch.emphasis = emphases[0]!;
    else if (emphases.length > 1) patch.emphasis = emphases;
    if (exits.length === 1) patch.exit = exits[0]!;
    else if (exits.length > 1) patch.exit = exits;
    if (
      patch.enter !== undefined ||
      patch.emphasis !== undefined ||
      patch.exit !== undefined
    ) {
      targetNode.set(patch);
    }
  };

  for (const [targetNode, list] of wholeEffectsByNode) {
    const topLevelNode = topLevelOwnerByNode.get(targetNode) ?? targetNode;
    const anchorStep = anchorStepIndexByTopLevel.get(topLevelNode);
    applySlotAdjustedPatch(targetNode, list, anchorStep);
  }

  const paragraphNodesByKey = new Map<
    string,
    {
      readonly pNode: CommonNode;
      readonly anchorStep: number;
      readonly clickStep?: number;
      readonly startAfterTimingId?: number;
      readonly startAfterOffsetMs?: number;
    }
  >();
  const paragraphPublishTimingIds = new Set<number>();

  for (const [targetNode, byPara] of paragraphEffectsByNode) {
    if (
      targetNode instanceof RichTextContainer &&
      !targetNode.children.some(
        (c) => c instanceof CommonNode && c.type === "paragraph",
      ) &&
      targetNode.children.length > 0
    ) {
      const p0 = createParagraphNode({});
      const rawChildren = [...targetNode.children];
      p0.content(
        ...rawChildren.map((child) =>
          child instanceof CommonNode
            ? child
            : (child as { readonly text: string }).text,
        ),
      );
      targetNode.content(p0);
    }
    const pNodes = targetNode.children.filter(
      (c): c is CommonNode => c instanceof CommonNode && c.type === "paragraph",
    );
    const stepByParagraph = new Map<CommonNode, number>();
    for (const [pIndex, pList] of byPara) {
      const pNode = pNodes[pIndex];
      if (pNode === undefined) continue;
      const anchorStep = Math.min(...pList.map((e) => e.stepIndex));
      applySlotAdjustedPatch(pNode, pList, anchorStep);
      const clickStep = clickNumberByStepIndex.get(anchorStep);
      if (clickStep !== undefined) {
        stepByParagraph.set(pNode, clickStep);
      }
      const anchorEff =
        pList.find((e) => e.stepIndex === anchorStep) ?? pList[0];
      paragraphNodesByKey.set(`${targetNode.id}:${pIndex}`, {
        pNode,
        anchorStep,
        ...(clickStep === undefined ? {} : { clickStep }),
        ...(anchorEff?.startAfterTimingId !== undefined &&
        !internalChainTimingIds.has(anchorEff.timingId)
          ? {
              startAfterTimingId: anchorEff.startAfterTimingId,
              ...(anchorEff.startAfterOffsetMs === undefined
                ? {}
                : { startAfterOffsetMs: anchorEff.startAfterOffsetMs }),
            }
          : {}),
      });
    }
    if (stepByParagraph.size > 0 && targetNode instanceof RichTextContainer) {
      const originalPChildren = [...targetNode.children];
      for (const child of originalPChildren) {
        if (child instanceof CommonNode) {
          child.remove();
        }
      }
      const wrappedPChildren: RichTextItem[] = [];
      let activeStepHost: CommonContainer | undefined;
      let activeClickStep: number | undefined;
      for (const child of originalPChildren) {
        if (child instanceof CommonNode) {
          const clickStep = stepByParagraph.get(child);
          if (clickStep !== undefined) {
            if (activeStepHost === undefined || activeClickStep !== clickStep) {
              activeStepHost = createMotionStepNode({
                n: clickStep,
                stepIdentity: createMotionStepIdentity(),
              });
              activeClickStep = clickStep;
              wrappedPChildren.push(activeStepHost);
            }
            activeStepHost.append(child);
          } else {
            activeStepHost = undefined;
            activeClickStep = undefined;
            wrappedPChildren.push(child);
          }
        } else {
          activeStepHost = undefined;
          activeClickStep = undefined;
          wrappedPChildren.push((child as { readonly text: string }).text);
        }
      }
      targetNode.content(...wrappedPChildren);
    }
  }

  // 9. Reconstruct Step & MotionScope choreography wrappers on slideNode.children
  // Identify every startAfter dependency timingId referenced by any surviving effect
  // (excluding internal same-target auto-chain links), and map each timingId to a MotionScopeHandle.
  const referencedTimingIds = new Set<number>();
  for (const eff of liftedCollapsed) {
    if (
      eff.startAfterTimingId !== undefined &&
      !internalChainTimingIds.has(eff.timingId) &&
      !omittedNodeIds.has(eff.startAfterTimingId)
    ) {
      referencedTimingIds.add(eff.startAfterTimingId);
      const sourceEff = liftedCollapsed.find(
        (candidate) => candidate.timingId === eff.startAfterTimingId,
      );
      if (sourceEff?.targetRange?.type === "paragraph") {
        paragraphPublishTimingIds.add(eff.startAfterTimingId);
      }
    }
  }
  const handleByTimingId = new Map<number, MotionScopeHandle>();
  for (const tid of referencedTimingIds) {
    handleByTimingId.set(tid, createMotionScopeHandle());
  }

  for (const pInfo of paragraphNodesByKey.values()) {
    if (pInfo.startAfterTimingId === undefined) continue;
    const sourceHandle = handleByTimingId.get(pInfo.startAfterTimingId);
    if (sourceHandle === undefined) continue;
    const offset = pInfo.startAfterOffsetMs ?? 0;
    const moment = paragraphPublishTimingIds.has(pInfo.startAfterTimingId)
      ? sourceHandle.cue("done")
      : sourceHandle.end;
    pInfo.pNode.set({
      startAt: offset > 0 ? after(moment, offset) : moment,
    });
  }

  for (const tid of paragraphPublishTimingIds) {
    const sourceEff = liftedCollapsed.find((e) => e.timingId === tid);
    const sourceNode =
      sourceEff === undefined
        ? undefined
        : objectBySpid.get(sourceEff.targetId);
    const handle = handleByTimingId.get(tid);
    if (
      sourceNode !== undefined &&
      sourceEff?.targetRange?.type === "paragraph" &&
      handle !== undefined
    ) {
      const pInfo = paragraphNodesByKey.get(
        `${sourceNode.id}:${sourceEff.targetRange.start}`,
      );
      pInfo?.pNode.set({ publishEnd: handle.cue("done") });
    }
  }

  // Build ancestor chain for each timingId -> all MotionScopeHandles that should publish its end
  const ancestorsOfTimingId = (timingId: number): number[] => {
    const chain: number[] = [];
    let current: number | undefined = timingId;
    while (current !== undefined) {
      if (
        handleByTimingId.has(current) &&
        !paragraphPublishTimingIds.has(current)
      ) {
        chain.unshift(current);
      }
      current = rawNodesById.get(current)?.parentTimingId;
    }
    return chain;
  };

  type TopLevelMotionAssignment = {
    readonly clickStep?: number;
    readonly sourceScopeIds: readonly number[];
    readonly startAfter?: MotionTiming;
  };
  const assignmentByTopLevel = new Map<CommonNode, TopLevelMotionAssignment>();

  for (const eff of liftedCollapsed) {
    if (eff.targetRange?.type === "paragraph") continue;
    const targetNode = objectBySpid.get(eff.targetId);
    if (targetNode === undefined) continue;
    const topLevelNode = topLevelOwnerByNode.get(targetNode) ?? targetNode;
    const anchorStep = anchorStepIndexByTopLevel.get(topLevelNode);
    if (eff.stepIndex !== anchorStep) continue;
    if (assignmentByTopLevel.has(topLevelNode)) continue;

    const clickStep = clickNumberByStepIndex.get(eff.stepIndex);
    const sourceScopeIds = ancestorsOfTimingId(eff.timingId);
    let startAfter: MotionTiming | undefined;
    if (
      eff.startAfterTimingId !== undefined &&
      !internalChainTimingIds.has(eff.timingId)
    ) {
      const sourceHandle = handleByTimingId.get(eff.startAfterTimingId);
      if (sourceHandle !== undefined) {
        const offset = eff.startAfterOffsetMs ?? 0;
        const moment = paragraphPublishTimingIds.has(eff.startAfterTimingId)
          ? sourceHandle.cue("done")
          : sourceHandle.end;
        startAfter = offset > 0 ? after(moment, offset) : moment;
      }
    }
    assignmentByTopLevel.set(topLevelNode, {
      ...(clickStep === undefined ? {} : { clickStep }),
      sourceScopeIds,
      ...(startAfter === undefined ? {} : { startAfter }),
    });
  }

  const originalChildren = [...context.slideNode.children];
  const anyMotionScopesNeeded = referencedTimingIds.size > 0;

  // Determine whether any referenced sourceScope has non-contiguous members in slideNode.children.
  // If contiguous, keep exact spTree order without zIndex. If non-contiguous, preserve visual z-order
  // via style.zIndex when grouping a MotionScope's members together.
  const indexByChild = new Map<CommonNode, number>();
  let childOrd = 0;
  for (const child of originalChildren) {
    if (child instanceof CommonNode) {
      indexByChild.set(child, childOrd++);
    }
  }

  let needsScopeReordering = false;
  if (anyMotionScopesNeeded) {
    for (const scopeTimingId of referencedTimingIds) {
      const indices = originalChildren
        .filter((c): c is CommonNode => c instanceof CommonNode)
        .filter((c) =>
          assignmentByTopLevel.get(c)?.sourceScopeIds.includes(scopeTimingId),
        )
        .map((c) => indexByChild.get(c)!)
        .sort((a, b) => a - b);
      if (
        indices.length > 1 &&
        indices[indices.length - 1]! - indices[0]! + 1 !== indices.length
      ) {
        needsScopeReordering = true;
      }
    }
  }

  if (needsScopeReordering) {
    for (const [node, ord] of indexByChild) {
      node.style({ zIndex: ord });
    }
  }

  // Group children by step & scope handles cleanly:
  // First detach all children from slideNode so we can re-append wrapped subtrees.
  if (assignmentByTopLevel.size === 0) {
    return diagnostics.length === 0 ? {} : { diagnostics };
  }

  for (const child of originalChildren) {
    if (child instanceof CommonNode || child instanceof CommonContainer) {
      child.remove();
    }
  }

  // Build ordered wrapper tree in spTree order (reordering only when a referenced scope is split)
  const orderedNodes: CommonNode[] = [...originalChildren].filter(
    (c): c is CommonNode => c instanceof CommonNode,
  );
  if (needsScopeReordering) {
    // Stable sort to cluster nodes sharing sourceScopeIds at their earliest spTree position
    const earliestForScope = new Map<number, number>();
    for (const tid of referencedTimingIds) {
      const minOrd = Math.min(
        ...orderedNodes
          .filter((n) =>
            assignmentByTopLevel.get(n)?.sourceScopeIds.includes(tid),
          )
          .map((n) => indexByChild.get(n)!),
      );
      earliestForScope.set(tid, minOrd);
    }
    orderedNodes.sort((left, right) => {
      const leftAssign = assignmentByTopLevel.get(left);
      const rightAssign = assignmentByTopLevel.get(right);
      const leftKey =
        leftAssign?.sourceScopeIds[0] !== undefined
          ? (earliestForScope.get(leftAssign.sourceScopeIds[0]) ??
            indexByChild.get(left)!)
          : indexByChild.get(left)!;
      const rightKey =
        rightAssign?.sourceScopeIds[0] !== undefined
          ? (earliestForScope.get(rightAssign.sourceScopeIds[0]) ??
            indexByChild.get(right)!)
          : indexByChild.get(right)!;
      return (
        leftKey - rightKey || indexByChild.get(left)! - indexByChild.get(right)!
      );
    });
  }

  // Wrap each contiguous block of nodes sharing (clickStep, sourceScopeIds, startAfter)
  // Note: because a MotionScopeHandle can appear only once on a slide, track created scope hosts!
  const createdScopeHosts = new Map<number, CommonContainer>();

  // Re-append nodes into Step + MotionScope wrappers
  for (const node of orderedNodes) {
    const assign = assignmentByTopLevel.get(node);
    if (assign === undefined) {
      slideContainer.append(node);
      continue;
    }

    let wrapped: CommonNode = node;

    if (assign.startAfter !== undefined) {
      const consumerHandle = createMotionScopeHandle();
      const consumerScope = createMotionScopeNode({
        scope: consumerHandle,
        startAt: assign.startAfter,
      });
      consumerScope.append(wrapped);
      wrapped = consumerScope;
    }

    // Wrap inside source scopes (innermost -> outermost), reusing already-mounted scope host
    let appendedToExistingHost = false;
    for (const scopeId of [...assign.sourceScopeIds].reverse()) {
      const existingHost = createdScopeHosts.get(scopeId);
      if (existingHost !== undefined) {
        existingHost.append(wrapped);
        appendedToExistingHost = true;
        break;
      }
      const handle = handleByTimingId.get(scopeId)!;
      const scopeHost = createMotionScopeNode({ scope: handle });
      scopeHost.append(wrapped);
      createdScopeHosts.set(scopeId, scopeHost);
      wrapped = scopeHost;
    }

    if (appendedToExistingHost) {
      continue;
    }

    if (assign.clickStep !== undefined) {
      // Try coalescing with previous top-level child if it is already `<Step n={clickStep}>`
      const lastChild = slideContainer.children.at(-1);
      if (
        lastChild instanceof CommonContainer &&
        lastChild.kind === "motion-step" &&
        lastChild.props["n"] === assign.clickStep
      ) {
        lastChild.append(wrapped);
      } else {
        const stepHost = createMotionStepNode({
          n: assign.clickStep,
          stepIdentity: createMotionStepIdentity(),
        });
        stepHost.append(wrapped);
        slideContainer.append(stepHost);
      }
    } else {
      slideContainer.append(wrapped);
    }
  }

  return diagnostics.length === 0 ? {} : { diagnostics };
}
