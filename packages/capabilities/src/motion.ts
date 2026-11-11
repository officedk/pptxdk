import {
  CapabilityError,
  pxToEmu,
  xmlEscape,
  type CapabilityInput,
  type GeometryBox,
  type SlideCapabilityModule,
  type TransparentCapabilityModule,
} from "@react-pptx/core";

import type {
  MotionDirection,
  MotionEffectObservation,
  MotionNativeProjection,
  MotionObservation,
  MotionStartDependency,
} from "./motion-api.ts";
import {
  maximumTimeNodeValue,
  motionScopeAnnotation,
  motionStepAnnotation,
} from "./motion-api.ts";
import {
  addMotionPlanStartDelay,
  composeMotionPlanStarts,
  compileMotionPlan,
  motionPlanStartLayers,
  type MotionPlanConditionPlacement,
  type MotionPlanNode,
  type MotionPlanStart,
  type MotionPlanStartGate,
  type MotionPlanStep,
} from "./motion-plan.ts";
import { registerMotionObservation } from "./motion-query.ts";

export {
  after,
  breathe,
  fillColor,
  fade,
  fly,
  floatDown,
  floatUp,
  repeat,
  spin,
  spinIn,
  stretch,
  wipe,
  zoom,
} from "./motion-api.ts";
export { Box, MotionScope, Step, useMotionScope } from "./motion-elements.ts";
export type { BoxProps } from "./motion-elements.ts";
export { observeMotion } from "./motion-query.ts";
export type {
  EffectInstanceObservation,
  MotionCueObservation,
  MotionCueRef,
  MotionDecision,
  MotionDependencyObservation,
  MotionDifference,
  MotionEffectRef,
  MotionIntervalRef,
  MotionMembership,
  MotionMomentObservation,
  MotionMomentRef,
  MotionQuery,
  MotionScopeQueryMember,
  MotionScopeQueryObservation,
  MotionScopeRef,
  MotionStepQuery,
  MotionStepRef,
  MotionTime,
} from "./motion-query.ts";
export type {
  Easing,
  Effect,
  EffectOptions,
  EffectPreset,
  BreatheEffect,
  BreatheOptions,
  FillColorOptions,
  EmphasisEffect,
  EmphasisEffectDeclaration,
  EntranceEffect,
  EntranceEffectDeclaration,
  MotionCompletionState,
  MotionDirection,
  MotionGranularity,
  MotionEffectObservation,
  MotionNativeProjection,
  MotionMoment,
  MotionMomentOffset,
  MotionObservation,
  MotionScopeHandle,
  MotionScopeMode,
  MotionScopeObservation,
  MotionScopeProps,
  MotionProps,
  MotionStartDependency,
  MotionTargetRange,
  MotionTiming,
  MotionStepObservation,
  RepeatCount,
  RepeatedEffect,
  SpinOptions,
  StepProps,
} from "./motion-api.ts";

type XmlAttributes = Record<string, string | number | boolean | undefined>;

function invalidMotionTime(field: string): never {
  throw new CapabilityError(
    "invalid-motion-time",
    `${field} must be an exact non-negative integer in milliseconds, and must not exceed ST_TLTime range when written`,
  );
}

// Total accumulated duration can be expressed by repeat; only fields actually projected to ST_TLTime are constrained by its upper limit.
function validateKnownTime(value: number | undefined, field: string): void {
  if (value !== undefined && (!Number.isSafeInteger(value) || value < 0)) {
    invalidMotionTime(field);
  }
}

function timeNodeValue(value: number, field: string): number {
  validateKnownTime(value, field);
  if (value > maximumTimeNodeValue) invalidMotionTime(field);
  return value;
}

function validateKnownMotionTimes(observation: MotionObservation): void {
  for (const scope of observation.scopes ?? []) {
    validateKnownTime(scope.delayMs, "orchestration scope delay");
    validateKnownTime(scope.eachMs, "stagger interval");
    validateKnownTime(scope.startMs, "orchestration scope start time");
    validateKnownTime(scope.endMs, "orchestration scope completion time");
    validateKnownTime(
      scope.startAfter?.offsetMs,
      "scope completion dependency offset",
    );
    for (const dependency of scope.startAfterAll ?? []) {
      validateKnownTime(dependency.offsetMs, "scope start constraint offset");
    }
  }
  for (const step of observation.steps) {
    for (const effect of step.effects) {
      validateKnownTime(effect.durationMs, "animation duration");
      validateKnownTime(effect.delayMs, "animation delay");
      validateKnownTime(effect.intervalMs, "animation iteration interval");
      validateKnownTime(effect.startMs, "animation start time");
      validateKnownTime(effect.endMs, "animation completion time");
      validateKnownTime(
        effect.startAfter?.offsetMs,
        "animation completion dependency offset",
      );
      for (const dependency of effect.startAfterAll ?? []) {
        validateKnownTime(
          dependency.offsetMs,
          "animation start constraint offset",
        );
      }
    }
  }
}

function element(
  name: string,
  attributes: XmlAttributes | undefined,
  ...children: readonly (string | undefined)[]
): string {
  const attrs = Object.entries(attributes ?? {})
    .filter(([, value]) => value !== undefined)
    .map(([key, value]) => {
      const normalized =
        typeof value === "boolean" ? (value ? "1" : "0") : value;
      return ` ${key}="${xmlEscape(String(normalized))}"`;
    })
    .join("");
  return `<${name}${attrs}>${children.filter(Boolean).join("")}</${name}>`;
}

class TimingIds {
  #next = 1;

  take(): number {
    return this.#next++;
  }
}

type TimingEffect = MotionEffectObservation & {
  readonly effectIndex: number;
  readonly partIndex?: number;
  readonly targetTextBody: boolean;
  readonly targetFilled: boolean;
  readonly targetBox?: GeometryBox;
  readonly targetRotationDegrees?: number;
  readonly targetHasBackground: boolean;
  readonly targetParagraphCount?: number;
  readonly targetMembers?: readonly MotionNativeMember[];
  readonly buildGroupId?: number;
  readonly groupSpinCenter?: { readonly x: number; readonly y: number };
};

type MotionNativeMember = {
  readonly id: number;
  readonly kind: string;
  readonly name: string;
  readonly testId?: string;
  readonly box?: GeometryBox;
  readonly members?: readonly MotionNativeMember[];
};

type MotionTimingWrite = {
  readonly xml: string;
  readonly nativeProjections: readonly MotionNativeProjection[];
};

function timingTarget(
  effect: TimingEffect,
  range: MotionEffectObservation["target"]["range"] = effect.target.range,
  background = false,
): string {
  return element(
    "p:tgtEl",
    undefined,
    element(
      "p:spTgt",
      { spid: effect.target.id },
      background ? element("p:bg", undefined) : undefined,
      range?.type === "paragraph"
        ? element(
            "p:txEl",
            undefined,
            element("p:pRg", { st: range.start, end: range.end }),
          )
        : undefined,
      range?.type === "chart"
        ? element(
            "p:graphicEl",
            undefined,
            element("a:chart", {
              seriesIdx: range.seriesIndex,
              categoryIdx: range.categoryIndex,
              bldStep: range.step,
            }),
          )
        : undefined,
    ),
  );
}

function attributeNames(value: string): string {
  return element(
    "p:attrNameLst",
    undefined,
    element("p:attrName", undefined, value),
  );
}

function visibilitySet(
  targetElement: string,
  visibility: "visible" | "hidden",
  ids: TimingIds,
): string {
  return element(
    "p:set",
    undefined,
    element(
      "p:cBhvr",
      undefined,
      element(
        "p:cTn",
        { id: ids.take(), dur: 1, fill: "hold" },
        element("p:stCondLst", undefined, element("p:cond", { delay: 0 })),
      ),
      targetElement,
      attributeNames("style.visibility"),
    ),
    element("p:to", undefined, element("p:strVal", { val: visibility })),
  );
}

function direction(effect: MotionEffectObservation): MotionDirection {
  return typeof effect.from === "string" ? effect.from : "bottom";
}

function wholeTextRanges(effect: TimingEffect): readonly {
  readonly type: "paragraph";
  readonly start: number;
  readonly end: number;
}[] {
  if (
    !effect.targetTextBody ||
    effect.target.range !== undefined ||
    effect.preset === "breathe" ||
    effect.preset === "fillColor" ||
    effect.by === "word" ||
    effect.by === "letter" ||
    effect.targetParagraphCount === undefined ||
    effect.targetParagraphCount < 1
  ) {
    return [];
  }
  return Array.from({ length: effect.targetParagraphCount }, (_, index) => ({
    type: "paragraph" as const,
    start: index,
    end: index,
  }));
}

function isWholeTextProjection(effect: TimingEffect): boolean {
  return wholeTextRanges(effect).length > 0;
}

function hasWholeTextBackgroundBehaviors(effect: TimingEffect): boolean {
  return isWholeTextProjection(effect) && effect.targetHasBackground;
}

function formulaNumber(value: number): string {
  return Number(value.toPrecision(15)).toString();
}

function formulaOffset(value: string, offset: number): string {
  if (offset === 0) return value;
  return `${value}${offset < 0 ? "-" : "+"}${formulaNumber(Math.abs(offset))}`;
}

function geometryCenter(
  box: GeometryBox,
  dimensions: { readonly width: number; readonly height: number },
): { readonly x: number; readonly y: number } {
  if (dimensions.width <= 0 || dimensions.height <= 0) {
    throw new CapabilityError(
      "invalid-motion-page-size",
      "Slide dimensions must be greater than 0",
    );
  }
  return {
    x: (box.x + box.width / 2) / dimensions.width,
    y: (box.y + box.height / 2) / dimensions.height,
  };
}

/** Common rotation center is calculated using integer EMUs from serialized xfrm and sldSz to avoid differing rounding results on readback. */
function serializedGeometryCenter(
  box: GeometryBox,
  dimensions: { readonly width: number; readonly height: number },
): { readonly x: number; readonly y: number } {
  const widthEmu = pxToEmu(dimensions.width);
  const heightEmu = pxToEmu(dimensions.height);
  if (widthEmu <= 0 || heightEmu <= 0) {
    throw new CapabilityError(
      "invalid-motion-page-size",
      "Slide dimensions must be greater than 0",
    );
  }
  return {
    x: (pxToEmu(box.x) + pxToEmu(box.width) / 2) / widthEmu,
    y: (pxToEmu(box.y) + pxToEmu(box.height) / 2) / heightEmu,
  };
}

function targetBox(effect: TimingEffect, purpose: string): GeometryBox {
  if (effect.targetBox === undefined) {
    throw new CapabilityError(
      "motion-target-geometry",
      `${purpose} requires native object bounds obtained from the same layout`,
    );
  }
  return effect.targetBox;
}

function rotationFormula(effect: TimingEffect): {
  readonly angle: string;
  readonly degrees: number;
} {
  const relativeBy = typeof effect.by === "number";
  const degrees = relativeBy
    ? effect.by
    : typeof effect.from === "number"
      ? effect.from
      : 360;
  return {
    angle: `(${formulaNumber(degrees)}*pi/180)*${relativeBy ? "$" : "(1-$)"}`,
    degrees,
  };
}

function centerRotationDelta(
  effect: TimingEffect,
  center: { readonly x: number; readonly y: number },
  dimensions: { readonly width: number; readonly height: number },
  scaleToCenter = false,
): readonly [string, string] {
  const widthEmu = pxToEmu(dimensions.width);
  const heightEmu = pxToEmu(dimensions.height);
  const ratioX = formulaNumber(heightEmu / widthEmu);
  const ratioY = formulaNumber(widthEmu / heightEmu);
  const centerX = formulaNumber(center.x);
  const centerY = formulaNumber(center.y);
  const { angle } = rotationFormula(effect);
  const deltaX = scaleToCenter
    ? `${centerX}+$*((#ppt_x-${centerX})*cos(${angle})-(#ppt_y-${centerY})*${ratioX}*sin(${angle}))-#ppt_x`
    : `${centerX}+((#ppt_x-${centerX})*cos(${angle})-(#ppt_y-${centerY})*${ratioX}*sin(${angle}))-#ppt_x`;
  const deltaY = scaleToCenter
    ? `${centerY}+$*((#ppt_x-${centerX})*${ratioY}*sin(${angle})+(#ppt_y-${centerY})*cos(${angle}))-#ppt_y`
    : `${centerY}+((#ppt_x-${centerX})*${ratioY}*sin(${angle})+(#ppt_y-${centerY})*cos(${angle}))-#ppt_y`;
  return [deltaX, deltaY];
}

function flyOffset(
  effect: TimingEffect,
  dimensions: { readonly width: number; readonly height: number },
): { readonly x: number; readonly y: number } {
  const box = effect.targetBox;
  if (box === undefined) {
    throw new CapabilityError(
      "motion-target-geometry",
      "Whole-text entrance requires native object bounds obtained from the same layout",
    );
  }
  const value = direction(effect);
  switch (value) {
    case "left":
      return { x: -(box.x + box.width) / dimensions.width, y: 0 };
    case "right":
      return { x: (dimensions.width - box.x) / dimensions.width, y: 0 };
    case "top":
      return { x: 0, y: -(box.y + box.height) / dimensions.height };
    case "bottom":
      return { x: 0, y: (dimensions.height - box.y) / dimensions.height };
    default: {
      const exhaustiveCheck: never = value;
      throw new Error(
        `Unsupported direction: ${JSON.stringify(exhaustiveCheck)}`,
      );
    }
  }
}

function behaviorIds(
  effect: TimingEffect,
  ids: TimingIds,
  dimensions: { readonly width: number; readonly height: number },
): readonly string[] {
  const ranges = wholeTextRanges(effect);
  const wholeText = ranges.length > 0;
  const target = () => timingTarget(effect);
  const paragraphTargets = () =>
    ranges.map((range) => timingTarget(effect, range));
  const backgroundTarget = () => timingTarget(effect, undefined, true);
  const duration = effect.durationMs;
  validateKnownTime(duration, "animation duration");
  if (effect.preset === "breathe") {
    timeNodeValue(Math.round(duration / 2), "breathe half-period");
  } else {
    timeNodeValue(duration, "animation behavior duration");
  }
  if (effect.repeatCount !== undefined && effect.repeatCount !== "forever") {
    timeNodeValue(effect.repeatCount * 1000, "animation repeat count");
  }
  if (effect.intervalMs !== undefined) {
    timeNodeValue(effect.intervalMs, "animation iteration interval");
  }
  const show = (targetElement: string) =>
    visibilitySet(targetElement, "visible", ids);
  const filter = (name: string, targetElement: string) =>
    element(
      "p:animEffect",
      { transition: "in", filter: name },
      element(
        "p:cBhvr",
        undefined,
        element("p:cTn", { id: ids.take(), dur: duration }),
        targetElement,
      ),
    );
  const value = (item: string | number) =>
    typeof item === "number"
      ? element("p:fltVal", { val: item })
      : element("p:strVal", { val: item });
  const motion = (
    name: string,
    from: string | number,
    to: string | number,
    additive: boolean | "sum",
    targetElement: string,
    progressFormula?: string,
  ) =>
    element(
      "p:anim",
      { calcmode: "lin", valueType: "num" },
      element(
        "p:cBhvr",
        { additive: additive === true ? "base" : additive || undefined },
        element("p:cTn", { id: ids.take(), dur: duration, fill: "hold" }),
        targetElement,
        attributeNames(name),
      ),
      element(
        "p:tavLst",
        undefined,
        element(
          "p:tav",
          {
            tm: 0,
            ...(progressFormula === undefined ? {} : { fmla: progressFormula }),
          },
          element("p:val", undefined, value(from)),
        ),
        element(
          "p:tav",
          { tm: 100000 },
          element("p:val", undefined, value(to)),
        ),
      ),
    );

  const singleMotion = (
    name: string,
    from: string | number,
    to: string | number,
    additive: boolean | "sum",
    targetElement = target(),
  ) => motion(name, from, to, additive, targetElement);
  // Common center displacement is already a full progress function; the formula receives 0->1 numeric progress to avoid function values being re-interpolated as endpoints.
  const progressMotion = (
    name: string,
    formula: string,
    targetElement = target(),
  ) => motion(name, 0, 1, "sum", targetElement, formula);
  const rotation = (targetElement: string, additive: boolean | "sum" = false) =>
    element(
      "p:animRot",
      typeof effect.by === "number"
        ? { by: Math.round(effect.by * 60000) }
        : {
            from: Math.round(
              (typeof effect.from === "number" ? effect.from : 360) * 60000,
            ),
            to: 0,
          },
      element(
        "p:cBhvr",
        {
          additive:
            additive === false
              ? undefined
              : additive === true
                ? "base"
                : additive,
        },
        element("p:cTn", { id: ids.take(), dur: duration, fill: "hold" }),
        targetElement,
        attributeNames("r"),
      ),
    );
  // Paragraph building uses real pRg and independent background entrance, preventing whole-object visible target from prematurely unlocking subsequent paragraphs.
  const showOnce = () =>
    wholeText && effect.kind === "entrance"
      ? [...paragraphTargets(), ...backgroundTargets()]
          .map((targetElement) => show(targetElement))
          .join("")
      : show(wholeText ? timingTarget(effect, undefined) : target());
  const filterTargets = () => (wholeText ? paragraphTargets() : [target()]);
  const positionTargets = () => (wholeText ? paragraphTargets() : [target()]);
  const backgroundTargets = () =>
    wholeText && effect.targetHasBackground ? [backgroundTarget()] : [];

  switch (effect.preset) {
    case "fade":
      return [
        showOnce(),
        ...filterTargets().map((targetElement) =>
          filter("fade", targetElement),
        ),
        ...backgroundTargets().map((targetElement) =>
          filter("fade", targetElement),
        ),
      ];
    case "wipe": {
      const filters: Record<MotionDirection, string> = {
        top: "wipe(up)",
        right: "wipe(right)",
        bottom: "wipe(down)",
        left: "wipe(left)",
      };
      return [
        showOnce(),
        ...filterTargets().map((targetElement) =>
          filter(filters[direction(effect)], targetElement),
        ),
        ...backgroundTargets().map((targetElement) =>
          filter(filters[direction(effect)], targetElement),
        ),
      ];
    }
    case "zoom": {
      const targets = positionTargets();
      if (!wholeText) {
        return [
          showOnce(),
          singleMotion("ppt_w", 0, "#ppt_w", false),
          singleMotion("ppt_h", 0, "#ppt_h", false),
          filter("fade", target()),
        ];
      }
      const center = geometryCenter(
        targetBox(effect, "whole-text zoom"),
        dimensions,
      );
      const geometry = targets.flatMap((targetElement) => [
        singleMotion(
          "ppt_x",
          `${formulaNumber(center.x)}-#ppt_x`,
          0,
          "sum",
          targetElement,
        ),
        singleMotion(
          "ppt_y",
          `${formulaNumber(center.y)}-#ppt_y`,
          0,
          "sum",
          targetElement,
        ),
        singleMotion("ppt_w", 0, "#ppt_w", false, targetElement),
        singleMotion("ppt_h", 0, "#ppt_h", false, targetElement),
      ]);
      const backgroundGeometry = backgroundTargets().flatMap(
        (targetElement) => [
          singleMotion(
            "ppt_x",
            `${formulaNumber(center.x)}-#ppt_x`,
            0,
            "sum",
            targetElement,
          ),
          singleMotion(
            "ppt_y",
            `${formulaNumber(center.y)}-#ppt_y`,
            0,
            "sum",
            targetElement,
          ),
          singleMotion("ppt_w", 0, "#ppt_w", false, targetElement),
          singleMotion("ppt_h", 0, "#ppt_h", false, targetElement),
        ],
      );
      return [
        showOnce(),
        ...geometry,
        ...backgroundGeometry,
        ...filterTargets().map((targetElement) =>
          filter("fade", targetElement),
        ),
        ...backgroundTargets().map((targetElement) =>
          filter("fade", targetElement),
        ),
      ];
    }
    case "fly": {
      const flyFrom: Record<MotionDirection, readonly [string, string]> = {
        left: ["0-#ppt_w/2", "#ppt_y"],
        right: ["1+#ppt_w/2", "#ppt_y"],
        top: ["#ppt_x", "0-#ppt_h/2"],
        bottom: ["#ppt_x", "1+#ppt_h/2"],
      };
      if (!wholeText) {
        const [x, y] = flyFrom[direction(effect)];
        return [
          showOnce(),
          singleMotion("ppt_x", x, "#ppt_x", true),
          singleMotion("ppt_y", y, "#ppt_y", true),
        ];
      }
      const offset = flyOffset(effect, dimensions);
      const geometries = [...positionTargets(), ...backgroundTargets()].flatMap(
        (targetElement) => [
          singleMotion(
            "ppt_x",
            formulaOffset("#ppt_x", offset.x),
            "#ppt_x",
            true,
            targetElement,
          ),
          singleMotion(
            "ppt_y",
            formulaOffset("#ppt_y", offset.y),
            "#ppt_y",
            true,
            targetElement,
          ),
        ],
      );
      return [showOnce(), ...geometries];
    }
    case "spinIn": {
      if (!wholeText) {
        return [
          showOnce(),
          singleMotion("ppt_w", 0, "#ppt_w", false),
          singleMotion("ppt_h", 0, "#ppt_h", false),
          singleMotion("style.rotation", 360, 0, false),
          filter("fade", target()),
        ];
      }
      const box = targetBox(effect, "whole-text spin-in");
      const center = serializedGeometryCenter(box, dimensions);
      const position = centerRotationDelta(effect, center, dimensions, true);
      const geometry = positionTargets().flatMap((targetElement) => [
        singleMotion("ppt_w", 0, "#ppt_w", false, targetElement),
        singleMotion("ppt_h", 0, "#ppt_h", false, targetElement),
        // Paragraph rotation is expressed relative to text box original angle; object angle must not be added to endpoints again.
        singleMotion("style.rotation", 360, 0, false, targetElement),
        progressMotion("ppt_x", position[0], targetElement),
        progressMotion("ppt_y", position[1], targetElement),
      ]);
      return [
        showOnce(),
        ...geometry,
        ...backgroundTargets().flatMap((targetElement) => [
          progressMotion("ppt_x", position[0], targetElement),
          progressMotion("ppt_y", position[1], targetElement),
          singleMotion("ppt_w", 0, "#ppt_w", false, targetElement),
          singleMotion("ppt_h", 0, "#ppt_h", false, targetElement),
          rotation(targetElement, "sum"),
          filter("fade", targetElement),
        ]),
        ...filterTargets().map((targetElement) =>
          filter("fade", targetElement),
        ),
      ];
    }
    case "stretch": {
      if (!wholeText) {
        return [
          showOnce(),
          singleMotion("ppt_w", 0, "#ppt_w", false),
          singleMotion("ppt_h", "#ppt_h", "#ppt_h", false),
        ];
      }
      const center = geometryCenter(
        targetBox(effect, "whole-text stretch"),
        dimensions,
      );
      return [
        showOnce(),
        ...[...positionTargets(), ...backgroundTargets()].flatMap(
          (targetElement) => [
            singleMotion(
              "ppt_x",
              `${formulaNumber(center.x)}-#ppt_x`,
              0,
              "sum",
              targetElement,
            ),
            singleMotion("ppt_w", 0, "#ppt_w", false, targetElement),
            singleMotion("ppt_h", "#ppt_h", "#ppt_h", false, targetElement),
          ],
        ),
      ];
    }
    case "floatUp":
    case "floatDown": {
      if (!wholeText) {
        return [
          showOnce(),
          filter("fade", target()),
          singleMotion("ppt_x", "#ppt_x", "#ppt_x", false),
          singleMotion(
            "ppt_y",
            effect.preset === "floatDown" ? "#ppt_y-.1" : "#ppt_y+.1",
            "#ppt_y",
            false,
          ),
        ];
      }
      const fromY = effect.preset === "floatDown" ? -0.1 : 0.1;
      const targets = [...positionTargets(), ...backgroundTargets()];
      return [
        showOnce(),
        ...filterTargets().map((targetElement) =>
          filter("fade", targetElement),
        ),
        ...backgroundTargets().map((targetElement) =>
          filter("fade", targetElement),
        ),
        ...targets.flatMap((targetElement) => [
          singleMotion("ppt_x", 0, 0, "sum", targetElement),
          singleMotion("ppt_y", fromY, 0, "sum", targetElement),
        ]),
      ];
    }
    case "spin": {
      if (!wholeText && effect.groupSpinCenter === undefined) {
        return [
          rotation(target(), typeof effect.by === "number" ? "sum" : false),
        ];
      }
      const center =
        effect.groupSpinCenter ??
        serializedGeometryCenter(
          targetBox(effect, "whole-text spin"),
          dimensions,
        );
      const position = centerRotationDelta(effect, center, dimensions);
      if (wholeText) {
        return [
          ...positionTargets().flatMap((targetElement) => [
            rotation(targetElement, "sum"),
            progressMotion("ppt_x", position[0], targetElement),
            progressMotion("ppt_y", position[1], targetElement),
          ]),
          ...backgroundTargets().flatMap((targetElement) => [
            rotation(targetElement, "sum"),
            progressMotion("ppt_x", position[0], targetElement),
            progressMotion("ppt_y", position[1], targetElement),
          ]),
        ];
      }
      return [
        rotation(target(), "sum"),
        progressMotion("ppt_x", position[0]),
        progressMotion("ppt_y", position[1]),
      ];
    }
    case "fillColor":
    case "breathe":
      return [
        element(
          "p:animClr",
          { clrSpc: "rgb" },
          element(
            "p:cBhvr",
            undefined,
            element("p:cTn", {
              id: ids.take(),
              dur:
                effect.preset === "breathe"
                  ? Math.round(duration / 2)
                  : duration,
              fill: "hold",
            }),
            target(),
            attributeNames("fillcolor"),
          ),
          element(
            "p:from",
            undefined,
            element("a:srgbClr", {
              val: effect.fromColor?.slice(1),
            }),
          ),
          element(
            "p:to",
            undefined,
            element("a:srgbClr", {
              val: effect.toColor?.slice(1),
            }),
          ),
        ),
      ];
    default: {
      const exhaustiveCheck: never = effect.preset;
      throw new Error(`Unsupported preset: ${JSON.stringify(exhaustiveCheck)}`);
    }
  }
}

function effectTree(
  effect: TimingEffect,
  nodeType: string,
  startCondition: string,
  ids: TimingIds,
  effectId: number,
  buildGroupId: number | undefined,
  dimensions: { readonly width: number; readonly height: number },
): string {
  if (effect.startAfter !== undefined) {
    validateKnownTime(
      effect.startAfter.offsetMs,
      "completion dependency offset",
    );
  }
  const preset = {
    fade: { id: 10, class: "entr" },
    wipe: { id: 22, class: "entr" },
    zoom: { id: 53, class: "entr" },
    fly: { id: 2, class: "entr" },
    floatUp: { id: 42, class: "entr" },
    floatDown: { id: 47, class: "entr" },
    spinIn: { id: 49, class: "entr" },
    stretch: { id: 17, class: "entr" },
    spin: { id: 8, class: "emph" },
    breathe: { id: 0, class: "emph" },
    fillColor: { id: 9, class: "emph" },
  }[effect.preset];
  const subtype =
    effect.preset === "zoom"
      ? 16
      : effect.preset === "stretch"
        ? 10
        : effect.preset === "wipe" || effect.preset === "fly"
          ? { top: 1, right: 2, bottom: 4, left: 8 }[direction(effect)]
          : 0;
  const easing = effect.easing;
  const properties: XmlAttributes = {
    id: effectId,
    presetID: preset.id,
    presetClass: preset.class,
    presetSubtype: subtype,
    repeatCount:
      effect.repeatCount === "forever"
        ? "indefinite"
        : effect.repeatCount === undefined
          ? undefined
          : effect.repeatCount * 1000,
    autoRev: effect.preset === "breathe" ? true : undefined,
    accel:
      easing?.accelerate === undefined || easing.accelerate === 0
        ? undefined
        : Math.round(easing.accelerate * 100000),
    decel:
      easing?.decelerate === undefined || easing.decelerate === 0
        ? undefined
        : Math.round(easing.decelerate * 100000),
    fill: "hold",
    grpId: effect.buildGroupId ?? (effect.targetTextBody ? 0 : buildGroupId),
    nodeType,
  };
  return element(
    "p:par",
    undefined,
    element(
      "p:cTn",
      properties,
      element("p:stCondLst", undefined, startCondition),
      effect.by !== "word" && effect.by !== "letter"
        ? undefined
        : element(
            "p:iterate",
            { type: effect.by === "word" ? "wd" : "lt" },
            // Single dur only describes current word or letter; cannot masquerade as natural completion time of the entire paragraph.
            element("p:tmAbs", { val: effect.intervalMs }),
          ),
      element(
        "p:childTnLst",
        undefined,
        ...behaviorIds(effect, ids, dimensions),
      ),
    ),
  );
}

function effectKey(effect: TimingEffect): string {
  return `${effect.target.id}:${effect.effectIndex}:${effect.partIndex ?? ""}`;
}

function addNativeStartDelay(
  start: MotionPlanStart,
  delayMs: number,
): MotionPlanStart | undefined {
  const delayed = addMotionPlanStartDelay(start, delayMs);
  if (delayed === undefined) return undefined;
  const { base, delays } = motionPlanStartLayers(delayed);
  const baseValue =
    base.kind === "parent-start"
      ? base.delayMs
      : base.kind === "end" || base.kind === "event"
        ? base.dependency.offsetMs
        : base.kind === "condition-end"
          ? base.delayMs
          : undefined;
  return [baseValue, ...delays].every(
    (value) =>
      value === undefined ||
      (Number.isSafeInteger(value) &&
        value >= 0 &&
        value <= maximumTimeNodeValue),
  )
    ? delayed
    : undefined;
}

function dependencyTimingId(
  dependency: MotionStartDependency,
  stepKey: string,
  effectNodeIds: ReadonlyMap<string, number>,
  scopeNodeIds: ReadonlyMap<string, number>,
  targetGroupIds: ReadonlyMap<string, number>,
): number | undefined {
  if ("scopeId" in dependency) {
    return scopeNodeIds.get(
      `${stepKey}:${dependency.point}:${dependency.scopeId}`,
    );
  }
  if ("group" in dependency) {
    return targetGroupIds.get(`${stepKey}:${dependency.targetId}`);
  }
  return effectNodeIds.get(
    `${dependency.targetId}:${dependency.effectIndex}:${dependency.partIndex ?? ""}`,
  );
}

function nativeStartCondition(
  start: MotionPlanStart,
  stepKey: string,
  effectNodeIds: ReadonlyMap<string, number>,
  scopeNodeIds: ReadonlyMap<string, number>,
  targetGroupIds: ReadonlyMap<string, number>,
  conditionNodeIds: ReadonlyMap<string, number>,
): string {
  if (start.kind === "parent-start") {
    return element("p:cond", {
      delay: timeNodeValue(start.delayMs, "animation start delay"),
    });
  }
  if (start.kind === "condition-join") {
    throw new Error(
      "MotionPlan condition join must be emitted before its consumers",
    );
  }
  if (start.kind === "delay-layer") {
    throw new Error(
      "MotionPlan delay layer must be emitted through native time container",
    );
  }
  if (start.kind === "condition-end") {
    const conditionId = conditionNodeIds.get(`${stepKey}:${start.id}`);
    if (conditionId === undefined) {
      throw new Error(
        "MotionPlan condition join is missing native time node identity",
      );
    }
    return element(
      "p:cond",
      {
        evt: "end",
        delay: timeNodeValue(start.delayMs, "condition join completion offset"),
      },
      element("p:tn", { val: conditionId }),
    );
  }
  const dependencyId = dependencyTimingId(
    start.dependency,
    stepKey,
    effectNodeIds,
    scopeNodeIds,
    targetGroupIds,
  );
  if (dependencyId === undefined) {
    throw new Error(
      "MotionPlan completion dependency lacks corresponding native timing node",
    );
  }
  return element(
    "p:cond",
    {
      evt:
        start.kind === "event" || start.dependency.point === "start"
          ? "begin"
          : "end",
      delay: timeNodeValue(
        start.dependency.offsetMs,
        "completion dependency offset",
      ),
    },
    element("p:tn", { val: dependencyId }),
  );
}

function endSync(): string {
  // Container waits for all actual child nodes to end naturally; aggregate duration value is not estimated from children.
  return element(
    "p:endSync",
    { evt: "end", delay: 0 },
    element("p:rtn", { val: "all" }),
  );
}

function timingContainer(
  id: number,
  startCondition: string,
  children: readonly string[],
  nodeType?: string,
): string {
  const isEmpty = children.length === 0;
  return element(
    "p:par",
    undefined,
    element(
      "p:cTn",
      { id, ...(isEmpty ? { dur: 0 } : {}), fill: "hold", nodeType },
      element("p:stCondLst", undefined, startCondition),
      isEmpty ? undefined : endSync(),
      // An empty scope's END matches its unknown START; omit empty list to maintain valid XML.
      isEmpty ? undefined : element("p:childTnLst", undefined, ...children),
    ),
  );
}

function conditionJoinTree(
  start: Extract<MotionPlanStart, { readonly kind: "condition-join" }>,
  stepKey: string,
  ids: TimingIds,
  effectNodeIds: ReadonlyMap<string, number>,
  scopeNodeIds: ReadonlyMap<string, number>,
  targetGroupIds: ReadonlyMap<string, number>,
  conditionNodeIds: ReadonlyMap<string, number>,
): string {
  const id = conditionNodeIds.get(`${stepKey}:${start.id}`);
  if (id === undefined) {
    throw new Error(
      "MotionPlan condition join is missing native time node identity",
    );
  }
  return endCollectorTree(
    id,
    element("p:cond", { delay: 0 }),
    start.gates,
    stepKey,
    ids,
    effectNodeIds,
    scopeNodeIds,
    targetGroupIds,
    conditionNodeIds,
  );
}

function gateStart(gate: MotionPlanStartGate): MotionPlanStart {
  switch (gate.kind) {
    case "parent-start":
      return { kind: "parent-start", delayMs: gate.delayMs };
    case "end":
      return { kind: "end", dependency: gate.dependency };
    case "begin":
      return { kind: "event", dependency: gate.dependency };
    case "plan-start":
      return gate.start;
    default: {
      const exhaustiveCheck: never = gate;
      throw new Error(`Unsupported gate: ${JSON.stringify(exhaustiveCheck)}`);
    }
  }
}

function endCollectorTree(
  id: number,
  startCondition: string,
  gates: readonly MotionPlanStartGate[],
  stepKey: string,
  ids: TimingIds,
  effectNodeIds: ReadonlyMap<string, number>,
  scopeNodeIds: ReadonlyMap<string, number>,
  targetGroupIds: ReadonlyMap<string, number>,
  conditionNodeIds: ReadonlyMap<string, number>,
  nodeType?: string,
): string {
  const gateTree = (gate: MotionPlanStartGate): string => {
    const { base, delays } = motionPlanStartLayers(gateStart(gate));
    if (base.kind === "condition-join") {
      throw new Error(
        "MotionPlan native condition gate cannot embed undefined condition joins",
      );
    }
    const container = (
      start: MotionPlanStart,
      children: readonly string[] = [],
    ) =>
      timingContainer(
        ids.take(),
        nativeStartCondition(
          start,
          stepKey,
          effectNodeIds,
          scopeNodeIds,
          targetGroupIds,
          conditionNodeIds,
        ),
        children,
      );
    const lastDelay = delays.at(-1);
    let children = [
      container(
        lastDelay === undefined
          ? base
          : { kind: "parent-start", delayMs: lastDelay },
      ),
    ];
    for (let layerIndex = delays.length - 2; layerIndex >= 0; layerIndex -= 1) {
      const delayMs = delays[layerIndex];
      if (delayMs !== undefined) {
        children = [container({ kind: "parent-start", delayMs }, children)];
      }
    }
    return delays.length === 0 ? children[0]! : container(base, children);
  };
  const leaves = gates.map(gateTree);
  return timingContainer(id, startCondition, leaves, nodeType);
}

function mainSequenceWithGroup(
  id: number,
  children: readonly string[],
): string {
  // mainSeq standard step->with-group->entry-node hierarchy allows native import to identify step entry roles.
  // Group starts immediately from step entry without duplicating plan start, nor creating Shape or completion sources.
  return element(
    "p:par",
    undefined,
    element(
      "p:cTn",
      { id, fill: "hold" },
      element("p:stCondLst", undefined, element("p:cond", { delay: 0 })),
      element("p:childTnLst", undefined, ...children),
    ),
  );
}

function timingXml(
  steps: readonly MotionPlanStep[],
  dimensions: { readonly width: number; readonly height: number },
): MotionTimingWrite {
  const ids = new TimingIds();
  const rootId = ids.take();
  const sequenceId = ids.take();
  const effectNodeIds = new Map<string, number>();
  const scopeNodeIds = new Map<string, number>();
  const targetGroupIds = new Map<string, number>();
  const conditionNodeIds = new Map<string, number>();
  const referencedScopeNodes = new Set<string>();
  const referencedScopeStartNodes = new Set<string>();
  const referencedTargetGroups = new Set<string>();
  const keepDependencyContainer = (
    dependency: MotionStartDependency,
    stepKey: string,
  ): void => {
    if ("scopeId" in dependency) {
      const key = `${stepKey}:${dependency.point}:${dependency.scopeId}`;
      if (dependency.point === "start") {
        referencedScopeStartNodes.add(key);
      } else {
        referencedScopeNodes.add(key);
      }
    } else if ("group" in dependency) {
      referencedTargetGroups.add(`${stepKey}:${dependency.targetId}`);
    }
  };
  const collectStartReferences = (
    start: MotionPlanStart,
    stepKey: string,
  ): void => {
    if (start.kind === "delay-layer") {
      collectStartReferences(start.start, stepKey);
    } else if (start.kind === "end" || start.kind === "event") {
      keepDependencyContainer(start.dependency, stepKey);
    } else if (start.kind === "condition-join") {
      referencedScopeStartNodes.add(`${stepKey}:start:${start.hostScopeId}`);
      referencedScopeStartNodes.add(
        `${stepKey}:start:${start.consumerParentScopeId}`,
      );
      for (const gate of start.gates) {
        collectStartReferences(gateStart(gate), stepKey);
      }
    }
  };

  // Preserve aggregate nodes only for real end dependencies; containers without consumers can project start onto members.
  const collectReferencedContainers = (
    nodes: readonly MotionPlanNode[],
    stepKey: string,
  ): void => {
    for (const node of nodes) {
      collectStartReferences(node.nativeStart, stepKey);
      if (node.kind === "completion-collector") {
        for (const gate of node.gates) {
          collectStartReferences(gateStart(gate), stepKey);
        }
      }
      if (node.kind === "scope") {
        for (const placement of node.conditionDefinitionsBefore ?? []) {
          referencedScopeStartNodes.add(
            `${stepKey}:start:${placement.join.hostScopeId}`,
          );
          referencedScopeStartNodes.add(
            `${stepKey}:start:${placement.join.consumerParentScopeId}`,
          );
          for (const gate of placement.join.gates) {
            collectStartReferences(gateStart(gate), stepKey);
          }
        }
        collectReferencedContainers(node.children, stepKey);
      }
    }
  };
  for (const step of steps) {
    collectReferencedContainers(step.nodes, step.key);
    for (const placement of step.conditionDefinitionsBefore ?? []) {
      for (const gate of placement.join.gates) {
        collectStartReferences(gateStart(gate), step.key);
      }
    }
  }

  const registerNodeIds = (
    nodes: readonly MotionPlanNode[],
    stepKey: string,
  ): void => {
    for (const node of nodes) {
      const { base: startBase } = motionPlanStartLayers(node.nativeStart);
      if (startBase.kind === "condition-join") {
        conditionNodeIds.set(`${stepKey}:${startBase.id}`, ids.take());
      }
      if (node.kind === "scope") {
        for (const placement of node.conditionDefinitionsBefore ?? []) {
          conditionNodeIds.set(`${stepKey}:${placement.join.id}`, ids.take());
        }
        const scopeTimingId = ids.take();
        scopeNodeIds.set(`${stepKey}:start:${node.id}`, scopeTimingId);
        if (node.splitCompletion !== true) {
          scopeNodeIds.set(`${stepKey}:end:${node.id}`, scopeTimingId);
        }
        registerNodeIds(node.children, stepKey);
        continue;
      }
      if (node.kind === "completion-collector") {
        scopeNodeIds.set(`${stepKey}:end:${node.scopeId}`, ids.take());
        continue;
      }
      if (node.effects.length > 1) {
        targetGroupIds.set(`${stepKey}:${node.targetId}`, ids.take());
      }
      for (const effect of node.effects) {
        effectNodeIds.set(effectKey(effect), ids.take());
      }
    }
  };
  for (const step of steps) {
    for (const placement of step.conditionDefinitionsBefore ?? []) {
      conditionNodeIds.set(`${step.key}:${placement.join.id}`, ids.take());
    }
    registerNodeIds(step.nodes, step.key);
  }
  const chartBuildGroupIds = new Map<number, number>();
  for (const step of steps) {
    for (const effect of step.effects) {
      if (
        effect.target.range?.type === "chart" &&
        !chartBuildGroupIds.has(effect.target.id)
      ) {
        // bldGraphic grpId must be associated with an actually existing timing node of that chart.
        const id = effectNodeIds.get(effectKey(effect));
        if (id !== undefined) chartBuildGroupIds.set(effect.target.id, id);
      }
    }
  }
  const slideEffects = steps.flatMap((step) =>
    step.effects.map((effect) => effect as TimingEffect),
  );
  const nativeProjections: MotionNativeProjection[] = [];
  const nativeTextEffects: {
    readonly effect: TimingEffect;
    readonly declaration: TimingEffect;
  }[] = [];
  const collectNativeTextEffect = (
    effect: TimingEffect,
    declaration: TimingEffect,
  ): void => {
    if (effect.targetTextBody) {
      nativeTextEffects.push({ effect, declaration });
    }
  };
  const memberIds = (members: readonly MotionNativeMember[]): Set<number> => {
    const result = new Set<number>();
    const pending = [...members];
    while (pending.length > 0) {
      const member = pending.pop();
      if (member === undefined) continue;
      result.add(member.id);
      pending.push(...(member.members ?? []));
    }
    return result;
  };
  const leafMembers = (
    members: readonly MotionNativeMember[],
  ): MotionNativeMember[] =>
    members.flatMap((member) =>
      member.members === undefined || member.members.length === 0
        ? [member]
        : leafMembers(member.members),
    );
  const renderedEffect = (
    effect: TimingEffect,
    nodeType: string,
    startCondition: string,
    effectId: number,
    chartGroupId: number | undefined,
  ): string[] => {
    const members = effect.targetMembers;
    const descendantIds =
      members === undefined ? new Set<number>() : memberIds(members);
    // Children can be triggered by subsequent clicks; checking only current Step misses still-active parent Groups.
    const hasAnimatedMember = slideEffects.some(
      ({ target }) =>
        target.id !== effect.target.id && descendantIds.has(target.id),
    );
    if (
      effect.preset === "spin" &&
      members !== undefined &&
      hasAnimatedMember
    ) {
      const box = effect.targetBox;
      if (box === undefined) {
        throw new CapabilityError(
          "motion-target-geometry",
          "Native Group rotation projection requires combination bounds obtained from the same layout",
        );
      }
      const center = serializedGeometryCenter(box, dimensions);
      const projectedMembers = leafMembers(members);
      if (projectedMembers.length === 0) {
        throw new CapabilityError(
          "empty-group-projection",
          "Native Group rotation requires at least one real member target",
        );
      }
      const nativeEffects: MotionNativeProjection["nativeEffects"][number][] =
        [];
      const {
        targetMembers: _targetMembers,
        buildGroupId: _buildGroupId,
        ...groupEffect
      } = effect;
      const rendered = projectedMembers.map((member, index) => {
        const targetEffectId = index === 0 ? effectId : ids.take();
        nativeEffects.push({ targetId: member.id, timingId: targetEffectId });
        const memberEffect = slideEffects.find(
          ({ target }) => target.id === member.id,
        );
        // Text members must keep their own paragraph count and background role to continue whole-object transformation in Group projection.
        const targetTextBody = memberEffect?.targetTextBody ?? false;
        const projected: TimingEffect = {
          ...groupEffect,
          target: {
            id: member.id,
            name: member.name,
            ...(member.testId === undefined ? {} : { testId: member.testId }),
          },
          targetTextBody,
          targetFilled: memberEffect?.targetFilled ?? false,
          targetHasBackground: memberEffect?.targetHasBackground ?? false,
          ...(member.box === undefined ? {} : { targetBox: member.box }),
          ...(memberEffect?.targetRotationDegrees === undefined
            ? {}
            : { targetRotationDegrees: memberEffect.targetRotationDegrees }),
          ...(memberEffect?.targetParagraphCount === undefined
            ? {}
            : { targetParagraphCount: memberEffect.targetParagraphCount }),
          ...(targetTextBody ? { buildGroupId: 1 } : {}),
          groupSpinCenter: center,
        };
        collectNativeTextEffect(projected, effect);
        return effectTree(
          projected,
          index === 0 ? nodeType : "withEffect",
          startCondition,
          ids,
          targetEffectId,
          undefined,
          dimensions,
        );
      });
      nativeProjections.push({
        logicalTargetId: effect.target.id,
        logicalEffectIndex: effect.effectIndex,
        preset: effect.preset,
        nativeEffects,
      });
      return rendered;
    }
    const textBuildGroupId = effect.targetTextBody
      ? isWholeTextProjection(effect)
        ? 1
        : 0
      : undefined;
    const nativeEffect: TimingEffect = {
      ...effect,
      ...(textBuildGroupId === undefined
        ? {}
        : { buildGroupId: textBuildGroupId }),
    };
    collectNativeTextEffect(nativeEffect, effect);
    return [
      effectTree(
        nativeEffect,
        nodeType,
        startCondition,
        ids,
        effectId,
        chartGroupId,
        dimensions,
      ),
    ];
  };
  // Empty steps preserve query declarations only without writing empty childTnLst into native timing tree.
  const stepsWithNativeNodes = steps.filter(({ nodes }) => nodes.length > 0);
  const stepGroups = stepsWithNativeNodes.map((step) => {
    const hasConditionJoin = (start: MotionPlanStart): boolean =>
      start.kind === "condition-join" ||
      (start.kind === "delay-layer" && hasConditionJoin(start.start));
    const hasPrivateTimingDefinitions = (
      nodes: readonly MotionPlanNode[],
    ): boolean =>
      nodes.some((node) => {
        if (
          node.kind === "completion-collector" ||
          hasConditionJoin(node.nativeStart)
        ) {
          return true;
        }
        return (
          node.kind === "scope" &&
          ((node.conditionDefinitionsBefore?.length ?? 0) > 0 ||
            hasPrivateTimingDefinitions(node.children))
        );
      });
    // Condition join must be defined before first consumer, and completion collector must collect all original ENDs;
    // current writer puts each gate and consumer on the same step host. As long as this step contains such
    // private definitions, preserve the entire shared host to avoid changing verified parent nodes and definition order when promoting roots.
    const preserveStepHost =
      (step.conditionDefinitionsBefore?.length ?? 0) > 0 ||
      hasPrivateTimingDefinitions(step.nodes);
    const stepId = ids.take();
    const trigger =
      step.trigger === "click"
        ? element("p:cond", { delay: "indefinite" })
        : `${element("p:cond", { delay: "indefinite" })}${element(
            "p:cond",
            { evt: "onBegin", delay: 0 },
            element("p:tn", { val: sequenceId }),
          )}`;
    const childDelay = (effect: TimingEffect): number =>
      timeNodeValue(
        effect.delayMs + (effect.partIndex ?? 0) * effect.durationMs,
        "animated object member delay",
      );
    type RenderedRootKind =
      | "effect"
      | "scope"
      | "target-group"
      | "condition-definition"
      | "completion-definition"
      | "delay-layer";
    type RenderedRoot = {
      readonly xml: string;
      readonly kind: RenderedRootKind;
      readonly stepEntryRole: boolean;
      readonly promoteToStep: boolean;
    };
    const renderedRoot = (
      xml: string,
      kind: RenderedRootKind,
      stepEntryRole = false,
      promoteToStep = false,
    ): RenderedRoot => ({ xml, kind, stepEntryRole, promoteToStep });
    const effectRoots = (
      xml: readonly string[],
      carriesStepEntryRole: boolean,
    ): RenderedRoot[] =>
      xml.map((item, index) =>
        renderedRoot(item, "effect", index === 0 && carriesStepEntryRole),
      );
    const wrapDelayLayers = (
      start: MotionPlanStart,
      children: readonly RenderedRoot[],
      nodeType?: string,
    ): RenderedRoot[] => {
      const { base, delays } = motionPlanStartLayers(start);
      if (delays.length === 0) return [...children];
      const rendered: RenderedRoot[] = [];
      let outerStart = base;
      if (outerStart.kind === "condition-join") {
        rendered.push(
          renderedRoot(
            conditionJoinTree(
              outerStart,
              step.key,
              ids,
              effectNodeIds,
              scopeNodeIds,
              targetGroupIds,
              conditionNodeIds,
            ),
            "condition-definition",
          ),
        );
        outerStart = {
          kind: "condition-end",
          id: outerStart.id,
          delayMs: 0,
        };
      }
      let nestedChildren = children.map(({ xml }) => xml);
      for (
        let layerIndex = delays.length - 2;
        layerIndex >= 0;
        layerIndex -= 1
      ) {
        const delayMs = delays[layerIndex];
        if (delayMs === undefined) continue;
        nestedChildren = [
          timingContainer(
            ids.take(),
            nativeStartCondition(
              { kind: "parent-start", delayMs },
              step.key,
              effectNodeIds,
              scopeNodeIds,
              targetGroupIds,
              conditionNodeIds,
            ),
            nestedChildren,
          ),
        ];
      }
      rendered.push(
        renderedRoot(
          timingContainer(
            ids.take(),
            nativeStartCondition(
              outerStart,
              step.key,
              effectNodeIds,
              scopeNodeIds,
              targetGroupIds,
              conditionNodeIds,
            ),
            nestedChildren,
            nodeType,
          ),
          "delay-layer",
          nodeType !== undefined,
        ),
      );
      return rendered;
    };
    const renderNodes = (
      nodes: readonly MotionPlanNode[],
      inheritedStart: MotionPlanStart | undefined,
      isTopLevel: boolean,
      firstRootNodeType?: string,
      conditionDefinitionsBefore: readonly MotionPlanConditionPlacement[] = [],
    ): RenderedRoot[] => {
      const rendered: RenderedRoot[] = [];
      let pendingRootNodeType = firstRootNodeType;
      for (const [index, node] of nodes.entries()) {
        const nodeId =
          node.kind === "scope"
            ? `scope:${node.id}`
            : node.kind === "completion-collector"
              ? `completion:${node.scopeId}`
              : `target:${node.targetId}`;
        for (const placement of conditionDefinitionsBefore) {
          if (placement.join.beforeNodeId !== nodeId) continue;
          rendered.push(
            renderedRoot(
              conditionJoinTree(
                placement.join,
                step.key,
                ids,
                effectNodeIds,
                scopeNodeIds,
                targetGroupIds,
                conditionNodeIds,
              ),
              "condition-definition",
            ),
          );
        }
        const defaultNodeType =
          index === 0 && isTopLevel && step.trigger === "click"
            ? "clickEffect"
            : index === 0
              ? "afterEffect"
              : "withEffect";
        const nodeType = defaultNodeType;
        let effectiveStart = composeMotionPlanStarts(
          inheritedStart,
          node.nativeStart,
        );
        if (effectiveStart === undefined) {
          throw new Error(
            "Native timing projection must preserve delay layering",
          );
        }
        if (effectiveStart.kind === "delay-layer") {
          const { delays } = motionPlanStartLayers(effectiveStart);
          const lastDelay = delays.at(-1);
          if (lastDelay === undefined) {
            throw new Error(
              "MotionPlan delay layer is missing actual delay field",
            );
          }
          const delayedNode: MotionPlanNode = {
            ...node,
            nativeStart: { kind: "parent-start", delayMs: lastDelay },
          };
          const delayedRoots = wrapDelayLayers(
            effectiveStart,
            renderNodes([delayedNode], undefined, false),
            pendingRootNodeType,
          );
          rendered.push(...delayedRoots);
          if (delayedRoots.some(({ stepEntryRole }) => stepEntryRole)) {
            pendingRootNodeType = undefined;
          }
          continue;
        }
        if (effectiveStart.kind === "condition-join") {
          rendered.push(
            renderedRoot(
              conditionJoinTree(
                effectiveStart,
                step.key,
                ids,
                effectNodeIds,
                scopeNodeIds,
                targetGroupIds,
                conditionNodeIds,
              ),
              "condition-definition",
            ),
          );
          effectiveStart = {
            kind: "condition-end",
            id: effectiveStart.id,
            delayMs: 0,
          };
        }
        if (node.kind === "scope") {
          const scopeId = scopeNodeIds.get(`${step.key}:start:${node.id}`);
          if (scopeId === undefined) {
            throw new Error(
              "MotionPlan scope is missing native time node identity",
            );
          }
          const needsCompletionNode =
            node.splitCompletion !== true &&
            referencedScopeNodes.has(`${step.key}:end:${node.id}`);
          const needsStartNode =
            node.preserveStart === true ||
            referencedScopeStartNodes.has(`${step.key}:start:${node.id}`);
          // Only root node of the plan step itself can reuse mainSeq step; when transparent outer layers are promoted, their descendants retain original logical depth.
          const promoteScopeToStep =
            isTopLevel &&
            !preserveStepHost &&
            (needsCompletionNode || needsStartNode);
          const childrenCanLift = node.children.every((child) => {
            const childStart = composeMotionPlanStarts(
              effectiveStart,
              child.nativeStart,
            );
            if (childStart === undefined) return false;
            if (
              child.kind === "scope" ||
              child.kind === "completion-collector" ||
              child.effects.length > 1
            ) {
              return true;
            }
            const [effect] = child.effects;
            return (
              effect !== undefined &&
              addNativeStartDelay(childStart, childDelay(effect)) !== undefined
            );
          });
          if (needsCompletionNode || needsStartNode || !childrenCanLift) {
            const condition = nativeStartCondition(
              effectiveStart,
              step.key,
              effectNodeIds,
              scopeNodeIds,
              targetGroupIds,
              conditionNodeIds,
            );
            const children = renderNodes(
              node.children,
              undefined,
              false,
              promoteScopeToStep
                ? (pendingRootNodeType ?? defaultNodeType)
                : undefined,
              node.conditionDefinitionsBefore,
            );
            const renderedScope = renderedRoot(
              timingContainer(
                scopeId,
                condition,
                children.map(({ xml }) => xml),
                promoteScopeToStep ? undefined : pendingRootNodeType,
              ),
              "scope",
              promoteScopeToStep
                ? pendingRootNodeType !== undefined &&
                    children.some(({ stepEntryRole }) => stepEntryRole)
                : pendingRootNodeType !== undefined,
              promoteScopeToStep,
            );
            rendered.push(renderedScope);
            if (
              pendingRootNodeType !== undefined &&
              renderedScope.stepEntryRole
            ) {
              pendingRootNodeType = undefined;
            }
          } else {
            const liftedRoots = renderNodes(
              node.children,
              effectiveStart,
              // Physically removing identity-less wrapper does not convert descendants into author-declared step roots.
              false,
              pendingRootNodeType,
              node.conditionDefinitionsBefore,
            );
            rendered.push(...liftedRoots);
            if (
              pendingRootNodeType !== undefined &&
              liftedRoots.some(({ stepEntryRole }) => stepEntryRole)
            ) {
              pendingRootNodeType = undefined;
            }
          }
          continue;
        }

        if (node.kind === "completion-collector") {
          const collectorId = scopeNodeIds.get(
            `${step.key}:end:${node.scopeId}`,
          );
          if (collectorId === undefined) {
            throw new Error(
              "MotionPlan completion collection is missing native time node identity",
            );
          }
          const condition = nativeStartCondition(
            effectiveStart,
            step.key,
            effectNodeIds,
            scopeNodeIds,
            targetGroupIds,
            conditionNodeIds,
          );
          rendered.push(
            renderedRoot(
              endCollectorTree(
                collectorId,
                condition,
                node.gates,
                step.key,
                ids,
                effectNodeIds,
                scopeNodeIds,
                targetGroupIds,
                conditionNodeIds,
              ),
              "completion-definition",
            ),
          );
          continue;
        }

        const condition = nativeStartCondition(
          effectiveStart,
          step.key,
          effectNodeIds,
          scopeNodeIds,
          targetGroupIds,
          conditionNodeIds,
        );
        const targetGroupKey = `${step.key}:${node.targetId}`;
        const effectStarts = node.effects.map((effect) =>
          addNativeStartDelay(effectiveStart, childDelay(effect)),
        );
        const keepTargetGroup =
          node.effects.length > 1 &&
          (referencedTargetGroups.has(targetGroupKey) ||
            effectStarts.some((start) => start === undefined));
        if (keepTargetGroup) {
          const groupId = targetGroupIds.get(targetGroupKey);
          if (groupId === undefined) {
            throw new Error(
              "MotionPlan object member is missing native aggregate node identity",
            );
          }
          const children = node.effects.flatMap((effect, effectIndex) => {
            const effectId = effectNodeIds.get(effectKey(effect));
            if (effectId === undefined) {
              throw new Error(
                "MotionPlan effect is missing native time node identity",
              );
            }
            const buildGroupId =
              effect.target.range?.type === "chart"
                ? chartBuildGroupIds.get(effect.target.id)
                : undefined;
            return renderedEffect(
              effect,
              effectIndex === 0 ? "afterEffect" : "withEffect",
              element("p:cond", { delay: childDelay(effect) }),
              effectId,
              buildGroupId,
            );
          });
          // Native import identifies step entry by fixed depth; aggregate container and first real effect retain their respective roles.
          rendered.push(
            renderedRoot(
              timingContainer(
                groupId,
                condition,
                children,
                pendingRootNodeType,
              ),
              "target-group",
              pendingRootNodeType !== undefined && children.length > 0,
            ),
          );
          pendingRootNodeType = undefined;
          continue;
        }

        for (const [effectIndex, effect] of node.effects.entries()) {
          const effectId = effectNodeIds.get(effectKey(effect));
          if (effectId === undefined) {
            throw new Error(
              "MotionPlan effect is missing native time node identity",
            );
          }
          const localStart = effectStarts[effectIndex];
          if (localStart === undefined) {
            throw new Error(
              "Native timing projection lost encodable animation delay layering",
            );
          }
          let effectStart = localStart;
          const rootNodeType =
            effectIndex === 0
              ? (pendingRootNodeType ?? nodeType)
              : "withEffect";
          let effectNodeType = rootNodeType;
          if (localStart.kind === "delay-layer") {
            const lastDelay = motionPlanStartLayers(localStart).delays.at(-1);
            if (lastDelay === undefined) {
              throw new Error(
                "MotionPlan effect delay layer is missing actual delay field",
              );
            }
            effectStart = { kind: "parent-start", delayMs: lastDelay };
            effectNodeType = effectIndex === 0 ? "afterEffect" : "withEffect";
          }
          const effectCondition = nativeStartCondition(
            effectStart,
            step.key,
            effectNodeIds,
            scopeNodeIds,
            targetGroupIds,
            conditionNodeIds,
          );
          const buildGroupId =
            effect.target.range?.type === "chart"
              ? chartBuildGroupIds.get(effect.target.id)
              : undefined;
          const effectTrees = renderedEffect(
            effect,
            effectNodeType,
            effectCondition,
            effectId,
            buildGroupId,
          );
          rendered.push(
            ...(localStart.kind === "delay-layer"
              ? wrapDelayLayers(
                  localStart,
                  effectRoots(effectTrees, false),
                  rootNodeType,
                )
              : effectRoots(
                  effectTrees,
                  effectIndex === 0 && pendingRootNodeType !== undefined,
                )),
          );
          pendingRootNodeType = undefined;
        }
      }
      return rendered;
    };
    // nodeType only marks true step entry node; application timing conditions remain entirely from the plan.
    const stepChildren = renderNodes(
      step.nodes,
      undefined,
      true,
      step.trigger === "click" ? "clickEffect" : "afterEffect",
      step.conditionDefinitionsBefore,
    );
    const stepTimingGroups: string[] = [];
    if (
      stepChildren.some(
        ({ kind, promoteToStep }) => kind === "scope" && promoteToStep,
      )
    ) {
      let groupChildren: string[] = [];
      const flushGroup = () => {
        if (groupChildren.length === 0) return;
        stepTimingGroups.push(mainSequenceWithGroup(ids.take(), groupChildren));
        groupChildren = [];
      };
      for (const root of stepChildren) {
        if (root.kind === "scope" && root.promoteToStep) {
          flushGroup();
          stepTimingGroups.push(root.xml);
        } else {
          groupChildren.push(root.xml);
        }
      }
      flushGroup();
    } else if (stepChildren.length > 0) {
      stepTimingGroups.push(
        mainSequenceWithGroup(
          ids.take(),
          stepChildren.map(({ xml }) => xml),
        ),
      );
    }
    return element(
      "p:par",
      undefined,
      element(
        "p:cTn",
        { id: stepId, fill: "hold" },
        element("p:stCondLst", undefined, trigger),
        element("p:childTnLst", undefined, ...stepTimingGroups),
      ),
    );
  });
  const slideTarget = element(
    "p:tgtEl",
    undefined,
    element("p:sldTgt", undefined),
  );
  const paragraphBuilds = new Map<
    string,
    {
      readonly effect: TimingEffect;
      readonly groupId: number;
      readonly build?: "p" | "allAtOnce";
      readonly animBg?: boolean;
    }
  >();
  const chartBuildTargets = new Map<number, TimingEffect>();
  const wholeTextArrays = new Map<number, Set<number>>();
  const wholeTextEntranceTargets = new Set<number>();
  const wholeTextDeclarations = new Map<number, Set<number>>();
  for (const { effect, declaration } of nativeTextEffects) {
    // Array building only counts declarations on Text itself; Group projection retains its logical identity.
    if (
      !declaration.targetTextBody ||
      declaration.target.id !== effect.target.id
    ) {
      continue;
    }
    if (!isWholeTextProjection(declaration)) continue;
    if (declaration.kind === "entrance") {
      wholeTextEntranceTargets.add(effect.target.id);
    }
    const declarations =
      wholeTextDeclarations.get(effect.target.id) ?? new Set<number>();
    declarations.add(declaration.effectIndex);
    wholeTextDeclarations.set(effect.target.id, declarations);
  }
  for (const [targetId, declarations] of wholeTextDeclarations) {
    if (declarations.size > 1) wholeTextArrays.set(targetId, declarations);
  }
  for (const { effect, declaration } of nativeTextEffects) {
    const groupId = effect.buildGroupId ?? 0;
    const key = `${effect.target.id}:${groupId}`;
    const previous = paragraphBuilds.get(key);
    const isDeclaredTextEffect =
      declaration.targetTextBody && declaration.target.id === effect.target.id;
    // Whole-text entrance builds by paragraph, preserving native responsibilities of pre-click hiding and per-paragraph delays.
    const build =
      groupId === 1
        ? wholeTextEntranceTargets.has(effect.target.id) ||
          wholeTextArrays.has(effect.target.id)
          ? "p"
          : "allAtOnce"
        : isDeclaredTextEffect &&
            (declaration.by === "paragraph" ||
              declaration.target.range?.type === "paragraph")
          ? "p"
          : previous?.build;
    // bldP aggregates by actually written text targets and build groups; does not alter author's orchestration identity.
    const animBg =
      groupId === 1 &&
      (previous?.animBg === true || hasWholeTextBackgroundBehaviors(effect));
    paragraphBuilds.set(key, {
      effect,
      groupId,
      ...(build === undefined ? {} : { build }),
      ...(effect.targetHasBackground ? { animBg } : {}),
    });
  }
  for (const step of steps) {
    for (const effect of step.effects) {
      const timed = effect as TimingEffect;
      if (!timed.targetTextBody && timed.target.range?.type === "chart") {
        chartBuildTargets.set(timed.target.id, timed);
      }
    }
  }
  const buildList = [
    ...[...paragraphBuilds.values()].map(({ effect, groupId, build, animBg }) =>
      element("p:bldP", {
        spid: effect.target.id,
        grpId: groupId,
        build,
        animBg,
      }),
    ),
    ...[...chartBuildTargets.values()].map((effect) => {
      const by =
        effect.target.range?.type === "chart"
          ? effect.target.range.step
          : undefined;
      return element(
        "p:bldGraphic",
        {
          spid: effect.target.id,
          grpId: chartBuildGroupIds.get(effect.target.id),
        },
        by === "series" || by === "category"
          ? element("p:bldSub", undefined, element("a:bldChart", { bld: by }))
          : element("p:bldAsOne", undefined),
      );
    }),
  ];
  const xml = element(
    "p:timing",
    undefined,
    element(
      "p:tnLst",
      undefined,
      element(
        "p:par",
        undefined,
        element(
          "p:cTn",
          {
            id: rootId,
            dur: "indefinite",
            restart: "never",
            nodeType: "tmRoot",
          },
          element(
            "p:childTnLst",
            undefined,
            element(
              "p:seq",
              { concurrent: true, nextAc: "seek" },
              element(
                "p:cTn",
                { id: sequenceId, dur: "indefinite", nodeType: "mainSeq" },
                element("p:childTnLst", undefined, ...stepGroups),
              ),
              element(
                "p:prevCondLst",
                undefined,
                element("p:cond", { evt: "onPrev", delay: 0 }, slideTarget),
              ),
              element(
                "p:nextCondLst",
                undefined,
                element("p:cond", { evt: "onNext", delay: 0 }, slideTarget),
              ),
            ),
          ),
        ),
      ),
    ),
    buildList.length === 0
      ? undefined
      : element("p:bldLst", undefined, ...buildList),
  );
  return { xml, nativeProjections };
}

/** Slide-level motion module writes object timing and returns observation describing declarations and determinable relative times only. */
export const motionCapability: SlideCapabilityModule = Object.freeze({
  declaration: "motion",
  slideProp: "__motion",
  slot: "timing",
  always: true,
  transparentDeclarations: [
    {
      declaration: "rp:motion-scope",
      compileDeclaration(input) {
        return [motionScopeAnnotation(input.props)];
      },
    } satisfies TransparentCapabilityModule,
    {
      declaration: "rp:motion-step",
      compileDeclaration(input) {
        return motionStepAnnotation(input.props);
      },
    } satisfies TransparentCapabilityModule,
  ],
  compileSlide(_input: CapabilityInput, context) {
    const plan = compileMotionPlan(context.contributions);
    validateKnownMotionTimes(plan.observation);
    const hasEffects = plan.steps.some((step) => step.effects.length > 0);
    if (!hasEffects) {
      registerMotionObservation(plan.observation, plan.query);
      return {
        xml: "",
        extensions: [],
        ...(plan.observation.scopes === undefined
          ? {}
          : { observation: plan.observation }),
      };
    }
    const timing = timingXml(plan.steps, context.dimensions);
    const observation = {
      ...plan.observation,
      ...(timing.nativeProjections.length === 0
        ? {}
        : { nativeProjections: timing.nativeProjections }),
    } satisfies MotionObservation;
    // Query identity must bind to final public observation object containing native projections.
    registerMotionObservation(observation, plan.query);
    return {
      xml: timing.xml,
      extensions: [],
      observation,
    };
  },
});
