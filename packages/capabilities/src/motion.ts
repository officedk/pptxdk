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
import { maximumTimeNodeValue, motionScopeAnnotation } from "./motion-api.ts";
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
  Box,
  MotionScope,
  Step,
  fade,
  fly,
  floatUp,
  repeat,
  spin,
  spinIn,
  stretch,
  useMotionScope,
  wipe,
  zoom,
} from "./motion-api.ts";
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
  BoxProps,
  Easing,
  Effect,
  EffectOptions,
  EffectPreset,
  BreatheEffect,
  BreatheOptions,
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
  StepProps,
} from "./motion-api.ts";

type XmlAttributes = Record<string, string | number | boolean | undefined>;

function invalidMotionTime(field: string): never {
  throw new CapabilityError(
    "invalid-motion-time",
    `${field} 必须是精确的非负整数毫秒，且实际写入时不得超过 ST_TLTime 范围`,
  );
}

// 累计总时长可由 repeat 表达；只有实际投影到 ST_TLTime 的字段受其上限约束。
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
    validateKnownTime(scope.delayMs, "编排作用域延迟");
    validateKnownTime(scope.eachMs, "交错间隔");
    validateKnownTime(scope.startMs, "编排作用域起点");
    validateKnownTime(scope.endMs, "编排作用域完成时刻");
    validateKnownTime(scope.startAfter?.offsetMs, "作用域完成依赖偏移");
    for (const dependency of scope.startAfterAll ?? []) {
      validateKnownTime(dependency.offsetMs, "作用域起点约束偏移");
    }
  }
  for (const step of observation.steps) {
    for (const effect of step.effects) {
      validateKnownTime(effect.durationMs, "动画时长");
      validateKnownTime(effect.delayMs, "动画延迟");
      validateKnownTime(effect.intervalMs, "动画迭代间隔");
      validateKnownTime(effect.startMs, "动画起点");
      validateKnownTime(effect.endMs, "动画完成时刻");
      validateKnownTime(effect.startAfter?.offsetMs, "动画完成依赖偏移");
      for (const dependency of effect.startAfterAll ?? []) {
        validateKnownTime(dependency.offsetMs, "动画起点约束偏移");
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
    throw new CapabilityError("invalid-motion-page-size", "页面尺寸必须大于 0");
  }
  return {
    x: (box.x + box.width / 2) / dimensions.width,
    y: (box.y + box.height / 2) / dimensions.height,
  };
}

/** 共同旋转中心按写入 xfrm 与 sldSz 的整数 EMU 计算，避免读回两种舍入结果。 */
function serializedGeometryCenter(
  box: GeometryBox,
  dimensions: { readonly width: number; readonly height: number },
): { readonly x: number; readonly y: number } {
  const widthEmu = pxToEmu(dimensions.width);
  const heightEmu = pxToEmu(dimensions.height);
  if (widthEmu <= 0 || heightEmu <= 0) {
    throw new CapabilityError("invalid-motion-page-size", "页面尺寸必须大于 0");
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
      `${purpose}需要同次布局得到的原生对象范围`,
    );
  }
  return effect.targetBox;
}

function rotationFormula(effect: TimingEffect): {
  readonly angle: string;
  readonly degrees: number;
} {
  const degrees = typeof effect.from === "number" ? effect.from : 360;
  return {
    angle: `(${formulaNumber(degrees)}*pi/180)*(1-$)`,
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
      "整体文字入场需要同次布局得到的原生对象范围",
    );
  }
  switch (direction(effect)) {
    case "left":
      return { x: -(box.x + box.width) / dimensions.width, y: 0 };
    case "right":
      return { x: (dimensions.width - box.x) / dimensions.width, y: 0 };
    case "top":
      return { x: 0, y: -(box.y + box.height) / dimensions.height };
    case "bottom":
      return { x: 0, y: (dimensions.height - box.y) / dimensions.height };
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
  validateKnownTime(duration, "动画时长");
  if (effect.preset === "breathe") {
    timeNodeValue(Math.round(duration / 2), "呼吸半周期");
  } else {
    timeNodeValue(duration, "动画行为时长");
  }
  if (effect.repeatCount !== undefined && effect.repeatCount !== "forever") {
    timeNodeValue(effect.repeatCount * 1000, "动画重复次数");
  }
  if (effect.intervalMs !== undefined) {
    timeNodeValue(effect.intervalMs, "动画迭代间隔");
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
  // 共同中心位移已是完整进度函数；公式接收 0→1 数值进度，避免函数值被当作端点再次插值。
  const progressMotion = (
    name: string,
    formula: string,
    targetElement = target(),
  ) => motion(name, 0, 1, "sum", targetElement, formula);
  const rotation = (targetElement: string, additive: boolean | "sum" = false) =>
    element(
      "p:animRot",
      {
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
  // 段落构建使用真实 pRg 和独立背景入口，避免整对象可见目标提前解锁后续段落。
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
        targetBox(effect, "整体文字缩放"),
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
      const box = targetBox(effect, "整体文字旋入");
      const center = serializedGeometryCenter(box, dimensions);
      const position = centerRotationDelta(effect, center, dimensions, true);
      const geometry = positionTargets().flatMap((targetElement) => [
        singleMotion("ppt_w", 0, "#ppt_w", false, targetElement),
        singleMotion("ppt_h", 0, "#ppt_h", false, targetElement),
        // 段落旋转相对文本框原角度表达，不能再把对象角度加进端点。
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
        targetBox(effect, "整体文字横向伸展"),
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
    case "floatUp": {
      if (!wholeText) {
        return [
          showOnce(),
          filter("fade", target()),
          singleMotion("ppt_x", "#ppt_x", "#ppt_x", false),
          singleMotion("ppt_y", "#ppt_y+.1", "#ppt_y", false),
        ];
      }
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
          singleMotion("ppt_y", 0.1, 0, "sum", targetElement),
        ]),
      ];
    }
    case "spin": {
      if (!wholeText && effect.groupSpinCenter === undefined) {
        return [rotation(target())];
      }
      const center =
        effect.groupSpinCenter ??
        serializedGeometryCenter(targetBox(effect, "整体文字自转"), dimensions);
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
              dur: Math.round(duration / 2),
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
    validateKnownTime(effect.startAfter.offsetMs, "完成依赖偏移");
  }
  const preset = {
    fade: { id: 10, class: "entr" },
    wipe: { id: 22, class: "entr" },
    zoom: { id: 53, class: "entr" },
    fly: { id: 2, class: "entr" },
    floatUp: { id: 42, class: "entr" },
    spinIn: { id: 49, class: "entr" },
    stretch: { id: 17, class: "entr" },
    spin: { id: 8, class: "emph" },
    breathe: { id: 0, class: "emph" },
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
            // 单次 dur 只描述当前词或字，不能伪装成整段自然完成时刻。
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
      delay: timeNodeValue(start.delayMs, "动画起点延迟"),
    });
  }
  if (start.kind === "condition-join") {
    throw new Error("MotionPlan 条件集合必须先于其消费者写出");
  }
  if (start.kind === "delay-layer") {
    throw new Error("MotionPlan 延迟层必须通过原生时间容器写出");
  }
  if (start.kind === "condition-end") {
    const conditionId = conditionNodeIds.get(`${stepKey}:${start.id}`);
    if (conditionId === undefined) {
      throw new Error("MotionPlan 条件集合缺少原生时间节点身份");
    }
    return element(
      "p:cond",
      {
        evt: "end",
        delay: timeNodeValue(start.delayMs, "条件集合完成偏移"),
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
    throw new Error("MotionPlan 完成依赖没有对应的原生时序节点");
  }
  return element(
    "p:cond",
    {
      evt:
        start.kind === "event" || start.dependency.point === "start"
          ? "begin"
          : "end",
      delay: timeNodeValue(start.dependency.offsetMs, "完成依赖偏移"),
    },
    element("p:tn", { val: dependencyId }),
  );
}

function endSync(): string {
  // 容器等全部真实子节点自然结束；不从这些子项估算聚合数值时长。
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
      // 空作用域的 END 与其未知 START 相同；省略空列表以保持合法 XML。
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
    throw new Error("MotionPlan 条件集合缺少原生时间节点身份");
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
  const leaves = gates.map((gate) =>
    element(
      "p:par",
      undefined,
      element(
        "p:cTn",
        { id: ids.take(), dur: 0, fill: "hold" },
        element(
          "p:stCondLst",
          undefined,
          nativeStartCondition(
            gateStart(gate),
            stepKey,
            effectNodeIds,
            scopeNodeIds,
            targetGroupIds,
            conditionNodeIds,
          ),
        ),
      ),
    ),
  );
  return timingContainer(id, startCondition, leaves, nodeType);
}

function mainSequenceWithGroup(
  id: number,
  children: readonly string[],
): string {
  // mainSeq 的标准 step→with-group→入口节点层级供原生导入识别步首角色。
  // 分组从步首立即开始，不复制计划起点，也不创建 Shape 或完成来源。
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
        if (gate.kind === "end" || gate.kind === "begin") {
          keepDependencyContainer(gate.dependency, stepKey);
        }
      }
    }
  };

  // 只为真实 end 依赖保留聚合节点；无消费者的容器可把起点投影给成员。
  const collectReferencedContainers = (
    nodes: readonly MotionPlanNode[],
    stepKey: string,
  ): void => {
    for (const node of nodes) {
      collectStartReferences(node.nativeStart, stepKey);
      if (node.kind === "completion-collector") {
        for (const gate of node.gates) {
          if (gate.kind === "end" || gate.kind === "begin") {
            keepDependencyContainer(gate.dependency, stepKey);
          }
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
            if (gate.kind === "end" || gate.kind === "begin") {
              keepDependencyContainer(gate.dependency, stepKey);
            }
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
        if (gate.kind === "end" || gate.kind === "begin") {
          keepDependencyContainer(gate.dependency, step.key);
        }
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
        // bldGraphic 的 grpId 必须关联该图表实际存在的时间节点。
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
    // 子项可由后续点击启动；只检查当前 Step 会遗漏仍活动的父 Group。
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
          "原生 Group 自转投影需要同次布局得到的组合范围",
        );
      }
      const center = serializedGeometryCenter(box, dimensions);
      const projectedMembers = leafMembers(members);
      if (projectedMembers.length === 0) {
        throw new CapabilityError(
          "empty-group-projection",
          "原生 Group 自转需要至少一个真实成员目标",
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
        // 文字成员需沿用自己的段落数与背景职责，才能在 Group 投影中继续整对象变换。
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
  // 空步骤只保留查询声明，不为原生时序树写入空的 childTnLst。
  const stepsWithNativeNodes = steps.filter(({ nodes }) => nodes.length > 0);
  const stepGroups = stepsWithNativeNodes.map((step) => {
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
        "动画对象成员延迟",
      );
    const wrapDelayLayers = (
      start: MotionPlanStart,
      children: readonly string[],
      nodeType?: string,
    ): string[] => {
      const { base, delays } = motionPlanStartLayers(start);
      if (delays.length === 0) return [...children];
      const rendered: string[] = [];
      let outerStart = base;
      if (outerStart.kind === "condition-join") {
        rendered.push(
          conditionJoinTree(
            outerStart,
            step.key,
            ids,
            effectNodeIds,
            scopeNodeIds,
            targetGroupIds,
            conditionNodeIds,
          ),
        );
        outerStart = {
          kind: "condition-end",
          id: outerStart.id,
          delayMs: 0,
        };
      }
      let nestedChildren = [...children];
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
      );
      return rendered;
    };
    const renderNodes = (
      nodes: readonly MotionPlanNode[],
      inheritedStart: MotionPlanStart | undefined,
      isTopLevel: boolean,
      firstRootNodeType?: string,
      conditionDefinitionsBefore: readonly MotionPlanConditionPlacement[] = [],
    ): string[] => {
      const rendered: string[] = [];
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
            conditionJoinTree(
              placement.join,
              step.key,
              ids,
              effectNodeIds,
              scopeNodeIds,
              targetGroupIds,
              conditionNodeIds,
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
          throw new Error("原生时序投影需要保留延迟分层");
        }
        if (effectiveStart.kind === "delay-layer") {
          const { delays } = motionPlanStartLayers(effectiveStart);
          const lastDelay = delays.at(-1);
          if (lastDelay === undefined) {
            throw new Error("MotionPlan 延迟层缺少实际延迟字段");
          }
          const delayedNode: MotionPlanNode = {
            ...node,
            nativeStart: { kind: "parent-start", delayMs: lastDelay },
          };
          rendered.push(
            ...wrapDelayLayers(
              effectiveStart,
              renderNodes([delayedNode], undefined, false),
              pendingRootNodeType,
            ),
          );
          pendingRootNodeType = undefined;
          continue;
        }
        if (effectiveStart.kind === "condition-join") {
          rendered.push(
            conditionJoinTree(
              effectiveStart,
              step.key,
              ids,
              effectNodeIds,
              scopeNodeIds,
              targetGroupIds,
              conditionNodeIds,
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
            throw new Error("MotionPlan 作用域缺少原生时间节点身份");
          }
          const needsCompletionNode =
            node.splitCompletion !== true &&
            referencedScopeNodes.has(`${step.key}:end:${node.id}`);
          const needsStartNode =
            node.preserveStart === true ||
            referencedScopeStartNodes.has(`${step.key}:start:${node.id}`);
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
              undefined,
              node.conditionDefinitionsBefore,
            );
            rendered.push(
              timingContainer(
                scopeId,
                condition,
                children,
                pendingRootNodeType,
              ),
            );
          } else {
            rendered.push(
              ...renderNodes(
                node.children,
                effectiveStart,
                false,
                pendingRootNodeType,
                node.conditionDefinitionsBefore,
              ),
            );
          }
          if (rendered.length > 0) pendingRootNodeType = undefined;
          continue;
        }

        if (node.kind === "completion-collector") {
          const collectorId = scopeNodeIds.get(
            `${step.key}:end:${node.scopeId}`,
          );
          if (collectorId === undefined) {
            throw new Error("MotionPlan 完成集合缺少原生时间节点身份");
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
              pendingRootNodeType,
            ),
          );
          pendingRootNodeType = undefined;
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
            throw new Error("MotionPlan 对象成员缺少原生聚合节点身份");
          }
          const children = node.effects.flatMap((effect, effectIndex) => {
            const effectId = effectNodeIds.get(effectKey(effect));
            if (effectId === undefined) {
              throw new Error("MotionPlan 效果缺少原生时间节点身份");
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
          rendered.push(
            timingContainer(groupId, condition, children, pendingRootNodeType),
          );
          pendingRootNodeType = undefined;
          continue;
        }

        for (const [effectIndex, effect] of node.effects.entries()) {
          const effectId = effectNodeIds.get(effectKey(effect));
          if (effectId === undefined) {
            throw new Error("MotionPlan 效果缺少原生时间节点身份");
          }
          const localStart = effectStarts[effectIndex];
          if (localStart === undefined) {
            throw new Error("原生时序投影丢失可编码的动画延迟分层");
          }
          let effectStart = localStart;
          let effectNodeType =
            pendingRootNodeType ??
            (effectIndex === 0 ? nodeType : "withEffect");
          if (localStart.kind === "delay-layer") {
            const lastDelay = motionPlanStartLayers(localStart).delays.at(-1);
            if (lastDelay === undefined) {
              throw new Error("MotionPlan 效果延迟层缺少实际延迟字段");
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
                  effectTrees,
                  pendingRootNodeType ??
                    (effectIndex === 0 ? nodeType : "withEffect"),
                )
              : effectTrees),
          );
          pendingRootNodeType = undefined;
        }
      }
      return rendered;
    };
    // nodeType 仅标记真实步首节点，应用的时刻条件仍完全来自计划。
    const stepChildren = renderNodes(
      step.nodes,
      undefined,
      true,
      step.trigger === "click" ? "clickEffect" : "afterEffect",
      step.conditionDefinitionsBefore,
    );
    const stepTimingGroups =
      stepChildren.length === 0
        ? []
        : [mainSequenceWithGroup(ids.take(), stepChildren)];
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
    // 数组构建只统计 Text 自身的声明，Group 投影仍保留其逻辑身份。
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
    // Whole Text 入场按段落构建，保留点击前隐藏和逐段延迟的原生职责。
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
    // bldP 按真实写出的文字目标和构建组聚合；不改动作者编排身份。
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

/** 页面级动画模块写入对象时序，并返回仅描述声明与可确定相对时间的观察。 */
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
    // 查询身份必须绑定包含原生投影的最终公开观察对象。
    registerMotionObservation(observation, plan.query);
    return {
      xml: timing.xml,
      extensions: [],
      observation,
    };
  },
});
