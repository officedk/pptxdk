import {
  xmlEscape,
  type CapabilityInput,
  type SlideCapabilityModule,
  type TransparentCapabilityModule,
} from "@react-pptx/core";

import type {
  MotionDirection,
  MotionEffectObservation,
  MotionObservation,
  MotionStartDependency,
} from "./motion-api.ts";
import { motionScopeAnnotation } from "./motion-api.ts";
import { compileMotionPlan, type MotionPlanStep } from "./motion-plan.ts";

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
};

function timingTarget(effect: TimingEffect): string {
  const range = effect.target.range;
  return element(
    "p:tgtEl",
    undefined,
    element(
      "p:spTgt",
      { spid: effect.target.id },
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

function direction(effect: MotionEffectObservation): MotionDirection {
  return typeof effect.from === "string" ? effect.from : "bottom";
}

function behaviorIds(effect: TimingEffect, ids: TimingIds): readonly string[] {
  const target = () => timingTarget(effect);
  const duration = effect.durationMs;
  const show = () =>
    element(
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
        target(),
        attributeNames("style.visibility"),
      ),
      element("p:to", undefined, element("p:strVal", { val: "visible" })),
    );
  const filter = (name: string) =>
    element(
      "p:animEffect",
      { transition: "in", filter: name },
      element(
        "p:cBhvr",
        undefined,
        element("p:cTn", { id: ids.take(), dur: duration }),
        target(),
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
    additive: boolean,
  ) =>
    element(
      "p:anim",
      { calcmode: "lin", valueType: "num" },
      element(
        "p:cBhvr",
        { additive: additive ? "base" : undefined },
        element("p:cTn", { id: ids.take(), dur: duration, fill: "hold" }),
        target(),
        attributeNames(name),
      ),
      element(
        "p:tavLst",
        undefined,
        element("p:tav", { tm: 0 }, element("p:val", undefined, value(from))),
        element(
          "p:tav",
          { tm: 100000 },
          element("p:val", undefined, value(to)),
        ),
      ),
    );

  switch (effect.preset) {
    case "fade":
      return [show(), filter("fade")];
    case "wipe": {
      const filters: Record<MotionDirection, string> = {
        top: "wipe(up)",
        right: "wipe(right)",
        bottom: "wipe(down)",
        left: "wipe(left)",
      };
      return [show(), filter(filters[direction(effect)])];
    }
    case "zoom":
      return [
        show(),
        motion("ppt_w", 0, "#ppt_w", false),
        motion("ppt_h", 0, "#ppt_h", false),
        filter("fade"),
      ];
    case "fly": {
      const flyFrom: Record<MotionDirection, readonly [string, string]> = {
        left: ["0-#ppt_w/2", "#ppt_y"],
        right: ["1+#ppt_w/2", "#ppt_y"],
        top: ["#ppt_x", "0-#ppt_h/2"],
        bottom: ["#ppt_x", "1+#ppt_h/2"],
      };
      const [x, y] = flyFrom[direction(effect)];
      return [
        show(),
        motion("ppt_x", x, "#ppt_x", true),
        motion("ppt_y", y, "#ppt_y", true),
      ];
    }
    case "spinIn":
      return [
        show(),
        motion("ppt_w", 0, "#ppt_w", false),
        motion("ppt_h", 0, "#ppt_h", false),
        motion("style.rotation", 360, 0, false),
        filter("fade"),
      ];
    case "stretch":
      return [
        show(),
        motion("ppt_w", 0, "#ppt_w", false),
        motion("ppt_h", "#ppt_h", "#ppt_h", false),
      ];
    case "floatUp":
      return [
        show(),
        filter("fade"),
        motion("ppt_x", "#ppt_x", "#ppt_x", false),
        motion("ppt_y", "#ppt_y+.1", "#ppt_y", false),
      ];
    case "spin":
      return [
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
            undefined,
            element("p:cTn", { id: ids.take(), dur: duration, fill: "hold" }),
            target(),
            attributeNames("r"),
          ),
        ),
      ];
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
  delay: number,
  ids: TimingIds,
  effectId: number,
  buildGroupId?: number,
  dependencyId?: number,
): string {
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
    grpId: effect.targetTextBody ? 0 : buildGroupId,
    nodeType,
  };
  return element(
    "p:par",
    undefined,
    element(
      "p:cTn",
      properties,
      element(
        "p:stCondLst",
        undefined,
        effect.startAfter === undefined
          ? element("p:cond", { delay })
          : element(
              "p:cond",
              { evt: "end", delay: effect.startAfter.offsetMs },
              element("p:tn", { val: dependencyId }),
            ),
      ),
      effect.by !== "word" && effect.by !== "letter"
        ? undefined
        : element(
            "p:iterate",
            { type: effect.by === "word" ? "wd" : "lt" },
            // 单次 dur 只描述当前词或字，不能伪装成整段自然完成时刻。
            element("p:tmAbs", { val: effect.intervalMs }),
          ),
      element("p:childTnLst", undefined, ...behaviorIds(effect, ids)),
    ),
  );
}

function effectKey(effect: TimingEffect | MotionStartDependency): string {
  const targetId = "target" in effect ? effect.target.id : effect.targetId;
  return `${targetId}:${effect.effectIndex}:${"partIndex" in effect ? (effect.partIndex ?? "") : ""}`;
}

function timingXml(steps: readonly MotionPlanStep[]): string {
  const ids = new TimingIds();
  const rootId = ids.take();
  const sequenceId = ids.take();
  const effectNodeIds = new Map<string, number>();
  for (const step of steps) {
    for (const effect of step.effects) {
      effectNodeIds.set(effectKey(effect), ids.take());
    }
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
  const stepGroups = steps.map((step) => {
    const stepId = ids.take();
    const ordered = step.effects
      .map((effect) => effect as TimingEffect)
      .sort(
        (left, right) =>
          (left.startMs ?? Number.MAX_SAFE_INTEGER) -
            (right.startMs ?? Number.MAX_SAFE_INTEGER) ||
          left.target.id - right.target.id,
      );
    const groups: {
      start: number;
      effects: { readonly effect: TimingEffect; readonly nodeType: string }[];
    }[] = [];
    let ended = 0;
    ordered.forEach((effect, index) => {
      const current = groups.at(-1);
      if (effect.startAfter !== undefined) {
        groups.push({
          start: 0,
          effects: [{ effect, nodeType: "afterEffect" }],
        });
        return;
      }
      if (effect.startMs === undefined) {
        throw new Error("MotionPlan 数值起点缺失且没有原生完成关系");
      }
      if (index === 0) {
        groups.push({
          start: effect.startMs,
          effects: [
            {
              effect,
              nodeType:
                step.trigger === "click" ? "clickEffect" : "afterEffect",
            },
          ],
        });
      } else if (current !== undefined && effect.startMs < ended) {
        current.effects.push({ effect, nodeType: "withEffect" });
      } else {
        groups.push({
          start: ended,
          effects: [{ effect, nodeType: "afterEffect" }],
        });
      }
      if (effect.endMs !== undefined) ended = Math.max(ended, effect.endMs);
    });
    const trigger =
      step.trigger === "click"
        ? element("p:cond", { delay: "indefinite" })
        : `${element("p:cond", { delay: "indefinite" })}${element(
            "p:cond",
            { evt: "onBegin", delay: 0 },
            element("p:tn", { val: sequenceId }),
          )}`;
    const groupXml = groups.map((group) => {
      const groupId = ids.take();
      return element(
        "p:par",
        undefined,
        element(
          "p:cTn",
          { id: groupId, fill: "hold" },
          element(
            "p:stCondLst",
            undefined,
            element("p:cond", { delay: group.start }),
          ),
          element(
            "p:childTnLst",
            undefined,
            ...group.effects.map(({ effect, nodeType }) => {
              const dependencyId =
                effect.startAfter === undefined
                  ? undefined
                  : effectNodeIds.get(effectKey(effect.startAfter));
              const effectId = effectNodeIds.get(effectKey(effect));
              const buildGroupId =
                effect.target.range?.type === "chart"
                  ? chartBuildGroupIds.get(effect.target.id)
                  : undefined;
              if (effectId === undefined) {
                throw new Error("MotionPlan 效果缺少原生时间节点身份");
              }
              if (
                effect.startAfter !== undefined &&
                dependencyId === undefined
              ) {
                throw new Error("MotionPlan 单前驱完成事件不存在");
              }
              return effectTree(
                effect,
                nodeType,
                effect.startMs === undefined ? 0 : effect.startMs - group.start,
                ids,
                effectId,
                buildGroupId,
                dependencyId,
              );
            }),
          ),
        ),
      );
    });
    return element(
      "p:par",
      undefined,
      element(
        "p:cTn",
        { id: stepId, fill: "hold" },
        element("p:stCondLst", undefined, trigger),
        element("p:childTnLst", undefined, ...groupXml),
      ),
    );
  });
  const slideTarget = element(
    "p:tgtEl",
    undefined,
    element("p:sldTgt", undefined),
  );
  const buildTargets = new Map<number, TimingEffect>();
  for (const step of steps) {
    for (const effect of step.effects) {
      const timed = effect as TimingEffect;
      if (timed.targetTextBody || timed.target.range?.type === "chart") {
        const previous = buildTargets.get(timed.target.id);
        // 原生构建粒度属于整个文字对象，整体效果不能覆盖逐段构建要求。
        buildTargets.set(
          timed.target.id,
          timed.targetTextBody &&
            (timed.by === "paragraph" || previous?.by === "paragraph")
            ? { ...timed, by: "paragraph" }
            : timed,
        );
      }
    }
  }
  return element(
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
    buildTargets.size === 0
      ? undefined
      : element(
          "p:bldLst",
          undefined,
          ...[...buildTargets.values()].map((effect) => {
            if (effect.target.range?.type === "chart") {
              const by = effect.target.range.step;
              return element(
                "p:bldGraphic",
                {
                  spid: effect.target.id,
                  grpId: chartBuildGroupIds.get(effect.target.id),
                },
                by === "series" || by === "category"
                  ? element(
                      "p:bldSub",
                      undefined,
                      element("a:bldChart", { bld: by }),
                    )
                  : element("p:bldAsOne", undefined),
              );
            }
            return element("p:bldP", {
              spid: effect.target.id,
              grpId: 0,
              build: effect.by === "paragraph" ? "p" : undefined,
              animBg: effect.targetFilled ? true : undefined,
            });
          }),
        ),
  );
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
    const hasEffects = plan.steps.some((step) => step.effects.length > 0);
    if (!hasEffects) {
      return {
        xml: "",
        extensions: [],
        ...(plan.observation.scopes === undefined
          ? {}
          : { observation: plan.observation }),
      };
    }
    return {
      xml: timingXml(plan.steps),
      extensions: [],
      observation: plan.observation satisfies MotionObservation,
    };
  },
});
