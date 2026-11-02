import {
  CapabilityError,
  xmlEscape,
  type CapabilityContribution,
  type CapabilityInput,
  type SlideCapabilityModule,
} from "@react-pptx/core";

import type {
  Effect,
  MotionDirection,
  MotionEffectObservation,
  MotionObservation,
  MotionStepObservation,
} from "./motion-api.ts";

export {
  Step,
  fade,
  fly,
  floatUp,
  spin,
  spinIn,
  stretch,
  wipe,
  zoom,
} from "./motion-api.ts";
export type {
  Easing,
  Effect,
  EffectOptions,
  EffectPreset,
  EmphasisEffect,
  EntranceEffect,
  MotionDirection,
  MotionEffectObservation,
  MotionObservation,
  MotionProps,
  MotionStepObservation,
  StepProps,
} from "./motion-api.ts";

type StepDeclaration =
  | { readonly mode: "automatic"; readonly identity: string }
  | { readonly mode: "explicit"; readonly number: number };

type MotionValue = {
  readonly target: {
    readonly kind: string;
    readonly name: string;
    readonly testId?: string;
    readonly textBody: boolean;
    readonly filled: boolean;
  };
  readonly step?: StepDeclaration;
  readonly effects: readonly Effect[];
};

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

type TimedEffect = MotionEffectObservation & {
  readonly stepKey: string;
  readonly order: number;
  readonly targetTextBody: boolean;
  readonly targetFilled: boolean;
};

type TimedStep = {
  readonly key: string;
  readonly trigger: "auto" | "click";
  readonly number?: number;
  readonly order: number;
  readonly effects: TimedEffect[];
};

function motionValue(value: unknown): value is MotionValue {
  return (
    isRecord(value) &&
    isRecord(value["target"]) &&
    Array.isArray(value["effects"])
  );
}

function motionSteps(
  contributions: readonly CapabilityContribution[],
): readonly MotionStepObservation[] {
  const steps = new Map<string, TimedStep>();
  const objects = contributions
    .filter(({ capability }) => capability === "motion")
    .slice()
    .sort((left, right) => left.targetId - right.targetId);

  for (const contribution of objects) {
    if (!motionValue(contribution.value)) {
      throw new CapabilityError(
        "invalid-motion",
        `motion 注解在 ${contribution.path} 中无效`,
      );
    }
    const value = contribution.value;
    const step = value.step;
    const trigger = step === undefined ? "auto" : "click";
    const key =
      step === undefined
        ? "slide-entry"
        : step.mode === "explicit"
          ? `explicit:${step.number}`
          : `automatic:${step.identity}`;
    const stepOrder = contribution.targetId;
    const existing = steps.get(key);
    const compiled: TimedStep = existing ?? {
      key,
      trigger,
      ...(step?.mode === "explicit" ? { number: step.number } : {}),
      order: stepOrder,
      effects: [],
    };
    value.effects.forEach((effect, index) => {
      const durationMs = Math.round(effect.duration);
      const delayMs = Math.round(effect.delay);
      compiled.effects.push({
        target: {
          id: contribution.targetId,
          name: value.target.name,
          ...(value.target.testId === undefined
            ? {}
            : { testId: value.target.testId }),
        },
        kind: effect.kind,
        preset: effect.preset,
        durationMs,
        delayMs,
        ...(effect.from === undefined ? {} : { from: effect.from }),
        ...(effect.easing === undefined
          ? {}
          : {
              easing: Object.freeze({
                ...(effect.easing.accelerate === undefined
                  ? {}
                  : { accelerate: effect.easing.accelerate }),
                ...(effect.easing.decelerate === undefined
                  ? {}
                  : { decelerate: effect.easing.decelerate }),
              }),
            }),
        startMs: delayMs,
        endMs: delayMs + durationMs,
        stepKey: key,
        order: index,
        targetTextBody: value.target.textBody,
        targetFilled: value.target.filled,
      });
    });
    steps.set(key, compiled);
  }

  return [...steps.values()]
    .sort((left, right) => {
      if (left.trigger !== right.trigger)
        return left.trigger === "auto" ? -1 : 1;
      return left.order - right.order;
    })
    .map((step) => ({
      trigger: step.trigger,
      ...(step.number === undefined ? {} : { number: step.number }),
      effects: step.effects
        .slice()
        .sort(
          (left, right) =>
            left.startMs - right.startMs ||
            left.target.id - right.target.id ||
            left.order - right.order,
        )
        .map(
          ({
            stepKey: _key,
            order: _order,
            targetTextBody: _text,
            targetFilled: _fill,
            ...effect
          }) => effect,
        ),
    }));
}

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
  readonly targetTextBody: boolean;
  readonly targetFilled: boolean;
};

function timingTarget(effect: TimingEffect): string {
  return element(
    "p:tgtEl",
    undefined,
    element("p:spTgt", { spid: effect.target.id }),
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
  }
}

function effectTree(
  effect: TimingEffect,
  nodeType: string,
  delay: number,
  ids: TimingIds,
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
    id: ids.take(),
    presetID: preset.id,
    presetClass: preset.class,
    presetSubtype: subtype,
    accel:
      easing?.accelerate === undefined || easing.accelerate === 0
        ? undefined
        : Math.round(easing.accelerate * 100000),
    decel:
      easing?.decelerate === undefined || easing.decelerate === 0
        ? undefined
        : Math.round(easing.decelerate * 100000),
    fill: "hold",
    grpId: effect.targetTextBody ? 0 : undefined,
    nodeType,
  };
  return element(
    "p:par",
    undefined,
    element(
      "p:cTn",
      properties,
      element("p:stCondLst", undefined, element("p:cond", { delay })),
      element("p:childTnLst", undefined, ...behaviorIds(effect, ids)),
    ),
  );
}

function timingXml(steps: readonly MotionStepObservation[]): string {
  const ids = new TimingIds();
  const rootId = ids.take();
  const sequenceId = ids.take();
  const stepGroups = steps.map((step) => {
    const stepId = ids.take();
    const ordered = step.effects
      .map((effect) => effect as TimingEffect)
      .sort(
        (left, right) =>
          left.startMs - right.startMs || left.target.id - right.target.id,
      );
    const groups: {
      start: number;
      effects: { readonly effect: TimingEffect; readonly nodeType: string }[];
    }[] = [];
    let ended = 0;
    ordered.forEach((effect, index) => {
      const current = groups.at(-1);
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
      ended = Math.max(ended, effect.endMs);
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
            ...group.effects.map(({ effect, nodeType }) =>
              effectTree(effect, nodeType, effect.startMs - group.start, ids),
            ),
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
  const effectTargets = new Map<number, TimingEffect>();
  for (const step of steps) {
    for (const effect of step.effects) {
      const timed = effect as TimingEffect;
      if (timed.targetTextBody) effectTargets.set(timed.target.id, timed);
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
    effectTargets.size === 0
      ? undefined
      : element(
          "p:bldLst",
          undefined,
          ...[...effectTargets.values()].map((effect) =>
            element("p:bldP", {
              spid: effect.target.id,
              grpId: 0,
              animBg: effect.targetFilled ? true : undefined,
            }),
          ),
        ),
  );
}

/** 页面级动画模块写入对象时序，并返回仅描述声明与可确定相对时间的观察。 */
export const motionCapability: SlideCapabilityModule = Object.freeze({
  declaration: "motion",
  slideProp: "__motion",
  slot: "timing",
  always: true,
  compileSlide(_input: CapabilityInput, context) {
    const steps = motionSteps(context.contributions);
    if (steps.length === 0) {
      return { xml: "", extensions: [] };
    }
    return {
      xml: timingXml(steps),
      extensions: [],
      observation: Object.freeze({ steps }) satisfies MotionObservation,
    };
  },
});
