import {
  CapabilityError,
  type CapabilityInput,
  type Extension,
  type SlideCapabilityModule,
} from "@react-pptx/core";

/** 转场时长，单位为毫秒；写入 p14:dur 时按最近整数毫秒取整。 */
export type TransitionDuration = number;
export type TransitionDirection = "left" | "right" | "top" | "bottom";
export type StripsDirection = "leftUp" | "rightUp" | "leftDown" | "rightDown";
export type TransitionSpeed = "fast" | "med" | "slow";

type TransitionOptions = { readonly duration?: TransitionDuration };

/** 不依赖扩展命名空间的 Transitional 基础转场。 */
export type StandardTransition =
  | {
      readonly type: "transition";
      readonly kind: "fade";
      readonly duration?: TransitionDuration;
    }
  | {
      readonly type: "transition";
      readonly kind: "strips";
      readonly duration?: TransitionDuration;
      readonly direction: StripsDirection;
    }
  | {
      readonly type: "transition";
      readonly kind: "push";
      readonly duration?: TransitionDuration;
      readonly from: TransitionDirection;
    };

/** p14 Flip 转场；必须给出供其他消费方选择的基础格式 fallback。 */
export type FlipTransition = {
  readonly type: "transition";
  readonly kind: "flip";
  readonly duration?: TransitionDuration;
  readonly direction: "left" | "right";
  readonly fallback: StandardTransition;
};

/** p159 Morph 转场；必须给出供其他消费方选择的基础格式 fallback。 */
export type MorphTransition = {
  readonly type: "transition";
  readonly kind: "morph";
  readonly duration?: TransitionDuration;
  readonly by: "object" | "word" | "character";
  readonly fallback: StandardTransition;
};

/** 页面切换声明；输出目标不允许必需机制时，构建会失败。 */
export type Transition = StandardTransition | FlipTransition | MorphTransition;

type TransitionRecord = Record<string, unknown> & {
  readonly type: unknown;
  readonly kind: unknown;
};

const stripDirections: Record<StripsDirection, string> = {
  leftUp: "lu",
  rightUp: "ru",
  leftDown: "ld",
  rightDown: "rd",
};
const pushDirections: Record<TransitionDirection, string> = {
  left: "l",
  right: "r",
  top: "u",
  bottom: "d",
};
const p14Namespace = "http://schemas.microsoft.com/office/powerpoint/2010/main";
const p159Namespace =
  "http://schemas.microsoft.com/office/powerpoint/2015/09/main";
const markupCompatibilityNamespace =
  "http://schemas.openxmlformats.org/markup-compatibility/2006";
const morphOptions = {
  object: "byObject",
  word: "byWord",
  character: "byChar",
} as const;

function declaration<Kind extends Transition["kind"]>(
  kind: Kind,
  options: object,
): Transition & { readonly kind: Kind } {
  return Object.freeze({
    ...options,
    type: "transition",
    kind,
  }) as Transition & {
    readonly kind: Kind;
  };
}

/** 创建按页面切换声明使用的基础与扩展转场值。 */
export const transitions = Object.freeze({
  /** 使用基础 fade 效果；默认采用 slow 速度。 */
  fade: (options: TransitionOptions = {}): StandardTransition =>
    declaration("fade", options),
  /** 使用朝向指定对角方向的基础 strips 效果。 */
  strips: (
    options: TransitionOptions & { readonly direction?: StripsDirection } = {},
  ): StandardTransition =>
    declaration("strips", {
      ...options,
      direction:
        options.direction === undefined ? "rightDown" : options.direction,
    }),
  /** 使用从页面边缘进入的基础 push 效果。 */
  push: (
    options: TransitionOptions & { readonly from?: TransitionDirection } = {},
  ): StandardTransition =>
    declaration("push", {
      ...options,
      from: options.from === undefined ? "bottom" : options.from,
    }),
  /** 使用 p14 Flip；必须显式提供基础格式 fallback。 */
  flip: (
    options: TransitionOptions & {
      readonly direction?: "left" | "right";
      readonly fallback: StandardTransition;
    },
  ): FlipTransition =>
    declaration("flip", {
      ...options,
      direction: options.direction === undefined ? "right" : options.direction,
    }) as FlipTransition,
  /** 使用 p159 Morph；必须显式提供基础格式 fallback。 */
  morph: (
    options: TransitionOptions & {
      readonly by?: "object" | "word" | "character";
      readonly fallback: StandardTransition;
    },
  ): MorphTransition =>
    declaration("morph", {
      ...options,
      by: options.by === undefined ? "object" : options.by,
    }) as MorphTransition,
});

function fail(code: string, message: string): never {
  throw new CapabilityError(code, message);
}

function asRecord(value: unknown): TransitionRecord {
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    return fail(
      "invalid-transition",
      "Slide.transition 必须是转场 helper 返回的声明值",
    );
  }
  return value as TransitionRecord;
}

function unknownParameter(
  value: TransitionRecord,
  allowed: readonly string[],
): string | undefined {
  return Object.keys(value).find((key) => !allowed.includes(key));
}

function durationOf(value: TransitionRecord): number | undefined {
  const duration = value["duration"];
  if (duration === undefined) return undefined;
  if (
    typeof duration !== "number" ||
    !Number.isFinite(duration) ||
    duration < 0 ||
    duration > Number.MAX_SAFE_INTEGER
  ) {
    return fail(
      "invalid-transition-duration",
      "转场 duration 必须是 0 到 Number.MAX_SAFE_INTEGER 之间的有限毫秒数",
    );
  }
  return duration;
}

function speed(duration: number | undefined): TransitionSpeed {
  if (duration === undefined || duration > 875) return "slow";
  return duration > 600 ? "med" : "fast";
}

function transitionElement(
  value: TransitionRecord,
  duration: number | undefined,
  useDurationExtension: boolean,
): string {
  const kind = value["kind"];
  const transitionSpeed = speed(duration);
  const speedAttribute = ` spd="${transitionSpeed}"`;
  const durationAttribute =
    useDurationExtension && duration !== undefined
      ? ` p14:dur="${Math.round(duration)}"`
      : "";
  if (kind === "fade")
    return `<p:transition${speedAttribute}${durationAttribute}><p:fade/></p:transition>`;
  if (kind === "strips") {
    const direction = value["direction"];
    if (
      typeof direction !== "string" ||
      !Object.hasOwn(stripDirections, direction)
    ) {
      return fail(
        "invalid-transition-direction",
        "strips.direction 必须是四种对角方向之一",
      );
    }
    return `<p:transition${speedAttribute}${durationAttribute}><p:strips dir="${stripDirections[direction as StripsDirection]}"/></p:transition>`;
  }
  if (kind === "push") {
    const from = value["from"];
    if (typeof from !== "string" || !Object.hasOwn(pushDirections, from)) {
      return fail(
        "invalid-transition-direction",
        "push.from 必须是 left、right、top 或 bottom",
      );
    }
    return `<p:transition${speedAttribute}${durationAttribute}><p:push dir="${pushDirections[from as TransitionDirection]}"/></p:transition>`;
  }
  if (kind === "flip") {
    const direction = value["direction"];
    if (direction !== "left" && direction !== "right") {
      return fail(
        "invalid-transition-direction",
        "flip.direction 必须是 left 或 right",
      );
    }
    return `<p:transition${speedAttribute}${durationAttribute}><p14:flip dir="${direction === "left" ? "l" : "r"}"/></p:transition>`;
  }
  if (kind === "morph") {
    const by = value["by"];
    if (by !== "object" && by !== "word" && by !== "character") {
      return fail(
        "invalid-morph-option",
        "morph.by 必须是 object、word 或 character",
      );
    }
    return `<p:transition${speedAttribute}${durationAttribute}><p159:morph option="${morphOptions[by]}"/></p:transition>`;
  }
  return fail("unsupported-transition", `不支持转场类型 ${String(kind)}`);
}

function standardTransition(
  value: TransitionRecord,
  inheritedDuration: number | undefined,
): {
  readonly xml: string;
  readonly observation: Readonly<Record<string, unknown>>;
} {
  const kind = value["kind"];
  if (kind !== "fade" && kind !== "strips" && kind !== "push") {
    return fail(
      "invalid-transition-fallback",
      "fallback 必须是 fade、strips 或 push",
    );
  }
  const allowed =
    kind === "fade"
      ? ["type", "kind", "duration"]
      : kind === "strips"
        ? ["type", "kind", "duration", "direction"]
        : ["type", "kind", "duration", "from"];
  const unknown = unknownParameter(value, allowed);
  if (unknown !== undefined) {
    return fail("unknown-parameter", `转场 fallback 含未知参数 ${unknown}`);
  }
  const explicitDuration = durationOf(value);
  const effectiveDuration = explicitDuration ?? inheritedDuration;
  return {
    xml: transitionElement(value, effectiveDuration, false),
    observation: Object.freeze({
      kind,
      ...(kind === "strips" ? { direction: value["direction"] } : {}),
      ...(kind === "push" ? { from: value["from"] } : {}),
      speed: speed(effectiveDuration),
    }),
  };
}

function compileTransition(value: unknown): {
  readonly xml: string;
  readonly extensions: readonly Extension[];
  readonly observation: Readonly<Record<string, unknown>>;
} {
  const transition = asRecord(value);
  if (transition["type"] !== "transition") {
    return fail(
      "invalid-transition",
      "Slide.transition 缺少有效的转场声明标记",
    );
  }
  const kind = transition["kind"];
  if (
    kind !== "fade" &&
    kind !== "strips" &&
    kind !== "push" &&
    kind !== "flip" &&
    kind !== "morph"
  ) {
    return fail("unsupported-transition", `不支持转场类型 ${String(kind)}`);
  }
  const allowed =
    kind === "fade"
      ? ["type", "kind", "duration"]
      : kind === "strips"
        ? ["type", "kind", "duration", "direction"]
        : kind === "push"
          ? ["type", "kind", "duration", "from"]
          : kind === "flip"
            ? ["type", "kind", "duration", "direction", "fallback"]
            : ["type", "kind", "duration", "by", "fallback"];
  const unknown = unknownParameter(transition, allowed);
  if (unknown !== undefined) {
    return fail("unknown-parameter", `转场声明含未知参数 ${unknown}`);
  }
  const duration = durationOf(transition);
  const requiresFallback = kind === "flip" || kind === "morph";
  let fallback: ReturnType<typeof standardTransition> | undefined;
  if (requiresFallback) {
    if (!Object.hasOwn(transition, "fallback")) {
      return fail(
        "missing-transition-fallback",
        `${kind} 转场必须显式声明 Transitional fallback`,
      );
    }
    const fallbackValue = asRecord(transition["fallback"]);
    if (fallbackValue["type"] !== "transition") {
      return fail("invalid-transition-fallback", "转场 fallback 声明无效");
    }
    fallback = standardTransition(fallbackValue, duration);
  } else if (duration !== undefined) {
    fallback = standardTransition(transition, duration);
  }
  const effectExtensions: Extension[] =
    kind === "flip" ? ["p14"] : kind === "morph" ? ["p159"] : [];
  const extensions = [
    ...new Set<Extension>([
      ...(duration === undefined ? [] : ["p14" as const]),
      ...effectExtensions,
    ]),
  ];
  const standardXml = transitionElement(transition, duration, false);
  const observation: Record<string, unknown> = {
    kind,
    ...(kind === "strips" ? { direction: transition["direction"] } : {}),
    ...(kind === "push" ? { from: transition["from"] } : {}),
    ...(kind === "flip" ? { direction: transition["direction"] } : {}),
    ...(kind === "morph" ? { by: transition["by"] } : {}),
    speed: speed(duration),
    ...(duration === undefined ? {} : { durationMs: Math.round(duration) }),
    ...(fallback === undefined ? {} : { fallback: fallback.observation }),
  };
  if (extensions.length === 0) {
    return {
      xml: standardXml,
      extensions,
      observation: Object.freeze(observation),
    };
  }
  const exactXml = transitionElement(transition, duration, true);
  const fallbackXml = fallback?.xml ?? standardXml;
  const requires = extensions.join(" ");
  const declarations = [
    ...(extensions.includes("p14") ? [`xmlns:p14="${p14Namespace}"`] : []),
    ...(extensions.includes("p159") ? [`xmlns:p159="${p159Namespace}"`] : []),
  ].join(" ");
  return {
    xml: `<mc:AlternateContent xmlns:mc="${markupCompatibilityNamespace}" ${declarations}><mc:Choice Requires="${requires}">${exactXml}</mc:Choice><mc:Fallback>${fallbackXml}</mc:Fallback></mc:AlternateContent>`,
    extensions,
    observation: Object.freeze(observation),
  };
}

export const transitionCapability: SlideCapabilityModule = Object.freeze({
  declaration: "transition",
  slideProp: "transition",
  slot: "transition",
  compileSlide(input: CapabilityInput) {
    const transition = compileTransition(input.props["transition"]);
    return {
      xml: transition.xml,
      extensions: transition.extensions,
      observation: transition.observation,
    };
  },
});
