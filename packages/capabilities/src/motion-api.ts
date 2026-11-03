import {
  CapabilityError,
  type BoxProps as CoreBoxProps,
  type CapabilityAnnotation,
} from "@react-pptx/core";
import {
  createContext,
  createElement,
  useContext,
  useId,
  type ReactNode,
} from "react";

export type MotionDirection = "left" | "right" | "top" | "bottom";
export type Easing = {
  readonly accelerate?: number;
  readonly decelerate?: number;
};

export type EffectPreset =
  | "fade"
  | "wipe"
  | "zoom"
  | "fly"
  | "floatUp"
  | "spinIn"
  | "stretch"
  | "spin"
  | "breathe";

export type RepeatCount = number | "forever";

export type MotionGranularity =
  | "paragraph"
  | "word"
  | "letter"
  | "series"
  | "category";

export type MotionTargetRange =
  | { readonly type: "paragraph"; readonly start: number; readonly end: number }
  | {
      readonly type: "chart";
      readonly step: "gridLegend" | "series" | "category";
      readonly seriesIndex: number;
      readonly categoryIndex: number;
    };

/** 未知起点指向效果、同对象效果集合或真实作用域完成事件。 */
export type MotionStartDependency =
  | {
      readonly targetId: number;
      readonly effectIndex: number;
      readonly partIndex?: number;
      readonly point: "end";
      readonly offsetMs: number;
    }
  | {
      readonly targetId: number;
      readonly group: "effects";
      readonly point: "end";
      readonly offsetMs: number;
    }
  | {
      readonly scopeId: string;
      readonly point: "end";
      readonly offsetMs: number;
    };

export type Effect = {
  readonly type: "effect";
  readonly kind: "entrance" | "emphasis";
  readonly preset: EffectPreset;
  /** 单次原生行为时长；逐词或逐字展开的完成总长保持未知。 */
  readonly duration: number;
  readonly delay: number;
  readonly from?: MotionDirection | number;
  readonly fromColor?: string;
  readonly toColor?: string;
  readonly easing?: Easing;
  /** 文字按段落、词或字展开，图表按系列或分类展开。 */
  readonly by?: MotionGranularity;
  /** 逐词或逐字效果中，相邻原生单元的启动间隔。 */
  readonly intervalMs?: number;
};

export type MotionScopeMode = "parallel" | "sequence" | "stagger";

/** 作用域自然完成状态；无自然完成与时长未知保持区分。 */
export type MotionCompletionState =
  | "known"
  | "unknown"
  | "no-natural-completion";

/** 指向某个显式编排作用域边界或其公开命名时刻。 */
export type MotionMoment = {
  readonly type: "motion-moment";
  readonly scope: string;
  readonly point: "start" | "end" | "named";
  readonly name?: string;
};

export type MotionMomentOffset = {
  readonly type: "motion-offset";
  readonly moment: MotionMoment;
  readonly offsetMs: number;
};

export type MotionTiming = MotionMoment | MotionMomentOffset;

/** 一次 hook 调用代表一个组件实例的作用域身份；同名组件实例各有独立身份。 */
export type MotionScopeHandle = {
  readonly type: "motion-scope-handle";
  readonly identity: string;
  readonly start: MotionMoment;
  readonly end: MotionMoment;
  cue(name: string): MotionMoment;
};

export type EntranceEffect = Effect & { readonly kind: "entrance" };
export type EmphasisEffect = Effect & { readonly kind: "emphasis" };
export type RepeatedEffect<E extends Effect = Effect> = {
  readonly type: "repeated-effect";
  readonly effect: E;
  readonly count: RepeatCount;
};
export type EntranceEffectDeclaration =
  | EntranceEffect
  | RepeatedEffect<EntranceEffect>;
export type EmphasisEffectDeclaration =
  | EmphasisEffect
  | RepeatedEffect<EmphasisEffect>;

export type BreatheOptions = {
  /** 完整的暗→亮→暗周期，单位为偶数毫秒。 */
  readonly period: number;
  /** 周期起点的暗色；播放开始时动画从此颜色起步。 */
  readonly from: string;
  /** 半周期终点的亮色。 */
  readonly to: string;
  /** 相对所属触发起点的延迟，单位为毫秒。 */
  readonly delay?: number;
  /** 加速和减速各占完整周期的比例，合计不能超过 1。 */
  readonly easing?: Easing;
};

export type BreatheEffect = EmphasisEffect & {
  readonly preset: "breathe";
  readonly fromColor: string;
  readonly toColor: string;
};

/** repeat() 校验后的内部效果；作者必须通过 RepeatedEffect 声明轮数。 */
export type NormalizedEffect = Effect & {
  readonly repeatCount?: RepeatCount;
};

export type EffectOptions = {
  /** 单次效果行为时长，单位为毫秒；细粒度展开的完成总长不由此值表示。 */
  readonly duration?: number;
  /** 相对所属触发起点的延迟，单位为毫秒。 */
  readonly delay?: number;
  /** 加速和减速各占时长的比例，合计不能超过 1。 */
  readonly easing?: Easing;
  /** 文字逐词或逐字启动间隔；省略时为单次行为时长的十分之一。 */
  readonly intervalMs?: number;
  /** 按原生段落、词、字、图表系列或分类逐项展开。 */
  readonly by?: MotionGranularity;
};

export type MotionProps = {
  /** 对象进入效果；数组中的效果从同一触发起点并行开始。 */
  readonly enter?:
    | EntranceEffectDeclaration
    | readonly EntranceEffectDeclaration[];
  /** 对已显示对象应用的强调效果。 */
  readonly emphasis?: EmphasisEffectDeclaration;
  /** 设为 false 可禁用此对象及其后代的动效。 */
  readonly motion?: boolean;
};

export type MotionScopeDeclaration = {
  readonly identity: string;
  readonly parent?: string;
  readonly step?: StepDeclaration;
  readonly mode: MotionScopeMode;
  readonly delay: number;
  readonly each?: number;
  readonly startAt?: MotionTiming;
  readonly publishStart?: MotionMoment;
  readonly publishEnd?: MotionMoment;
};

export type MotionScopeProps = {
  /** 由当前组件实例的 useMotionScope() 创建。 */
  readonly scope: MotionScopeHandle;
  /** 子成员的组合关系，缺省为并行。 */
  readonly mode?: MotionScopeMode;
  /** stagger 每个直接成员的启动间隔，单位为毫秒。 */
  readonly each?: number;
  /** 相对父作用域起点的偏移，单位为毫秒。 */
  readonly delay?: number;
  /** 从另一个作用域时刻启动，可用 after() 表达偏移。 */
  readonly startAt?: MotionTiming;
  /** 将本作用域起点发布到同属一条编排的命名时刻。 */
  readonly publishStart?: MotionMoment;
  /** 将本作用域完成时刻发布到同属一条编排的命名时刻。 */
  readonly publishEnd?: MotionMoment;
  /** 设为 false 可禁用本作用域及其后代动效。 */
  readonly motion?: boolean;
  readonly children?: ReactNode;
};

/** 显式作用域的布局容器外观；Box 本身仍不生成原生对象。 */
export type BoxProps = CoreBoxProps & Pick<MotionProps, "enter" | "motion">;

export type StepProps = {
  /** 相同正整数的点击步合并；省略时每个 Step 独立触发。 */
  readonly n?: number;
  readonly children?: ReactNode;
};

type StepDeclaration =
  | { readonly mode: "automatic"; readonly identity: string }
  | { readonly mode: "explicit"; readonly number: number };

type MotionContextValue = {
  readonly enabled: boolean;
  readonly step?: StepDeclaration;
  readonly scopes: readonly MotionScopeDeclaration[];
  readonly defaultEnter?:
    | EntranceEffectDeclaration
    | readonly EntranceEffectDeclaration[];
};

const motionContext = createContext<MotionContextValue>({
  enabled: true,
  scopes: [],
});

function moment(
  scope: string,
  point: MotionMoment["point"],
  name?: string,
): MotionMoment {
  return Object.freeze({
    type: "motion-moment",
    scope,
    point,
    ...(name === undefined ? {} : { name }),
  });
}

/** 在任意时刻句柄之后安排一个明确的非负偏移。 */
export function after(
  timing: MotionTiming,
  offsetMs: number,
): MotionMomentOffset {
  const offset =
    timing.type === "motion-offset" ? timing.offsetMs + offsetMs : offsetMs;
  return Object.freeze({
    type: "motion-offset",
    moment: timing.type === "motion-offset" ? timing.moment : timing,
    offsetMs: offset,
  });
}

/** 为当前 React 组件实例创建显式编排作用域和可发布的命名时刻。 */
export function useMotionScope(): MotionScopeHandle {
  const identity = useId();
  const cue = (name: string): MotionMoment => moment(identity, "named", name);
  return Object.freeze({
    type: "motion-scope-handle",
    identity,
    start: moment(identity, "start"),
    end: moment(identity, "end"),
    cue,
  });
}

/** 显式编排作用域只声明时序关系，不创建布局或原生对象。 */
export function MotionScope({
  scope,
  mode = "parallel",
  each,
  delay = 0,
  startAt,
  publishStart,
  publishEnd,
  motion,
  children,
}: MotionScopeProps): ReactNode {
  const parent = useContext(motionContext);
  const parentScope = parent.scopes.at(-1);
  const declaration: MotionScopeDeclaration = Object.freeze({
    identity: scope.identity,
    ...(parentScope === undefined ? {} : { parent: parentScope.identity }),
    ...(parent.step === undefined ? {} : { step: parent.step }),
    mode,
    delay,
    ...(each === undefined ? {} : { each }),
    ...(startAt === undefined ? {} : { startAt }),
    ...(publishStart === undefined ? {} : { publishStart }),
    ...(publishEnd === undefined ? {} : { publishEnd }),
  });
  const context = Object.freeze({
    ...parent,
    enabled: parent.enabled && motion !== false,
    scopes: Object.freeze([...parent.scopes, declaration]),
  });
  return createElement(
    motionContext.Provider,
    { value: context },
    createElement("rp:motion-scope", { __motionScope: declaration }, children),
  );
}

/** 可见 Box 的自有背景和未自带入场的后代可继承默认入场。 */
export function Box({
  enter,
  motion,
  children,
  ...props
}: BoxProps): ReactNode {
  const parent = useContext(motionContext);
  const defaultEnter = enter === undefined ? parent.defaultEnter : enter;
  const context = Object.freeze({
    ...parent,
    enabled: parent.enabled && motion !== false,
    ...(defaultEnter === undefined ? {} : { defaultEnter }),
  });
  return createElement(
    motionContext.Provider,
    { value: context },
    createElement("rp:box", { ...props, __motionContext: context }, children),
  );
}

function declaration(
  kind: Effect["kind"],
  preset: EffectPreset,
  options: EffectOptions & object,
  defaults: { readonly duration: number },
  extra: Partial<Pick<Effect, "from">> = {},
): Effect {
  const easing = options.easing as unknown;
  const normalizedEasing =
    typeof easing === "object" && easing !== null && !Array.isArray(easing)
      ? Object.freeze({ ...(easing as Easing) })
      : easing;
  return Object.freeze({
    ...options,
    type: "effect",
    kind,
    preset,
    duration:
      options.duration === undefined ? defaults.duration : options.duration,
    delay: options.delay === undefined ? 0 : options.delay,
    ...extra,
    ...(easing === undefined ? {} : { easing: normalizedEasing }),
  }) as Effect;
}

/** 创建纯淡入声明；缺省时长为 500 毫秒。 */
export function fade(options: EffectOptions = {}): EntranceEffect {
  return declaration("entrance", "fade", options, {
    duration: 500,
  }) as EntranceEffect;
}

/** 从指定边缘擦入；缺省方向为 bottom。 */
export function wipe(
  options: EffectOptions & { readonly from?: MotionDirection } = {},
): EntranceEffect {
  return declaration(
    "entrance",
    "wipe",
    options,
    { duration: 500 },
    {
      from: options.from === undefined ? "bottom" : options.from,
    },
  ) as EntranceEffect;
}

/** 从原点放大并淡入。 */
export function zoom(options: EffectOptions = {}): EntranceEffect {
  return declaration("entrance", "zoom", options, {
    duration: 500,
  }) as EntranceEffect;
}

/** 从指定页面边缘飞入；缺省方向为 bottom。 */
export function fly(
  options: EffectOptions & { readonly from?: MotionDirection } = {},
): EntranceEffect {
  return declaration(
    "entrance",
    "fly",
    options,
    { duration: 500 },
    {
      from: options.from === undefined ? "bottom" : options.from,
    },
  ) as EntranceEffect;
}

/** 淡入并向上浮动。 */
export function floatUp(options: EffectOptions = {}): EntranceEffect {
  return declaration("entrance", "floatUp", options, {
    duration: 500,
  }) as EntranceEffect;
}

/** 放大、旋转并淡入。 */
export function spinIn(options: EffectOptions = {}): EntranceEffect {
  return declaration("entrance", "spinIn", options, {
    duration: 500,
  }) as EntranceEffect;
}

/** 从中心线展开。 */
export function stretch(options: EffectOptions = {}): EntranceEffect {
  return declaration("entrance", "stretch", options, {
    duration: 500,
  }) as EntranceEffect;
}

/** 从指定起始角度旋转到对象原生朝向；缺省从 360 度开始。 */
export function spin(
  options: EffectOptions & { readonly from?: number } = {},
): EmphasisEffect {
  return declaration(
    "emphasis",
    "spin",
    options,
    { duration: 1000 },
    {
      from: options.from === undefined ? 360 : options.from,
    },
  ) as EmphasisEffect;
}

// 保留结构化 JavaScript 输入中的额外选项，交由正式效果校验诊断。
const breatheUnknownParameters = Symbol("breathe unknown parameters");
const breatheOptionNames = new Set([
  "period",
  "delay",
  "from",
  "to",
  "easing",
  "by",
  "intervalMs",
]);

/** 声明 Shape 填充色暗→亮的半周期；原生 autoRev 完成反向，形成完整周期。 */
export function breathe(options: BreatheOptions): BreatheEffect {
  const { period, delay = 0, from, to } = options;
  const optionRecord = options as unknown as Record<string, unknown>;
  const by = optionRecord["by"];
  const intervalMs = optionRecord["intervalMs"];
  const unknownParameters = Object.keys(optionRecord).filter(
    (key) => !breatheOptionNames.has(key),
  );
  const rawEasing: unknown = options.easing;
  const easing = isRecord(rawEasing)
    ? Object.freeze({ ...rawEasing })
    : rawEasing;
  return Object.freeze({
    type: "effect",
    kind: "emphasis",
    preset: "breathe",
    duration: period,
    delay,
    fromColor: from,
    toColor: to,
    ...(easing === undefined ? {} : { easing: easing as Easing }),
    ...(by === undefined ? {} : { by }),
    ...(intervalMs === undefined ? {} : { intervalMs }),
    ...(unknownParameters.length === 0
      ? {}
      : { [breatheUnknownParameters]: unknownParameters }),
  }) as unknown as BreatheEffect;
}

/** 将受支持效果包装为指定总轮数；forever 只适用于 breathe。 */
export function repeat<E extends Effect>(
  effect: E,
  count: RepeatCount,
): RepeatedEffect<E> {
  return Object.freeze({ type: "repeated-effect", effect, count });
}

/** 一次点击触发一组对象；相同编号的声明共用点击步。 */
export function Step({ n, children }: StepProps): ReactNode {
  const parent = useContext(motionContext);
  const identity = useId();
  const step: StepDeclaration =
    n === undefined
      ? { mode: "automatic", identity }
      : { mode: "explicit", number: n };
  return createElement(
    motionContext.Provider,
    { value: { ...parent, step } },
    children,
  );
}

/** 将当前 Step 与禁用状态传给正式原生对象宿主节点。 */
export function motionElement(
  type: string,
  props: MotionProps & { readonly children?: ReactNode },
): ReactNode {
  const parent = useContext(motionContext);
  const enabled = parent.enabled && props.motion !== false;
  const context = Object.freeze({ ...parent, enabled });
  const host = createElement(
    type,
    {
      ...props,
      __motionContext: context,
    },
    props.children,
  );
  if (enabled === parent.enabled) return host;
  return createElement(motionContext.Provider, { value: context }, host);
}

type MotionTarget = {
  readonly kind: string;
  readonly name: string;
  readonly testId?: string;
  readonly textBody: boolean;
  readonly hasTextCharacters?: boolean;
  readonly filled: boolean;
  readonly paragraphCount?: number;
  readonly chartSeriesCount?: number;
  readonly chartCategoryCount?: number;
};

type MotionValue = {
  readonly target: MotionTarget;
  readonly step?: StepDeclaration;
  readonly scopes: readonly MotionScopeDeclaration[];
  readonly effects: readonly NormalizedEffect[];
};

export type MotionEffectObservation = {
  readonly target: {
    readonly id: number;
    readonly name: string;
    readonly testId?: string;
    readonly range?: MotionTargetRange;
  };
  readonly kind: Effect["kind"];
  readonly preset: EffectPreset;
  /** 单次效果时长；breathe 表示完整往返周期，不代表细粒度展开总长。 */
  readonly durationMs: number;
  /** 完整呼吸周期的总轮数；forever 表示本页持续且无自然完成。 */
  readonly repeatCount?: RepeatCount;
  readonly delayMs: number;
  readonly by?: MotionGranularity;
  readonly intervalMs?: number;
  readonly from?: MotionDirection | number;
  readonly fromColor?: string;
  readonly toColor?: string;
  readonly easing?: Easing;
  readonly startMs?: number;
  readonly endMs?: number;
  readonly completion?: Exclude<MotionCompletionState, "known">;
  readonly startAfter?: MotionStartDependency;
  /** 未知起点需要多项完成事件共同约束时，列出全部真实来源及各自偏移。 */
  readonly startAfterAll?: readonly MotionStartDependency[];
};

export type MotionStepObservation = {
  readonly trigger: "auto" | "click";
  readonly number?: number;
  readonly effects: readonly MotionEffectObservation[];
};

export type MotionScopeObservation = {
  readonly id: string;
  readonly parentId?: string;
  readonly mode: MotionScopeMode;
  readonly eachMs?: number;
  readonly delayMs: number;
  readonly startMs?: number;
  readonly endMs?: number;
  readonly completion: MotionCompletionState;
  readonly startAfter?: MotionStartDependency;
  /** 未知起点需要多项完成事件共同约束时，列出全部真实来源及各自偏移。 */
  readonly startAfterAll?: readonly MotionStartDependency[];
  readonly members: readonly (
    | { readonly kind: "scope"; readonly id: string }
    | { readonly kind: "object"; readonly id: number }
  )[];
};

export type MotionObservation = {
  readonly steps: readonly MotionStepObservation[];
  readonly scopes?: readonly MotionScopeObservation[];
};

function fail(code: string, message: string): never {
  throw new CapabilityError(code, message);
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

const directions = new Set<MotionDirection>(["left", "right", "top", "bottom"]);
const presets = new Set<EffectPreset>([
  "fade",
  "wipe",
  "zoom",
  "fly",
  "floatUp",
  "spinIn",
  "stretch",
  "spin",
  "breathe",
]);
const presetKinds: Readonly<Record<EffectPreset, Effect["kind"]>> = {
  fade: "entrance",
  wipe: "entrance",
  zoom: "entrance",
  fly: "entrance",
  floatUp: "entrance",
  spinIn: "entrance",
  stretch: "entrance",
  spin: "emphasis",
  breathe: "emphasis",
};
const minimumRotationUnits = -2_147_483_648;
const maximumRotationUnits = 2_147_483_647;
const effectKeys = [
  "type",
  "kind",
  "preset",
  "duration",
  "delay",
  "from",
  "easing",
  "by",
  "intervalMs",
  "fromColor",
  "toColor",
] as const;

const motionGranularities = new Set<MotionGranularity>([
  "paragraph",
  "word",
  "letter",
  "series",
  "category",
]);

// ST_TLTime 使用 unsignedInt；repeatCount 的单位是千分之一轮。
export const maximumTimeNodeValue = 4_294_967_295;
const maxRepeatCount = Math.floor(maximumTimeNodeValue / 1000);

function normalizeColor(value: unknown, name: string): string {
  if (typeof value !== "string" || !/^#[0-9a-f]{6}$/iu.test(value)) {
    return fail("invalid-breathe-color", `${name} 必须是 #RRGGBB 颜色`);
  }
  return value.toUpperCase();
}

function relativeLuminance(color: string): number {
  const channels = [1, 3, 5].map(
    (offset) => Number.parseInt(color.slice(offset, offset + 2), 16) / 255,
  );
  const linear = channels.map((channel) =>
    channel <= 0.04045 ? channel / 12.92 : ((channel + 0.055) / 1.055) ** 2.4,
  );
  return (
    0.2126 * (linear[0] ?? 0) +
    0.7152 * (linear[1] ?? 0) +
    0.0722 * (linear[2] ?? 0)
  );
}

function validatedEffect(
  value: unknown,
  expected: Effect["kind"],
): NormalizedEffect {
  let repeatCount: RepeatCount | undefined;
  let repeatDeclared = false;
  let effectValue = value;
  if (isRecord(value) && value["type"] === "repeated-effect") {
    repeatDeclared = true;
    const unknownRepeatField = Object.keys(value).find(
      (key) => key !== "type" && key !== "effect" && key !== "count",
    );
    if (unknownRepeatField !== undefined) {
      return fail(
        "unknown-parameter",
        `repeat 声明含未实现参数 ${unknownRepeatField}`,
      );
    }
    if (!isRecord(value["effect"]) || value["effect"]["type"] !== "effect") {
      return fail("invalid-repeat", "repeat 必须直接包装单个效果声明");
    }
    repeatCount = value["count"] as RepeatCount;
    effectValue = value["effect"];
  }
  if (!isRecord(effectValue) || effectValue["type"] !== "effect") {
    return fail(
      "invalid-motion",
      expected === "entrance"
        ? "enter 必须是 fade()、wipe() 等入场效果声明"
        : "emphasis 必须是 spin() 等强调效果声明",
    );
  }
  const unsupportedBreatheParameters = (
    effectValue as {
      readonly [breatheUnknownParameters]?: readonly string[];
    }
  )[breatheUnknownParameters];
  if (unsupportedBreatheParameters?.length) {
    return fail(
      "unknown-parameter",
      `breathe 声明含未实现参数 ${unsupportedBreatheParameters[0]}`,
    );
  }
  const unknown = Object.keys(effectValue).find(
    (key) => !(effectKeys as readonly string[]).includes(key),
  );
  if (unknown !== undefined) {
    return fail("unknown-parameter", `动画效果含未实现参数 ${unknown}`);
  }
  if (effectValue["kind"] !== expected) {
    return fail(
      "invalid-motion",
      expected === "entrance"
        ? `${String(effectValue["preset"])}() 是强调效果，应传给 emphasis`
        : `${String(effectValue["preset"])}() 是入场效果，应传给 enter`,
    );
  }
  const preset = effectValue["preset"];
  if (typeof preset !== "string" || !presets.has(preset as EffectPreset)) {
    return fail("unsupported-effect", `不支持动画效果 ${String(preset)}`);
  }
  if (
    preset !== "breathe" &&
    presetKinds[preset as EffectPreset] !== expected
  ) {
    return fail(
      "invalid-motion",
      expected === "entrance"
        ? `${preset}() 是强调效果，应传给 emphasis`
        : `${preset}() 是入场效果，应传给 enter`,
    );
  }
  const duration = effectValue["duration"];
  const delay = effectValue["delay"];
  if (
    typeof duration !== "number" ||
    !Number.isFinite(duration) ||
    duration <= 0 ||
    Math.round(duration) < 1 ||
    !Number.isSafeInteger(Math.round(duration))
  ) {
    return fail("invalid-duration", "动画时长必须能写成有限正整数毫秒");
  }
  if (
    typeof delay !== "number" ||
    !Number.isFinite(delay) ||
    delay < 0 ||
    !Number.isSafeInteger(Math.round(delay)) ||
    !Number.isFinite(Math.round(delay) + Math.round(duration))
  ) {
    return fail("invalid-delay", "动画延迟必须能写成有限非负整数毫秒");
  }
  if (preset === "breathe" && Math.round(delay) > maximumTimeNodeValue) {
    return fail("invalid-delay", "breathe 延迟超出原生计时值范围");
  }
  const by = effectValue["by"];
  const interval = effectValue["intervalMs"];
  if (
    by !== undefined &&
    (typeof by !== "string" ||
      !motionGranularities.has(by as MotionGranularity))
  ) {
    return fail("invalid-motion", "by 必须是受支持的文字或图表粒度");
  }
  if (
    interval !== undefined &&
    ((by !== "word" && by !== "letter") ||
      typeof interval !== "number" ||
      !Number.isSafeInteger(interval) ||
      interval < 1)
  ) {
    return fail(
      "invalid-motion-interval",
      "intervalMs 仅适用于逐词或逐字效果，且必须是正整数毫秒",
    );
  }
  if (
    (preset === "wipe" || preset === "fly") &&
    (typeof effectValue["from"] !== "string" ||
      !directions.has(effectValue["from"] as MotionDirection))
  ) {
    return fail("invalid-direction", `${preset} 的 from 必须是有效方向`);
  }
  if (preset === "spin") {
    const from = effectValue["from"];
    const units = typeof from === "number" ? Math.round(from * 60000) : NaN;
    if (
      typeof from !== "number" ||
      !Number.isFinite(from) ||
      !Number.isSafeInteger(units) ||
      units < minimumRotationUnits ||
      units > maximumRotationUnits
    ) {
      return fail(
        "invalid-rotation",
        "spin 的 from 换算后必须落在有符号 32 位角度范围内",
      );
    }
  }
  if (
    preset !== "wipe" &&
    preset !== "fly" &&
    preset !== "spin" &&
    effectValue["from"] !== undefined
  ) {
    return fail("invalid-motion", `${preset} 不接受 from 参数`);
  }
  const easing = effectValue["easing"];
  if (easing !== undefined) {
    if (!isRecord(easing)) {
      return fail("invalid-easing", "easing 必须是对象");
    }
    const unknownEasing = Object.keys(easing).find(
      (key) => key !== "accelerate" && key !== "decelerate",
    );
    if (unknownEasing !== undefined) {
      return fail("unknown-parameter", `easing 含未实现参数 ${unknownEasing}`);
    }
    const accelerate = easing["accelerate"] ?? 0;
    const decelerate = easing["decelerate"] ?? 0;
    if (
      typeof accelerate !== "number" ||
      !Number.isFinite(accelerate) ||
      typeof decelerate !== "number" ||
      !Number.isFinite(decelerate) ||
      accelerate < 0 ||
      decelerate < 0 ||
      accelerate + decelerate > 1
    ) {
      return fail(
        "invalid-easing",
        "easing 的加速、减速比例必须为非负数，且合计不超过 1",
      );
    }
  }
  let fromColor: string | undefined;
  let toColor: string | undefined;
  if (preset === "breathe") {
    if (expected !== "emphasis") {
      return fail("invalid-motion", "breathe() 必须传给 emphasis");
    }
    if (
      typeof duration !== "number" ||
      !Number.isSafeInteger(duration) ||
      duration < 2 ||
      duration % 2 !== 0 ||
      duration / 2 > maximumTimeNodeValue
    ) {
      return fail(
        "invalid-breathe-period",
        "breathe 的 period 必须是大于 1 的偶数毫秒，且半周期可写入原生时长",
      );
    }
    fromColor = normalizeColor(effectValue["fromColor"], "breathe.from");
    toColor = normalizeColor(effectValue["toColor"], "breathe.to");
    if (relativeLuminance(toColor) <= relativeLuminance(fromColor)) {
      return fail(
        "invalid-breathe-colors",
        "breathe.to 必须比 breathe.from 明亮",
      );
    }
  } else if (
    effectValue["fromColor"] !== undefined ||
    effectValue["toColor"] !== undefined
  ) {
    return fail("invalid-motion", `${String(preset)} 不接受呼吸灯颜色参数`);
  }

  if (repeatDeclared) {
    if (preset !== "spin" && preset !== "breathe") {
      return fail(
        "unsupported-repeat",
        `${String(preset)} 尚不支持重复；当前仅支持 spin 和 breathe`,
      );
    }
    if (repeatCount === "forever") {
      if (preset !== "breathe") {
        return fail("unsupported-repeat", "forever 当前仅支持完整呼吸周期");
      }
    } else if (
      typeof repeatCount !== "number" ||
      !Number.isSafeInteger(repeatCount) ||
      repeatCount < 1 ||
      repeatCount > maxRepeatCount
    ) {
      return fail(
        "invalid-repeat-count",
        `repeat 次数必须是 1 到 ${maxRepeatCount} 的总轮数`,
      );
    }
    if (
      typeof repeatCount === "number" &&
      (!Number.isSafeInteger(
        Math.round(delay) + Math.round(duration) * repeatCount,
      ) ||
        Math.round(delay) > maximumTimeNodeValue ||
        (preset === "breathe" ? duration / 2 : Math.round(duration)) >
          maximumTimeNodeValue ||
        repeatCount * 1000 > maximumTimeNodeValue)
    ) {
      return fail(
        "invalid-repeat-duration",
        "重复后的总时长必须保持为精确整数毫秒且单轮不超过原生计时范围",
      );
    }
    if (by !== undefined) {
      return fail(
        "unsupported-repeat-combination",
        "重复效果当前不能与逐段、逐词、逐字或图表粒度组合",
      );
    }
  }

  return {
    ...(effectValue as unknown as Effect),
    ...(repeatCount === undefined ? {} : { repeatCount }),
    ...(fromColor === undefined ? {} : { fromColor }),
    ...(toColor === undefined ? {} : { toColor }),
  };
}

function stepDeclaration(value: unknown): StepDeclaration | undefined {
  if (value === undefined) return undefined;
  if (!isRecord(value)) return fail("invalid-step", "Step 声明无效");
  if (value["mode"] === "automatic") {
    if (typeof value["identity"] !== "string" || value["identity"] === "") {
      return fail("invalid-step", "自动点击步缺少实例身份");
    }
    return value as unknown as StepDeclaration;
  }
  const number = value["number"];
  if (
    value["mode"] !== "explicit" ||
    typeof number !== "number" ||
    !Number.isSafeInteger(number) ||
    number < 1
  ) {
    return fail("invalid-step", "显式点击步编号必须是正整数");
  }
  return value as unknown as StepDeclaration;
}

function listed(value: unknown, many: boolean): readonly unknown[] {
  if (value === undefined) return [];
  return many && Array.isArray(value) ? value : [value];
}

function validatedMoment(value: unknown): MotionMoment {
  if (!isRecord(value) || value["type"] !== "motion-moment") {
    return fail("invalid-moment", "时刻句柄必须来自 useMotionScope()。");
  }
  const scope = value["scope"];
  const point = value["point"];
  const name = value["name"];
  if (
    typeof scope !== "string" ||
    scope === "" ||
    (point !== "start" && point !== "end" && point !== "named") ||
    (point === "named"
      ? typeof name !== "string" || name.trim() === ""
      : name !== undefined)
  ) {
    return fail("invalid-moment", "时刻句柄的作用域、类型或名称无效");
  }
  return value as unknown as MotionMoment;
}

function validatedTiming(value: unknown): MotionTiming {
  if (isRecord(value) && value["type"] === "motion-offset") {
    const offset = value["offsetMs"];
    if (
      typeof offset !== "number" ||
      !Number.isSafeInteger(offset) ||
      offset < 0
    ) {
      return fail("invalid-delay", "时刻偏移必须是有限非负整数毫秒");
    }
    return Object.freeze({
      type: "motion-offset",
      moment: validatedMoment(value["moment"]),
      offsetMs: offset,
    });
  }
  return validatedMoment(value);
}

function validatedPublisher(value: unknown): MotionMoment | undefined {
  if (value === undefined) return undefined;
  const momentValue = validatedMoment(value);
  if (momentValue.point !== "named") {
    return fail(
      "invalid-moment",
      "publishStart 和 publishEnd 需要命名时刻句柄",
    );
  }
  return momentValue;
}

/** 将显式作用域宿主编成 motion 页面能力的透明声明。 */
export function motionScopeAnnotation(
  input: Readonly<Record<string, unknown>>,
): CapabilityAnnotation {
  const unknown = Object.keys(input).find((key) => key !== "__motionScope");
  if (unknown !== undefined) {
    return fail("unknown-parameter", `MotionScope 含未实现参数 ${unknown}`);
  }
  const value = input["__motionScope"];
  if (!isRecord(value)) {
    return fail("invalid-scope", "MotionScope 缺少作用域声明");
  }
  const identity = value["identity"];
  const parent = value["parent"];
  const step = value["step"];
  const mode = value["mode"];
  const delay = value["delay"];
  const each = value["each"];
  let normalizedStep: StepDeclaration | undefined;
  try {
    normalizedStep = step === undefined ? undefined : stepDeclaration(step);
  } catch {
    return fail("invalid-scope", "MotionScope 的点击步身份无效");
  }
  if (
    typeof identity !== "string" ||
    identity === "" ||
    (parent !== undefined && typeof parent !== "string") ||
    (step !== undefined && normalizedStep === undefined)
  ) {
    return fail("invalid-scope", "MotionScope 的实例身份无效");
  }
  if (mode !== "parallel" && mode !== "sequence" && mode !== "stagger") {
    return fail(
      "invalid-scope",
      "MotionScope 的 mode 必须是 parallel、sequence 或 stagger",
    );
  }
  if (
    typeof delay !== "number" ||
    !Number.isFinite(delay) ||
    delay < 0 ||
    !Number.isSafeInteger(Math.round(delay))
  ) {
    return fail("invalid-delay", "MotionScope 的 delay 必须是有限非负整数毫秒");
  }
  if (
    (mode === "stagger" &&
      (typeof each !== "number" ||
        !Number.isFinite(each) ||
        each < 0 ||
        !Number.isSafeInteger(Math.round(each)))) ||
    (mode !== "stagger" && each !== undefined)
  ) {
    return fail(
      "invalid-scope",
      "stagger 需要非负整数 each，其他 mode 不接受 each",
    );
  }
  const startAt =
    value["startAt"] === undefined
      ? undefined
      : validatedTiming(value["startAt"]);
  const publishStart = validatedPublisher(value["publishStart"]);
  const publishEnd = validatedPublisher(value["publishEnd"]);
  const scope: MotionScopeDeclaration = Object.freeze({
    identity,
    ...(typeof parent === "string" ? { parent } : {}),
    ...(normalizedStep === undefined ? {} : { step: normalizedStep }),
    mode,
    delay: Math.round(delay),
    ...(typeof each === "number" ? { each: Math.round(each) } : {}),
    ...(startAt === undefined ? {} : { startAt }),
    ...(publishStart === undefined ? {} : { publishStart }),
    ...(publishEnd === undefined ? {} : { publishEnd }),
  });
  return {
    capability: "motion",
    value: Object.freeze({ type: "scope", scope }),
  };
}

function contextValue(value: unknown): MotionContextValue {
  if (value === undefined) return { enabled: true, scopes: [] };
  if (
    !isRecord(value) ||
    typeof value["enabled"] !== "boolean" ||
    !Array.isArray(value["scopes"])
  ) {
    return fail("invalid-motion", "动效上下文无效");
  }
  const step = stepDeclaration(value["step"]);
  const defaultEnter = value["defaultEnter"];
  return {
    enabled: value["enabled"],
    scopes: value["scopes"] as MotionScopeDeclaration[],
    ...(step === undefined ? {} : { step }),
    ...(defaultEnter === undefined
      ? {}
      : {
          defaultEnter: defaultEnter as
            | EntranceEffectDeclaration
            | readonly EntranceEffectDeclaration[],
        }),
  };
}

/** 将对象动画声明转成由 motion 页面能力消费的局部注解。 */
export function motionAnnotations(
  props: Readonly<Record<string, unknown>>,
  target: MotionTarget,
): readonly CapabilityAnnotation[] {
  if (props["motion"] !== undefined && typeof props["motion"] !== "boolean") {
    return fail("invalid-motion", "motion 必须是布尔值");
  }
  const context = contextValue(props["__motionContext"]);
  if (!context.enabled || props["motion"] === false) return [];

  const step = context.step;
  const enter =
    props["enter"] === undefined ? context.defaultEnter : props["enter"];
  const effects = [
    ...listed(enter, true).map((value) => validatedEffect(value, "entrance")),
    ...listed(props["emphasis"], false).map((value) =>
      validatedEffect(value, "emphasis"),
    ),
  ];
  if (
    effects.some(
      ({ preset }) =>
        preset === "breathe" && (target.kind !== "shape" || !target.filled),
    )
  ) {
    return fail(
      "invalid-breathe-target",
      "breathe() 需要带实色填充的原生 Shape；Group 需由实际成员分别声明效果",
    );
  }
  if (effects.length === 0) return [];
  const annotation: CapabilityAnnotation = {
    capability: "motion",
    value: Object.freeze({
      target: Object.freeze({ ...target }),
      ...(step === undefined ? {} : { step }),
      scopes: context.scopes,
      effects: Object.freeze(effects),
    } satisfies MotionValue),
  };
  return [annotation];
}
