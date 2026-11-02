import { CapabilityError, type CapabilityAnnotation } from "@react-pptx/core";
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
  | "spin";

export type Effect = {
  readonly type: "effect";
  readonly kind: "entrance" | "emphasis";
  readonly preset: EffectPreset;
  readonly duration: number;
  readonly delay: number;
  readonly from?: MotionDirection | number;
  readonly easing?: Easing;
};

export type EntranceEffect = Effect & { readonly kind: "entrance" };
export type EmphasisEffect = Effect & { readonly kind: "emphasis" };

export type EffectOptions = {
  /** 时长，单位为毫秒。 */
  readonly duration?: number;
  /** 相对所属触发起点的延迟，单位为毫秒。 */
  readonly delay?: number;
  /** 加速和减速各占时长的比例，合计不能超过 1。 */
  readonly easing?: Easing;
};

export type MotionProps = {
  /** 对象进入效果；数组中的效果从同一触发起点并行开始。 */
  readonly enter?: EntranceEffect | readonly EntranceEffect[];
  /** 对已显示对象应用的强调效果。 */
  readonly emphasis?: EmphasisEffect;
  /** 设为 false 可禁用此对象及其后代的动效。 */
  readonly motion?: boolean;
};

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
};

const motionContext = createContext<MotionContextValue>({ enabled: true });

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
  const host = createElement(
    type,
    {
      ...props,
      __motionEnabled: enabled,
      ...(parent.step === undefined ? {} : { __motionStep: parent.step }),
    },
    props.children,
  );
  if (enabled === parent.enabled) return host;
  return createElement(
    motionContext.Provider,
    { value: { ...parent, enabled } },
    host,
  );
}

type MotionTarget = {
  readonly kind: string;
  readonly name: string;
  readonly testId?: string;
  readonly textBody: boolean;
  readonly filled: boolean;
};

type MotionValue = {
  readonly target: MotionTarget;
  readonly step?: StepDeclaration;
  readonly effects: readonly Effect[];
};

export type MotionEffectObservation = {
  readonly target: {
    readonly id: number;
    readonly name: string;
    readonly testId?: string;
  };
  readonly kind: Effect["kind"];
  readonly preset: EffectPreset;
  readonly durationMs: number;
  readonly delayMs: number;
  readonly from?: MotionDirection | number;
  readonly easing?: Easing;
  readonly startMs: number;
  readonly endMs: number;
};

export type MotionStepObservation = {
  readonly trigger: "auto" | "click";
  readonly number?: number;
  readonly effects: readonly MotionEffectObservation[];
};

export type MotionObservation = {
  readonly steps: readonly MotionStepObservation[];
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
] as const;

function validatedEffect(value: unknown, expected: Effect["kind"]): Effect {
  if (!isRecord(value) || value["type"] !== "effect") {
    return fail(
      "invalid-motion",
      expected === "entrance"
        ? "enter 必须是 fade()、wipe() 等入场效果声明"
        : "emphasis 必须是 spin() 等强调效果声明",
    );
  }
  const unknown = Object.keys(value).find(
    (key) => !(effectKeys as readonly string[]).includes(key),
  );
  if (unknown !== undefined) {
    return fail("unknown-parameter", `动画效果含未实现参数 ${unknown}`);
  }
  if (value["kind"] !== expected) {
    return fail(
      "invalid-motion",
      expected === "entrance"
        ? `${String(value["preset"])}() 是强调效果，应传给 emphasis`
        : `${String(value["preset"])}() 是入场效果，应传给 enter`,
    );
  }
  const preset = value["preset"];
  if (typeof preset !== "string" || !presets.has(preset as EffectPreset)) {
    return fail("unsupported-effect", `不支持动画效果 ${String(preset)}`);
  }
  if (presetKinds[preset as EffectPreset] !== expected) {
    return fail(
      "invalid-motion",
      expected === "entrance"
        ? `${preset}() 是强调效果，应传给 emphasis`
        : `${preset}() 是入场效果，应传给 enter`,
    );
  }
  const duration = value["duration"];
  const delay = value["delay"];
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
  if (
    (preset === "wipe" || preset === "fly") &&
    (typeof value["from"] !== "string" ||
      !directions.has(value["from"] as MotionDirection))
  ) {
    return fail("invalid-direction", `${preset} 的 from 必须是有效方向`);
  }
  if (preset === "spin") {
    const from = value["from"];
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
    value["from"] !== undefined
  ) {
    return fail("invalid-motion", `${preset} 不接受 from 参数`);
  }
  const easing = value["easing"];
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
  return value as unknown as Effect;
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

/** 将对象动画声明转成由 motion 页面能力消费的局部注解。 */
export function motionAnnotations(
  props: Readonly<Record<string, unknown>>,
  target: MotionTarget,
): readonly CapabilityAnnotation[] {
  if (props["motion"] !== undefined && typeof props["motion"] !== "boolean") {
    return fail("invalid-motion", "motion 必须是布尔值");
  }
  if (props["__motionEnabled"] === false || props["motion"] === false)
    return [];

  const step = stepDeclaration(props["__motionStep"]);
  const effects = [
    ...listed(props["enter"], true).map((value) =>
      validatedEffect(value, "entrance"),
    ),
    ...listed(props["emphasis"], false).map((value) =>
      validatedEffect(value, "emphasis"),
    ),
  ];
  if (effects.length === 0) return [];
  const annotation: CapabilityAnnotation = {
    capability: "motion",
    value: Object.freeze({
      target: Object.freeze({ ...target }),
      ...(step === undefined ? {} : { step }),
      effects: Object.freeze(effects),
    } satisfies MotionValue),
  };
  return [annotation];
}
