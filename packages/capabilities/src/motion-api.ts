import {
  CapabilityError,
  type CapabilityAnnotation,
  type CapabilityChild,
  type CapabilityContribution,
  type GeometryBox,
} from "@react-pptx/core/capability";
import { CommonContainer } from "@react-pptx/core/document";
import type { ReactNode } from "react";

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
  | "floatDown"
  | "spinIn"
  | "stretch"
  | "spin"
  | "breathe"
  | "fillColor";

export type RepeatCount = number | "forever";

export type MotionGranularity =
  | "paragraph"
  | "word"
  | "letter"
  | "series"
  | "category"
  | "seriesElements"
  | "categoryElements";

export type MotionTargetRange =
  | { readonly type: "paragraph"; readonly start: number; readonly end: number }
  | {
      readonly type: "chart";
      readonly step:
        | "gridLegend"
        | "series"
        | "category"
        | "ptInSeries"
        | "ptInCategory";
      readonly seriesIndex: number;
      readonly categoryIndex: number;
    };

/** Unknown start constraint points to the start or end time of an effect, effect collection for the same object, or real scope. */
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
      readonly point: "start" | "end";
      readonly offsetMs: number;
    };

export type Effect = {
  readonly type: "effect";
  readonly kind: "entrance" | "emphasis";
  readonly preset: EffectPreset;
  /** Duration of a single native behavior; total completion length when unfolded per-word or per-letter remains unknown. */
  readonly duration: number;
  readonly delay: number;
  readonly from?: MotionDirection | number;
  readonly fromColor?: string;
  readonly toColor?: string;
  readonly easing?: Easing;
  /** Text or chart granularity; in spin, represents relative angle in degrees. */
  readonly by?: MotionGranularity | number;
  /** Launch interval between adjacent native units in per-word or per-letter effects. */
  readonly intervalMs?: number;
};

export type MotionScopeMode = "parallel" | "sequence" | "stagger";

/** Natural completion state of a scope; distinguishes between no natural completion and unknown duration. */
export type MotionCompletionState =
  | "known"
  | "unknown"
  | "no-natural-completion";

/** Points to an explicit orchestration scope boundary or its publicly named moment. */
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

/** A hook invocation represents the scope identity of a component instance; instances with the same component name each have an independent identity. */
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

/** One-time fill color transition; terminal state is preserved until slide exit, applicable only to Shapes with solid fill. */
export type FillColorOptions = Pick<
  EffectOptions,
  "duration" | "delay" | "easing"
> & {
  readonly from: string;
  readonly to: string;
};

export type BreatheOptions = {
  /** Complete dark->light->dark cycle in even milliseconds. */
  readonly period: number;
  /** Dark color at cycle start; playback begins from this color. */
  readonly from: string;
  /** Light color at half-cycle end. */
  readonly to: string;
  /** Delay relative to the trigger start, in milliseconds. */
  readonly delay?: number;
  /** Proportions of full cycle for acceleration and deceleration; sum cannot exceed 1. */
  readonly easing?: Easing;
};

export type BreatheEffect = EmphasisEffect & {
  readonly preset: "breathe";
  readonly fromColor: string;
  readonly toColor: string;
};

/** Internal effect validated after repeat(); authors must declare round count via RepeatedEffect. */
export type NormalizedEffect = Effect & {
  readonly repeatCount?: RepeatCount;
};

export type EffectOptions = {
  /** Single effect behavior duration in milliseconds; fine-grained unfolding completion length is not represented by this value. */
  readonly duration?: number;
  /** Delay relative to the trigger start, in milliseconds. */
  readonly delay?: number;
  /** Proportions of duration for acceleration and deceleration; sum cannot exceed 1. */
  readonly easing?: Easing;
  /** Text per-word or per-letter launch interval; defaults to one tenth of single behavior duration when omitted. */
  readonly intervalMs?: number;
  /** Unfold item by item by native paragraph, word, letter, chart series, category, series points, or category points. */
  readonly by?: MotionGranularity;
};

export type SpinOptions = Omit<EffectOptions, "by"> & {
  /** String unfolds by text granularity; number represents relative angle, negative rotates in opposite direction. */
  readonly by?: MotionGranularity | number;
  /** Rotates back to object native orientation from specified angle. */
  readonly from?: number;
};

export type MotionProps = {
  /** Object entrance effects; effects in array start in parallel from the same trigger start. */
  readonly enter?:
    | EntranceEffectDeclaration
    | readonly EntranceEffectDeclaration[];
  /** Emphasis effect applied to already displayed object. */
  readonly emphasis?: EmphasisEffectDeclaration;
  /** Set to false to disable motion for this object and its descendants. */
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
  /** Created by useMotionScope() of the current component instance. */
  readonly scope: MotionScopeHandle;
  /** Composition relationship of child members, defaults to parallel. */
  readonly mode?: MotionScopeMode;
  /** Stagger launch interval for each direct member, in milliseconds. */
  readonly each?: number;
  /** Offset relative to parent scope start, in milliseconds. */
  readonly delay?: number;
  /** Launch from another scope moment; after() can be used to express offset. */
  readonly startAt?: MotionTiming;
  /** Publish this scope's start time to named moment belonging to the same orchestration. */
  readonly publishStart?: MotionMoment;
  /** Publish this scope's completion time to named moment belonging to the same orchestration. */
  readonly publishEnd?: MotionMoment;
  /** Set to false to disable motion for this scope and its descendants. */
  readonly motion?: boolean;
  readonly children?: ReactNode;
};

export type StepProps = {
  /** Merge click steps with the same positive integer; each Step triggers independently when omitted. */
  readonly n?: number;
  readonly children?: ReactNode;
};

export type StepDeclaration =
  | { readonly mode: "automatic"; readonly identity: string }
  | { readonly mode: "explicit"; readonly number: number };

/** Plain authoring scope props; React children stay on the JSX entry. */
export type MotionScopeAuthorProps = Omit<MotionScopeProps, "children">;
export type MotionStepAuthorProps = Pick<StepProps, "n">;

type MotionStepNodeProps = MotionStepAuthorProps & {
  /** Build-local identity of an automatic Step, created once when the host node is created. */
  readonly stepIdentity: string;
};

type MotionScopeNodeProps = MotionScopeAuthorProps;

/** Creates the editable orchestration scope shared by the command factory and the JSX entry. */
export function createMotionScopeNode(
  props: MotionScopeNodeProps,
): CommonContainer<"motion-scope", MotionScopeNodeProps> {
  return new CommonContainer("motion-scope", "rp:motion-scope", props);
}

/** Creates the editable click-step container shared by the command factory and the JSX entry. */
export function createMotionStepNode(
  props: MotionStepNodeProps,
): CommonContainer<"motion-step", MotionStepNodeProps> {
  return new CommonContainer("motion-step", "rp:motion-step", props);
}

let nextMotionIdentity = 1;

function createMotionIdentity(kind: "scope" | "step"): string {
  return `react-pptx-${kind}-${nextMotionIdentity++}`;
}

/**
 * Creates a React-independent scope identity and its start/end/named moments.
 *
 * The returned handle keeps its identity when nodes are moved or rebuilt. React authoring keeps one
 * handle per component instance and passes it to `MotionScope`, while plain authoring can create it
 * directly. Cues are values derived from the same identity, not React keys or author ids.
 */
export function createMotionScopeHandle(): MotionScopeHandle {
  const identity = createMotionIdentity("scope");
  const moment = (point: MotionMoment["point"], name?: string): MotionMoment =>
    Object.freeze({
      type: "motion-moment",
      scope: identity,
      point,
      ...(name === undefined ? {} : { name }),
    });
  return Object.freeze({
    type: "motion-scope-handle",
    identity,
    start: moment("start"),
    end: moment("end"),
    cue(name: string): MotionMoment {
      if (typeof name !== "string" || name.trim() === "") {
        throw new TypeError(
          "MotionScopeHandle.cue(name) requires a non-empty name",
        );
      }
      return moment("named", name);
    },
  });
}

/** Creates the stable identity used by an automatic Step host without React. */
export function createMotionStepIdentity(): string {
  return createMotionIdentity("step");
}

/** Schedules an explicit non-negative offset after any moment handle. */
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

function declaration(
  kind: Effect["kind"],
  preset: EffectPreset,
  options: Omit<EffectOptions, "by"> & {
    readonly by?: Effect["by"];
  } & object,
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

/** Creates pure fade-in declaration; default duration is 500 ms. */
export function fade(options: EffectOptions = {}): EntranceEffect {
  return declaration("entrance", "fade", options, {
    duration: 500,
  }) as EntranceEffect;
}

/** Wipes in from specified edge; default direction is bottom. */
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

/** Zooms in from origin and fades in. */
export function zoom(options: EffectOptions = {}): EntranceEffect {
  return declaration("entrance", "zoom", options, {
    duration: 500,
  }) as EntranceEffect;
}

/** Flies in from specified slide edge; default direction is bottom. */
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

/** Fades in and floats up. */
export function floatUp(options: EffectOptions = {}): EntranceEffect {
  return declaration("entrance", "floatUp", options, {
    duration: 500,
  }) as EntranceEffect;
}

/** Fades in and floats down from above; default duration is 500 ms. */
export function floatDown(options: EffectOptions = {}): EntranceEffect {
  return declaration("entrance", "floatDown", options, {
    duration: 500,
  }) as EntranceEffect;
}

/** Zooms, rotates, and fades in. */
export function spinIn(options: EffectOptions = {}): EntranceEffect {
  return declaration("entrance", "spinIn", options, {
    duration: 500,
  }) as EntranceEffect;
}

/** Stretches from center line. */
export function stretch(options: EffectOptions = {}): EntranceEffect {
  return declaration("entrance", "stretch", options, {
    duration: 500,
  }) as EntranceEffect;
}

/** Rotates back to initial orientation from specified angle, or rotates relative to initial angle using signed by; default from is 360 degrees. */
export function spin(options: SpinOptions = {}): EmphasisEffect {
  return declaration(
    "emphasis",
    "spin",
    options,
    { duration: 1000 },
    typeof options.by === "number"
      ? options.from === undefined
        ? {}
        : { from: options.from }
      : { from: options.from === undefined ? 360 : options.from },
  ) as EmphasisEffect;
}

/** Changes from from to to and maintains terminal state; non-repeating, non-reversing, default duration is 500 ms. */
export function fillColor(options: FillColorOptions): EmphasisEffect {
  const { from, to, ...timing } = options;
  return Object.freeze({
    ...declaration("emphasis", "fillColor", timing, { duration: 500 }),
    kind: "emphasis" as const,
    fromColor: from,
    toColor: to,
  });
}

// Retains extra options from structured JavaScript input for formal effect validation diagnostics.
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

/** Declares dark->light half-cycle of Shape fill color; native autoRev reverses it to form a full cycle. */
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

/** Wraps supported effect with specified total rounds; forever applies only to breathe. */
export function repeat<E extends Effect>(
  effect: E,
  count: RepeatCount,
): RepeatedEffect<E> {
  return Object.freeze({ type: "repeated-effect", effect, count });
}

type MotionTarget = {
  readonly kind: string;
  readonly name: string;
  readonly authorId?: string;
  readonly textBody: boolean;
  readonly filled: boolean;
  readonly solidFilled?: boolean;
  readonly paragraphCount?: number;
  /** Original point count for every chart-wide series, in formal series order. */
  readonly chartSeriesPointCounts?: readonly number[];
  readonly box?: GeometryBox;
  readonly rotationDegrees?: number;
  readonly hasBackground?: boolean;
  readonly members?: readonly MotionNativeMember[];
};

type MotionNativeMember = {
  readonly id: number;
  readonly kind: string;
  readonly name: string;
  readonly authorId?: string;
  readonly box?: GeometryBox;
  readonly members?: readonly MotionNativeMember[];
};

type MotionValue = {
  readonly target: MotionTarget;
  readonly step?: StepDeclaration;
  readonly scopes: readonly MotionScopeDeclaration[];
  readonly effects: readonly NormalizedEffect[];
};

/** Local object declaration before inherited scope, step, default effect, and disabled ancestors are resolved. */
type LocalMotionValue = {
  readonly target: MotionTarget;
  readonly enterDeclared: boolean;
  readonly effects: readonly NormalizedEffect[];
};

type MotionFrame = {
  readonly enabled: boolean;
  readonly step?: StepDeclaration;
  readonly scopes: readonly MotionScopeDeclaration[];
  readonly defaultEnter?:
    | EntranceEffectDeclaration
    | readonly EntranceEffectDeclaration[];
};

type LocalScopeValue = {
  readonly type: "scope";
  readonly scope: MotionScopeDeclaration;
};

export type MotionEffectObservation = {
  readonly target: {
    readonly id: number;
    readonly name: string;
    readonly authorId?: string;
    readonly range?: MotionTargetRange;
  };
  readonly kind: Effect["kind"];
  readonly preset: EffectPreset;
  /** Single effect duration; breathe indicates full round-trip cycle, not representing total length of fine-grained unfolding. */
  readonly durationMs: number;
  /** Total rounds of complete breathe cycle; forever indicates persisting on current slide with no natural completion. */
  readonly repeatCount?: RepeatCount;
  readonly delayMs: number;
  /** Text or chart granularity; in spin, represents relative rotation angle in degrees. */
  readonly by?: MotionGranularity | number;
  readonly intervalMs?: number;
  readonly from?: MotionDirection | number;
  readonly fromColor?: string;
  readonly toColor?: string;
  readonly easing?: Easing;
  readonly startMs?: number;
  readonly endMs?: number;
  readonly completion?: Exclude<MotionCompletionState, "known">;
  readonly startAfter?: MotionStartDependency;
  /** When unknown start is jointly constrained by multiple real moments, lists all sources and relative offsets. */
  readonly startAfterAll?: readonly MotionStartDependency[];
};

export type MotionStepObservation = {
  readonly trigger: "auto" | "click";
  readonly number?: number;
  readonly effects: readonly MotionEffectObservation[];
};

/** Corresponding real member effects and timing identity in PPTX for a logical Group effect. */
export type MotionNativeProjection = {
  readonly logicalTargetId: number;
  readonly logicalEffectIndex: number;
  readonly preset: EffectPreset;
  readonly nativeEffects: readonly {
    readonly targetId: number;
    readonly timingId: number;
  }[];
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
  /** When unknown start is jointly constrained by multiple real moments, lists all sources and relative offsets. */
  readonly startAfterAll?: readonly MotionStartDependency[];
  readonly members: readonly (
    | { readonly kind: "scope"; readonly id: string }
    | { readonly kind: "object"; readonly id: number }
  )[];
};

export type MotionObservation = {
  readonly steps: readonly MotionStepObservation[];
  readonly scopes?: readonly MotionScopeObservation[];
  readonly nativeProjections?: readonly MotionNativeProjection[];
};

function fail(code: string, message: string): never {
  throw new CapabilityError(code, message);
}

/** Validates the local motion enable flag during capability parsing. */
export function validateMotionFlag(value: unknown): void {
  if (value !== undefined && typeof value !== "boolean") {
    fail("invalid-motion", "motion must be a boolean");
  }
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
  "floatDown",
  "spinIn",
  "stretch",
  "spin",
  "breathe",
  "fillColor",
]);
const presetKinds: Readonly<Record<EffectPreset, Effect["kind"]>> = {
  fade: "entrance",
  wipe: "entrance",
  zoom: "entrance",
  fly: "entrance",
  floatUp: "entrance",
  floatDown: "entrance",
  spinIn: "entrance",
  stretch: "entrance",
  spin: "emphasis",
  breathe: "emphasis",
  fillColor: "emphasis",
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
  "seriesElements",
  "categoryElements",
]);

// ST_TLTime uses unsignedInt; repeatCount unit is one thousandth of a round.
export const maximumTimeNodeValue = 4_294_967_295;
const maxRepeatCount = Math.floor(maximumTimeNodeValue / 1000);

function normalizeColor(
  value: unknown,
  name: string,
  code = "invalid-breathe-color",
): string {
  if (typeof value !== "string" || !/^#[0-9a-f]{6}$/iu.test(value)) {
    return fail(code, `${name} must be a #RRGGBB color`);
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
        `repeat declaration contains unimplemented parameter ${unknownRepeatField}`,
      );
    }
    if (!isRecord(value["effect"]) || value["effect"]["type"] !== "effect") {
      return fail(
        "invalid-repeat",
        "repeat must directly wrap a single effect declaration",
      );
    }
    repeatCount = value["count"] as RepeatCount;
    effectValue = value["effect"];
  }
  if (!isRecord(effectValue) || effectValue["type"] !== "effect") {
    return fail(
      "invalid-motion",
      expected === "entrance"
        ? "enter must be an entrance effect declaration such as fade() or wipe()"
        : "emphasis must be an emphasis effect declaration such as spin()",
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
      `breathe declaration contains unimplemented parameter ${unsupportedBreatheParameters[0]}`,
    );
  }
  const unknown = Object.keys(effectValue).find(
    (key) => !(effectKeys as readonly string[]).includes(key),
  );
  if (unknown !== undefined) {
    return fail(
      "unknown-parameter",
      `Animation effect contains unimplemented parameter ${unknown}`,
    );
  }
  if (effectValue["kind"] !== expected) {
    return fail(
      "invalid-motion",
      expected === "entrance"
        ? `${String(effectValue["preset"])}() is an emphasis effect and should be passed to emphasis`
        : `${String(effectValue["preset"])}() is an entrance effect and should be passed to enter`,
    );
  }
  const preset = effectValue["preset"];
  if (typeof preset !== "string" || !presets.has(preset as EffectPreset)) {
    return fail(
      "unsupported-effect",
      `Unsupported animation effect ${String(preset)}`,
    );
  }
  if (
    preset !== "breathe" &&
    presetKinds[preset as EffectPreset] !== expected
  ) {
    return fail(
      "invalid-motion",
      expected === "entrance"
        ? `${preset}() is an emphasis effect and should be passed to emphasis`
        : `${preset}() is an entrance effect and should be passed to enter`,
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
    return fail(
      "invalid-duration",
      "Animation duration must be writable as a finite positive integer in milliseconds",
    );
  }
  if (
    typeof delay !== "number" ||
    !Number.isFinite(delay) ||
    delay < 0 ||
    !Number.isSafeInteger(Math.round(delay)) ||
    !Number.isFinite(Math.round(delay) + Math.round(duration))
  ) {
    return fail(
      "invalid-delay",
      "Animation delay must be writable as a finite non-negative integer in milliseconds",
    );
  }
  if (preset === "breathe" && Math.round(delay) > maximumTimeNodeValue) {
    return fail(
      "invalid-delay",
      "breathe delay exceeds native timing value range",
    );
  }
  const by = effectValue["by"];
  const interval = effectValue["intervalMs"];
  if (preset === "spin" && typeof by === "number") {
    const units = Math.round(by * 60000);
    if (
      !Number.isFinite(by) ||
      !Number.isSafeInteger(units) ||
      units < minimumRotationUnits ||
      units > maximumRotationUnits
    ) {
      return fail(
        "invalid-rotation",
        "spin by converted value must fall within signed 32-bit angle range",
      );
    }
    if (effectValue["from"] !== undefined) {
      return fail(
        "invalid-motion",
        "spin from and by cannot be declared simultaneously",
      );
    }
  } else if (
    by !== undefined &&
    (typeof by !== "string" ||
      !motionGranularities.has(by as MotionGranularity))
  ) {
    return fail(
      "invalid-motion",
      "by must be a supported text or chart granularity",
    );
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
      "intervalMs applies only to per-word or per-letter effects, and must be a positive integer in milliseconds",
    );
  }
  if (
    (preset === "wipe" || preset === "fly") &&
    (typeof effectValue["from"] !== "string" ||
      !directions.has(effectValue["from"] as MotionDirection))
  ) {
    return fail(
      "invalid-direction",
      `${preset} from must be a valid direction`,
    );
  }
  if (preset === "spin") {
    const from = effectValue["from"];
    if (typeof by !== "number") {
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
          "spin from converted value must fall within signed 32-bit angle range",
        );
      }
    }
  }
  if (
    preset !== "wipe" &&
    preset !== "fly" &&
    preset !== "spin" &&
    effectValue["from"] !== undefined
  ) {
    return fail("invalid-motion", `${preset} does not accept from parameter`);
  }
  const easing = effectValue["easing"];
  if (easing !== undefined) {
    if (!isRecord(easing)) {
      return fail("invalid-easing", "easing must be an object");
    }
    const unknownEasing = Object.keys(easing).find(
      (key) => key !== "accelerate" && key !== "decelerate",
    );
    if (unknownEasing !== undefined) {
      return fail(
        "unknown-parameter",
        `easing contains unimplemented parameter ${unknownEasing}`,
      );
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
        "easing accelerate and decelerate proportions must be non-negative and sum to at most 1",
      );
    }
  }
  let fromColor: string | undefined;
  let toColor: string | undefined;
  if (preset === "breathe") {
    if (expected !== "emphasis") {
      return fail("invalid-motion", "breathe() must be passed to emphasis");
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
        "breathe period must be an even integer in milliseconds greater than 1, and half-period must be writable to native duration",
      );
    }
    fromColor = normalizeColor(effectValue["fromColor"], "breathe.from");
    toColor = normalizeColor(effectValue["toColor"], "breathe.to");
    if (relativeLuminance(toColor) <= relativeLuminance(fromColor)) {
      return fail(
        "invalid-breathe-colors",
        "breathe.to must be brighter than breathe.from",
      );
    }
  } else if (preset === "fillColor") {
    fromColor = normalizeColor(
      effectValue["fromColor"],
      "fillColor.from",
      "invalid-fill-color",
    );
    toColor = normalizeColor(
      effectValue["toColor"],
      "fillColor.to",
      "invalid-fill-color",
    );
    if (by !== undefined || interval !== undefined) {
      return fail(
        "invalid-motion",
        "fillColor does not support text or chart granularity",
      );
    }
  } else if (
    effectValue["fromColor"] !== undefined ||
    effectValue["toColor"] !== undefined
  ) {
    return fail(
      "invalid-motion",
      `${String(preset)} does not accept breathe color parameters`,
    );
  }

  if (repeatDeclared) {
    if (preset !== "spin" && preset !== "breathe") {
      return fail(
        "unsupported-repeat",
        `${String(preset)} does not support repetition yet; currently only spin and breathe are supported`,
      );
    }
    if (repeatCount === "forever") {
      if (preset !== "breathe") {
        return fail(
          "unsupported-repeat",
          "forever currently only supports complete breathe cycles",
        );
      }
    } else if (
      typeof repeatCount !== "number" ||
      !Number.isSafeInteger(repeatCount) ||
      repeatCount < 1 ||
      repeatCount > maxRepeatCount
    ) {
      return fail(
        "invalid-repeat-count",
        `repeat count must be total rounds between 1 and ${maxRepeatCount}`,
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
        "Repeated total duration must remain an exact integer in milliseconds and single round must not exceed native timing range",
      );
    }
    if (by !== undefined && typeof by !== "number") {
      return fail(
        "unsupported-repeat-combination",
        "Repeated effects cannot currently be combined with per-paragraph, per-word, per-letter, or chart granularity",
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

function stepDeclarationFromProps(
  props: Readonly<Record<string, unknown>>,
): StepDeclaration {
  const number = props["n"];
  if (number === undefined) {
    const identity = props["stepIdentity"];
    if (typeof identity !== "string" || identity === "") {
      return fail(
        "invalid-step",
        "Automatic click step lacks instance identity",
      );
    }
    return Object.freeze({ mode: "automatic" as const, identity });
  }
  if (
    typeof number !== "number" ||
    !Number.isSafeInteger(number) ||
    number < 1
  ) {
    return fail(
      "invalid-step",
      "Explicit click step number must be a positive integer",
    );
  }
  return Object.freeze({ mode: "explicit" as const, number });
}

function listed(value: unknown, many: boolean): readonly unknown[] {
  if (value === undefined) return [];
  return many && Array.isArray(value) ? value : [value];
}

function validatedMoment(value: unknown): MotionMoment {
  if (!isRecord(value) || value["type"] !== "motion-moment") {
    return fail(
      "invalid-moment",
      "Moment handle must come from a motion scope handle.",
    );
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
    return fail(
      "invalid-moment",
      "Moment handle scope, type, or name is invalid",
    );
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
      return fail(
        "invalid-delay",
        "Moment offset must be a finite non-negative integer in milliseconds",
      );
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
      "publishStart and publishEnd require named moment handles",
    );
  }
  return momentValue;
}

function isMotionScopeHandle(value: unknown): value is MotionScopeHandle {
  return (
    isRecord(value) &&
    value["type"] === "motion-scope-handle" &&
    typeof value["identity"] === "string" &&
    value["identity"] !== ""
  );
}

function localScopeDeclaration(
  props: Readonly<Record<string, unknown>>,
): MotionScopeDeclaration {
  const allowedParameters = new Set([
    "scope",
    "mode",
    "each",
    "delay",
    "startAt",
    "publishStart",
    "publishEnd",
    "motion",
  ]);
  const unknown = Object.keys(props).find((key) => !allowedParameters.has(key));
  if (unknown !== undefined) {
    return fail(
      "unknown-parameter",
      `MotionScope contains unimplemented parameter ${unknown}`,
    );
  }
  validateMotionFlag(props["motion"]);
  const handle = props["scope"];
  if (!isMotionScopeHandle(handle)) {
    return fail(
      "invalid-scope",
      "MotionScope requires a handle created by createMotionScopeHandle() or useMotionScope()",
    );
  }
  const mode = props["mode"] ?? "parallel";
  const delay = props["delay"] ?? 0;
  const each = props["each"];
  if (mode !== "parallel" && mode !== "sequence" && mode !== "stagger") {
    return fail(
      "invalid-scope",
      "MotionScope mode must be parallel, sequence, or stagger",
    );
  }
  if (
    typeof delay !== "number" ||
    !Number.isFinite(delay) ||
    delay < 0 ||
    !Number.isSafeInteger(Math.round(delay))
  ) {
    return fail(
      "invalid-delay",
      "MotionScope delay must be a finite non-negative integer in milliseconds",
    );
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
      "stagger requires non-negative integer each; other modes do not accept each",
    );
  }
  const startAt =
    props["startAt"] === undefined
      ? undefined
      : validatedTiming(props["startAt"]);
  const publishStart = validatedPublisher(props["publishStart"]);
  const publishEnd = validatedPublisher(props["publishEnd"]);
  return Object.freeze({
    identity: handle.identity,
    mode,
    delay: Math.round(delay),
    ...(typeof each === "number" ? { each: Math.round(each) } : {}),
    ...(startAt === undefined ? {} : { startAt }),
    ...(publishStart === undefined ? {} : { publishStart }),
    ...(publishEnd === undefined ? {} : { publishEnd }),
  });
}

/** Compiles explicit scope host into a local declaration resolved against the current document tree. */
export function motionScopeAnnotation(
  input: Readonly<Record<string, unknown>>,
): CapabilityAnnotation {
  const scope = localScopeDeclaration(input);
  return {
    capability: "motion",
    value: Object.freeze({ type: "scope", scope }),
    resolveContributions: resolveMotionContributions,
  };
}

/** Validates Step wrapper parameters; step ancestry is derived from the build-time tree. */
export function motionStepAnnotation(
  input: Readonly<Record<string, unknown>>,
): readonly CapabilityAnnotation[] {
  const unknown = Object.keys(input).find(
    (key) => key !== "n" && key !== "stepIdentity",
  );
  if (unknown !== undefined) {
    return fail(
      "unknown-parameter",
      `Step contains unimplemented parameter ${unknown}`,
    );
  }
  stepDeclarationFromProps(input);
  return [];
}

function localMotionValue(
  props: Readonly<Record<string, unknown>>,
  target: MotionTarget,
): LocalMotionValue {
  const enter = props["enter"];
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
      "breathe() requires native Shape with solid fill; Group requires effects declared on real members separately",
    );
  }
  if (
    effects.some(
      ({ preset }) =>
        preset === "fillColor" &&
        (target.kind !== "shape" || !target.solidFilled),
    )
  ) {
    return fail(
      "invalid-fill-color-target",
      "fillColor() requires native Shape with solid fill",
    );
  }
  return Object.freeze({
    target: Object.freeze({ ...target }),
    enterDeclared: enter !== undefined,
    effects: Object.freeze(effects),
  });
}

/** Converts object animation declarations into local annotations consumed by motion slide capability. */
export function motionAnnotations(
  props: Readonly<Record<string, unknown>>,
  target: MotionTarget,
): readonly CapabilityAnnotation[] {
  validateMotionFlag(props["motion"]);
  if (props["motion"] === false) return [];
  const value = localMotionValue(props, target);
  return [
    {
      capability: "motion",
      optional: value.effects.length === 0,
      value,
      // Attach the resolver even for an empty local declaration: an inherited Box default may still
      // make this object's motion required when the timing writer is trimmed.
      resolveContributions: resolveMotionContributions,
    },
  ];
}

function isLocalMotionValue(value: unknown): value is LocalMotionValue {
  return (
    isRecord(value) &&
    isRecord(value["target"]) &&
    typeof value["enterDeclared"] === "boolean" &&
    Array.isArray(value["effects"])
  );
}

function isLocalScopeValue(value: unknown): value is LocalScopeValue {
  return (
    isRecord(value) &&
    value["type"] === "scope" &&
    isRecord(value["scope"]) &&
    typeof value["scope"]["identity"] === "string"
  );
}

function entranceEffects(
  value: unknown,
  path: string,
): readonly NormalizedEffect[] {
  try {
    return listed(value, true).map((effect) =>
      validatedEffect(effect, "entrance"),
    );
  } catch (error) {
    if (error instanceof CapabilityError) {
      throw new CapabilityError(error.code, error.message, path);
    }
    throw error;
  }
}

function motionContainer(declaration: string): boolean {
  return (
    declaration === "rp:box" ||
    declaration === "rp:group" ||
    declaration === "rp:motion-scope"
  );
}

/**
 * Resolves local motion declarations against the current declaration tree.
 *
 * Scope parent chains, nearest click step, inherited Box default entrance, and disabled ancestor
 * chains are derived from this build's structure. Nothing here is stored on the editable nodes, so
 * moving or appending content changes the next build's choreography without stale React context.
 *
 * The build passes only this capability's contributions, then uses the effective result for both
 * timing output and missing-capability checks.
 */
export function resolveMotionContributions(
  children: readonly CapabilityChild[],
  contributions: readonly CapabilityContribution[],
): readonly CapabilityContribution[] {
  const byDeclaration = new WeakMap<
    object,
    readonly CapabilityContribution[]
  >();
  for (const contribution of contributions) {
    const declaration = contribution.declaration;
    if (declaration === undefined) continue;
    const existing = byDeclaration.get(declaration);
    byDeclaration.set(
      declaration,
      existing === undefined ? [contribution] : [...existing, contribution],
    );
  }

  const resolved: CapabilityContribution[] = [];
  const visit = (node: CapabilityChild, frame: MotionFrame): void => {
    if (!("props" in node)) return;
    const localContributions = byDeclaration.get(node) ?? [];
    let childFrame = frame;

    if (node.type === "rp:motion-scope") {
      const enabled = frame.enabled && node.props["motion"] !== false;
      const scopeContribution = localContributions.find((contribution) =>
        isLocalScopeValue(contribution.value),
      );
      if (scopeContribution !== undefined) {
        const local = scopeContribution.value as LocalScopeValue;
        const parent = frame.scopes.at(-1);
        const scope: MotionScopeDeclaration = Object.freeze({
          ...local.scope,
          ...(parent === undefined ? {} : { parent: parent.identity }),
          ...(frame.step === undefined ? {} : { step: frame.step }),
        });
        resolved.push(
          Object.freeze({
            ...scopeContribution,
            value: Object.freeze({ type: "scope" as const, scope }),
          }),
        );
        childFrame = {
          ...frame,
          enabled,
          scopes: Object.freeze([...frame.scopes, scope]),
        };
      } else {
        childFrame = { ...frame, enabled };
      }
    } else if (node.type === "rp:motion-step") {
      childFrame = { ...frame, step: stepDeclarationFromProps(node.props) };
    } else {
      const localContribution = localContributions.find((contribution) =>
        isLocalMotionValue(contribution.value),
      );
      let defaultEnter = frame.defaultEnter;
      if (node.type === "rp:box" && node.props["enter"] !== undefined) {
        defaultEnter = node.props["enter"] as
          | EntranceEffectDeclaration
          | readonly EntranceEffectDeclaration[];
      }
      const enabled = frame.enabled && node.props["motion"] !== false;
      if (localContribution !== undefined && enabled) {
        const local = localContribution.value as LocalMotionValue;
        const inherited = local.enterDeclared
          ? []
          : entranceEffects(defaultEnter, localContribution.path);
        const effects = Object.freeze([...inherited, ...local.effects]);
        if (effects.length > 0) {
          resolved.push(
            Object.freeze({
              ...localContribution,
              optional: false,
              value: Object.freeze({
                target: local.target,
                ...(frame.step === undefined ? {} : { step: frame.step }),
                scopes: frame.scopes,
                effects,
              } satisfies MotionValue),
            }),
          );
        }
      }
      childFrame = {
        ...frame,
        ...(motionContainer(node.type) ? { enabled } : {}),
        ...(node.type === "rp:box" ? { defaultEnter } : {}),
      };
    }

    for (const child of node.children) visit(child, childFrame);
  };

  for (const node of children) visit(node, { enabled: true, scopes: [] });
  return resolved;
}
