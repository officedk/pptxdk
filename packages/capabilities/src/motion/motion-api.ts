import {
  CapabilityError,
  type CapabilityAnnotation,
  type CapabilityChild,
  type CapabilityContribution,
  type Rect,
} from "@pptxdk/core/capability";
import { CommonContainer } from "@pptxdk/core/document";
import type { ReactNode } from "react";

import {
  breatheOptionsShape,
  easingSchema,
  effectDelaySchema,
  effectDurationSchema,
  effectPresetSchema,
  entranceEffectShape,
  intervalMsSchema,
  motionDirectionSchema,
  motionGranularitySchema,
  motionHexColorSchema,
  motionMomentOffsetSchema,
  motionMomentSchema,
  motionScopeAuthorPropKeys,
  motionScopeHandleSchema,
  motionScopeModeSchema,
  motionStepAuthorPropKeys,
  motionStepAuthorPropsSchema,
  repeatCountSchema,
  spinAngleSchema,
  maxRepeatCount,
  maximumTimeNodeValue,
  type BreatheEffect,
  type BreatheOptions,
  type Easing,
  type Effect,
  type EffectOptions,
  type EffectPreset,
  type EmphasisEffect,
  type EntranceEffect,
  type FillColorOptions,
  type MotionCompletionState,
  type MotionDirection,
  type MotionGranularity,
  type MotionMoment,
  type MotionMomentOffset,
  type MotionScopeAuthorProps,
  type MotionScopeHandle,
  type MotionScopeMode,
  type MotionStepAuthorProps,
  type MotionTargetRange,
  type MotionTiming,
  type NormalizedEffect,
  type RepeatCount,
  type RepeatedEffect,
  type SpinOptions,
} from "./motion-schemas.ts";

export {
  breatheEffectSchema,
  breatheOptionsSchema,
  easingSchema,
  effectDelaySchema,
  effectDurationSchema,
  effectOptionsSchema,
  effectPresetSchema,
  effectSchema,
  emphasisEffectDeclarationSchema,
  emphasisEffectSchema,
  entranceEffectDeclarationSchema,
  entranceEffectSchema,
  fillColorOptionsSchema,
  motionCompletionStateSchema,
  motionDirectionSchema,
  motionGranularitySchema,
  motionHexColorSchema,
  motionMomentOffsetSchema,
  motionMomentSchema,
  motionPropsSchema,
  motionScopeAuthorPropsSchema,
  motionScopeHandleSchema,
  motionScopeModeSchema,
  motionStepAuthorPropsSchema,
  motionTargetRangeSchema,
  motionTimingSchema,
  normalizedEffectSchema,
  ownerMotionEffectSchema,
  ownerMotionSchema,
  ownerMotionStepSchema,
  ownerMotionTargetSchema,
  repeatCountSchema,
  repeatedEffectSchema,
  spinAngleSchema,
  spinOptionsSchema,
  maximumTimeNodeValue,
} from "./motion-schemas.ts";
export type {
  BreatheEffect,
  BreatheOptions,
  Easing,
  Effect,
  EffectOptions,
  EffectPreset,
  EmphasisEffect,
  EmphasisEffectDeclaration,
  EntranceEffect,
  EntranceEffectDeclaration,
  FillColorOptions,
  MotionCompletionState,
  MotionDirection,
  MotionGranularity,
  MotionMoment,
  MotionMomentOffset,
  MotionProps,
  MotionScopeAuthorProps,
  MotionScopeHandle,
  MotionScopeMode,
  MotionStepAuthorProps,
  MotionTargetRange,
  MotionTiming,
  NormalizedEffect,
  OwnerMotion,
  OwnerMotionEffect,
  OwnerMotionStep,
  OwnerMotionTarget,
  RepeatCount,
  RepeatedEffect,
  SpinOptions,
} from "./motion-schemas.ts";

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
  /** Internal parallel grouping scope for explicit-motion containers; omitted from public scope observations. */
  readonly implicit?: true;
};

export type MotionScopeProps = MotionScopeAuthorProps & {
  readonly children?: ReactNode;
};

export type StepProps = MotionStepAuthorProps & {
  readonly children?: ReactNode;
};

export type StepDeclaration =
  | { readonly mode: "automatic"; readonly identity: string }
  | { readonly mode: "explicit"; readonly number: number };

type MotionStepNodeProps = MotionStepAuthorProps & {
  /** Build-local identity of an automatic Step, created once when the host node is created. */
  readonly stepIdentity: string;
};

type MotionScopeNodeProps = MotionScopeAuthorProps;

/** Creates the editable orchestration scope shared by the command factory and the JSX entry. */
export function createMotionScopeNode(
  props: MotionScopeNodeProps,
): CommonContainer<"motion-scope", MotionScopeNodeProps> {
  return new CommonContainer("motion-scope", "motion-scope", props);
}

/** Creates the editable click-step container shared by the command factory and the JSX entry. */
export function createMotionStepNode(
  props: MotionStepNodeProps,
): CommonContainer<"motion-step", MotionStepNodeProps> {
  return new CommonContainer("motion-step", "motion-step", props);
}

let nextMotionIdentity = 1;

function createMotionIdentity(kind: "scope" | "step"): string {
  return `pptxdk-${kind}-${nextMotionIdentity++}`;
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
  extra: Partial<Pick<Effect, "from" | "to">> = {},
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

/** Rotates between specified `from`/`to` endpoint angles (defaulting to 360 -> 0, or 0 -> `to` when only `to` is declared), or rotates relative to initial angle using signed `by`. */
export function spin(options: SpinOptions = {}): EmphasisEffect {
  const rawFrom = (options as { readonly from?: number }).from;
  const rawTo = (options as { readonly to?: number }).to;
  return declaration(
    "emphasis",
    "spin",
    options,
    { duration: 1000 },
    typeof options.by === "number"
      ? {
          ...(rawFrom === undefined ? {} : { from: rawFrom }),
          ...(rawTo === undefined ? {} : { to: rawTo }),
        }
      : {
          from:
            rawFrom === undefined ? (rawTo === undefined ? 360 : 0) : rawFrom,
          ...(rawTo === undefined ? {} : { to: rawTo }),
        },
  ) as EmphasisEffect;
}

/** Changes from from to to and maintains terminal state; non-repeating, non-reversing, default duration is 500 ms. */
export function fillColor(options: FillColorOptions): EmphasisEffect {
  const { from, to, ...timing } = options;
  return Object.freeze({
    ...declaration("emphasis", "fillColor", timing, { duration: 500 }),
    kind: "emphasis" as const,
    preset: "fillColor" as const,
    fromColor: from,
    toColor: to,
  });
}

// Retains extra options from structured JavaScript input for formal effect validation diagnostics.
const breatheUnknownParameters = Symbol("breathe unknown parameters");
const breatheOptionNames = new Set([
  ...Object.keys(breatheOptionsShape),
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
  readonly box?: Rect;
  readonly rotationDegrees?: number;
  readonly hasBackground?: boolean;
  readonly members?: readonly MotionNativeMember[];
};

type MotionNativeMember = {
  readonly id: number;
  readonly kind: string;
  readonly name: string;
  readonly authorId?: string;
  readonly box?: Rect;
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
  readonly defaultEnter?: readonly NormalizedEffect[];
};

type LocalScopeValue = {
  readonly type: "scope";
  readonly scope: MotionScopeDeclaration;
  readonly enabled: boolean;
};

type LocalStepValue = {
  readonly type: "step";
  readonly step: StepDeclaration;
};

type LocalContainerValue = {
  readonly type: "container";
  readonly enabled: boolean;
  readonly defaultEnter?: readonly NormalizedEffect[];
  readonly groupsAsParallelUnit?: true;
};

export type ResolvedMotionEffect = {
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
  readonly to?: number;
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

export type ResolvedMotionStep = {
  readonly trigger: "auto" | "click";
  readonly number?: number;
  readonly effects: readonly ResolvedMotionEffect[];
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

export type ResolvedMotionScope = {
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

export type ResolvedMotion = {
  readonly steps: readonly ResolvedMotionStep[];
  readonly scopes?: readonly ResolvedMotionScope[];
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
const effectKeys: readonly string[] = Object.keys(entranceEffectShape);

function normalizeColor(
  value: unknown,
  name: string,
  code = "invalid-breathe-color",
): string {
  const parsed = motionHexColorSchema.safeParse(value);
  if (!parsed.success) {
    return fail(code, `${name} must be a #RRGGBB color`);
  }
  return parsed.data;
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
    (key) => !effectKeys.includes(key),
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
  const parsedPreset = effectPresetSchema.safeParse(effectValue["preset"]);
  if (!parsedPreset.success) {
    return fail(
      "unsupported-effect",
      `Unsupported animation effect ${String(effectValue["preset"])}`,
    );
  }
  const preset = parsedPreset.data;
  if (preset !== "breathe" && presetKinds[preset] !== expected) {
    return fail(
      "invalid-motion",
      expected === "entrance"
        ? `${preset}() is an emphasis effect and should be passed to emphasis`
        : `${preset}() is an entrance effect and should be passed to enter`,
    );
  }
  const parsedDuration = effectDurationSchema.safeParse(
    effectValue["duration"],
  );
  if (!parsedDuration.success) {
    return fail(
      "invalid-duration",
      "Animation duration must be writable as a finite positive integer in milliseconds",
    );
  }
  const duration = parsedDuration.data;
  const parsedDelay = effectDelaySchema.safeParse(effectValue["delay"]);
  if (
    !parsedDelay.success ||
    !Number.isFinite(Math.round(parsedDelay.data) + Math.round(duration))
  ) {
    return fail(
      "invalid-delay",
      "Animation delay must be writable as a finite non-negative integer in milliseconds",
    );
  }
  const delay = parsedDelay.data;
  if (preset === "breathe" && Math.round(delay) > maximumTimeNodeValue) {
    return fail(
      "invalid-delay",
      "breathe delay exceeds native timing value range",
    );
  }
  const by = effectValue["by"];
  const interval = effectValue["intervalMs"];
  if (preset === "spin" && typeof by === "number") {
    if (!spinAngleSchema.safeParse(by).success) {
      return fail(
        "invalid-rotation",
        "spin by converted value must fall within signed 32-bit angle range",
      );
    }
    if (effectValue["from"] !== undefined || effectValue["to"] !== undefined) {
      return fail(
        "invalid-motion",
        "spin from/to and by cannot be declared simultaneously",
      );
    }
  } else if (
    by !== undefined &&
    !motionGranularitySchema.safeParse(by).success
  ) {
    return fail(
      "invalid-motion",
      "by must be a supported text or chart granularity",
    );
  }
  if (
    interval !== undefined &&
    ((by !== "word" && by !== "letter") ||
      !intervalMsSchema.safeParse(interval).success)
  ) {
    return fail(
      "invalid-motion-interval",
      "intervalMs applies only to per-word or per-letter effects, and must be a positive integer in milliseconds",
    );
  }
  if (
    (preset === "wipe" || preset === "fly") &&
    !motionDirectionSchema.safeParse(effectValue["from"]).success
  ) {
    return fail(
      "invalid-direction",
      `${preset} from must be a valid direction`,
    );
  }
  if (preset === "spin") {
    const from = effectValue["from"];
    const to = effectValue["to"];
    if (typeof by !== "number" && !spinAngleSchema.safeParse(from).success) {
      return fail(
        "invalid-rotation",
        "spin from converted value must fall within signed 32-bit angle range",
      );
    }
    if (
      typeof by !== "number" &&
      to !== undefined &&
      !spinAngleSchema.safeParse(to).success
    ) {
      return fail(
        "invalid-rotation",
        "spin to converted value must fall within signed 32-bit angle range",
      );
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
  if (preset !== "spin" && effectValue["to"] !== undefined) {
    return fail("invalid-motion", `${preset} does not accept to parameter`);
  }
  const easing = effectValue["easing"];
  if (easing !== undefined) {
    if (!isRecord(easing)) {
      return fail("invalid-easing", "easing must be an object");
    }
    const parsedEasing = easingSchema.safeParse(easing);
    if (!parsedEasing.success) {
      const issue = parsedEasing.error.issues[0];
      if (issue?.code === "unrecognized_keys" && issue.keys[0] !== undefined) {
        return fail(
          "unknown-parameter",
          `easing contains unimplemented parameter ${issue.keys[0]}`,
        );
      }
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
    } else if (!repeatCountSchema.safeParse(repeatCount).success) {
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
    if (!repeatCountSchema.safeParse(repeatCount).success) {
      return fail(
        "invalid-repeat-count",
        `repeat count must be total rounds between 1 and ${maxRepeatCount}`,
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
  const parsed = motionStepAuthorPropsSchema.safeParse({ n: number });
  if (!parsed.success || parsed.data.n === undefined) {
    return fail(
      "invalid-step",
      "Explicit click step number must be a positive integer",
    );
  }
  return Object.freeze({ mode: "explicit" as const, number: parsed.data.n });
}

function listed(value: unknown, many: boolean): readonly unknown[] {
  if (value === undefined) return [];
  return many && Array.isArray(value) ? value : [value];
}

function validatedMoment(value: unknown): MotionMoment {
  const parsed = motionMomentSchema.safeParse(value);
  if (!parsed.success) {
    if (!isRecord(value) || value["type"] !== "motion-moment") {
      return fail(
        "invalid-moment",
        "Moment handle must come from a motion scope handle.",
      );
    }
    return fail(
      "invalid-moment",
      "Moment handle scope, type, or name is invalid",
    );
  }
  return parsed.data;
}

function validatedTiming(value: unknown): MotionTiming {
  if (isRecord(value) && value["type"] === "motion-offset") {
    const moment = validatedMoment(value["moment"]);
    const parsed = motionMomentOffsetSchema.safeParse({
      type: "motion-offset",
      moment,
      offsetMs: value["offsetMs"],
    });
    if (!parsed.success) {
      return fail(
        "invalid-delay",
        "Moment offset must be a finite non-negative integer in milliseconds",
      );
    }
    return Object.freeze(parsed.data);
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

const motionScopeAllowedParameters = new Set(motionScopeAuthorPropKeys);

const motionStepAllowedParameters = new Set([
  ...motionStepAuthorPropKeys,
  "stepIdentity",
]);

function localScopeDeclaration(
  props: Readonly<Record<string, unknown>>,
): MotionScopeDeclaration {
  const unknown = Object.keys(props).find(
    (key) => !motionScopeAllowedParameters.has(key),
  );
  if (unknown !== undefined) {
    return fail(
      "unknown-parameter",
      `MotionScope contains unimplemented parameter ${unknown}`,
    );
  }
  validateMotionFlag(props["motion"]);
  const parsedHandle = motionScopeHandleSchema.safeParse(props["scope"]);
  if (!parsedHandle.success) {
    return fail(
      "invalid-scope",
      "MotionScope requires a handle created by createMotionScopeHandle() or useMotionScope()",
    );
  }
  const handle = parsedHandle.data;
  const parsedMode = motionScopeModeSchema.safeParse(
    props["mode"] ?? "parallel",
  );
  if (!parsedMode.success) {
    return fail(
      "invalid-scope",
      "MotionScope mode must be parallel, sequence, or stagger",
    );
  }
  const mode = parsedMode.data;
  const parsedDelay = effectDelaySchema.safeParse(props["delay"] ?? 0);
  if (!parsedDelay.success) {
    return fail(
      "invalid-delay",
      "MotionScope delay must be a finite non-negative integer in milliseconds",
    );
  }
  const delay = parsedDelay.data;
  const each = props["each"];
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
    value: Object.freeze({
      type: "scope",
      scope,
      enabled: input["motion"] !== false,
    } satisfies LocalScopeValue),
    resolveContributions: resolveMotionContributions,
  };
}

/** Validates Step wrapper parameters and emits a transparent step annotation resolved by tree order. */
export function motionStepAnnotation(
  input: Readonly<Record<string, unknown>>,
): readonly CapabilityAnnotation[] {
  const unknown = Object.keys(input).find(
    (key) => !motionStepAllowedParameters.has(key),
  );
  if (unknown !== undefined) {
    return fail(
      "unknown-parameter",
      `Step contains unimplemented parameter ${unknown}`,
    );
  }
  const step = stepDeclarationFromProps(input);
  return [
    {
      capability: "motion",
      optional: true,
      value: Object.freeze({ type: "step", step } satisfies LocalStepValue),
      resolveContributions: resolveMotionContributions,
    },
  ];
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

/**
 * Converts container motion scope overrides and optional native target animation declarations into
 * local annotations consumed by motion slide capability without inspecting raw container props later.
 */
export function motionContainerAnnotations(
  props: Readonly<Record<string, unknown>>,
  target: MotionTarget | undefined,
  options: {
    readonly propagatesDefaultEnter?: boolean;
    readonly groupsAsParallelUnit?: boolean;
  } = {},
): readonly CapabilityAnnotation[] {
  validateMotionFlag(props["motion"]);
  const enabled = props["motion"] !== false;
  const propagatesDefaultEnter = options.propagatesDefaultEnter === true;
  const hasExplicitMotion =
    enabled &&
    (props["enter"] !== undefined ||
      props["emphasis"] !== undefined ||
      props["motion"] === true);
  const groupsAsParallelUnit =
    options.groupsAsParallelUnit === true && hasExplicitMotion;
  let defaultEnter: readonly NormalizedEffect[] | undefined;
  if (enabled && propagatesDefaultEnter && props["enter"] !== undefined) {
    defaultEnter = Object.freeze(
      listed(props["enter"], true).map((effect) =>
        validatedEffect(effect, "entrance"),
      ),
    );
  }
  const annotations: CapabilityAnnotation[] = [];
  if (!enabled || defaultEnter !== undefined || groupsAsParallelUnit) {
    annotations.push({
      capability: "motion",
      optional: true,
      value: Object.freeze({
        type: "container",
        enabled,
        ...(defaultEnter === undefined ? {} : { defaultEnter }),
        ...(groupsAsParallelUnit
          ? { groupsAsParallelUnit: true as const }
          : {}),
      } satisfies LocalContainerValue),
      resolveContributions: resolveMotionContributions,
    });
  }
  if (enabled && target !== undefined) {
    annotations.push(...motionAnnotations(props, target));
  }
  return annotations;
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
    typeof value["scope"]["identity"] === "string" &&
    typeof value["enabled"] === "boolean"
  );
}

function isLocalStepValue(value: unknown): value is LocalStepValue {
  return isRecord(value) && value["type"] === "step" && isRecord(value["step"]);
}

function isLocalContainerValue(value: unknown): value is LocalContainerValue {
  return (
    isRecord(value) &&
    value["type"] === "container" &&
    typeof value["enabled"] === "boolean"
  );
}

/**
 * Resolves local motion declarations against the current declaration tree.
 *
 * Scope parent chains, nearest click step, inherited container default entrance, and disabled
 * ancestor chains are derived from submitted capability annotations along the build's structure.
 * Nothing here is stored on the editable nodes or read from raw container props, so moving or
 * appending content changes the next build's choreography without stale React context.
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
    if (!("children" in node)) return;
    const localContributions = byDeclaration.get(node) ?? [];
    let childFrame = frame;

    const stepContribution = localContributions.find((contribution) =>
      isLocalStepValue(contribution.value),
    );
    const scopeContribution = localContributions.find((contribution) =>
      isLocalScopeValue(contribution.value),
    );
    const containerContribution = localContributions.find((contribution) =>
      isLocalContainerValue(contribution.value),
    );
    const localContribution = localContributions.find((contribution) =>
      isLocalMotionValue(contribution.value),
    );

    if (stepContribution !== undefined) {
      const localStep = stepContribution.value as LocalStepValue;
      childFrame = { ...childFrame, step: localStep.step };
    }

    if (scopeContribution !== undefined) {
      const local = scopeContribution.value as LocalScopeValue;
      const enabled = childFrame.enabled && local.enabled;
      const parent = childFrame.scopes.at(-1);
      const scope: MotionScopeDeclaration = Object.freeze({
        ...local.scope,
        ...(parent === undefined ? {} : { parent: parent.identity }),
        ...(childFrame.step === undefined ? {} : { step: childFrame.step }),
      });
      resolved.push(
        Object.freeze({
          ...scopeContribution,
          value: Object.freeze({ type: "scope" as const, scope }),
        }),
      );
      childFrame = {
        ...childFrame,
        enabled,
        scopes: Object.freeze([...childFrame.scopes, scope]),
      };
    }

    if (
      containerContribution !== undefined ||
      localContribution !== undefined
    ) {
      const container = containerContribution?.value as
        | LocalContainerValue
        | undefined;
      const enabled =
        childFrame.enabled && (container === undefined || container.enabled);
      // Group explicit-motion containers (e.g., Box with enter/motion) inside an outer MotionScope
      // as a single parallel unit before emitting the container's own background shape and children.
      if (
        container?.groupsAsParallelUnit === true &&
        enabled &&
        childFrame.scopes.length > 0
      ) {
        const parent = childFrame.scopes.at(-1);
        const implicitScope: MotionScopeDeclaration = Object.freeze({
          identity: createMotionIdentity("scope"),
          ...(parent === undefined ? {} : { parent: parent.identity }),
          ...(childFrame.step === undefined ? {} : { step: childFrame.step }),
          mode: "parallel",
          delay: 0,
          implicit: true,
        });
        resolved.push(
          Object.freeze({
            ...containerContribution!,
            optional: true,
            value: Object.freeze({
              type: "scope" as const,
              scope: implicitScope,
            }),
          }),
        );
        childFrame = {
          ...childFrame,
          scopes: Object.freeze([...childFrame.scopes, implicitScope]),
        };
      }
      if (localContribution !== undefined && enabled) {
        const local = localContribution.value as LocalMotionValue;
        const inherited = local.enterDeclared
          ? []
          : (childFrame.defaultEnter ?? []);
        const effects = Object.freeze([...inherited, ...local.effects]);
        if (effects.length > 0) {
          resolved.push(
            Object.freeze({
              ...localContribution,
              optional: false,
              value: Object.freeze({
                target: local.target,
                ...(childFrame.step === undefined
                  ? {}
                  : { step: childFrame.step }),
                scopes: childFrame.scopes,
                effects,
              } satisfies MotionValue),
            }),
          );
        }
      }
      if (container !== undefined) {
        childFrame = {
          ...childFrame,
          enabled,
          ...(container.defaultEnter === undefined
            ? {}
            : { defaultEnter: container.defaultEnter }),
        };
      }
    }

    for (const child of node.children) visit(child, childFrame);
  };

  for (const node of children) visit(node, { enabled: true, scopes: [] });
  return resolved;
}
