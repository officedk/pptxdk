import { colorValueSchema } from "@pptxdk/core/capability";
import { z } from "zod";

type AssertExact<A, B> =
  (<T>() => T extends A ? 1 : 2) extends <T>() => T extends B ? 1 : 2
    ? true
    : never;

/** ST_TLTime uses unsignedInt; repeatCount unit is one thousandth of a round. */
export const maximumTimeNodeValue = 4_294_967_295;
export const maxRepeatCount = Math.floor(maximumTimeNodeValue / 1000);
export const minimumRotationUnits = -2_147_483_648;
export const maximumRotationUnits = 2_147_483_647;

export type MotionDirection = "left" | "right" | "top" | "bottom";
const rawMotionDirectionSchema = z.enum(["left", "right", "top", "bottom"]);
export const motionDirectionSchema: z.ZodType<MotionDirection> =
  rawMotionDirectionSchema satisfies z.ZodType<
    AssertExact<
      MotionDirection,
      z.infer<typeof rawMotionDirectionSchema>
    > extends true
      ? MotionDirection
      : never
  >;

export type EntrancePreset =
  | "fade"
  | "wipe"
  | "zoom"
  | "fly"
  | "floatUp"
  | "floatDown"
  | "spinIn"
  | "stretch";
const rawEntrancePresetSchema = z.enum([
  "fade",
  "wipe",
  "zoom",
  "fly",
  "floatUp",
  "floatDown",
  "spinIn",
  "stretch",
]);
export const entrancePresetSchema: z.ZodType<EntrancePreset> =
  rawEntrancePresetSchema satisfies z.ZodType<
    AssertExact<
      EntrancePreset,
      z.infer<typeof rawEntrancePresetSchema>
    > extends true
      ? EntrancePreset
      : never
  >;

export type EmphasisPreset = "spin" | "breathe" | "fillColor";
const rawEmphasisPresetSchema = z.enum(["spin", "breathe", "fillColor"]);
export const emphasisPresetSchema: z.ZodType<EmphasisPreset> =
  rawEmphasisPresetSchema satisfies z.ZodType<
    AssertExact<
      EmphasisPreset,
      z.infer<typeof rawEmphasisPresetSchema>
    > extends true
      ? EmphasisPreset
      : never
  >;

export type EffectPreset = EntrancePreset | EmphasisPreset;
const rawEffectPresetSchema = z.enum([
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
export const effectPresetSchema: z.ZodType<EffectPreset> =
  rawEffectPresetSchema satisfies z.ZodType<
    AssertExact<
      EffectPreset,
      z.infer<typeof rawEffectPresetSchema>
    > extends true
      ? EffectPreset
      : never
  >;

export type RepeatCount = number | "forever";
const rawRepeatCountSchema = z.union([
  z.number().int().min(1).max(maxRepeatCount),
  z.literal("forever"),
]);
export const repeatCountSchema: z.ZodType<RepeatCount> =
  rawRepeatCountSchema satisfies z.ZodType<
    AssertExact<RepeatCount, z.infer<typeof rawRepeatCountSchema>> extends true
      ? RepeatCount
      : never
  >;

export type MotionGranularity =
  | "paragraph"
  | "word"
  | "letter"
  | "series"
  | "category"
  | "seriesElements"
  | "categoryElements";
const rawMotionGranularitySchema = z.enum([
  "paragraph",
  "word",
  "letter",
  "series",
  "category",
  "seriesElements",
  "categoryElements",
]);
export const motionGranularitySchema: z.ZodType<MotionGranularity> =
  rawMotionGranularitySchema satisfies z.ZodType<
    AssertExact<
      MotionGranularity,
      z.infer<typeof rawMotionGranularitySchema>
    > extends true
      ? MotionGranularity
      : never
  >;

export type MotionChartStep =
  | "gridLegend"
  | "series"
  | "category"
  | "ptInSeries"
  | "ptInCategory";
export const motionChartStepSchema = z.enum([
  "gridLegend",
  "series",
  "category",
  "ptInSeries",
  "ptInCategory",
]);

export type MotionTargetRange =
  | {
      readonly type: "paragraph";
      readonly start: number;
      readonly end: number;
    }
  | {
      readonly type: "chart";
      readonly step: MotionChartStep;
      readonly seriesIndex: number;
      readonly categoryIndex: number;
    };

const rawMotionTargetRangeSchema = z.discriminatedUnion("type", [
  z
    .strictObject({
      type: z.literal("paragraph"),
      start: z.number().int().nonnegative(),
      end: z.number().int().nonnegative(),
    })
    .readonly(),
  z
    .strictObject({
      type: z.literal("chart"),
      step: motionChartStepSchema,
      seriesIndex: z.number().int(),
      categoryIndex: z.number().int(),
    })
    .readonly(),
]);
export const motionTargetRangeSchema: z.ZodType<MotionTargetRange> =
  rawMotionTargetRangeSchema satisfies z.ZodType<
    AssertExact<
      MotionTargetRange,
      z.infer<typeof rawMotionTargetRangeSchema>
    > extends true
      ? MotionTargetRange
      : never
  >;

export type Easing = {
  readonly accelerate?: number | undefined;
  readonly decelerate?: number | undefined;
};

const rawEasingSchema = z
  .strictObject({
    accelerate: z.number().finite().min(0).max(1).optional(),
    decelerate: z.number().finite().min(0).max(1).optional(),
  })
  .readonly()
  .refine(
    (value) => (value.accelerate ?? 0) + (value.decelerate ?? 0) <= 1,
    "easing accelerate and decelerate proportions must be non-negative and sum to at most 1",
  );
export const easingSchema: z.ZodType<Easing> =
  rawEasingSchema satisfies z.ZodType<
    AssertExact<Easing, z.infer<typeof rawEasingSchema>> extends true
      ? Easing
      : never
  >;

/** Hex RGB color (`#RRGGBB`) validated against core's color value domain and normalized to uppercase. */
export const motionHexColorSchema: z.ZodType<string> = (
  colorValueSchema as unknown as z.ZodType<string>
)
  .pipe(z.string().regex(/^#[0-9a-f]{6}$/iu, "color must be a #RRGGBB color"))
  .transform((value) => value.toUpperCase());

export const effectDurationSchema = z
  .number()
  .finite()
  .positive()
  .refine(
    (value) =>
      Math.round(value) >= 1 && Number.isSafeInteger(Math.round(value)),
    "Animation duration must be writable as a finite positive integer in milliseconds",
  );

export const effectDelaySchema = z
  .number()
  .finite()
  .min(0)
  .refine(
    (value) => Number.isSafeInteger(Math.round(value)),
    "Animation delay must be writable as a finite non-negative integer in milliseconds",
  );

export const intervalMsSchema = z.number().int().min(1);

export const spinAngleSchema = z
  .number()
  .finite()
  .refine((value) => {
    const units = Math.round(value * 60000);
    return (
      Number.isSafeInteger(units) &&
      units >= minimumRotationUnits &&
      units <= maximumRotationUnits
    );
  }, "spin angle converted value must fall within signed 32-bit angle range");

export type EffectOptions = {
  readonly duration?: number | undefined;
  readonly delay?: number | undefined;
  readonly easing?: Easing | undefined;
  readonly intervalMs?: number | undefined;
  readonly by?: MotionGranularity | undefined;
};

const rawEffectOptionsSchema = z
  .strictObject({
    duration: effectDurationSchema.optional(),
    delay: effectDelaySchema.optional(),
    easing: rawEasingSchema.optional(),
    intervalMs: intervalMsSchema.optional(),
    by: rawMotionGranularitySchema.optional(),
  })
  .readonly();
export const effectOptionsSchema: z.ZodType<EffectOptions> =
  rawEffectOptionsSchema satisfies z.ZodType<
    AssertExact<
      EffectOptions,
      z.infer<typeof rawEffectOptionsSchema>
    > extends true
      ? EffectOptions
      : never
  >;

export type SpinDeltaOptions = Omit<EffectOptions, "by"> & {
  readonly by?: number | undefined;
  readonly from?: undefined;
  readonly to?: undefined;
};

export type SpinEndpointsOptions = EffectOptions & {
  readonly from?: number | undefined;
  readonly to?: number | undefined;
};

export type SpinOptions = SpinDeltaOptions | SpinEndpointsOptions;

export const spinOptionsSchema: z.ZodType<SpinOptions> = z
  .strictObject({
    duration: effectDurationSchema.optional(),
    delay: effectDelaySchema.optional(),
    easing: rawEasingSchema.optional(),
    intervalMs: intervalMsSchema.optional(),
    by: z.union([rawMotionGranularitySchema, spinAngleSchema]).optional(),
    from: spinAngleSchema.optional(),
    to: spinAngleSchema.optional(),
  })
  .readonly()
  .superRefine((value, ctx) => {
    if (
      typeof value.by === "number" &&
      (value.from !== undefined || value.to !== undefined)
    ) {
      ctx.addIssue({
        code: "custom",
        message: "spin() cannot combine from/to and by",
      });
    }
  })
  .transform((value) => value as SpinOptions);

export type FillColorOptions = {
  readonly duration?: number | undefined;
  readonly delay?: number | undefined;
  readonly easing?: Easing | undefined;
  readonly from: string;
  readonly to: string;
};

const rawFillColorOptionsSchema = z
  .strictObject({
    duration: effectDurationSchema.optional(),
    delay: effectDelaySchema.optional(),
    easing: rawEasingSchema.optional(),
    from: motionHexColorSchema,
    to: motionHexColorSchema,
  })
  .readonly();
export const fillColorOptionsSchema: z.ZodType<FillColorOptions> =
  rawFillColorOptionsSchema satisfies z.ZodType<
    AssertExact<
      FillColorOptions,
      z.infer<typeof rawFillColorOptionsSchema>
    > extends true
      ? FillColorOptions
      : never
  >;

export type BreatheOptions = {
  readonly period: number;
  readonly from: string;
  readonly to: string;
  readonly delay?: number | undefined;
  readonly easing?: Easing | undefined;
};

const breatheOptionsObjectSchema = z.strictObject({
  period: z
    .number()
    .int()
    .min(2)
    .refine(
      (value) => value % 2 === 0 && value / 2 <= maximumTimeNodeValue,
      "breathe period must be an even integer in milliseconds greater than 1",
    ),
  from: motionHexColorSchema,
  to: motionHexColorSchema,
  delay: effectDelaySchema
    .refine(
      (value) => Math.round(value) <= maximumTimeNodeValue,
      "breathe delay exceeds native timing value range",
    )
    .optional(),
  easing: rawEasingSchema.optional(),
});
export const breatheOptionsShape = breatheOptionsObjectSchema.shape;
const rawBreatheOptionsSchema = breatheOptionsObjectSchema.readonly();
export const breatheOptionsSchema: z.ZodType<BreatheOptions> =
  rawBreatheOptionsSchema satisfies z.ZodType<
    AssertExact<
      BreatheOptions,
      z.infer<typeof rawBreatheOptionsSchema>
    > extends true
      ? BreatheOptions
      : never
  >;

const effectCommonFields = {
  type: z.literal("effect"),
  duration: effectDurationSchema,
  delay: effectDelaySchema,
  from: z.union([rawMotionDirectionSchema, spinAngleSchema]).optional(),
  to: spinAngleSchema.optional(),
  fromColor: motionHexColorSchema.optional(),
  toColor: motionHexColorSchema.optional(),
  easing: rawEasingSchema.optional(),
  by: z.union([rawMotionGranularitySchema, spinAngleSchema]).optional(),
  intervalMs: intervalMsSchema.optional(),
};

type EffectCommon = {
  readonly type: "effect";
  readonly duration: number;
  readonly delay: number;
  readonly from?: MotionDirection | number | undefined;
  readonly to?: number | undefined;
  readonly fromColor?: string | undefined;
  readonly toColor?: string | undefined;
  readonly easing?: Easing | undefined;
  readonly by?: MotionGranularity | number | undefined;
  readonly intervalMs?: number | undefined;
};

export type EntranceEffect = EffectCommon & {
  readonly kind: "entrance";
  readonly preset: EntrancePreset;
};

const entranceEffectObjectSchema = z.strictObject({
  ...effectCommonFields,
  kind: z.literal("entrance"),
  preset: entrancePresetSchema,
});
export const entranceEffectShape = entranceEffectObjectSchema.shape;
const rawEntranceEffectSchema = entranceEffectObjectSchema.readonly();
export const entranceEffectSchema: z.ZodType<EntranceEffect> =
  rawEntranceEffectSchema satisfies z.ZodType<
    AssertExact<
      EntranceEffect,
      z.infer<typeof rawEntranceEffectSchema>
    > extends true
      ? EntranceEffect
      : never
  >;

export type EmphasisEffect = EffectCommon & {
  readonly kind: "emphasis";
  readonly preset: EmphasisPreset;
};

const emphasisEffectObjectSchema = z.strictObject({
  ...effectCommonFields,
  kind: z.literal("emphasis"),
  preset: emphasisPresetSchema,
});
const rawEmphasisEffectSchema = emphasisEffectObjectSchema.readonly();
export const emphasisEffectSchema: z.ZodType<EmphasisEffect> =
  rawEmphasisEffectSchema satisfies z.ZodType<
    AssertExact<
      EmphasisEffect,
      z.infer<typeof rawEmphasisEffectSchema>
    > extends true
      ? EmphasisEffect
      : never
  >;

export type BreatheEffect = EffectCommon & {
  readonly kind: "emphasis";
  readonly preset: "breathe";
  readonly fromColor: string;
  readonly toColor: string;
};

const rawBreatheEffectSchema = z
  .strictObject({
    ...effectCommonFields,
    kind: z.literal("emphasis"),
    preset: z.literal("breathe"),
    fromColor: motionHexColorSchema,
    toColor: motionHexColorSchema,
  })
  .readonly();
export const breatheEffectSchema: z.ZodType<BreatheEffect> =
  rawBreatheEffectSchema satisfies z.ZodType<
    AssertExact<
      BreatheEffect,
      z.infer<typeof rawBreatheEffectSchema>
    > extends true
      ? BreatheEffect
      : never
  >;

export type Effect = EntranceEffect | EmphasisEffect;
const rawEffectSchema = z.discriminatedUnion("kind", [
  rawEntranceEffectSchema,
  rawEmphasisEffectSchema,
]);
export const effectSchema: z.ZodType<Effect> =
  rawEffectSchema satisfies z.ZodType<
    AssertExact<Effect, z.infer<typeof rawEffectSchema>> extends true
      ? Effect
      : never
  >;

export type RepeatedEffect<E extends Effect = Effect> = {
  readonly type: "repeated-effect";
  readonly effect: E;
  readonly count: RepeatCount;
};

export const repeatedEffectSchema: z.ZodType<RepeatedEffect> = z
  .strictObject({
    type: z.literal("repeated-effect"),
    effect: rawEffectSchema,
    count: rawRepeatCountSchema,
  })
  .readonly();

export type EntranceEffectDeclaration =
  | EntranceEffect
  | RepeatedEffect<EntranceEffect>;

const rawEntranceEffectDeclarationSchema = z.union([
  rawEntranceEffectSchema,
  z
    .strictObject({
      type: z.literal("repeated-effect"),
      effect: rawEntranceEffectSchema,
      count: rawRepeatCountSchema,
    })
    .readonly(),
]);
export const entranceEffectDeclarationSchema: z.ZodType<EntranceEffectDeclaration> =
  rawEntranceEffectDeclarationSchema satisfies z.ZodType<
    AssertExact<
      EntranceEffectDeclaration,
      z.infer<typeof rawEntranceEffectDeclarationSchema>
    > extends true
      ? EntranceEffectDeclaration
      : never
  >;

export type EmphasisEffectDeclaration =
  | EmphasisEffect
  | RepeatedEffect<EmphasisEffect>;

const rawEmphasisEffectDeclarationSchema = z.union([
  rawEmphasisEffectSchema,
  z
    .strictObject({
      type: z.literal("repeated-effect"),
      effect: rawEmphasisEffectSchema,
      count: rawRepeatCountSchema,
    })
    .readonly(),
]);
export const emphasisEffectDeclarationSchema: z.ZodType<EmphasisEffectDeclaration> =
  rawEmphasisEffectDeclarationSchema satisfies z.ZodType<
    AssertExact<
      EmphasisEffectDeclaration,
      z.infer<typeof rawEmphasisEffectDeclarationSchema>
    > extends true
      ? EmphasisEffectDeclaration
      : never
  >;

export type NormalizedEffect =
  | (EntranceEffect & { readonly repeatCount?: RepeatCount | undefined })
  | (EmphasisEffect & { readonly repeatCount?: RepeatCount | undefined });

const rawNormalizedEffectSchema = z.discriminatedUnion("kind", [
  entranceEffectObjectSchema
    .extend({ repeatCount: rawRepeatCountSchema.optional() })
    .readonly(),
  emphasisEffectObjectSchema
    .extend({ repeatCount: rawRepeatCountSchema.optional() })
    .readonly(),
]);
export const normalizedEffectSchema: z.ZodType<NormalizedEffect> =
  rawNormalizedEffectSchema satisfies z.ZodType<
    AssertExact<
      NormalizedEffect,
      z.infer<typeof rawNormalizedEffectSchema>
    > extends true
      ? NormalizedEffect
      : never
  >;

export type MotionScopeMode = "parallel" | "sequence" | "stagger";
const rawMotionScopeModeSchema = z.enum(["parallel", "sequence", "stagger"]);
export const motionScopeModeSchema: z.ZodType<MotionScopeMode> =
  rawMotionScopeModeSchema satisfies z.ZodType<
    AssertExact<
      MotionScopeMode,
      z.infer<typeof rawMotionScopeModeSchema>
    > extends true
      ? MotionScopeMode
      : never
  >;

export type MotionCompletionState =
  | "known"
  | "unknown"
  | "no-natural-completion";
const rawMotionCompletionStateSchema = z.enum([
  "known",
  "unknown",
  "no-natural-completion",
]);
export const motionCompletionStateSchema: z.ZodType<MotionCompletionState> =
  rawMotionCompletionStateSchema satisfies z.ZodType<
    AssertExact<
      MotionCompletionState,
      z.infer<typeof rawMotionCompletionStateSchema>
    > extends true
      ? MotionCompletionState
      : never
  >;

export type MotionMoment = {
  readonly type: "motion-moment";
  readonly scope: string;
  readonly point: "start" | "end" | "named";
  readonly name?: string | undefined;
};

const rawMotionMomentSchema = z
  .strictObject({
    type: z.literal("motion-moment"),
    scope: z.string().min(1),
    point: z.enum(["start", "end", "named"]),
    name: z.string().min(1).optional(),
  })
  .readonly()
  .refine(
    (value) =>
      value.point === "named"
        ? value.name !== undefined && value.name.trim() !== ""
        : value.name === undefined,
    "Moment handle scope, type, or name is invalid",
  );
export const motionMomentSchema: z.ZodType<MotionMoment> =
  rawMotionMomentSchema satisfies z.ZodType<
    AssertExact<
      MotionMoment,
      z.infer<typeof rawMotionMomentSchema>
    > extends true
      ? MotionMoment
      : never
  >;

export type MotionMomentOffset = {
  readonly type: "motion-offset";
  readonly moment: MotionMoment;
  readonly offsetMs: number;
};

const rawMotionMomentOffsetSchema = z
  .strictObject({
    type: z.literal("motion-offset"),
    moment: rawMotionMomentSchema,
    offsetMs: z.number().int().min(0),
  })
  .readonly();
export const motionMomentOffsetSchema: z.ZodType<MotionMomentOffset> =
  rawMotionMomentOffsetSchema satisfies z.ZodType<
    AssertExact<
      MotionMomentOffset,
      z.infer<typeof rawMotionMomentOffsetSchema>
    > extends true
      ? MotionMomentOffset
      : never
  >;

export type MotionTiming = MotionMoment | MotionMomentOffset;
const rawMotionTimingSchema = z.union([
  rawMotionMomentSchema,
  rawMotionMomentOffsetSchema,
]);
export const motionTimingSchema: z.ZodType<MotionTiming> =
  rawMotionTimingSchema satisfies z.ZodType<
    AssertExact<
      MotionTiming,
      z.infer<typeof rawMotionTimingSchema>
    > extends true
      ? MotionTiming
      : never
  >;

export type MotionScopeHandle = {
  readonly type: "motion-scope-handle";
  readonly identity: string;
  readonly start: MotionMoment;
  readonly end: MotionMoment;
  cue(name: string): MotionMoment;
};

/** Runtime-identity-preserving schema for handles created by `createMotionScopeHandle()` / `useMotionScope()`. */
export const motionScopeHandleSchema: z.ZodType<MotionScopeHandle> =
  z.custom<MotionScopeHandle>(
    (value) =>
      typeof value === "object" &&
      value !== null &&
      !Array.isArray(value) &&
      (value as { readonly type?: unknown }).type === "motion-scope-handle" &&
      typeof (value as { readonly identity?: unknown }).identity === "string" &&
      (value as { readonly identity: string }).identity !== "" &&
      rawMotionMomentSchema.safeParse(
        (value as { readonly start?: unknown }).start,
      ).success &&
      rawMotionMomentSchema.safeParse((value as { readonly end?: unknown }).end)
        .success &&
      typeof (value as { readonly cue?: unknown }).cue === "function",
    "MotionScope requires a handle created by createMotionScopeHandle() or useMotionScope()",
  );

export type MotionProps = {
  readonly enter?:
    | EntranceEffectDeclaration
    | readonly EntranceEffectDeclaration[]
    | undefined;
  readonly emphasis?: EmphasisEffectDeclaration | undefined;
  readonly motion?: boolean | undefined;
};

const motionPropsObjectSchema = z.strictObject({
  enter: z
    .union([
      rawEntranceEffectDeclarationSchema,
      z.array(rawEntranceEffectDeclarationSchema).readonly(),
    ])
    .optional(),
  emphasis: rawEmphasisEffectDeclarationSchema.optional(),
  motion: z.boolean().optional(),
});
export const motionPropsShape = motionPropsObjectSchema.shape;
const rawMotionPropsSchema = motionPropsObjectSchema.readonly();
export const motionPropsSchema: z.ZodType<MotionProps> =
  rawMotionPropsSchema satisfies z.ZodType<
    AssertExact<MotionProps, z.infer<typeof rawMotionPropsSchema>> extends true
      ? MotionProps
      : never
  >;

export type MotionScopeAuthorProps = {
  readonly scope: MotionScopeHandle;
  readonly mode?: MotionScopeMode | undefined;
  readonly each?: number | undefined;
  readonly delay?: number | undefined;
  readonly startAt?: MotionTiming | undefined;
  readonly publishStart?: MotionMoment | undefined;
  readonly publishEnd?: MotionMoment | undefined;
  readonly motion?: boolean | undefined;
};

const motionScopeAuthorPropsObjectSchema = z.strictObject({
  scope: motionScopeHandleSchema,
  mode: rawMotionScopeModeSchema.optional(),
  each: z.number().int().min(0).optional(),
  delay: effectDelaySchema.optional(),
  startAt: rawMotionTimingSchema.optional(),
  publishStart: rawMotionMomentSchema.optional(),
  publishEnd: rawMotionMomentSchema.optional(),
  motion: z.boolean().optional(),
});
const rawMotionScopeAuthorPropsSchema =
  motionScopeAuthorPropsObjectSchema.readonly();
export const motionScopeAuthorPropKeys: readonly string[] = Object.keys(
  motionScopeAuthorPropsObjectSchema.shape,
);
export const motionScopeAuthorPropsSchema: z.ZodType<MotionScopeAuthorProps> =
  rawMotionScopeAuthorPropsSchema satisfies z.ZodType<
    AssertExact<
      MotionScopeAuthorProps,
      z.infer<typeof rawMotionScopeAuthorPropsSchema>
    > extends true
      ? MotionScopeAuthorProps
      : never
  >;

export type MotionStepAuthorProps = {
  readonly n?: number | undefined;
};

const motionStepAuthorPropsObjectSchema = z.strictObject({
  n: z.number().int().min(1).optional(),
});
const rawMotionStepAuthorPropsSchema =
  motionStepAuthorPropsObjectSchema.readonly();
export const motionStepAuthorPropKeys: readonly string[] = Object.keys(
  motionStepAuthorPropsObjectSchema.shape,
);
export const motionStepAuthorPropsSchema: z.ZodType<MotionStepAuthorProps> =
  rawMotionStepAuthorPropsSchema satisfies z.ZodType<
    AssertExact<
      MotionStepAuthorProps,
      z.infer<typeof rawMotionStepAuthorPropsSchema>
    > extends true
      ? MotionStepAuthorProps
      : never
  >;

/** Structured target and step timing declaration retained directly on `SlideMaster` and `SlideLayout` owners. */
export type OwnerMotionTarget =
  | { readonly placeholder: string }
  | { readonly spid: number };

const rawOwnerMotionTargetSchema = z.union([
  z.strictObject({ placeholder: z.string().min(1) }).readonly(),
  z.strictObject({ spid: z.number().int().positive() }).readonly(),
]);
/** Structured target and step timing declaration retained directly on `SlideMaster` and `SlideLayout` owners. */
export const ownerMotionTargetSchema: z.ZodType<OwnerMotionTarget> =
  rawOwnerMotionTargetSchema satisfies z.ZodType<
    AssertExact<
      OwnerMotionTarget,
      z.infer<typeof rawOwnerMotionTargetSchema>
    > extends true
      ? OwnerMotionTarget
      : never
  >;

export type OwnerMotionEffect = {
  readonly target: OwnerMotionTarget;
  readonly effect: EntranceEffectDeclaration | EmphasisEffectDeclaration;
  readonly range?: MotionTargetRange | undefined;
};

const rawOwnerMotionEffectSchema = z
  .strictObject({
    target: rawOwnerMotionTargetSchema,
    effect: z.union([
      rawEntranceEffectDeclarationSchema,
      rawEmphasisEffectDeclarationSchema,
    ]),
    range: rawMotionTargetRangeSchema.optional(),
  })
  .readonly();
export const ownerMotionEffectSchema: z.ZodType<OwnerMotionEffect> =
  rawOwnerMotionEffectSchema satisfies z.ZodType<
    AssertExact<
      OwnerMotionEffect,
      z.infer<typeof rawOwnerMotionEffectSchema>
    > extends true
      ? OwnerMotionEffect
      : never
  >;

export type OwnerMotionStep = {
  readonly trigger: "auto" | "click";
  readonly effects: readonly OwnerMotionEffect[];
};

const rawOwnerMotionStepSchema = z
  .strictObject({
    trigger: z.enum(["auto", "click"]),
    effects: z.array(rawOwnerMotionEffectSchema).readonly(),
  })
  .readonly();
export const ownerMotionStepSchema: z.ZodType<OwnerMotionStep> =
  rawOwnerMotionStepSchema satisfies z.ZodType<
    AssertExact<
      OwnerMotionStep,
      z.infer<typeof rawOwnerMotionStepSchema>
    > extends true
      ? OwnerMotionStep
      : never
  >;

export type OwnerMotion = {
  readonly steps: readonly OwnerMotionStep[];
};

const rawOwnerMotionSchema = z
  .strictObject({
    steps: z.array(rawOwnerMotionStepSchema).readonly(),
  })
  .readonly();
export const ownerMotionSchema: z.ZodType<OwnerMotion> =
  rawOwnerMotionSchema satisfies z.ZodType<
    AssertExact<OwnerMotion, z.infer<typeof rawOwnerMotionSchema>> extends true
      ? OwnerMotion
      : never
  >;
