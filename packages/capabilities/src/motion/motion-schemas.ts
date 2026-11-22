import { colorValueSchema } from "@pptxdk/core/capability";
import { z } from "zod";

/** ST_TLTime uses unsignedInt; repeatCount unit is one thousandth of a round. */
export const maximumTimeNodeValue = 4_294_967_295;
export const maxRepeatCount = Math.floor(maximumTimeNodeValue / 1000);
export const minimumRotationUnits = -2_147_483_648;
export const maximumRotationUnits = 2_147_483_647;

export const MotionDirectionSchema = z.enum(["left", "right", "top", "bottom"]);
export type MotionDirection = z.infer<typeof MotionDirectionSchema>;

export const EntrancePresetSchema = z.enum([
  "fade",
  "wipe",
  "zoom",
  "fly",
  "floatUp",
  "floatDown",
  "spinIn",
  "stretch",
]);

export const EmphasisPresetSchema = z.enum(["spin", "breathe", "fillColor"]);

export const EffectPresetSchema = z.enum([
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
export type EffectPreset = z.infer<typeof EffectPresetSchema>;

export const RepeatCountSchema = z.union([
  z.number().int().min(1).max(maxRepeatCount),
  z.literal("forever"),
]);
export type RepeatCount = z.infer<typeof RepeatCountSchema>;

export const MotionGranularitySchema = z.enum([
  "paragraph",
  "word",
  "letter",
  "series",
  "category",
  "seriesElements",
  "categoryElements",
]);
export type MotionGranularity = z.infer<typeof MotionGranularitySchema>;

export const MotionChartStepSchema = z.enum([
  "gridLegend",
  "series",
  "category",
  "ptInSeries",
  "ptInCategory",
]);

export const MotionTargetRangeSchema = z.discriminatedUnion("type", [
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
      step: MotionChartStepSchema,
      seriesIndex: z.number().int(),
      categoryIndex: z.number().int(),
    })
    .readonly(),
]);
export type MotionTargetRange = z.infer<typeof MotionTargetRangeSchema>;

export const EasingSchema = z
  .strictObject({
    accelerate: z.number().finite().min(0).max(1).optional(),
    decelerate: z.number().finite().min(0).max(1).optional(),
  })
  .readonly()
  .refine(
    (value) => (value.accelerate ?? 0) + (value.decelerate ?? 0) <= 1,
    "easing accelerate and decelerate proportions must be non-negative and sum to at most 1",
  );
export type Easing = z.infer<typeof EasingSchema>;

/** Hex RGB color (`#RRGGBB`) validated against core's color value domain and normalized to uppercase. */
export const MotionHexColorSchema = colorValueSchema
  .pipe(z.string().regex(/^#[0-9a-f]{6}$/iu, "color must be a #RRGGBB color"))
  .transform((value) => value.toUpperCase());

export const EffectDurationSchema = z
  .number()
  .finite()
  .positive()
  .refine(
    (value) =>
      Math.round(value) >= 1 && Number.isSafeInteger(Math.round(value)),
    "Animation duration must be writable as a finite positive integer in milliseconds",
  );

export const EffectDelaySchema = z
  .number()
  .finite()
  .min(0)
  .refine(
    (value) => Number.isSafeInteger(Math.round(value)),
    "Animation delay must be writable as a finite non-negative integer in milliseconds",
  );

export const IntervalMsSchema = z.number().int().min(1);

export const SpinAngleSchema = z
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

export const EffectOptionsSchema = z
  .strictObject({
    duration: EffectDurationSchema.optional(),
    delay: EffectDelaySchema.optional(),
    easing: EasingSchema.optional(),
    intervalMs: IntervalMsSchema.optional(),
    by: MotionGranularitySchema.optional(),
  })
  .readonly();
export type EffectOptions = z.infer<typeof EffectOptionsSchema>;

export const SpinOptionsSchema = z
  .strictObject({
    duration: EffectDurationSchema.optional(),
    delay: EffectDelaySchema.optional(),
    easing: EasingSchema.optional(),
    intervalMs: IntervalMsSchema.optional(),
    by: z.union([MotionGranularitySchema, SpinAngleSchema]).optional(),
    from: SpinAngleSchema.optional(),
  })
  .readonly();
export type SpinOptions = z.infer<typeof SpinOptionsSchema>;

export const FillColorOptionsSchema = z
  .strictObject({
    duration: EffectDurationSchema.optional(),
    delay: EffectDelaySchema.optional(),
    easing: EasingSchema.optional(),
    from: MotionHexColorSchema,
    to: MotionHexColorSchema,
  })
  .readonly();
export type FillColorOptions = z.infer<typeof FillColorOptionsSchema>;

export const BreatheOptionsSchema = z
  .strictObject({
    period: z
      .number()
      .int()
      .min(2)
      .refine(
        (value) => value % 2 === 0 && value / 2 <= maximumTimeNodeValue,
        "breathe period must be an even integer in milliseconds greater than 1",
      ),
    from: MotionHexColorSchema,
    to: MotionHexColorSchema,
    delay: EffectDelaySchema.refine(
      (value) => Math.round(value) <= maximumTimeNodeValue,
      "breathe delay exceeds native timing value range",
    ).optional(),
    easing: EasingSchema.optional(),
  })
  .readonly();
export type BreatheOptions = z.infer<typeof BreatheOptionsSchema>;

const effectCommonFields = {
  type: z.literal("effect"),
  duration: EffectDurationSchema,
  delay: EffectDelaySchema,
  from: z.union([MotionDirectionSchema, SpinAngleSchema]).optional(),
  fromColor: MotionHexColorSchema.optional(),
  toColor: MotionHexColorSchema.optional(),
  easing: EasingSchema.optional(),
  by: z.union([MotionGranularitySchema, SpinAngleSchema]).optional(),
  intervalMs: IntervalMsSchema.optional(),
};

const entranceEffectObjectSchema = z.strictObject({
  ...effectCommonFields,
  kind: z.literal("entrance"),
  preset: EntrancePresetSchema,
});

export const EntranceEffectSchema = entranceEffectObjectSchema.readonly();
export type EntranceEffect = z.infer<typeof EntranceEffectSchema>;

const emphasisEffectObjectSchema = z.strictObject({
  ...effectCommonFields,
  kind: z.literal("emphasis"),
  preset: EmphasisPresetSchema,
});

export const EmphasisEffectSchema = emphasisEffectObjectSchema.readonly();
export type EmphasisEffect = z.infer<typeof EmphasisEffectSchema>;

export const BreatheEffectSchema = z
  .strictObject({
    ...effectCommonFields,
    kind: z.literal("emphasis"),
    preset: z.literal("breathe"),
    fromColor: MotionHexColorSchema,
    toColor: MotionHexColorSchema,
  })
  .readonly();
export type BreatheEffect = z.infer<typeof BreatheEffectSchema>;

export const EffectSchema = z.discriminatedUnion("kind", [
  EntranceEffectSchema,
  EmphasisEffectSchema,
]);
export type Effect = z.infer<typeof EffectSchema>;

export const RepeatedEffectSchema = z
  .strictObject({
    type: z.literal("repeated-effect"),
    effect: EffectSchema,
    count: RepeatCountSchema,
  })
  .readonly();

export type RepeatedEffect<E extends Effect = Effect> = {
  readonly type: "repeated-effect";
  readonly effect: E;
  readonly count: RepeatCount;
};

export const EntranceEffectDeclarationSchema = z.union([
  EntranceEffectSchema,
  z
    .strictObject({
      type: z.literal("repeated-effect"),
      effect: EntranceEffectSchema,
      count: RepeatCountSchema,
    })
    .readonly(),
]);
export type EntranceEffectDeclaration =
  | EntranceEffect
  | RepeatedEffect<EntranceEffect>;

export const EmphasisEffectDeclarationSchema = z.union([
  EmphasisEffectSchema,
  z
    .strictObject({
      type: z.literal("repeated-effect"),
      effect: EmphasisEffectSchema,
      count: RepeatCountSchema,
    })
    .readonly(),
]);
export type EmphasisEffectDeclaration =
  | EmphasisEffect
  | RepeatedEffect<EmphasisEffect>;

export const NormalizedEffectSchema = z.discriminatedUnion("kind", [
  entranceEffectObjectSchema
    .extend({ repeatCount: RepeatCountSchema.optional() })
    .readonly(),
  emphasisEffectObjectSchema
    .extend({ repeatCount: RepeatCountSchema.optional() })
    .readonly(),
]);
export type NormalizedEffect = z.infer<typeof NormalizedEffectSchema>;

export const MotionScopeModeSchema = z.enum([
  "parallel",
  "sequence",
  "stagger",
]);
export type MotionScopeMode = z.infer<typeof MotionScopeModeSchema>;

export const MotionCompletionStateSchema = z.enum([
  "known",
  "unknown",
  "no-natural-completion",
]);
export type MotionCompletionState = z.infer<typeof MotionCompletionStateSchema>;

export const MotionMomentSchema = z
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
export type MotionMoment = z.infer<typeof MotionMomentSchema>;

export const MotionMomentOffsetSchema = z
  .strictObject({
    type: z.literal("motion-offset"),
    moment: MotionMomentSchema,
    offsetMs: z.number().int().min(0),
  })
  .readonly();
export type MotionMomentOffset = z.infer<typeof MotionMomentOffsetSchema>;

export const MotionTimingSchema = z.union([
  MotionMomentSchema,
  MotionMomentOffsetSchema,
]);
export type MotionTiming = z.infer<typeof MotionTimingSchema>;

export type MotionScopeHandle = {
  readonly type: "motion-scope-handle";
  readonly identity: string;
  readonly start: MotionMoment;
  readonly end: MotionMoment;
  cue(name: string): MotionMoment;
};

/** Runtime-identity-preserving schema for handles created by `createMotionScopeHandle()` / `useMotionScope()`. */
export const MotionScopeHandleSchema = z.custom<MotionScopeHandle>(
  (value) =>
    typeof value === "object" &&
    value !== null &&
    !Array.isArray(value) &&
    (value as { readonly type?: unknown }).type === "motion-scope-handle" &&
    typeof (value as { readonly identity?: unknown }).identity === "string" &&
    (value as { readonly identity: string }).identity !== "" &&
    MotionMomentSchema.safeParse((value as { readonly start?: unknown }).start)
      .success &&
    MotionMomentSchema.safeParse((value as { readonly end?: unknown }).end)
      .success &&
    typeof (value as { readonly cue?: unknown }).cue === "function",
  "MotionScope requires a handle created by createMotionScopeHandle() or useMotionScope()",
);

export const MotionPropsSchema = z
  .strictObject({
    enter: z
      .union([
        EntranceEffectDeclarationSchema,
        z.array(EntranceEffectDeclarationSchema).readonly(),
      ])
      .optional(),
    emphasis: EmphasisEffectDeclarationSchema.optional(),
    motion: z.boolean().optional(),
  })
  .readonly();
export type MotionProps = z.infer<typeof MotionPropsSchema>;

export const MotionScopeAuthorPropsSchema = z
  .strictObject({
    scope: MotionScopeHandleSchema,
    mode: MotionScopeModeSchema.optional(),
    each: z.number().int().min(0).optional(),
    delay: EffectDelaySchema.optional(),
    startAt: MotionTimingSchema.optional(),
    publishStart: MotionMomentSchema.optional(),
    publishEnd: MotionMomentSchema.optional(),
    motion: z.boolean().optional(),
  })
  .readonly();
export type MotionScopeAuthorProps = z.infer<
  typeof MotionScopeAuthorPropsSchema
>;

export const MotionStepAuthorPropsSchema = z
  .strictObject({
    n: z.number().int().min(1).optional(),
  })
  .readonly();
export type MotionStepAuthorProps = z.infer<typeof MotionStepAuthorPropsSchema>;

/** Structured target and step timing declaration retained directly on `SlideMaster` and `SlideLayout` owners. */
export const OwnerMotionTargetSchema = z.union([
  z.strictObject({ placeholder: z.string().min(1) }).readonly(),
  z.strictObject({ spid: z.number().int().positive() }).readonly(),
]);
export type OwnerMotionTarget = z.infer<typeof OwnerMotionTargetSchema>;

export const OwnerMotionEffectSchema = z
  .strictObject({
    target: OwnerMotionTargetSchema,
    effect: z.union([
      EntranceEffectDeclarationSchema,
      EmphasisEffectDeclarationSchema,
    ]),
    range: MotionTargetRangeSchema.optional(),
  })
  .readonly();
export type OwnerMotionEffect = z.infer<typeof OwnerMotionEffectSchema>;

export const OwnerMotionStepSchema = z
  .strictObject({
    trigger: z.enum(["auto", "click"]),
    effects: z.array(OwnerMotionEffectSchema).readonly(),
  })
  .readonly();
export type OwnerMotionStep = z.infer<typeof OwnerMotionStepSchema>;

export const OwnerMotionSchema = z
  .strictObject({
    steps: z.array(OwnerMotionStepSchema).readonly(),
  })
  .readonly();
export type OwnerMotion = z.infer<typeof OwnerMotionSchema>;
