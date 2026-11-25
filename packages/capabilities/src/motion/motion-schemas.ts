import { colorValueSchema } from "@pptxdk/core/capability";
import { z } from "zod";

/** ST_TLTime uses unsignedInt; repeatCount unit is one thousandth of a round. */
export const maximumTimeNodeValue = 4_294_967_295;
export const maxRepeatCount = Math.floor(maximumTimeNodeValue / 1000);
export const minimumRotationUnits = -2_147_483_648;
export const maximumRotationUnits = 2_147_483_647;

export const motionDirectionSchema = z.enum(["left", "right", "top", "bottom"]);
export type MotionDirection = z.infer<typeof motionDirectionSchema>;

export const entrancePresetSchema = z.enum([
  "fade",
  "wipe",
  "zoom",
  "fly",
  "floatUp",
  "floatDown",
  "spinIn",
  "stretch",
]);

export const emphasisPresetSchema = z.enum(["spin", "breathe", "fillColor"]);

export const effectPresetSchema = z.enum([
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
export type EffectPreset = z.infer<typeof effectPresetSchema>;

export const repeatCountSchema = z.union([
  z.number().int().min(1).max(maxRepeatCount),
  z.literal("forever"),
]);
export type RepeatCount = z.infer<typeof repeatCountSchema>;

export const motionGranularitySchema = z.enum([
  "paragraph",
  "word",
  "letter",
  "series",
  "category",
  "seriesElements",
  "categoryElements",
]);
export type MotionGranularity = z.infer<typeof motionGranularitySchema>;

export const motionChartStepSchema = z.enum([
  "gridLegend",
  "series",
  "category",
  "ptInSeries",
  "ptInCategory",
]);

export const motionTargetRangeSchema = z.discriminatedUnion("type", [
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
export type MotionTargetRange = z.infer<typeof motionTargetRangeSchema>;

export const easingSchema = z
  .strictObject({
    accelerate: z.number().finite().min(0).max(1).optional(),
    decelerate: z.number().finite().min(0).max(1).optional(),
  })
  .readonly()
  .refine(
    (value) => (value.accelerate ?? 0) + (value.decelerate ?? 0) <= 1,
    "easing accelerate and decelerate proportions must be non-negative and sum to at most 1",
  );
export type Easing = z.infer<typeof easingSchema>;

/** Hex RGB color (`#RRGGBB`) validated against core's color value domain and normalized to uppercase. */
export const motionHexColorSchema = colorValueSchema
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

export const effectOptionsSchema = z
  .strictObject({
    duration: effectDurationSchema.optional(),
    delay: effectDelaySchema.optional(),
    easing: easingSchema.optional(),
    intervalMs: intervalMsSchema.optional(),
    by: motionGranularitySchema.optional(),
  })
  .readonly();
export type EffectOptions = z.infer<typeof effectOptionsSchema>;

export const spinOptionsSchema = z
  .strictObject({
    duration: effectDurationSchema.optional(),
    delay: effectDelaySchema.optional(),
    easing: easingSchema.optional(),
    intervalMs: intervalMsSchema.optional(),
    by: z.union([motionGranularitySchema, spinAngleSchema]).optional(),
    from: spinAngleSchema.optional(),
  })
  .readonly();
export type SpinOptions = z.infer<typeof spinOptionsSchema>;

export const fillColorOptionsSchema = z
  .strictObject({
    duration: effectDurationSchema.optional(),
    delay: effectDelaySchema.optional(),
    easing: easingSchema.optional(),
    from: motionHexColorSchema,
    to: motionHexColorSchema,
  })
  .readonly();
export type FillColorOptions = z.infer<typeof fillColorOptionsSchema>;

export const breatheOptionsSchema = z
  .strictObject({
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
    easing: easingSchema.optional(),
  })
  .readonly();
export type BreatheOptions = z.infer<typeof breatheOptionsSchema>;

const effectCommonFields = {
  type: z.literal("effect"),
  duration: effectDurationSchema,
  delay: effectDelaySchema,
  from: z.union([motionDirectionSchema, spinAngleSchema]).optional(),
  fromColor: motionHexColorSchema.optional(),
  toColor: motionHexColorSchema.optional(),
  easing: easingSchema.optional(),
  by: z.union([motionGranularitySchema, spinAngleSchema]).optional(),
  intervalMs: intervalMsSchema.optional(),
};

const entranceEffectObjectSchema = z.strictObject({
  ...effectCommonFields,
  kind: z.literal("entrance"),
  preset: entrancePresetSchema,
});

export const entranceEffectSchema = entranceEffectObjectSchema.readonly();
export type EntranceEffect = z.infer<typeof entranceEffectSchema>;

const emphasisEffectObjectSchema = z.strictObject({
  ...effectCommonFields,
  kind: z.literal("emphasis"),
  preset: emphasisPresetSchema,
});

export const emphasisEffectSchema = emphasisEffectObjectSchema.readonly();
export type EmphasisEffect = z.infer<typeof emphasisEffectSchema>;

export const breatheEffectSchema = z
  .strictObject({
    ...effectCommonFields,
    kind: z.literal("emphasis"),
    preset: z.literal("breathe"),
    fromColor: motionHexColorSchema,
    toColor: motionHexColorSchema,
  })
  .readonly();
export type BreatheEffect = z.infer<typeof breatheEffectSchema>;

export const effectSchema = z.discriminatedUnion("kind", [
  entranceEffectSchema,
  emphasisEffectSchema,
]);
export type Effect = z.infer<typeof effectSchema>;

export const repeatedEffectSchema = z
  .strictObject({
    type: z.literal("repeated-effect"),
    effect: effectSchema,
    count: repeatCountSchema,
  })
  .readonly();

export type RepeatedEffect<E extends Effect = Effect> = {
  readonly type: "repeated-effect";
  readonly effect: E;
  readonly count: RepeatCount;
};

export const entranceEffectDeclarationSchema = z.union([
  entranceEffectSchema,
  z
    .strictObject({
      type: z.literal("repeated-effect"),
      effect: entranceEffectSchema,
      count: repeatCountSchema,
    })
    .readonly(),
]);
export type EntranceEffectDeclaration =
  | EntranceEffect
  | RepeatedEffect<EntranceEffect>;

export const emphasisEffectDeclarationSchema = z.union([
  emphasisEffectSchema,
  z
    .strictObject({
      type: z.literal("repeated-effect"),
      effect: emphasisEffectSchema,
      count: repeatCountSchema,
    })
    .readonly(),
]);
export type EmphasisEffectDeclaration =
  | EmphasisEffect
  | RepeatedEffect<EmphasisEffect>;

export const normalizedEffectSchema = z.discriminatedUnion("kind", [
  entranceEffectObjectSchema
    .extend({ repeatCount: repeatCountSchema.optional() })
    .readonly(),
  emphasisEffectObjectSchema
    .extend({ repeatCount: repeatCountSchema.optional() })
    .readonly(),
]);
export type NormalizedEffect = z.infer<typeof normalizedEffectSchema>;

export const motionScopeModeSchema = z.enum([
  "parallel",
  "sequence",
  "stagger",
]);
export type MotionScopeMode = z.infer<typeof motionScopeModeSchema>;

export const motionCompletionStateSchema = z.enum([
  "known",
  "unknown",
  "no-natural-completion",
]);
export type MotionCompletionState = z.infer<typeof motionCompletionStateSchema>;

export const motionMomentSchema = z
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
export type MotionMoment = z.infer<typeof motionMomentSchema>;

export const motionMomentOffsetSchema = z
  .strictObject({
    type: z.literal("motion-offset"),
    moment: motionMomentSchema,
    offsetMs: z.number().int().min(0),
  })
  .readonly();
export type MotionMomentOffset = z.infer<typeof motionMomentOffsetSchema>;

export const motionTimingSchema = z.union([
  motionMomentSchema,
  motionMomentOffsetSchema,
]);
export type MotionTiming = z.infer<typeof motionTimingSchema>;

export type MotionScopeHandle = {
  readonly type: "motion-scope-handle";
  readonly identity: string;
  readonly start: MotionMoment;
  readonly end: MotionMoment;
  cue(name: string): MotionMoment;
};

/** Runtime-identity-preserving schema for handles created by `createMotionScopeHandle()` / `useMotionScope()`. */
export const motionScopeHandleSchema = z.custom<MotionScopeHandle>(
  (value) =>
    typeof value === "object" &&
    value !== null &&
    !Array.isArray(value) &&
    (value as { readonly type?: unknown }).type === "motion-scope-handle" &&
    typeof (value as { readonly identity?: unknown }).identity === "string" &&
    (value as { readonly identity: string }).identity !== "" &&
    motionMomentSchema.safeParse((value as { readonly start?: unknown }).start)
      .success &&
    motionMomentSchema.safeParse((value as { readonly end?: unknown }).end)
      .success &&
    typeof (value as { readonly cue?: unknown }).cue === "function",
  "MotionScope requires a handle created by createMotionScopeHandle() or useMotionScope()",
);

export const motionPropsSchema = z
  .strictObject({
    enter: z
      .union([
        entranceEffectDeclarationSchema,
        z.array(entranceEffectDeclarationSchema).readonly(),
      ])
      .optional(),
    emphasis: emphasisEffectDeclarationSchema.optional(),
    motion: z.boolean().optional(),
  })
  .readonly();
export type MotionProps = z.infer<typeof motionPropsSchema>;

export const motionScopeAuthorPropsSchema = z
  .strictObject({
    scope: motionScopeHandleSchema,
    mode: motionScopeModeSchema.optional(),
    each: z.number().int().min(0).optional(),
    delay: effectDelaySchema.optional(),
    startAt: motionTimingSchema.optional(),
    publishStart: motionMomentSchema.optional(),
    publishEnd: motionMomentSchema.optional(),
    motion: z.boolean().optional(),
  })
  .readonly();
export type MotionScopeAuthorProps = z.infer<
  typeof motionScopeAuthorPropsSchema
>;

export const motionStepAuthorPropsSchema = z
  .strictObject({
    n: z.number().int().min(1).optional(),
  })
  .readonly();
export type MotionStepAuthorProps = z.infer<typeof motionStepAuthorPropsSchema>;

/** Structured target and step timing declaration retained directly on `SlideMaster` and `SlideLayout` owners. */
export const ownerMotionTargetSchema = z.union([
  z.strictObject({ placeholder: z.string().min(1) }).readonly(),
  z.strictObject({ spid: z.number().int().positive() }).readonly(),
]);
export type OwnerMotionTarget = z.infer<typeof ownerMotionTargetSchema>;

export const ownerMotionEffectSchema = z
  .strictObject({
    target: ownerMotionTargetSchema,
    effect: z.union([
      entranceEffectDeclarationSchema,
      emphasisEffectDeclarationSchema,
    ]),
    range: motionTargetRangeSchema.optional(),
  })
  .readonly();
export type OwnerMotionEffect = z.infer<typeof ownerMotionEffectSchema>;

export const ownerMotionStepSchema = z
  .strictObject({
    trigger: z.enum(["auto", "click"]),
    effects: z.array(ownerMotionEffectSchema).readonly(),
  })
  .readonly();
export type OwnerMotionStep = z.infer<typeof ownerMotionStepSchema>;

export const ownerMotionSchema = z
  .strictObject({
    steps: z.array(ownerMotionStepSchema).readonly(),
  })
  .readonly();
export type OwnerMotion = z.infer<typeof ownerMotionSchema>;
