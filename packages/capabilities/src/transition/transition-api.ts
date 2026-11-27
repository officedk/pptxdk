import { z } from "zod";

type AssertExact<A, B> =
  (<T>() => T extends A ? 1 : 2) extends <T>() => T extends B ? 1 : 2
    ? true
    : never;

export type TransitionDuration = number;
/** Transition duration in milliseconds; rounded to the nearest integer millisecond when written to p14:dur. */
export const transitionDurationSchema: z.ZodType<TransitionDuration> = z
  .number()
  .finite()
  .min(0)
  .max(Number.MAX_SAFE_INTEGER);

export const advanceAfterSchema = z.number().finite().min(0).max(4_294_967_295);

export type TransitionDirection = "left" | "right" | "top" | "bottom";
const rawTransitionDirectionSchema = z.enum(["left", "right", "top", "bottom"]);
export const transitionDirectionSchema: z.ZodType<TransitionDirection> =
  rawTransitionDirectionSchema satisfies z.ZodType<
    AssertExact<
      TransitionDirection,
      z.infer<typeof rawTransitionDirectionSchema>
    > extends true
      ? TransitionDirection
      : never
  >;

export type StripsDirection = "leftUp" | "rightUp" | "leftDown" | "rightDown";
const rawStripsDirectionSchema = z.enum([
  "leftUp",
  "rightUp",
  "leftDown",
  "rightDown",
]);
export const stripsDirectionSchema: z.ZodType<StripsDirection> =
  rawStripsDirectionSchema satisfies z.ZodType<
    AssertExact<
      StripsDirection,
      z.infer<typeof rawStripsDirectionSchema>
    > extends true
      ? StripsDirection
      : never
  >;

export type TransitionSpeed = "fast" | "med" | "slow";
export const transitionSpeedSchema: z.ZodType<TransitionSpeed> = z.enum([
  "fast",
  "med",
  "slow",
]);

export const flipDirectionSchema = z.enum(["left", "right"]);
export const morphBySchema = z.enum(["object", "word", "character"]);

export type TransitionKind = "fade" | "strips" | "push" | "flip" | "morph";
const rawTransitionKindSchema = z.enum([
  "fade",
  "strips",
  "push",
  "flip",
  "morph",
]);
export const transitionKindSchema: z.ZodType<TransitionKind> =
  rawTransitionKindSchema satisfies z.ZodType<
    AssertExact<
      TransitionKind,
      z.infer<typeof rawTransitionKindSchema>
    > extends true
      ? TransitionKind
      : never
  >;

const fadeTransitionObjectSchema = z.strictObject({
  type: z.literal("transition"),
  kind: z.literal("fade"),
  duration: transitionDurationSchema.optional(),
});
export const fadeTransitionSchema = fadeTransitionObjectSchema.readonly();

const stripsTransitionObjectSchema = z.strictObject({
  type: z.literal("transition"),
  kind: z.literal("strips"),
  duration: transitionDurationSchema.optional(),
  direction: rawStripsDirectionSchema,
});
export const stripsTransitionSchema = stripsTransitionObjectSchema.readonly();

const pushTransitionObjectSchema = z.strictObject({
  type: z.literal("transition"),
  kind: z.literal("push"),
  duration: transitionDurationSchema.optional(),
  from: rawTransitionDirectionSchema,
});
export const pushTransitionSchema = pushTransitionObjectSchema.readonly();

export type StandardTransition =
  | {
      readonly type: "transition";
      readonly kind: "fade";
      readonly duration?: TransitionDuration | undefined;
    }
  | {
      readonly type: "transition";
      readonly kind: "strips";
      readonly duration?: TransitionDuration | undefined;
      readonly direction: StripsDirection;
    }
  | {
      readonly type: "transition";
      readonly kind: "push";
      readonly duration?: TransitionDuration | undefined;
      readonly from: TransitionDirection;
    };

const rawStandardTransitionSchema = z.discriminatedUnion("kind", [
  fadeTransitionSchema,
  stripsTransitionSchema,
  pushTransitionSchema,
]);
/** Transitional baseline transition that does not depend on extension namespaces. */
export const standardTransitionSchema: z.ZodType<StandardTransition> =
  rawStandardTransitionSchema satisfies z.ZodType<
    AssertExact<
      StandardTransition,
      z.infer<typeof rawStandardTransitionSchema>
    > extends true
      ? StandardTransition
      : never
  >;

/** p14 Flip transition; must provide a baseline format fallback for other consumers. */
export type FlipTransition = {
  readonly type: "transition";
  readonly kind: "flip";
  readonly duration?: TransitionDuration | undefined;
  readonly direction: "left" | "right";
  readonly fallback: StandardTransition;
};

const flipTransitionObjectSchema = z.strictObject({
  type: z.literal("transition"),
  kind: z.literal("flip"),
  duration: transitionDurationSchema.optional(),
  direction: flipDirectionSchema,
  fallback: rawStandardTransitionSchema,
});
const rawFlipTransitionSchema = flipTransitionObjectSchema.readonly();
/** p14 Flip transition; must provide a baseline format fallback for other consumers. */
export const flipTransitionSchema: z.ZodType<FlipTransition> =
  rawFlipTransitionSchema satisfies z.ZodType<
    AssertExact<
      FlipTransition,
      z.infer<typeof rawFlipTransitionSchema>
    > extends true
      ? FlipTransition
      : never
  >;

/** p159 Morph transition; must provide a baseline format fallback for other consumers. */
export type MorphTransition = {
  readonly type: "transition";
  readonly kind: "morph";
  readonly duration?: TransitionDuration | undefined;
  readonly by: "object" | "word" | "character";
  readonly fallback: StandardTransition;
};

const morphTransitionObjectSchema = z.strictObject({
  type: z.literal("transition"),
  kind: z.literal("morph"),
  duration: transitionDurationSchema.optional(),
  by: morphBySchema,
  fallback: rawStandardTransitionSchema,
});
const rawMorphTransitionSchema = morphTransitionObjectSchema.readonly();
/** p159 Morph transition; must provide a baseline format fallback for other consumers. */
export const morphTransitionSchema: z.ZodType<MorphTransition> =
  rawMorphTransitionSchema satisfies z.ZodType<
    AssertExact<
      MorphTransition,
      z.infer<typeof rawMorphTransitionSchema>
    > extends true
      ? MorphTransition
      : never
  >;

export type Transition = StandardTransition | FlipTransition | MorphTransition;

const rawTransitionSchema = z.discriminatedUnion("kind", [
  fadeTransitionSchema,
  stripsTransitionSchema,
  pushTransitionSchema,
  rawFlipTransitionSchema,
  rawMorphTransitionSchema,
]);
/** Slide, layout, or master transition declaration; build fails when target forbids required mechanisms. */
export const transitionSchema: z.ZodType<Transition> =
  rawTransitionSchema satisfies z.ZodType<
    AssertExact<Transition, z.infer<typeof rawTransitionSchema>> extends true
      ? Transition
      : never
  >;

export type FadeTransitionOptions = {
  readonly duration?: TransitionDuration | undefined;
};
const rawFadeTransitionOptionsSchema = fadeTransitionObjectSchema
  .pick({ duration: true })
  .readonly();
export const fadeTransitionOptionsSchema: z.ZodType<FadeTransitionOptions> =
  rawFadeTransitionOptionsSchema satisfies z.ZodType<
    AssertExact<
      FadeTransitionOptions,
      z.infer<typeof rawFadeTransitionOptionsSchema>
    > extends true
      ? FadeTransitionOptions
      : never
  >;

export type StripsTransitionOptions = {
  readonly duration?: TransitionDuration | undefined;
  readonly direction?: StripsDirection | undefined;
};
const rawStripsTransitionOptionsSchema = stripsTransitionObjectSchema
  .pick({ duration: true, direction: true })
  .partial({ direction: true })
  .readonly();
export const stripsTransitionOptionsSchema: z.ZodType<StripsTransitionOptions> =
  rawStripsTransitionOptionsSchema satisfies z.ZodType<
    AssertExact<
      StripsTransitionOptions,
      z.infer<typeof rawStripsTransitionOptionsSchema>
    > extends true
      ? StripsTransitionOptions
      : never
  >;

export type PushTransitionOptions = {
  readonly duration?: TransitionDuration | undefined;
  readonly from?: TransitionDirection | undefined;
};
const rawPushTransitionOptionsSchema = pushTransitionObjectSchema
  .pick({ duration: true, from: true })
  .partial({ from: true })
  .readonly();
export const pushTransitionOptionsSchema: z.ZodType<PushTransitionOptions> =
  rawPushTransitionOptionsSchema satisfies z.ZodType<
    AssertExact<
      PushTransitionOptions,
      z.infer<typeof rawPushTransitionOptionsSchema>
    > extends true
      ? PushTransitionOptions
      : never
  >;

export type FlipTransitionOptions = {
  readonly duration?: TransitionDuration | undefined;
  readonly direction?: "left" | "right" | undefined;
  readonly fallback: StandardTransition;
};
const rawFlipTransitionOptionsSchema = flipTransitionObjectSchema
  .pick({ duration: true, direction: true, fallback: true })
  .partial({ direction: true })
  .readonly();
export const flipTransitionOptionsSchema: z.ZodType<FlipTransitionOptions> =
  rawFlipTransitionOptionsSchema satisfies z.ZodType<
    AssertExact<
      FlipTransitionOptions,
      z.infer<typeof rawFlipTransitionOptionsSchema>
    > extends true
      ? FlipTransitionOptions
      : never
  >;

export type MorphTransitionOptions = {
  readonly duration?: TransitionDuration | undefined;
  readonly by?: "object" | "word" | "character" | undefined;
  readonly fallback: StandardTransition;
};
const rawMorphTransitionOptionsSchema = morphTransitionObjectSchema
  .pick({ duration: true, by: true, fallback: true })
  .partial({ by: true })
  .readonly();
export const morphTransitionOptionsSchema: z.ZodType<MorphTransitionOptions> =
  rawMorphTransitionOptionsSchema satisfies z.ZodType<
    AssertExact<
      MorphTransitionOptions,
      z.infer<typeof rawMorphTransitionOptionsSchema>
    > extends true
      ? MorphTransitionOptions
      : never
  >;

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

/** Creates baseline and extended transition values used for slide, layout, and master transition declarations. */
export const transitions = Object.freeze({
  /** Uses baseline fade effect; defaults to slow speed. */
  fade: (options: FadeTransitionOptions = {}): StandardTransition =>
    declaration("fade", options),
  /** Uses baseline strips effect toward the specified diagonal direction. */
  strips: (options: StripsTransitionOptions = {}): StandardTransition =>
    declaration("strips", {
      ...options,
      direction:
        options.direction === undefined ? "rightDown" : options.direction,
    }),
  /** Uses baseline push effect entering from the slide edge. */
  push: (options: PushTransitionOptions = {}): StandardTransition =>
    declaration("push", {
      ...options,
      from: options.from === undefined ? "bottom" : options.from,
    }),
  /** Uses p14 Flip with an explicit baseline fallback for targets that exclude extensions. */
  flip: (options: FlipTransitionOptions): FlipTransition =>
    declaration("flip", {
      ...options,
      direction: options.direction === undefined ? "right" : options.direction,
    }),
  /** Uses p159 Morph with an explicit baseline fallback for targets that exclude extensions. */
  morph: (options: MorphTransitionOptions): MorphTransition =>
    declaration("morph", {
      ...options,
      by: options.by === undefined ? "object" : options.by,
    }),
});
