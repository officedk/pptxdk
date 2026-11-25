import { z } from "zod";

/** Transition duration in milliseconds; rounded to the nearest integer millisecond when written to p14:dur. */
export const transitionDurationSchema = z
  .number()
  .finite()
  .min(0)
  .max(Number.MAX_SAFE_INTEGER);
export type TransitionDuration = z.infer<typeof transitionDurationSchema>;

export const advanceAfterSchema = z.number().finite().min(0).max(4_294_967_295);

export const transitionDirectionSchema = z.enum([
  "left",
  "right",
  "top",
  "bottom",
]);
export type TransitionDirection = z.infer<typeof transitionDirectionSchema>;

export const stripsDirectionSchema = z.enum([
  "leftUp",
  "rightUp",
  "leftDown",
  "rightDown",
]);
export type StripsDirection = z.infer<typeof stripsDirectionSchema>;

export const transitionSpeedSchema = z.enum(["fast", "med", "slow"]);
export type TransitionSpeed = z.infer<typeof transitionSpeedSchema>;

export const flipDirectionSchema = z.enum(["left", "right"]);
export const morphBySchema = z.enum(["object", "word", "character"]);

export const transitionKindSchema = z.enum([
  "fade",
  "strips",
  "push",
  "flip",
  "morph",
]);
export type TransitionKind = z.infer<typeof transitionKindSchema>;

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
  direction: stripsDirectionSchema,
});
export const stripsTransitionSchema = stripsTransitionObjectSchema.readonly();

const pushTransitionObjectSchema = z.strictObject({
  type: z.literal("transition"),
  kind: z.literal("push"),
  duration: transitionDurationSchema.optional(),
  from: transitionDirectionSchema,
});
export const pushTransitionSchema = pushTransitionObjectSchema.readonly();

/** Transitional baseline transition that does not depend on extension namespaces. */
export const standardTransitionSchema = z.discriminatedUnion("kind", [
  fadeTransitionSchema,
  stripsTransitionSchema,
  pushTransitionSchema,
]);
export type StandardTransition = z.infer<typeof standardTransitionSchema>;

const flipTransitionObjectSchema = z.strictObject({
  type: z.literal("transition"),
  kind: z.literal("flip"),
  duration: transitionDurationSchema.optional(),
  direction: flipDirectionSchema,
  fallback: standardTransitionSchema,
});
/** p14 Flip transition; must provide a baseline format fallback for other consumers. */
export const flipTransitionSchema = flipTransitionObjectSchema.readonly();
export type FlipTransition = z.infer<typeof flipTransitionSchema>;

const morphTransitionObjectSchema = z.strictObject({
  type: z.literal("transition"),
  kind: z.literal("morph"),
  duration: transitionDurationSchema.optional(),
  by: morphBySchema,
  fallback: standardTransitionSchema,
});
/** p159 Morph transition; must provide a baseline format fallback for other consumers. */
export const morphTransitionSchema = morphTransitionObjectSchema.readonly();
export type MorphTransition = z.infer<typeof morphTransitionSchema>;

/** Slide, layout, or master transition declaration; build fails when target forbids required mechanisms. */
export const transitionSchema = z.discriminatedUnion("kind", [
  fadeTransitionSchema,
  stripsTransitionSchema,
  pushTransitionSchema,
  flipTransitionSchema,
  morphTransitionSchema,
]);
export type Transition = z.infer<typeof transitionSchema>;

export const fadeTransitionOptionsSchema = fadeTransitionObjectSchema
  .pick({ duration: true })
  .readonly();
export type FadeTransitionOptions = z.infer<typeof fadeTransitionOptionsSchema>;

export const stripsTransitionOptionsSchema = stripsTransitionObjectSchema
  .pick({ duration: true, direction: true })
  .partial({ direction: true })
  .readonly();
export type StripsTransitionOptions = z.infer<
  typeof stripsTransitionOptionsSchema
>;

export const pushTransitionOptionsSchema = pushTransitionObjectSchema
  .pick({ duration: true, from: true })
  .partial({ from: true })
  .readonly();
export type PushTransitionOptions = z.infer<typeof pushTransitionOptionsSchema>;

export const flipTransitionOptionsSchema = flipTransitionObjectSchema
  .pick({ duration: true, direction: true, fallback: true })
  .partial({ direction: true })
  .readonly();
export type FlipTransitionOptions = z.infer<typeof flipTransitionOptionsSchema>;

export const morphTransitionOptionsSchema = morphTransitionObjectSchema
  .pick({ duration: true, by: true, fallback: true })
  .partial({ by: true })
  .readonly();
export type MorphTransitionOptions = z.infer<
  typeof morphTransitionOptionsSchema
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
