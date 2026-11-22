import { z } from "zod";

/** Transition duration in milliseconds; rounded to the nearest integer millisecond when written to p14:dur. */
export const TransitionDurationSchema = z
  .number()
  .finite()
  .min(0)
  .max(Number.MAX_SAFE_INTEGER);
export type TransitionDuration = z.infer<typeof TransitionDurationSchema>;

export const AdvanceAfterSchema = z.number().finite().min(0).max(4_294_967_295);

export const TransitionDirectionSchema = z.enum([
  "left",
  "right",
  "top",
  "bottom",
]);
export type TransitionDirection = z.infer<typeof TransitionDirectionSchema>;

export const StripsDirectionSchema = z.enum([
  "leftUp",
  "rightUp",
  "leftDown",
  "rightDown",
]);
export type StripsDirection = z.infer<typeof StripsDirectionSchema>;

export const TransitionSpeedSchema = z.enum(["fast", "med", "slow"]);
export type TransitionSpeed = z.infer<typeof TransitionSpeedSchema>;

export const FlipDirectionSchema = z.enum(["left", "right"]);
export const MorphBySchema = z.enum(["object", "word", "character"]);

export const FadeTransitionSchema = z
  .strictObject({
    type: z.literal("transition"),
    kind: z.literal("fade"),
    duration: TransitionDurationSchema.optional(),
  })
  .readonly();

export const StripsTransitionSchema = z
  .strictObject({
    type: z.literal("transition"),
    kind: z.literal("strips"),
    duration: TransitionDurationSchema.optional(),
    direction: StripsDirectionSchema,
  })
  .readonly();

export const PushTransitionSchema = z
  .strictObject({
    type: z.literal("transition"),
    kind: z.literal("push"),
    duration: TransitionDurationSchema.optional(),
    from: TransitionDirectionSchema,
  })
  .readonly();

/** Transitional baseline transition that does not depend on extension namespaces. */
export const StandardTransitionSchema = z.discriminatedUnion("kind", [
  FadeTransitionSchema,
  StripsTransitionSchema,
  PushTransitionSchema,
]);
export type StandardTransition = z.infer<typeof StandardTransitionSchema>;

/** p14 Flip transition; must provide a baseline format fallback for other consumers. */
export const FlipTransitionSchema = z
  .strictObject({
    type: z.literal("transition"),
    kind: z.literal("flip"),
    duration: TransitionDurationSchema.optional(),
    direction: FlipDirectionSchema,
    fallback: StandardTransitionSchema,
  })
  .readonly();
export type FlipTransition = z.infer<typeof FlipTransitionSchema>;

/** p159 Morph transition; must provide a baseline format fallback for other consumers. */
export const MorphTransitionSchema = z
  .strictObject({
    type: z.literal("transition"),
    kind: z.literal("morph"),
    duration: TransitionDurationSchema.optional(),
    by: MorphBySchema,
    fallback: StandardTransitionSchema,
  })
  .readonly();
export type MorphTransition = z.infer<typeof MorphTransitionSchema>;

/** Slide, layout, or master transition declaration; build fails when target forbids required mechanisms. */
export const TransitionSchema = z.discriminatedUnion("kind", [
  FadeTransitionSchema,
  StripsTransitionSchema,
  PushTransitionSchema,
  FlipTransitionSchema,
  MorphTransitionSchema,
]);
export type Transition = z.infer<typeof TransitionSchema>;

type TransitionOptions = { readonly duration?: TransitionDuration };

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
  fade: (options: TransitionOptions = {}): StandardTransition =>
    declaration("fade", options),
  /** Uses baseline strips effect toward the specified diagonal direction. */
  strips: (
    options: TransitionOptions & { readonly direction?: StripsDirection } = {},
  ): StandardTransition =>
    declaration("strips", {
      ...options,
      direction:
        options.direction === undefined ? "rightDown" : options.direction,
    }),
  /** Uses baseline push effect entering from the slide edge. */
  push: (
    options: TransitionOptions & { readonly from?: TransitionDirection } = {},
  ): StandardTransition =>
    declaration("push", {
      ...options,
      from: options.from === undefined ? "bottom" : options.from,
    }),
  /** Uses p14 Flip with an explicit baseline fallback for targets that exclude extensions. */
  flip: (
    options: TransitionOptions & {
      readonly direction?: "left" | "right";
      readonly fallback: StandardTransition;
    },
  ): FlipTransition =>
    declaration("flip", {
      ...options,
      direction: options.direction === undefined ? "right" : options.direction,
    }),
  /** Uses p159 Morph with an explicit baseline fallback for targets that exclude extensions. */
  morph: (
    options: TransitionOptions & {
      readonly by?: "object" | "word" | "character";
      readonly fallback: StandardTransition;
    },
  ): MorphTransition =>
    declaration("morph", {
      ...options,
      by: options.by === undefined ? "object" : options.by,
    }),
});
