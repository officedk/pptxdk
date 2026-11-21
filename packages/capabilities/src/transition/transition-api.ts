/** Transition duration in milliseconds; rounded to the nearest integer millisecond when written to p14:dur. */
export type TransitionDuration = number;
export type TransitionDirection = "left" | "right" | "top" | "bottom";
export type StripsDirection = "leftUp" | "rightUp" | "leftDown" | "rightDown";
export type TransitionSpeed = "fast" | "med" | "slow";

type TransitionOptions = { readonly duration?: TransitionDuration };

/** Transitional baseline transition that does not depend on extension namespaces. */
export type StandardTransition =
  | {
      readonly type: "transition";
      readonly kind: "fade";
      readonly duration?: TransitionDuration;
    }
  | {
      readonly type: "transition";
      readonly kind: "strips";
      readonly duration?: TransitionDuration;
      readonly direction: StripsDirection;
    }
  | {
      readonly type: "transition";
      readonly kind: "push";
      readonly duration?: TransitionDuration;
      readonly from: TransitionDirection;
    };

/** p14 Flip transition; must provide a baseline format fallback for other consumers. */
export type FlipTransition = {
  readonly type: "transition";
  readonly kind: "flip";
  readonly duration?: TransitionDuration;
  readonly direction: "left" | "right";
  readonly fallback: StandardTransition;
};

/** p159 Morph transition; must provide a baseline format fallback for other consumers. */
export type MorphTransition = {
  readonly type: "transition";
  readonly kind: "morph";
  readonly duration?: TransitionDuration;
  readonly by: "object" | "word" | "character";
  readonly fallback: StandardTransition;
};

/** Slide transition declaration; build fails when target does not permit required mechanisms. */
export type Transition = StandardTransition | FlipTransition | MorphTransition;

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

/** Creates baseline and extended transition values used for slide transition declarations. */
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
