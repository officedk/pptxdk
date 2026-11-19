import {
  CapabilityError,
  type CapabilityInput,
  type Extension,
  type ResolvedBoundSlideCapability,
  type ResolvedSlide,
  type SlideCapabilityModule,
  type SlideEmitOutput,
} from "@react-pptx/core/capability";

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

type TransitionRecord = Record<string, unknown> & {
  readonly type: unknown;
  readonly kind: unknown;
};

const stripDirections: Record<StripsDirection, string> = {
  leftUp: "lu",
  rightUp: "ru",
  leftDown: "ld",
  rightDown: "rd",
};
const pushDirections: Record<TransitionDirection, string> = {
  left: "l",
  right: "r",
  top: "u",
  bottom: "d",
};
const p14Namespace = "http://schemas.microsoft.com/office/powerpoint/2010/main";
const p159Namespace =
  "http://schemas.microsoft.com/office/powerpoint/2015/09/main";
const markupCompatibilityNamespace =
  "http://schemas.openxmlformats.org/markup-compatibility/2006";
const morphOptions = {
  object: "byObject",
  word: "byWord",
  character: "byChar",
} as const;

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
  /** Uses p14 Flip; must explicitly provide a baseline format fallback. */
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
  /** Uses p159 Morph; must explicitly provide a baseline format fallback. */
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

function fail(code: string, message: string): never {
  throw new CapabilityError(code, message);
}

function asRecord(value: unknown): TransitionRecord {
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    return fail(
      "invalid-transition",
      "Slide.transition must be a declaration value returned by transition helpers",
    );
  }
  return value as TransitionRecord;
}

function unknownParameter(
  value: TransitionRecord,
  allowed: readonly string[],
): string | undefined {
  return Object.keys(value).find((key) => !allowed.includes(key));
}

function durationOf(value: TransitionRecord): number | undefined {
  const duration = value["duration"];
  if (duration === undefined) return undefined;
  if (
    typeof duration !== "number" ||
    !Number.isFinite(duration) ||
    duration < 0 ||
    duration > Number.MAX_SAFE_INTEGER
  ) {
    return fail(
      "invalid-transition-duration",
      "Transition duration must be a finite millisecond value between 0 and Number.MAX_SAFE_INTEGER",
    );
  }
  return duration;
}

function speed(duration: number | undefined): TransitionSpeed {
  if (duration === undefined || duration > 875) return "slow";
  return duration > 600 ? "med" : "fast";
}

function transitionElement(
  value: TransitionRecord,
  duration: number | undefined,
  useDurationExtension: boolean,
  advanceAfter?: number,
): string {
  const kind = value["kind"];
  const transitionSpeed = speed(duration);
  const speedAttribute = ` spd="${transitionSpeed}"`;
  const durationAttribute =
    useDurationExtension && duration !== undefined
      ? ` p14:dur="${Math.round(duration)}"`
      : "";
  const advanceAttribute =
    advanceAfter === undefined ? "" : ` advTm="${Math.round(advanceAfter)}"`;
  if (kind === "fade")
    return `<p:transition${speedAttribute}${durationAttribute}${advanceAttribute}><p:fade/></p:transition>`;
  if (kind === "strips") {
    const direction = value["direction"];
    if (
      typeof direction !== "string" ||
      !Object.hasOwn(stripDirections, direction)
    ) {
      return fail(
        "invalid-transition-direction",
        "strips.direction must be one of the four diagonal directions",
      );
    }
    return `<p:transition${speedAttribute}${durationAttribute}${advanceAttribute}><p:strips dir="${stripDirections[direction as StripsDirection]}"/></p:transition>`;
  }
  if (kind === "push") {
    const from = value["from"];
    if (typeof from !== "string" || !Object.hasOwn(pushDirections, from)) {
      return fail(
        "invalid-transition-direction",
        "push.from must be left, right, top, or bottom",
      );
    }
    return `<p:transition${speedAttribute}${durationAttribute}${advanceAttribute}><p:push dir="${pushDirections[from as TransitionDirection]}"/></p:transition>`;
  }
  if (kind === "flip") {
    const direction = value["direction"];
    if (direction !== "left" && direction !== "right") {
      return fail(
        "invalid-transition-direction",
        "flip.direction must be left or right",
      );
    }
    return `<p:transition${speedAttribute}${durationAttribute}${advanceAttribute}><p14:flip dir="${direction === "left" ? "l" : "r"}"/></p:transition>`;
  }
  if (kind === "morph") {
    const by = value["by"];
    if (by !== "object" && by !== "word" && by !== "character") {
      return fail(
        "invalid-morph-option",
        "morph.by must be object, word, or character",
      );
    }
    return `<p:transition${speedAttribute}${durationAttribute}${advanceAttribute}><p159:morph option="${morphOptions[by]}"/></p:transition>`;
  }
  return fail(
    "unsupported-transition",
    `Unsupported transition kind ${String(kind)}`,
  );
}

function standardTransition(
  value: TransitionRecord,
  inheritedDuration: number | undefined,
  advanceAfter?: number,
): {
  readonly xml: string;
  readonly observation: Readonly<Record<string, unknown>>;
} {
  const kind = value["kind"];
  if (kind !== "fade" && kind !== "strips" && kind !== "push") {
    return fail(
      "invalid-transition-fallback",
      "fallback must be fade, strips, or push",
    );
  }
  const allowed =
    kind === "fade"
      ? ["type", "kind", "duration"]
      : kind === "strips"
        ? ["type", "kind", "duration", "direction"]
        : ["type", "kind", "duration", "from"];
  const unknown = unknownParameter(value, allowed);
  if (unknown !== undefined) {
    return fail(
      "unknown-parameter",
      `Transition fallback contains unknown parameter ${unknown}`,
    );
  }
  const explicitDuration = durationOf(value);
  const effectiveDuration = explicitDuration ?? inheritedDuration;
  return {
    xml: transitionElement(value, effectiveDuration, false, advanceAfter),
    observation: Object.freeze({
      kind,
      ...(kind === "strips" ? { direction: value["direction"] } : {}),
      ...(kind === "push" ? { from: value["from"] } : {}),
      speed: speed(effectiveDuration),
    }),
  };
}

function compileTransition(
  value: unknown,
  advanceAfter?: number,
): {
  readonly xml: string;
  readonly extensions: readonly Extension[];
  readonly matchSourceKinds: boolean;
  readonly observation: Readonly<Record<string, unknown>>;
} {
  const transition = asRecord(value);
  if (transition["type"] !== "transition") {
    return fail(
      "invalid-transition",
      "Slide.transition is missing a valid transition declaration tag",
    );
  }
  const kind = transition["kind"];
  if (
    kind !== "fade" &&
    kind !== "strips" &&
    kind !== "push" &&
    kind !== "flip" &&
    kind !== "morph"
  ) {
    return fail(
      "unsupported-transition",
      `Unsupported transition kind ${String(kind)}`,
    );
  }
  const allowed =
    kind === "fade"
      ? ["type", "kind", "duration"]
      : kind === "strips"
        ? ["type", "kind", "duration", "direction"]
        : kind === "push"
          ? ["type", "kind", "duration", "from"]
          : kind === "flip"
            ? ["type", "kind", "duration", "direction", "fallback"]
            : ["type", "kind", "duration", "by", "fallback"];
  const unknown = unknownParameter(transition, allowed);
  if (unknown !== undefined) {
    return fail(
      "unknown-parameter",
      `Transition declaration contains unknown parameter ${unknown}`,
    );
  }
  const duration = durationOf(transition);
  const requiresFallback = kind === "flip" || kind === "morph";
  let fallback: ReturnType<typeof standardTransition> | undefined;
  if (requiresFallback) {
    if (!Object.hasOwn(transition, "fallback")) {
      return fail(
        "missing-transition-fallback",
        `${kind} transition must explicitly declare a Transitional fallback`,
      );
    }
    const fallbackValue = asRecord(transition["fallback"]);
    if (fallbackValue["type"] !== "transition") {
      return fail(
        "invalid-transition-fallback",
        "Invalid transition fallback declaration",
      );
    }
    fallback = standardTransition(fallbackValue, duration, advanceAfter);
  } else if (duration !== undefined) {
    fallback = standardTransition(transition, duration, advanceAfter);
  }
  const effectExtensions: Extension[] =
    kind === "flip" ? ["p14"] : kind === "morph" ? ["p159"] : [];
  const extensions = [
    ...new Set<Extension>([
      ...(duration === undefined ? [] : ["p14" as const]),
      ...effectExtensions,
    ]),
  ];
  const standardXml = transitionElement(
    transition,
    duration,
    false,
    advanceAfter,
  );
  const observation: Record<string, unknown> = {
    kind,
    ...(kind === "strips" ? { direction: transition["direction"] } : {}),
    ...(kind === "push" ? { from: transition["from"] } : {}),
    ...(kind === "flip" ? { direction: transition["direction"] } : {}),
    ...(kind === "morph" ? { by: transition["by"] } : {}),
    speed: speed(duration),
    ...(duration === undefined ? {} : { durationMs: Math.round(duration) }),
    ...(fallback === undefined ? {} : { fallback: fallback.observation }),
    ...(advanceAfter === undefined
      ? {}
      : { advanceAfterMs: Math.round(advanceAfter) }),
  };
  if (extensions.length === 0) {
    return {
      xml: standardXml,
      extensions,
      matchSourceKinds: kind === "morph",
      observation: Object.freeze(observation),
    };
  }
  const exactXml = transitionElement(transition, duration, true, advanceAfter);
  const fallbackXml = fallback?.xml ?? standardXml;
  const requires = extensions.join(" ");
  const declarations = [
    ...(extensions.includes("p14") ? [`xmlns:p14="${p14Namespace}"`] : []),
    ...(extensions.includes("p159") ? [`xmlns:p159="${p159Namespace}"`] : []),
  ].join(" ");
  return {
    xml: `<mc:AlternateContent xmlns:mc="${markupCompatibilityNamespace}" ${declarations}><mc:Choice Requires="${requires}">${exactXml}</mc:Choice><mc:Fallback>${fallbackXml}</mc:Fallback></mc:AlternateContent>`,
    extensions,
    matchSourceKinds: kind === "morph",
    observation: Object.freeze(observation),
  };
}

/**
 * Bound transition module: validates transition & advanceAfter once per slide during bind,
 * submits generic continuity matching collaboration for Morph, and leaves static preview untouched.
 */
export const transitionCapability: SlideCapabilityModule<ResolvedBoundSlideCapability> =
  {
    declaration: "transition",
    slideProp: "transition",
    slideProps: ["transition", "advanceAfter"],
    slot: "transition",
    bindSlide(input: CapabilityInput) {
      const advanceAfter = input.props["advanceAfter"];
      if (
        advanceAfter !== undefined &&
        (typeof advanceAfter !== "number" ||
          !Number.isFinite(advanceAfter) ||
          advanceAfter < 0 ||
          advanceAfter > 4_294_967_295)
      ) {
        return fail(
          "invalid-advance-after",
          "Slide.advanceAfter must be a finite millisecond value between 0 and 4294967295",
        );
      }
      const transitionValue = input.props["transition"];
      const compiled =
        transitionValue === undefined
          ? {
              xml:
                advanceAfter === undefined
                  ? ""
                  : `<p:transition advTm="${Math.round(advanceAfter)}"/>`,
              extensions: [] as readonly Extension[],
              matchSourceKinds: false,
              observation:
                advanceAfter === undefined
                  ? undefined
                  : Object.freeze({ advanceAfterMs: Math.round(advanceAfter) }),
            }
          : compileTransition(transitionValue, advanceAfter);

      const resolved: ResolvedSlide = {
        extensions: compiled.extensions,
        ...(compiled.observation === undefined
          ? {}
          : { observation: compiled.observation }),
        emit(): SlideEmitOutput {
          return { xml: compiled.xml };
        },
      };

      return {
        ...(compiled.matchSourceKinds
          ? { continuity: { matchSourceKinds: true } }
          : {}),
        resolveSlide() {
          return resolved;
        },
      };
    },
  };
