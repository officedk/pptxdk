import {
  CapabilityError,
  type CapabilityInput,
  type Extension,
  type ResolvedSlide,
  type SlideCapabilityModule,
  type SlideEmitOutput,
} from "@pptxdk/core/capability";
import { defineShowMechanism } from "@pptxdk/import";
import type { z } from "zod";

import {
  advanceAfterSchema,
  transitionKindSchema,
  transitionSchema,
  type FlipTransition,
  type MorphTransition,
  type StandardTransition,
  type StripsDirection,
  type Transition,
  type TransitionDirection,
  type TransitionSpeed,
} from "./transition-api.ts";
import { transitionCodegen } from "./transition-codegen.ts";
import { readSlideTransitionImport } from "./transition-import.ts";

export {
  advanceAfterSchema,
  fadeTransitionSchema,
  flipDirectionSchema,
  flipTransitionSchema,
  morphBySchema,
  morphTransitionSchema,
  pushTransitionSchema,
  standardTransitionSchema,
  stripsDirectionSchema,
  stripsTransitionSchema,
  transitionDirectionSchema,
  transitionDurationSchema,
  transitionSchema,
  transitionSpeedSchema,
  transitions,
  type FlipTransition,
  type MorphTransition,
  type StandardTransition,
  type StripsDirection,
  type Transition,
  type TransitionDirection,
  type TransitionDuration,
  type TransitionSpeed,
} from "./transition-api.ts";

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

type TransitionEmitInput<T extends Transition = Transition> = {
  readonly transition: T;
  readonly duration: number | undefined;
  readonly useDurationExtension: boolean;
  readonly advanceAfter?: number;
};

function fail(code: string, message: string): never {
  throw new CapabilityError(code, message);
}

function mapTransitionIssue(
  raw: Readonly<Record<string, unknown>>,
  issues: readonly z.core.$ZodIssue[],
): never {
  const unrecognized = issues.find(
    (issue) => issue.code === "unrecognized_keys",
  );
  if (unrecognized !== undefined && "keys" in unrecognized) {
    const key = unrecognized.keys[0] ?? "unknown";
    const prefix =
      unrecognized.path[0] === "fallback"
        ? "Transition fallback contains unknown parameter"
        : "Transition declaration contains unknown parameter";
    return fail("unknown-parameter", `${prefix} ${String(key)}`);
  }

  const issue = issues[0];
  if (issue === undefined) {
    return fail("invalid-transition", "Invalid transition declaration");
  }

  if (issue.path[0] === "fallback") {
    if (!Object.hasOwn(raw, "fallback")) {
      return fail(
        "missing-transition-fallback",
        `${String(raw["kind"])} transition must explicitly declare a Transitional fallback`,
      );
    }
    if (issue.path[1] === "duration") {
      return fail(
        "invalid-transition-duration",
        "Transition duration must be a finite millisecond value between 0 and Number.MAX_SAFE_INTEGER",
      );
    }
    if (issue.path[1] === "direction" || issue.path[1] === "from") {
      return fail(
        "invalid-transition-direction",
        "Transition fallback direction is invalid",
      );
    }
    return fail(
      "invalid-transition-fallback",
      "fallback must be fade, strips, or push",
    );
  }

  if (issue.path[0] === "duration") {
    return fail(
      "invalid-transition-duration",
      "Transition duration must be a finite millisecond value between 0 and Number.MAX_SAFE_INTEGER",
    );
  }
  if (issue.path[0] === "direction" || issue.path[0] === "from") {
    return fail(
      "invalid-transition-direction",
      `${String(raw["kind"])}.${String(issue.path[0])} is invalid`,
    );
  }
  if (issue.path[0] === "by") {
    return fail(
      "invalid-morph-option",
      "morph.by must be object, word, or character",
    );
  }
  return fail("invalid-transition", issue.message);
}

function parseTransitionValue(value: unknown): Transition {
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    return fail(
      "invalid-transition",
      "Slide.transition must be a declaration value returned by transition helpers",
    );
  }
  const raw = value as Readonly<Record<string, unknown>>;
  if (raw["type"] !== "transition") {
    return fail(
      "invalid-transition",
      "Slide.transition is missing a valid transition declaration tag",
    );
  }
  if (!transitionKindSchema.safeParse(raw["kind"]).success) {
    return fail(
      "unsupported-transition",
      `Unsupported transition kind ${String(raw["kind"])}`,
    );
  }
  const parsed = transitionSchema.safeParse(raw);
  if (!parsed.success) {
    return mapTransitionIssue(raw, parsed.error.issues);
  }
  return parsed.data;
}

function speed(duration: number | undefined): TransitionSpeed {
  if (duration === undefined || duration > 875) return "slow";
  return duration > 600 ? "med" : "fast";
}

function transitionAttributes(
  duration: number | undefined,
  useDurationExtension: boolean,
  advanceAfter?: number,
): string {
  const transitionSpeed = speed(duration);
  const speedAttribute = ` spd="${transitionSpeed}"`;
  const durationAttribute =
    useDurationExtension && duration !== undefined
      ? ` p14:dur="${Math.round(duration)}"`
      : "";
  const advanceAttribute =
    advanceAfter === undefined ? "" : ` advTm="${Math.round(advanceAfter)}"`;
  return `${speedAttribute}${durationAttribute}${advanceAttribute}`;
}

/** Native show mechanism for `<p:transition><p:fade/></p:transition>`. */
export const transitionFadeMechanism = defineShowMechanism({
  key: "transition-fade",
  role: "show",
  emit(
    model: TransitionEmitInput<Extract<StandardTransition, { kind: "fade" }>>,
    _context: undefined,
  ): SlideEmitOutput {
    const attrs = transitionAttributes(
      model.duration,
      model.useDurationExtension,
      model.advanceAfter,
    );
    return { xml: `<p:transition${attrs}><p:fade/></p:transition>` };
  },
  import: {
    target: "slide",
    slot: "transition",
    read: readSlideTransitionImport,
  },
});

/** Native show mechanism for `<p:transition><p:strips dir="..."/></p:transition>`. */
export const transitionStripsMechanism = defineShowMechanism({
  key: "transition-strips",
  role: "show",
  emit(
    model: TransitionEmitInput<Extract<StandardTransition, { kind: "strips" }>>,
    _context: undefined,
  ): SlideEmitOutput {
    const attrs = transitionAttributes(
      model.duration,
      model.useDurationExtension,
      model.advanceAfter,
    );
    return {
      xml: `<p:transition${attrs}><p:strips dir="${stripDirections[model.transition.direction]}"/></p:transition>`,
    };
  },
  import: {
    target: "slide",
    slot: "transition",
    matches: () => false,
    read: readSlideTransitionImport,
  },
});

/** Native show mechanism for `<p:transition><p:push dir="..."/></p:transition>`. */
export const transitionPushMechanism = defineShowMechanism({
  key: "transition-push",
  role: "show",
  emit(
    model: TransitionEmitInput<Extract<StandardTransition, { kind: "push" }>>,
    _context: undefined,
  ): SlideEmitOutput {
    const attrs = transitionAttributes(
      model.duration,
      model.useDurationExtension,
      model.advanceAfter,
    );
    return {
      xml: `<p:transition${attrs}><p:push dir="${pushDirections[model.transition.from]}"/></p:transition>`,
    };
  },
  import: {
    target: "slide",
    slot: "transition",
    matches: () => false,
    read: readSlideTransitionImport,
  },
});

/** Native show mechanism for `<p:transition><p14:flip dir="..."/></p:transition>`. */
export const transitionFlipMechanism = defineShowMechanism({
  key: "transition-flip",
  role: "show",
  emit(
    model: TransitionEmitInput<FlipTransition>,
    _context: undefined,
  ): SlideEmitOutput {
    const attrs = transitionAttributes(
      model.duration,
      model.useDurationExtension,
      model.advanceAfter,
    );
    return {
      xml: `<p:transition${attrs}><p14:flip dir="${model.transition.direction === "left" ? "l" : "r"}"/></p:transition>`,
    };
  },
  import: {
    target: "slide",
    slot: "transition",
    matches: () => false,
    read: readSlideTransitionImport,
  },
});

/** Native show mechanism for `<p:transition><p159:morph option="..."/></p:transition>`. */
export const transitionMorphMechanism = defineShowMechanism({
  key: "transition-morph",
  role: "show",
  emit(
    model: TransitionEmitInput<MorphTransition>,
    _context: undefined,
  ): SlideEmitOutput {
    const attrs = transitionAttributes(
      model.duration,
      model.useDurationExtension,
      model.advanceAfter,
    );
    return {
      xml: `<p:transition${attrs}><p159:morph option="${morphOptions[model.transition.by]}"/></p:transition>`,
    };
  },
  import: {
    target: "slide",
    slot: "transition",
    matches: () => false,
    read: readSlideTransitionImport,
  },
});

function serializeTransitionElement(
  transition: Transition,
  duration: number | undefined,
  useDurationExtension: boolean,
  advanceAfter?: number,
): string {
  const input = {
    duration,
    useDurationExtension,
    ...(advanceAfter === undefined ? {} : { advanceAfter }),
  };
  switch (transition.kind) {
    case "fade":
      return transitionFadeMechanism.emit({ ...input, transition }, undefined)
        .xml;
    case "strips":
      return transitionStripsMechanism.emit({ ...input, transition }, undefined)
        .xml;
    case "push":
      return transitionPushMechanism.emit({ ...input, transition }, undefined)
        .xml;
    case "flip":
      return transitionFlipMechanism.emit({ ...input, transition }, undefined)
        .xml;
    case "morph":
      return transitionMorphMechanism.emit({ ...input, transition }, undefined)
        .xml;
    default: {
      const exhaustive: never = transition;
      return exhaustive;
    }
  }
}

function standardObservation(
  fallback: StandardTransition,
  inheritedDuration: number | undefined,
  advanceAfter?: number,
): {
  readonly xml: string;
  readonly observation: Readonly<Record<string, unknown>>;
} {
  const effectiveDuration = fallback.duration ?? inheritedDuration;
  return {
    xml: serializeTransitionElement(
      fallback,
      effectiveDuration,
      false,
      advanceAfter,
    ),
    observation: Object.freeze({
      kind: fallback.kind,
      ...(fallback.kind === "strips" ? { direction: fallback.direction } : {}),
      ...(fallback.kind === "push" ? { from: fallback.from } : {}),
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
  const transition = parseTransitionValue(value);
  const duration = transition.duration;
  const fallback =
    transition.kind === "flip" || transition.kind === "morph"
      ? standardObservation(transition.fallback, duration, advanceAfter)
      : duration !== undefined
        ? standardObservation(transition, duration, advanceAfter)
        : undefined;

  const effectExtensions: Extension[] =
    transition.kind === "flip"
      ? ["p14"]
      : transition.kind === "morph"
        ? ["p159"]
        : [];
  const extensions = [
    ...new Set<Extension>([
      ...(duration === undefined ? [] : ["p14" as const]),
      ...effectExtensions,
    ]),
  ];
  const standardXml = serializeTransitionElement(
    transition,
    duration,
    false,
    advanceAfter,
  );
  const observation: Record<string, unknown> = {
    kind: transition.kind,
    ...(transition.kind === "strips"
      ? { direction: transition.direction }
      : {}),
    ...(transition.kind === "push" ? { from: transition.from } : {}),
    ...(transition.kind === "flip" ? { direction: transition.direction } : {}),
    ...(transition.kind === "morph" ? { by: transition.by } : {}),
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
      matchSourceKinds: transition.kind === "morph",
      observation: Object.freeze(observation),
    };
  }

  const exactXml = serializeTransitionElement(
    transition,
    duration,
    true,
    advanceAfter,
  );
  const fallbackXml = fallback?.xml ?? standardXml;
  const requires = extensions.join(" ");
  const declarations = [
    ...(extensions.includes("p14") ? [`xmlns:p14="${p14Namespace}"`] : []),
    ...(extensions.includes("p159") ? [`xmlns:p159="${p159Namespace}"`] : []),
  ].join(" ");
  return {
    xml: `<mc:AlternateContent xmlns:mc="${markupCompatibilityNamespace}" ${declarations}><mc:Choice Requires="${requires}">${exactXml}</mc:Choice><mc:Fallback>${fallbackXml}</mc:Fallback></mc:AlternateContent>`,
    extensions,
    matchSourceKinds: transition.kind === "morph",
    observation: Object.freeze(observation),
  };
}

/**
 * Bound transition module: validates transition & advanceAfter once per slide, layout, or master
 * during bind, submits generic continuity matching collaboration for Morph, and leaves static preview untouched.
 */
export const transitionCapability = {
  declaration: "transition",
  slideProp: "transition",
  slideProps: ["transition", "advanceAfter"],
  slot: "transition",
  codegen: transitionCodegen,
  mechanisms: [
    transitionFadeMechanism,
    transitionStripsMechanism,
    transitionPushMechanism,
    transitionFlipMechanism,
    transitionMorphMechanism,
  ] as const,
  bindSlide(input: CapabilityInput) {
    const rawAdvanceAfter = input.props["advanceAfter"];
    let advanceAfter: number | undefined;
    if (rawAdvanceAfter !== undefined) {
      const parsedAdvance = advanceAfterSchema.safeParse(rawAdvanceAfter);
      if (!parsedAdvance.success) {
        return fail(
          "invalid-advance-after",
          "Slide.advanceAfter must be a finite millisecond value between 0 and 4294967295",
        );
      }
      advanceAfter = parsedAdvance.data;
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
            matchSourceKinds: undefined as boolean | undefined,
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
      ...(compiled.matchSourceKinds === undefined
        ? {}
        : { continuity: { matchSourceKinds: compiled.matchSourceKinds } }),
      resolveSlide() {
        return resolved;
      },
    };
  },
} satisfies SlideCapabilityModule;
