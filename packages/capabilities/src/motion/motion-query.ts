import type { ResolvedSlide, SlideLocator } from "@pptxdk/core/capability";
import { BuildQueryError } from "@pptxdk/core/query";

import type {
  MotionEffectObservation,
  MotionMoment,
  MotionScopeDeclaration,
  MotionScopeMode,
  MotionStartDependency,
  MotionTiming,
} from "./motion-api.ts";

export type MotionTime =
  | { readonly state: "known"; readonly ms: number }
  | { readonly state: "unknown" }
  | { readonly state: "no-natural-completion" };

export type MotionDecision =
  | { readonly state: "known"; readonly value: boolean }
  | {
      readonly state: "unknown";
      readonly reason: "insufficient-facts" | "different-step";
    }
  | { readonly state: "no-natural-completion" };

export type MotionDifference =
  | { readonly state: "known"; readonly ms: number }
  | {
      readonly state: "unknown";
      readonly reason: "insufficient-facts" | "different-step";
      readonly minimumMs?: number;
      readonly maximumMs?: number;
    }
  | { readonly state: "no-natural-completion" };

export type MotionStepRef = object;
export type MotionEffectRef = {
  readonly targetId: number;
  readonly effectIndex: number;
  readonly partIndex?: number;
};
export type MotionScopeRef = { readonly id: string };
export type MotionCueRef = {
  readonly name: string;
  readonly scopeId: string;
  readonly point: "start" | "end";
};
export type MotionMomentRef =
  | {
      readonly kind: "effect";
      readonly effect: MotionEffectRef;
      readonly point: "start" | "end";
    }
  | {
      readonly kind: "scope";
      readonly scope: MotionScopeRef;
      readonly point: "start" | "end";
    };
export type MotionIntervalRef = MotionEffectRef | MotionScopeRef;

export type MotionMomentObservation =
  | ({ readonly ref: MotionMomentRef } & {
      readonly state: "known";
      readonly ms: number;
    })
  | ({ readonly ref: MotionMomentRef } & {
      readonly state: "unknown";
    })
  | ({ readonly ref: MotionMomentRef } & {
      readonly state: "no-natural-completion";
    });

export type EffectInstanceObservation = MotionEffectObservation & {
  readonly ref: MotionEffectRef;
  readonly effectIndex: number;
  readonly partIndex?: number;
  readonly scopes: readonly MotionScopeRef[];
  readonly start: MotionMomentObservation;
  readonly end: MotionMomentObservation;
};

export type MotionScopeQueryMember =
  | { readonly kind: "scope"; readonly ref: MotionScopeRef }
  | { readonly kind: "effect"; readonly ref: MotionEffectRef };

export type MotionScopeQueryObservation = {
  readonly ref: MotionScopeRef;
  readonly step: MotionStepRef;
  readonly parent?: MotionScopeRef;
  readonly mode: MotionScopeMode;
  readonly eachMs?: number;
  readonly delayMs: number;
  readonly members: readonly MotionScopeQueryMember[];
  readonly start: MotionMomentObservation;
  readonly end: MotionMomentObservation;
};

export type MotionCueObservation = {
  readonly ref: MotionCueRef;
  readonly name: string;
  readonly publishedBy: MotionScopeRef;
  readonly point: "start" | "end";
  readonly at: MotionMomentRef;
};

export type MotionDependencyObservation =
  | {
      readonly kind: "scope-start";
      readonly scope: MotionScopeRef;
      readonly source: MotionMomentRef;
      readonly offsetMs: number;
    }
  | {
      readonly kind: "cue-use";
      readonly scope: MotionScopeRef;
      readonly cue: MotionCueRef;
      readonly offsetMs: number;
    }
  | {
      readonly kind: "cue-publish";
      readonly cue: MotionCueRef;
      readonly scope: MotionScopeRef;
      readonly point: "start" | "end";
    }
  | {
      readonly kind: "start-after";
      readonly effect: MotionEffectRef;
      readonly source: MotionStartDependency;
    };

export type MotionMembership = {
  readonly complete: boolean;
  readonly items: readonly {
    readonly effect: MotionEffectRef;
    readonly result: MotionDecision;
  }[];
};

export type MotionStepQuery = {
  readonly ref: MotionStepRef;
  readonly trigger: "auto" | "click";
  readonly number?: number;
  readonly effects: readonly EffectInstanceObservation[];
  readonly scopes: readonly MotionScopeQueryObservation[];
  readonly dependencies: readonly MotionDependencyObservation[];
  readonly cues: readonly MotionCueObservation[];
  difference(from: MotionMomentRef, to: MotionMomentRef): MotionDifference;
  before(from: MotionMomentRef, to: MotionMomentRef): MotionDecision;
  overlap(a: MotionIntervalRef, b: MotionIntervalRef): MotionDecision;
  intervalsCovering(
    ms: number,
    effects?: readonly MotionEffectRef[],
  ): MotionMembership;
};

export type MotionSchedule = {
  readonly evidence: "generation";
  readonly steps: readonly MotionStepQuery[];
};

/** Transitional alias retained until Ticket 11. */
export type MotionQuery = MotionSchedule;

export type MotionFactReader = {
  exactOffset(from: string, to: string): number | undefined;
  lowerBound(from: string, to: string): number | undefined;
};

export type MotionQueryMomentSource = {
  readonly eventId?: string;
  readonly time: MotionTime;
};

export type MotionQueryEffectSource = {
  readonly observation: MotionEffectObservation;
  readonly targetId: number;
  readonly effectIndex: number;
  readonly partIndex?: number;
  readonly start: MotionQueryMomentSource;
  readonly end: MotionQueryMomentSource;
  readonly scopeIdentities: readonly string[];
};

export type MotionQueryScopeMemberSource =
  | { readonly kind: "scope"; readonly identity: string }
  | { readonly kind: "object"; readonly targetId: number };

export type MotionQueryScopeSource = {
  readonly identity: string;
  readonly id: string;
  readonly parentIdentity?: string;
  readonly mode: MotionScopeMode;
  readonly eachMs?: number;
  readonly delayMs: number;
  readonly members: readonly MotionQueryScopeMemberSource[];
  readonly start: MotionQueryMomentSource;
  readonly end: MotionQueryMomentSource;
  readonly declaration: MotionScopeDeclaration;
};

export type MotionQueryCueSource = {
  readonly key: string;
  readonly name: string;
  readonly scopeIdentity: string;
  readonly point: "start" | "end";
};

export type MotionQueryStepSource = {
  readonly key: string;
  readonly trigger: "auto" | "click";
  readonly number?: number;
  readonly facts: MotionFactReader;
  readonly effects: readonly MotionQueryEffectSource[];
  readonly scopes: readonly MotionQueryScopeSource[];
  readonly cues: readonly MotionQueryCueSource[];
};

type MotionReferenceKind = "step" | "effect" | "scope" | "moment" | "cue";

type MotionReferenceOwner = {
  readonly query: object;
  readonly step: object;
  readonly kind: MotionReferenceKind;
  readonly eventId?: string;
  readonly time: MotionTime;
  readonly start?: MotionMomentRef;
  readonly end?: MotionMomentRef;
};

// Handle ownership is preserved only in WeakMap; reports and public references do not carry build identity tokens.
const referenceOwners = new WeakMap<object, MotionReferenceOwner>();
const queryByObservation = new WeakMap<object, MotionQuery>();

function frozenRef<T extends object>(value: T, owner: MotionReferenceOwner): T {
  const ref = Object.freeze(value);
  referenceOwners.set(ref, owner);
  return ref;
}

function freezeMoment(
  ref: MotionMomentRef,
  time: MotionTime,
): MotionMomentObservation {
  return Object.freeze({ ...time, ref });
}

function projectEffectObservation(
  source: MotionEffectObservation,
): MotionEffectObservation {
  const target = Object.freeze({
    id: source.target.id,
    name: source.target.name,
    ...(source.target.authorId === undefined
      ? {}
      : { authorId: source.target.authorId }),
    ...(source.target.range === undefined
      ? {}
      : { range: Object.freeze({ ...source.target.range }) }),
  });
  return Object.freeze({
    target,
    kind: source.kind,
    preset: source.preset,
    durationMs: source.durationMs,
    ...(source.repeatCount === undefined
      ? {}
      : { repeatCount: source.repeatCount }),
    delayMs: source.delayMs,
    ...(source.by === undefined ? {} : { by: source.by }),
    ...(source.intervalMs === undefined
      ? {}
      : { intervalMs: source.intervalMs }),
    ...(source.from === undefined ? {} : { from: source.from }),
    ...(source.fromColor === undefined ? {} : { fromColor: source.fromColor }),
    ...(source.toColor === undefined ? {} : { toColor: source.toColor }),
    ...(source.easing === undefined
      ? {}
      : { easing: Object.freeze({ ...source.easing }) }),
    ...(source.startMs === undefined ? {} : { startMs: source.startMs }),
    ...(source.endMs === undefined ? {} : { endMs: source.endMs }),
    ...(source.completion === undefined
      ? {}
      : { completion: source.completion }),
    ...(source.startAfter === undefined
      ? {}
      : { startAfter: Object.freeze({ ...source.startAfter }) }),
    ...(source.startAfterAll === undefined
      ? {}
      : {
          startAfterAll: Object.freeze(
            source.startAfterAll.map((dependency) =>
              Object.freeze({ ...dependency }),
            ),
          ),
        }),
  });
}

function assertOwnedReference(
  ref: object,
  queryIdentity: object,
  kind: MotionReferenceKind,
): MotionReferenceOwner {
  const owner = referenceOwners.get(ref);
  if (owner === undefined || owner.kind !== kind) {
    throw new BuildQueryError(
      "invalid-reference",
      "Invalid motion query reference",
    );
  }
  if (owner.query !== queryIdentity) {
    throw new BuildQueryError(
      "foreign-reference",
      "Motion query reference belongs to another build or slide",
    );
  }
  return owner;
}

function unknownDecision(
  reason: "insufficient-facts" | "different-step",
): MotionDecision {
  return { state: "unknown", reason };
}

function timeDifference(
  from: MotionReferenceOwner,
  to: MotionReferenceOwner,
  facts: MotionFactReader,
): MotionDifference {
  if (
    from.time.state === "no-natural-completion" ||
    to.time.state === "no-natural-completion"
  ) {
    return { state: "no-natural-completion" };
  }
  if (from.eventId === undefined || to.eventId === undefined) {
    return { state: "unknown", reason: "insufficient-facts" };
  }
  const exact = facts.exactOffset(from.eventId, to.eventId);
  if (exact !== undefined) return { state: "known", ms: exact };
  const minimumMs = facts.lowerBound(from.eventId, to.eventId);
  const reverseMinimum = facts.lowerBound(to.eventId, from.eventId);
  const maximumMs = reverseMinimum === undefined ? undefined : -reverseMinimum;
  return {
    state: "unknown",
    reason: "insufficient-facts",
    ...(minimumMs === undefined ? {} : { minimumMs }),
    ...(maximumMs === undefined ? {} : { maximumMs }),
  };
}

function compareBefore(difference: MotionDifference): MotionDecision {
  if (difference.state === "known") {
    return { state: "known", value: difference.ms > 0 };
  }
  if (difference.state === "no-natural-completion") return difference;
  if (difference.minimumMs !== undefined && difference.minimumMs > 0) {
    return { state: "known", value: true };
  }
  if (difference.maximumMs !== undefined && difference.maximumMs <= 0) {
    return { state: "known", value: false };
  }
  return unknownDecision(difference.reason);
}

function combineOverlap(
  first: MotionDecision,
  second: MotionDecision,
): MotionDecision {
  if (
    (first.state === "known" && !first.value) ||
    (second.state === "known" && !second.value)
  ) {
    return { state: "known", value: false };
  }
  if (
    first.state === "known" &&
    first.value &&
    second.state === "known" &&
    second.value
  ) {
    return { state: "known", value: true };
  }
  if (
    (first.state === "unknown" && first.reason === "different-step") ||
    (second.state === "unknown" && second.reason === "different-step")
  ) {
    return unknownDecision("different-step");
  }
  if (
    first.state === "no-natural-completion" ||
    second.state === "no-natural-completion"
  ) {
    return { state: "no-natural-completion" };
  }
  return unknownDecision("insufficient-facts");
}

function intervalMembership(
  ms: number,
  start: MotionTime,
  end: MotionTime,
): MotionDecision {
  if (start.state === "known" && ms < start.ms) {
    return { state: "known", value: false };
  }
  if (end.state === "known" && ms >= end.ms) {
    return { state: "known", value: false };
  }
  if (start.state === "known" && end.state === "known") {
    return { state: "known", value: start.ms <= ms && ms < end.ms };
  }
  if (start.state === "known" && end.state === "no-natural-completion") {
    return { state: "no-natural-completion" };
  }
  if (
    start.state === "no-natural-completion" ||
    end.state === "no-natural-completion"
  ) {
    return start.state === "known"
      ? { state: "no-natural-completion" }
      : unknownDecision("insufficient-facts");
  }
  return unknownDecision("insufficient-facts");
}

export function createMotionQuery(
  sources: readonly MotionQueryStepSource[],
): MotionQuery {
  const queryIdentity = {};
  const stepIdentities = new Map<string, object>();
  for (const source of sources) stepIdentities.set(source.key, {});

  const steps = sources.map((source) => {
    const stepIdentity = stepIdentities.get(source.key);
    if (stepIdentity === undefined)
      throw new Error("Missing motion step query reference");
    const stepRef = frozenRef(
      {},
      {
        query: queryIdentity,
        step: stepIdentity,
        kind: "step",
        time: { state: "unknown" },
      },
    );
    const scopesByIdentity = new Map<string, MotionScopeRef>();
    const effectsByTarget = new Map<number, MotionEffectRef[]>();
    const effectSourcesByRef = new Map<object, MotionQueryEffectSource>();
    const momentOwners = new Map<object, MotionReferenceOwner>();

    for (const scope of source.scopes) {
      const ref = frozenRef(
        { id: scope.id },
        {
          query: queryIdentity,
          step: stepIdentity,
          kind: "scope",
          time: scope.start.time,
        },
      );
      scopesByIdentity.set(scope.identity, ref);
    }

    for (const effect of source.effects) {
      const ref = frozenRef(
        {
          targetId: effect.targetId,
          effectIndex: effect.effectIndex,
          ...(effect.partIndex === undefined
            ? {}
            : { partIndex: effect.partIndex }),
        },
        {
          query: queryIdentity,
          step: stepIdentity,
          kind: "effect",
          time: effect.start.time,
        },
      );
      const instances = effectsByTarget.get(effect.targetId) ?? [];
      instances.push(ref);
      effectsByTarget.set(effect.targetId, instances);
      effectSourcesByRef.set(ref, effect);
    }

    const makeMoment = (
      kind: "effect" | "scope",
      ownerRef: MotionEffectRef | MotionScopeRef,
      point: "start" | "end",
      sourceMoment: MotionQueryMomentSource,
    ): MotionMomentObservation => {
      const ref: MotionMomentRef = Object.freeze(
        kind === "effect"
          ? { kind, effect: ownerRef as MotionEffectRef, point }
          : { kind, scope: ownerRef as MotionScopeRef, point },
      );
      const owner = {
        query: queryIdentity,
        step: stepIdentity,
        kind: "moment" as const,
        ...(sourceMoment.eventId === undefined
          ? {}
          : { eventId: sourceMoment.eventId }),
        time: sourceMoment.time,
      };
      referenceOwners.set(ref, owner);
      momentOwners.set(ref, owner);
      return freezeMoment(ref, sourceMoment.time);
    };

    const effectObservations = source.effects.map((effect) => {
      const ref = [...(effectsByTarget.get(effect.targetId) ?? [])].find(
        (candidate) =>
          candidate.effectIndex === effect.effectIndex &&
          candidate.partIndex === effect.partIndex,
      );
      if (ref === undefined)
        throw new Error("Missing motion effect instance reference");
      const start = makeMoment("effect", ref, "start", effect.start);
      const end = makeMoment("effect", ref, "end", effect.end);
      referenceOwners.set(ref, {
        query: queryIdentity,
        step: stepIdentity,
        kind: "effect",
        time: effect.start.time,
        start: start.ref,
        end: end.ref,
      });
      const scopes = effect.scopeIdentities.flatMap((identity) => {
        const scope = scopesByIdentity.get(identity);
        return scope === undefined ? [] : [scope];
      });
      // MotionPlan runtime records contain writer-private fields; query copies public observation fields only.
      const observation = projectEffectObservation(effect.observation);
      return Object.freeze({
        ...observation,
        ref,
        effectIndex: effect.effectIndex,
        ...(effect.partIndex === undefined
          ? {}
          : { partIndex: effect.partIndex }),
        scopes: Object.freeze(scopes),
        start,
        end,
      });
    });

    const scopeObservations = source.scopes.map((scope) => {
      const ref = scopesByIdentity.get(scope.identity);
      if (ref === undefined) throw new Error("Missing MotionScope reference");
      const start = makeMoment("scope", ref, "start", scope.start);
      const end = makeMoment("scope", ref, "end", scope.end);
      referenceOwners.set(ref, {
        query: queryIdentity,
        step: stepIdentity,
        kind: "scope",
        time: scope.start.time,
        start: start.ref,
        end: end.ref,
      });
      const parent =
        scope.parentIdentity === undefined
          ? undefined
          : scopesByIdentity.get(scope.parentIdentity);
      const members = scope.members.flatMap<MotionScopeQueryMember>(
        (member) => {
          if (member.kind === "scope") {
            const child = scopesByIdentity.get(member.identity);
            return child === undefined
              ? []
              : [Object.freeze({ kind: "scope" as const, ref: child })];
          }
          return (effectsByTarget.get(member.targetId) ?? [])
            .map((effectRef) => ({
              kind: "effect" as const,
              ref: effectRef,
            }))
            .map((effect) => Object.freeze(effect));
        },
      );
      return Object.freeze({
        ref,
        step: stepRef,
        ...(parent === undefined ? {} : { parent }),
        mode: scope.mode,
        ...(scope.eachMs === undefined ? {} : { eachMs: scope.eachMs }),
        delayMs: scope.delayMs,
        members: Object.freeze(members),
        start,
        end,
      });
    });

    const cuesByKey = new Map<string, MotionCueObservation>();
    const cues = source.cues.flatMap((cue) => {
      const scope = scopesByIdentity.get(cue.scopeIdentity);
      if (scope === undefined) return [];
      const at = scopeObservations.find(({ ref }) => ref === scope)?.[cue.point]
        .ref;
      if (at === undefined) return [];
      const ref = frozenRef(
        { name: cue.name, scopeId: scope.id, point: cue.point },
        {
          query: queryIdentity,
          step: stepIdentity,
          kind: "cue",
          time: { state: "unknown" },
        },
      );
      const observation = Object.freeze({
        ref,
        name: cue.name,
        publishedBy: scope,
        point: cue.point,
        at,
      });
      cuesByKey.set(cue.key, observation);
      return [observation];
    });

    const dependencies: MotionDependencyObservation[] = [];
    for (const scope of source.scopes) {
      const consumer = scopesByIdentity.get(scope.identity);
      if (consumer === undefined || scope.declaration.startAt === undefined)
        continue;
      const { moment, offsetMs } = timingParts(scope.declaration.startAt);
      if (moment.point === "named") {
        const cue = cuesByKey.get(momentKey(moment));
        if (cue !== undefined) {
          dependencies.push({
            kind: "cue-use",
            scope: consumer,
            cue: cue.ref,
            offsetMs,
          });
        }
      } else {
        const sourceScope = scopesByIdentity.get(moment.scope);
        const sourceObservation = scopeObservations.find(
          ({ ref }) => ref === sourceScope,
        );
        const sourceMoment = sourceObservation?.[moment.point].ref;
        if (sourceMoment !== undefined) {
          dependencies.push({
            kind: "scope-start",
            scope: consumer,
            source: sourceMoment,
            offsetMs,
          });
        }
      }
    }
    for (const cue of cues) {
      dependencies.push({
        kind: "cue-publish",
        cue: cue.ref,
        scope: cue.publishedBy,
        point: cue.point,
      });
    }
    for (const effect of effectObservations) {
      const ref = effect.ref;
      const scheduled = effectSourcesByRef.get(ref);
      if (scheduled?.observation.startAfter !== undefined) {
        dependencies.push({
          kind: "start-after",
          effect: ref,
          source: Object.freeze({ ...scheduled.observation.startAfter }),
        });
      }
      for (const source of scheduled?.observation.startAfterAll ?? []) {
        dependencies.push({
          kind: "start-after",
          effect: ref,
          source: Object.freeze({ ...source }),
        });
      }
    }

    const momentOwner = (ref: MotionMomentRef) =>
      momentOwners.get(ref) ??
      assertOwnedReference(ref, queryIdentity, "moment");
    const difference = (
      from: MotionMomentRef,
      to: MotionMomentRef,
    ): MotionDifference => {
      const first = momentOwner(from);
      const second = momentOwner(to);
      if (first.step !== stepIdentity || second.step !== stepIdentity) {
        return { state: "unknown", reason: "different-step" };
      }
      return timeDifference(first, second, source.facts);
    };
    const queryStep: MotionStepQuery = {
      ref: stepRef,
      trigger: source.trigger,
      ...(source.number === undefined ? {} : { number: source.number }),
      effects: Object.freeze(effectObservations),
      scopes: Object.freeze(scopeObservations),
      dependencies: Object.freeze(
        dependencies.map((item) => Object.freeze(item)),
      ),
      cues: Object.freeze(cues),
      difference,
      before(from, to) {
        return compareBefore(difference(from, to));
      },
      overlap(a, b) {
        const first = assertInterval(a, queryIdentity);
        const second = assertInterval(b, queryIdentity);
        if (first.step !== stepIdentity || second.step !== stepIdentity) {
          return unknownDecision("different-step");
        }
        const firstStart =
          first.start === undefined ? undefined : momentOwner(first.start);
        const firstEnd =
          first.end === undefined ? undefined : momentOwner(first.end);
        const secondStart =
          second.start === undefined ? undefined : momentOwner(second.start);
        const secondEnd =
          second.end === undefined ? undefined : momentOwner(second.end);
        if (
          firstStart === undefined ||
          firstEnd === undefined ||
          secondStart === undefined ||
          secondEnd === undefined
        ) {
          return { state: "no-natural-completion" };
        }
        const isEmpty = (
          start: MotionReferenceOwner,
          end: MotionReferenceOwner,
        ) => {
          const duration = timeDifference(start, end, source.facts);
          return duration.state === "known" && duration.ms === 0;
        };
        // A half-open interval with equal endpoints is empty; checking intersection inequalities alone would falsely treat it as overlapping.
        if (isEmpty(firstStart, firstEnd) || isEmpty(secondStart, secondEnd)) {
          return { state: "known", value: false };
        }
        return combineOverlap(
          compareBefore(timeDifference(firstStart, secondEnd, source.facts)),
          compareBefore(timeDifference(secondStart, firstEnd, source.facts)),
        );
      },
      intervalsCovering(ms, selected) {
        if (!Number.isFinite(ms) || ms < 0) {
          throw new BuildQueryError(
            "invalid-query",
            "Local time must be non-negative finite milliseconds",
          );
        }
        const references = selected ?? effectObservations.map(({ ref }) => ref);
        const items = references.map((effectRef) => {
          const owner = assertOwnedReference(
            effectRef,
            queryIdentity,
            "effect",
          );
          if (owner.step !== stepIdentity) {
            return Object.freeze({
              effect: effectRef,
              result: unknownDecision("different-step"),
            });
          }
          const effect = effectSourcesByRef.get(effectRef);
          if (effect === undefined)
            throw new Error("Missing motion query effect source");
          return Object.freeze({
            effect: effectRef,
            result: intervalMembership(ms, effect.start.time, effect.end.time),
          });
        });
        return Object.freeze({
          complete: items.every(({ result }) => result.state === "known"),
          items: Object.freeze(items),
        });
      },
    };
    return Object.freeze(queryStep);
  });

  return Object.freeze({ evidence: "generation", steps: Object.freeze(steps) });
}

function assertInterval(
  ref: MotionIntervalRef,
  queryIdentity: object,
): MotionReferenceOwner {
  const owner = referenceOwners.get(ref);
  if (
    owner === undefined ||
    (owner.kind !== "effect" && owner.kind !== "scope")
  ) {
    throw new BuildQueryError(
      "invalid-reference",
      "Invalid motion interval reference",
    );
  }
  if (owner.query !== queryIdentity) {
    throw new BuildQueryError(
      "foreign-reference",
      "Motion query reference belongs to another build or slide",
    );
  }
  return owner;
}

export function registerMotionObservation(
  observation: object,
  query: MotionQuery,
): void {
  queryByObservation.set(observation, query);
}

function isSlideLocator(
  slide: ResolvedSlide | SlideLocator,
): slide is SlideLocator {
  return "observation" in slide && typeof slide.get === "function";
}

/** Accepts only resolved slides or slide locators directly emitted by the current build; reread files will not masquerade as build results. */
export function motionSchedule(
  slide: ResolvedSlide | SlideLocator,
): MotionSchedule | undefined {
  const resolvedSlide = isSlideLocator(slide) ? slide.observation : slide;
  const capability = resolvedSlide.capabilities?.["motion"];
  if (capability === undefined) return undefined;
  if (typeof capability !== "object" || capability === null) {
    throw new BuildQueryError(
      "unbound-observation",
      "The motion slide observation does not belong to the current successful build",
    );
  }
  const query = queryByObservation.get(capability);
  if (query === undefined) {
    throw new BuildQueryError(
      "unbound-observation",
      "The motion slide observation does not belong to the current successful build",
    );
  }
  return query;
}

/** Transitional alias retained until Ticket 11. */
export const observeMotion = motionSchedule;

function timingParts(timing: MotionTiming): {
  readonly moment: MotionMoment;
  readonly offsetMs: number;
} {
  return timing.type === "motion-offset"
    ? { moment: timing.moment, offsetMs: timing.offsetMs }
    : { moment: timing, offsetMs: 0 };
}

function momentKey(moment: MotionMoment): string {
  return `${moment.scope}\u0000${moment.name ?? ""}`;
}
