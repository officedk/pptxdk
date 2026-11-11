import {
  CapabilityError,
  type CapabilityContribution,
  type GeometryBox,
} from "@react-pptx/core/capability";

import { maximumTimeNodeValue } from "./motion-api.ts";
import type {
  Effect,
  MotionMoment,
  MotionScopeDeclaration,
  MotionTiming,
  MotionEffectObservation,
  MotionGranularity,
  MotionStartDependency,
  MotionTargetRange,
  MotionObservation,
  MotionScopeObservation,
  MotionStepObservation,
  NormalizedEffect,
} from "./motion-api.ts";
import {
  createMotionQuery,
  type MotionQuery,
  type MotionQueryCueSource,
  type MotionQueryEffectSource,
  type MotionQueryMomentSource,
  type MotionQueryScopeSource,
  type MotionQueryStepSource,
} from "./motion-query.ts";

type StepDeclaration =
  | { readonly mode: "automatic"; readonly identity: string }
  | { readonly mode: "explicit"; readonly number: number };

type MotionTarget = {
  readonly kind: string;
  readonly name: string;
  readonly authorId?: string;
  readonly textBody: boolean;
  readonly filled: boolean;
  readonly paragraphCount?: number;
  readonly chartSeriesCount?: number;
  readonly chartCategoryCount?: number;
  readonly box?: GeometryBox;
  readonly rotationDegrees?: number;
  readonly hasBackground?: boolean;
  readonly members?: readonly MotionNativeMember[];
};

type MotionNativeMember = {
  readonly id: number;
  readonly kind: string;
  readonly name: string;
  readonly authorId?: string;
  readonly box?: GeometryBox;
  readonly members?: readonly MotionNativeMember[];
};

type MotionValue = {
  readonly target: MotionTarget;
  readonly step?: StepDeclaration;
  readonly scopes: readonly MotionScopeDeclaration[];
  readonly effects: readonly NormalizedEffect[];
};

type ScopeValue = {
  readonly type: "scope";
  readonly scope: MotionScopeDeclaration;
};

type ScheduledEffect = MotionEffectObservation & {
  readonly stepKey: string;
  readonly order: number;
  readonly effectIndex: number;
  readonly partIndex?: number;
  readonly targetTextBody: boolean;
  readonly targetFilled: boolean;
  readonly targetBox?: GeometryBox;
  readonly targetRotationDegrees?: number;
  readonly targetHasBackground: boolean;
  readonly targetParagraphCount?: number;
  readonly targetMembers?: readonly MotionNativeMember[];
  readonly iterate?: {
    readonly type: "wd" | "lt";
    readonly intervalMs: number;
  };
};

export type MotionPlanStep = MotionStepObservation & {
  readonly key: string;
  readonly order: number;
  readonly effects: ScheduledEffect[];
  readonly nodes: readonly MotionPlanNode[];
  /** Condition collection placed by first actual consumer node inside root with-group. */
  readonly conditionDefinitionsBefore?: readonly MotionPlanConditionPlacement[];
};

/** Real timeline containment tree; scope and multi-effect object nodes become native parallel containers. */
export type MotionPlanNode =
  | {
      readonly kind: "scope";
      readonly id: string;
      readonly nativeStart: MotionPlanStart;
      readonly children: readonly MotionPlanNode[];
      /** END collector takes over logical completion; real container preserves only logical start. */
      readonly splitCompletion?: true;
      /** Native container preserved for start layering even when child nodes are empty. */
      readonly preserveStart?: true;
      /** Condition collection placed by direct child anchor of real host. */
      readonly conditionDefinitionsBefore?: readonly MotionPlanConditionPlacement[];
    }
  | {
      readonly kind: "completion-collector";
      /** This native container represents only the logical scope END, not its START. */
      readonly scopeId: string;
      readonly nativeStart: MotionPlanStart;
      readonly gates: readonly MotionPlanStartGate[];
    }
  | {
      readonly kind: "target";
      readonly targetId: number;
      readonly nativeStart: MotionPlanStart;
      readonly effects: readonly ScheduledEffect[];
    };

/** Condition collection definition placement determined by plan; writer no longer infers host from consumer node. */
export type MotionPlanConditionPlacement = {
  readonly join: Extract<MotionPlanStart, { readonly kind: "condition-join" }>;
};

/** Resolved start form; PPTX writer handles serialization only, no longer inferring causality. */
export type MotionPlanStart =
  | { readonly kind: "parent-start"; readonly delayMs: number }
  | { readonly kind: "end"; readonly dependency: MotionStartDependency }
  | { readonly kind: "event"; readonly dependency: MotionStartDependency }
  | {
      readonly kind: "delay-layer";
      /** Previous native start; actual node starts after this node. */
      readonly start: MotionPlanStart;
      readonly delayMs: number;
    }
  | {
      readonly kind: "condition-join";
      readonly id: string;
      readonly hostScopeId: string;
      readonly consumerParentScopeId: string;
      readonly beforeNodeId: string;
      readonly gates: readonly MotionPlanStartGate[];
    }
  | {
      readonly kind: "condition-end";
      readonly id: string;
      readonly delayMs: number;
    };

export type MotionPlanStartGate =
  | {
      readonly kind: "parent-start";
      readonly delayMs: number;
      /** Preserves original logical source after exact-equivalence folding. */
      readonly sourceDependency?: MotionStartDependency;
    }
  | { readonly kind: "end"; readonly dependency: MotionStartDependency }
  | { readonly kind: "begin"; readonly dependency: MotionStartDependency }
  | {
      readonly kind: "plan-start";
      /** Reuses same MotionPlan start representation when candidate offsets are layered. */
      readonly start: MotionPlanStart;
      readonly sourceDependency?: MotionStartDependency;
    };

function planStartForGate(gate: MotionPlanStartGate): MotionPlanStart {
  switch (gate.kind) {
    case "parent-start":
      return { kind: "parent-start", delayMs: gate.delayMs };
    case "end":
      return { kind: "end", dependency: gate.dependency };
    case "begin":
      return { kind: "event", dependency: gate.dependency };
    case "plan-start":
      return gate.start;
  }
}

function gateSourceDependencies(
  gate: MotionPlanStartGate,
): readonly MotionStartDependency[] {
  if (gate.kind === "end" || gate.kind === "begin") {
    return [gate.dependency];
  }
  if (gate.kind !== "plan-start") return [];
  if (gate.sourceDependency !== undefined) return [gate.sourceDependency];
  const { base } = motionPlanStartLayers(gate.start);
  return base.kind === "end" || base.kind === "event" ? [base.dependency] : [];
}

export function motionPlanStartLayers(start: MotionPlanStart): {
  readonly base: MotionPlanStart;
  readonly delays: readonly number[];
} {
  let base = start;
  const reversedDelays: number[] = [];
  while (base.kind === "delay-layer") {
    reversedDelays.push(base.delayMs);
    base = base.start;
  }
  return { base, delays: reversedDelays.reverse() };
}

export function addMotionPlanStartDelay(
  start: MotionPlanStart,
  delayMs: number,
): MotionPlanStart | undefined {
  if (!Number.isSafeInteger(delayMs) || delayMs < 0) return undefined;
  if (delayMs === 0) return start;
  const { base, delays } = motionPlanStartLayers(start);
  const baseValue =
    base.kind === "parent-start"
      ? base.delayMs
      : base.kind === "end" || base.kind === "event"
        ? base.dependency.offsetMs
        : base.kind === "condition-end"
          ? base.delayMs
          : 0;
  const totalDelay = delays.reduce((total, delay) => total + delay, baseValue);
  if (!Number.isSafeInteger(totalDelay + delayMs)) return undefined;

  let lowered = start;
  let remaining = delayMs;
  while (remaining > 0) {
    const fieldDelay = Math.min(remaining, maximumTimeNodeValue);
    lowered = addMotionPlanStartDelayField(lowered, fieldDelay);
    remaining -= fieldDelay;
  }
  return lowered;
}

function addMotionPlanStartDelayField(
  start: MotionPlanStart,
  delayMs: number,
): MotionPlanStart {
  // Preserve underlying logical source; in particular, condition-join must not be rewritten into alias lacking original collection.
  if (start.kind === "delay-layer" || start.kind === "condition-join") {
    return { kind: "delay-layer", start, delayMs };
  }
  let delayed: MotionPlanStart;
  if (start.kind === "parent-start") {
    delayed = { kind: start.kind, delayMs: start.delayMs + delayMs };
  } else if (start.kind === "end" || start.kind === "event") {
    delayed = {
      kind: start.kind,
      dependency: {
        ...start.dependency,
        offsetMs: start.dependency.offsetMs + delayMs,
      },
    };
  } else {
    delayed = { ...start, delayMs: start.delayMs + delayMs };
  }
  const value =
    delayed.kind === "parent-start"
      ? delayed.delayMs
      : delayed.kind === "end" || delayed.kind === "event"
        ? delayed.dependency.offsetMs
        : delayed.kind === "condition-end"
          ? delayed.delayMs
          : 0;
  return value <= maximumTimeNodeValue
    ? delayed
    : { kind: "delay-layer", start, delayMs };
}

/** Fold only local delays proven to preserve parent threshold; preserve real container layers otherwise. */
export function composeMotionPlanStarts(
  inherited: MotionPlanStart | undefined,
  child: MotionPlanStart,
): MotionPlanStart | undefined {
  if (child.kind === "delay-layer") {
    const composedBase = composeMotionPlanStarts(inherited, child.start);
    return composedBase === undefined
      ? undefined
      : addMotionPlanStartDelay(composedBase, child.delayMs);
  }
  if (child.kind === "condition-join") {
    return inherited === undefined ? child : undefined;
  }
  if (inherited === undefined) return child;
  if (inherited.kind === "condition-join") {
    return child.kind === "parent-start"
      ? { kind: "condition-end", id: inherited.id, delayMs: child.delayMs }
      : undefined;
  }
  if (child.kind === "end" || child.kind === "event") {
    return inherited.kind === "parent-start" && inherited.delayMs === 0
      ? child
      : undefined;
  }
  if (child.kind === "condition-end") {
    return undefined;
  }
  return addMotionPlanStartDelay(inherited, child.delayMs);
}

/** Single slide result used jointly by writer and build report. */
export type MotionPlan = {
  readonly steps: readonly MotionPlanStep[];
  readonly observation: MotionObservation;
  readonly query: MotionQuery;
};

type TimeCandidate = {
  readonly eventId: string;
  readonly offsetMs: number;
  readonly dependency?: MotionStartDependency;
};

type TimeReference = {
  readonly eventId: string;
  readonly offsetMs: number;
};

type Completion =
  | {
      readonly state: "known";
      readonly timeMs: number;
      readonly eventId: string;
      readonly candidates: readonly TimeCandidate[];
      readonly dependency?: MotionStartDependency;
      readonly nativeStart?: MotionPlanStart;
      readonly reportedStartAfter?: MotionStartDependency;
      readonly reportedStartAfterAll?: readonly MotionStartDependency[];
    }
  | {
      readonly state: "unknown";
      readonly candidates: readonly TimeCandidate[];
      readonly eventId: string;
      readonly dependency?: MotionStartDependency;
      readonly nativeStart?: MotionPlanStart;
      readonly reportedStartAfter?: MotionStartDependency;
      readonly reportedStartAfterAll?: readonly MotionStartDependency[];
    }
  | { readonly state: "no-natural-completion" };

type ScopeMember =
  | { readonly kind: "scope"; readonly id: string; readonly order: number }
  | { readonly kind: "object"; readonly id: number; readonly order: number };

type ScopeState = {
  readonly declaration: MotionScopeDeclaration;
  readonly path: string;
  readonly order: number;
  readonly members: ScopeMember[];
};

type TargetState = {
  readonly id: number;
  readonly path: string;
  readonly order: number;
  readonly value: MotionValue;
  readonly stepKey: string;
  readonly trigger: "auto" | "click";
  readonly number?: number;
};

type StepState = {
  readonly key: string;
  readonly trigger: "auto" | "click";
  readonly number?: number;
  order: number;
};

type TimePoint = Completion;

type TimeExpression =
  | {
      readonly state: "known";
      readonly timeMs: number;
      readonly candidates: readonly TimeCandidate[];
      /** Referenced event identity stored separately from candidate source, allowing exact aliases to reuse existing timeline nodes. */
      readonly references: readonly TimeReference[];
    }
  | {
      readonly state: "unknown";
      readonly candidates: readonly TimeCandidate[];
      readonly references: readonly TimeReference[];
    }
  | { readonly state: "no-natural-completion" };

type EffectExpansion =
  | { readonly type: "single"; readonly count: 1 }
  | {
      readonly type: "iterate";
      readonly count: 1;
      readonly by: "word" | "letter";
      readonly intervalMs: number;
    }
  | {
      readonly type: "parts";
      readonly count: number;
      readonly ranges: readonly MotionTargetRange[];
    };

function fail(code: string, message: string, path?: string): never {
  throw new CapabilityError(code, message, path);
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function isMotionValue(value: unknown): value is MotionValue {
  return (
    isRecord(value) &&
    isRecord(value["target"]) &&
    Array.isArray(value["scopes"]) &&
    Array.isArray(value["effects"])
  );
}

function isScopeValue(value: unknown): value is ScopeValue {
  return (
    isRecord(value) &&
    value["type"] === "scope" &&
    isRecord(value["scope"]) &&
    typeof value["scope"]["identity"] === "string"
  );
}

function stepKey(step: StepDeclaration | undefined): string {
  if (step === undefined) return "slide-entry";
  return step.mode === "explicit"
    ? `explicit:${step.number}`
    : `automatic:${step.identity}`;
}

function stepInfo(step: StepDeclaration | undefined): Omit<StepState, "order"> {
  if (step === undefined) return { key: "slide-entry", trigger: "auto" };
  return step.mode === "explicit"
    ? { key: stepKey(step), trigger: "click", number: step.number }
    : { key: stepKey(step), trigger: "click" };
}

function momentKey(moment: MotionMoment): string {
  return `${moment.scope}\u0000${moment.name ?? ""}`;
}

function timingParts(timing: MotionTiming): {
  readonly moment: MotionMoment;
  readonly offsetMs: number;
} {
  return timing.type === "motion-offset"
    ? { moment: timing.moment, offsetMs: timing.offsetMs }
    : { moment: timing, offsetMs: 0 };
}

function eventSource(event: MotionStartDependency): string {
  if ("scopeId" in event) {
    return event.point === "start"
      ? `scope-start:${event.scopeId}`
      : `scope:${event.scopeId}`;
  }
  if ("group" in event) return `target:${event.targetId}`;
  return `effect:${event.targetId}:${event.effectIndex}:${event.partIndex ?? ""}`;
}

function dependencyAtEvent(
  dependency: MotionStartDependency,
  offsetMs = 0,
): MotionStartDependency {
  return { ...dependency, offsetMs };
}

function dependencyEventId(dependency: MotionStartDependency): string {
  return eventSource(dependency);
}

function eventTime(
  eventId: string,
  state: "known" | "unknown",
  timeMs?: number,
  dependency?: MotionStartDependency,
  nativeStart?: MotionPlanStart,
  reportedStartAfter?: MotionStartDependency,
  reportedStartAfterAll?: readonly MotionStartDependency[],
  candidates?: readonly TimeCandidate[],
): TimePoint {
  const candidate: TimeCandidate = {
    eventId,
    offsetMs: 0,
    ...(dependency === undefined
      ? {}
      : { dependency: dependencyAtEvent(dependency) }),
  };
  if (state === "known") {
    if (timeMs === undefined)
      throw new Error("Known moment lacks numeric value");
    return {
      state,
      timeMs,
      eventId,
      candidates: candidates ?? [candidate],
      ...(dependency === undefined ? {} : { dependency }),
      ...(nativeStart === undefined ? {} : { nativeStart }),
      ...(reportedStartAfter === undefined ? {} : { reportedStartAfter }),
      ...(reportedStartAfterAll === undefined ? {} : { reportedStartAfterAll }),
    };
  }
  return {
    state,
    eventId,
    candidates: candidates ?? [candidate],
    ...(dependency === undefined ? {} : { dependency }),
    ...(nativeStart === undefined ? {} : { nativeStart }),
    ...(reportedStartAfter === undefined ? {} : { reportedStartAfter }),
    ...(reportedStartAfterAll === undefined ? {} : { reportedStartAfterAll }),
  };
}

class MotionTimingFacts {
  private readonly successors = new Map<string, Map<string, number>>();
  // Exact time differences can be derived in both directions; reverse relations do not enter cycle check of necessary prerequisite graph.
  private readonly exactRelations = new Map<string, Map<string, number>>();
  private readonly exactTimes = new Map<string, number>();

  setExactTime(eventId: string, timeMs: number, path: string): void {
    const existing = this.exactTimes.get(eventId);
    if (existing !== undefined && existing !== timeMs) {
      fail(
        "motion-plan-conflict",
        "Conflicting numeric moments for the same choreography event",
        path,
      );
    }
    this.exactTimes.set(eventId, timeMs);
  }

  exactTime(eventId: string): number | undefined {
    return this.exactTimes.get(eventId);
  }

  addExactOffset(
    before: string,
    after: string,
    offsetMs: number,
    path: string,
  ): void {
    const exactBefore = this.exactTime(before);
    const exactAfter = this.exactTime(after);
    if (
      exactBefore !== undefined &&
      exactAfter !== undefined &&
      exactAfter - exactBefore !== offsetMs
    ) {
      fail(
        "motion-plan-conflict",
        "Exact choreography time difference conflicts with known moment",
        path,
      );
    }
    const existing = this.exactRelations.get(before)?.get(after);
    if (existing !== undefined && existing !== offsetMs) {
      fail(
        "motion-plan-conflict",
        "Conflicting exact time differences for the same choreography event",
        path,
      );
    }
    const existingReverse = this.exactRelations.get(after)?.get(before);
    if (existingReverse !== undefined && existingReverse !== -offsetMs) {
      fail(
        "motion-plan-conflict",
        "Conflicting exact time differences for the same choreography event",
        path,
      );
    }
    const edges = this.exactRelations.get(before) ?? new Map<string, number>();
    edges.set(after, offsetMs);
    this.exactRelations.set(before, edges);
    const reverseEdges =
      this.exactRelations.get(after) ?? new Map<string, number>();
    reverseEdges.set(before, -offsetMs);
    this.exactRelations.set(after, reverseEdges);
  }

  exactOffset(before: string, after: string): number | undefined {
    const exactBefore = this.exactTime(before);
    const exactAfter = this.exactTime(after);
    const exactDifference =
      exactBefore === undefined || exactAfter === undefined
        ? undefined
        : exactAfter - exactBefore;
    const relationDifference = this.exactPath(before, after);
    if (
      exactDifference !== undefined &&
      relationDifference !== undefined &&
      exactDifference !== relationDifference
    ) {
      fail(
        "motion-plan-conflict",
        "Exact choreography time difference conflicts with known moment",
      );
    }
    return relationDifference ?? exactDifference;
  }

  addLowerBound(
    before: string,
    after: string,
    minimumMs: number,
    path: string,
  ): void {
    if (before === after) {
      if (minimumMs > 0) {
        fail(
          "motion-cycle",
          "Choreography moment dependencies contain a cycle",
          path,
        );
      }
      return;
    }
    if (this.hasPath(after, before)) {
      fail(
        "motion-cycle",
        "Choreography moment dependencies contain a cycle",
        path,
      );
    }
    const beforeTime = this.exactTime(before);
    const afterTime = this.exactTime(after);
    if (
      beforeTime !== undefined &&
      afterTime !== undefined &&
      afterTime - beforeTime < minimumMs
    ) {
      fail(
        "motion-plan-conflict",
        "Choreography time difference fact conflicts with known moment",
        path,
      );
    }
    const edges = this.successors.get(before) ?? new Map<string, number>();
    edges.set(
      after,
      Math.max(edges.get(after) ?? Number.NEGATIVE_INFINITY, minimumMs),
    );
    this.successors.set(before, edges);
  }

  lowerBound(before: string, after: string): number | undefined {
    const exactDifference = this.exactOffset(before, after);
    const pathDifference = this.longestPath(before, after, new Set());
    if (exactDifference === undefined) return pathDifference;
    if (pathDifference === undefined) return exactDifference;
    if (pathDifference > exactDifference) {
      fail(
        "motion-plan-conflict",
        "Choreography time difference fact conflicts with exact time difference",
      );
    }
    return Math.max(exactDifference, pathDifference);
  }

  earliestTime(eventId: string, originEventId: string): number | undefined {
    const exact = this.exactTime(eventId);
    if (exact !== undefined) return exact;
    return this.lowerBound(originEventId, eventId);
  }

  private hasPath(from: string, to: string): boolean {
    if (from === to) return true;
    const pending = [from];
    const visited = new Set<string>();
    while (pending.length > 0) {
      const current = pending.pop();
      if (current === undefined || visited.has(current)) continue;
      visited.add(current);
      for (const next of this.successors.get(current)?.keys() ?? []) {
        if (next === to) return true;
        pending.push(next);
      }
    }
    return false;
  }

  private longestPath(
    from: string,
    to: string,
    visited: Set<string>,
  ): number | undefined {
    if (from === to) return 0;
    if (visited.has(from)) return undefined;
    visited.add(from);
    let longest: number | undefined;
    for (const [next, weight] of this.successors.get(from) ?? []) {
      const remaining = this.longestPath(next, to, visited);
      if (remaining !== undefined) {
        longest = Math.max(
          longest ?? Number.NEGATIVE_INFINITY,
          weight + remaining,
        );
      }
    }
    visited.delete(from);
    return longest;
  }

  private exactPath(from: string, to: string): number | undefined {
    if (from === to) return 0;
    const offsets = new Map<string, number>([[from, 0]]);
    const pending = [from];
    for (let index = 0; index < pending.length; index += 1) {
      const current = pending[index];
      if (current === undefined) continue;
      const currentOffset = offsets.get(current);
      if (currentOffset === undefined) continue;
      for (const [next, difference] of this.exactRelations.get(current) ?? []) {
        const candidate = currentOffset + difference;
        const existing = offsets.get(next);
        if (existing !== undefined) {
          if (existing !== candidate) {
            fail(
              "motion-plan-conflict",
              "Conflicting exact relations for the same choreography moment",
            );
          }
          continue;
        }
        offsets.set(next, candidate);
        pending.push(next);
      }
    }
    return offsets.get(to);
  }
}

function asExpression(time: TimePoint): TimeExpression {
  if (time.state === "no-natural-completion") return time;
  const reference = { eventId: time.eventId, offsetMs: 0 };
  return time.state === "known"
    ? {
        state: "known",
        timeMs: time.timeMs,
        candidates: time.candidates,
        references: [reference],
      }
    : {
        state: "unknown",
        candidates: time.candidates,
        references: [reference],
      };
}

function add(
  time: TimePoint | TimeExpression,
  offsetMs: number,
): TimeExpression {
  if (time.state === "no-natural-completion") return time;
  const candidates = time.candidates.map((candidate) => ({
    ...candidate,
    offsetMs: candidate.offsetMs + offsetMs,
  }));
  const references =
    "references" in time
      ? time.references
      : [{ eventId: time.eventId, offsetMs: 0 }];
  const shiftedReferences = references.map((reference) => ({
    ...reference,
    offsetMs: reference.offsetMs + offsetMs,
  }));
  return time.state === "known"
    ? {
        state: "known",
        timeMs: time.timeMs + offsetMs,
        candidates,
        references: shiftedReferences,
      }
    : { state: "unknown", candidates, references: shiftedReferences };
}

function commonCandidateOffset(
  candidates: readonly TimeCandidate[],
  reference: readonly TimeCandidate[],
): number | undefined {
  if (candidates.length === 0 || candidates.length !== reference.length) {
    return undefined;
  }
  const referenceByEvent = new Map(
    reference.map((candidate) => [candidate.eventId, candidate]),
  );
  if (referenceByEvent.size !== reference.length) return undefined;
  let commonOffset: number | undefined;
  for (const candidate of candidates) {
    const source = referenceByEvent.get(candidate.eventId);
    if (source === undefined) return undefined;
    const offset = candidate.offsetMs - source.offsetMs;
    if (!Number.isSafeInteger(offset)) return undefined;
    if (commonOffset !== undefined && commonOffset !== offset) {
      return undefined;
    }
    commonOffset = offset;
  }
  return commonOffset;
}

function candidateDominates(
  earlier: TimeCandidate,
  later: TimeCandidate,
  facts: MotionTimingFacts,
): boolean {
  const difference = facts.lowerBound(earlier.eventId, later.eventId);
  return (
    difference !== undefined &&
    difference + later.offsetMs - earlier.offsetMs >= 0
  );
}

function maximum(
  times: readonly (TimePoint | TimeExpression)[],
  facts: MotionTimingFacts,
): TimeExpression {
  if (times.some(({ state }) => state === "no-natural-completion")) {
    return { state: "no-natural-completion" };
  }
  const candidates = times.flatMap((time) =>
    time.state === "no-natural-completion" ? [] : time.candidates,
  );
  const uniqueByEvent = new Map<string, TimeCandidate>();
  for (const candidate of candidates) {
    const current = uniqueByEvent.get(candidate.eventId);
    if (current === undefined || candidate.offsetMs > current.offsetMs) {
      uniqueByEvent.set(candidate.eventId, candidate);
    }
  }
  const unique = [...uniqueByEvent.values()];
  const referencesByEvent = new Map<string, TimeReference>();
  for (const time of times) {
    if (time.state === "no-natural-completion") continue;
    const references =
      "references" in time
        ? time.references
        : [{ eventId: time.eventId, offsetMs: 0 }];
    for (const reference of references) {
      const current = referencesByEvent.get(reference.eventId);
      if (current === undefined || reference.offsetMs > current.offsetMs) {
        referencesByEvent.set(reference.eventId, reference);
      }
    }
  }
  const references = [...referencesByEvent.values()];
  const exactCandidates = unique.map((candidate) => {
    const timeMs = facts.exactTime(candidate.eventId);
    return timeMs === undefined ? undefined : timeMs + candidate.offsetMs;
  });
  if (exactCandidates.every((timeMs) => timeMs !== undefined)) {
    const latest = Math.max(
      0,
      ...exactCandidates.filter(
        (timeMs): timeMs is number => timeMs !== undefined,
      ),
    );
    return {
      state: "known",
      timeMs: latest,
      candidates: unique.filter(
        (_candidate, index) => exactCandidates[index] === latest,
      ),
      references,
    };
  }

  const reduced = unique.filter((candidate, index) => {
    const dominated = unique.some((other, otherIndex) => {
      if (index === otherIndex) return false;
      if (!candidateDominates(candidate, other, facts)) return false;
      const reverse = candidateDominates(other, candidate, facts);
      return !reverse || otherIndex < index;
    });
    return !dominated;
  });
  return { state: "unknown", candidates: reduced, references };
}

function materializeEvent(
  expression: TimeExpression,
  eventId: string,
  dependency: MotionStartDependency | undefined,
  facts: MotionTimingFacts,
  path: string,
  nativeStart?: MotionPlanStart,
  reportedStartAfter?: MotionStartDependency,
  reportedStartAfterAll?: readonly MotionStartDependency[],
  preservedCandidates?: readonly TimeCandidate[],
): TimePoint {
  if (expression.state === "no-natural-completion") return expression;
  for (const candidate of expression.candidates) {
    facts.addLowerBound(candidate.eventId, eventId, candidate.offsetMs, path);
  }
  if (expression.candidates.length === 1) {
    const [candidate] = expression.candidates;
    if (candidate !== undefined) {
      facts.addExactOffset(
        candidate.eventId,
        eventId,
        candidate.offsetMs,
        path,
      );
    }
  }
  if (expression.state === "known") {
    facts.setExactTime(eventId, expression.timeMs, path);
    return eventTime(
      eventId,
      "known",
      expression.timeMs,
      dependency,
      nativeStart,
      reportedStartAfter,
      reportedStartAfterAll,
      preservedCandidates,
    );
  }
  return eventTime(
    eventId,
    "unknown",
    undefined,
    dependency,
    nativeStart,
    reportedStartAfter,
    reportedStartAfterAll,
    preservedCandidates,
  );
}

function completion(
  start: TimePoint,
  children: readonly TimePoint[],
  dependency: MotionStartDependency,
  facts: MotionTimingFacts,
  path: string,
): TimePoint {
  if (
    start.state === "no-natural-completion" ||
    children.some(({ state }) => state === "no-natural-completion")
  ) {
    return { state: "no-natural-completion" };
  }
  if (children.length === 0) return start;
  const eventId = dependencyEventId(dependency);
  for (const member of [start, ...children]) {
    if (member.state !== "no-natural-completion") {
      facts.addLowerBound(member.eventId, eventId, 0, path);
    }
  }
  const end = maximum([start, ...children], facts);
  if (end.state === "no-natural-completion") return end;
  if (end.state === "known") {
    facts.setExactTime(eventId, end.timeMs, path);
  }
  if (end.candidates.length === 1) {
    const [winner] = end.candidates;
    if (winner !== undefined) {
      // The single winning real completion event remains exactly equivalent to overall end; absolute moment may still be unknown.
      facts.addLowerBound(winner.eventId, eventId, winner.offsetMs, path);
      facts.addExactOffset(winner.eventId, eventId, winner.offsetMs, path);
    }
  }
  if (end.state === "known") {
    return eventTime(eventId, "known", end.timeMs, dependency);
  }
  return eventTime(eventId, "unknown", undefined, dependency);
}

function expansionFor(
  target: MotionTarget,
  effect: Effect,
  path: string,
): EffectExpansion {
  if (effect.by === undefined || typeof effect.by === "number") {
    return { type: "single", count: 1 };
  }
  if (effect.by === "word" || effect.by === "letter") {
    if (!target.textBody || !target.paragraphCount) {
      fail(
        "invalid-motion-granularity",
        `by: "${effect.by}" requires native object with text`,
        path,
      );
    }
    return {
      type: "iterate",
      count: 1,
      by: effect.by,
      intervalMs:
        effect.intervalMs ?? Math.max(1, Math.round(effect.duration / 10)),
    };
  }
  if (effect.by === "paragraph") {
    const count = target.paragraphCount ?? 0;
    if (!target.textBody || count < 1) {
      fail(
        "invalid-motion-granularity",
        'by: "paragraph" requires native text object with paragraphs',
        path,
      );
    }
    return {
      type: "parts",
      count,
      ranges: Array.from({ length: count }, (_, index) => ({
        type: "paragraph" as const,
        start: index,
        end: index,
      })),
    };
  }
  const seriesCount = target.chartSeriesCount ?? 0;
  const categoryCount = target.chartCategoryCount ?? 0;
  if (target.kind !== "chart" || seriesCount < 1 || categoryCount < 1) {
    fail(
      "invalid-motion-granularity",
      `by: "${effect.by}" requires native chart with data`,
      path,
    );
  }
  const count = effect.by === "series" ? seriesCount : categoryCount;
  const ranges: MotionTargetRange[] = [
    {
      type: "chart",
      step: "gridLegend",
      seriesIndex: -3,
      categoryIndex: -3,
    },
    ...Array.from({ length: count }, (_, index) =>
      effect.by === "series"
        ? {
            type: "chart" as const,
            step: "series" as const,
            seriesIndex: index,
            categoryIndex: -4,
          }
        : {
            type: "chart" as const,
            step: "category" as const,
            seriesIndex: -4,
            categoryIndex: index,
          },
    ),
  ];
  return { type: "parts", count: ranges.length, ranges };
}

function effectEnd(
  base: TimeExpression,
  target: TargetState,
  effect: NormalizedEffect,
  effectIndex: number,
  expansion: EffectExpansion,
  facts: MotionTimingFacts,
): TimePoint {
  const partIndex =
    expansion.type === "parts" ? expansion.count - 1 : undefined;
  const startOffsetMs =
    Math.round(effect.delay) +
    (partIndex === undefined ? 0 : partIndex * Math.round(effect.duration));
  const startExpression = add(base, startOffsetMs);
  return effectTimingAt(
    startExpression,
    target,
    effect,
    effectIndex,
    partIndex,
    expansion,
    facts,
  ).end;
}

function effectTimingAt(
  startExpression: TimeExpression,
  target: TargetState,
  effect: NormalizedEffect,
  effectIndex: number,
  partIndex: number | undefined,
  expansion: EffectExpansion,
  facts: MotionTimingFacts,
): { readonly start: TimePoint; readonly end: TimePoint } {
  if (startExpression.state === "no-natural-completion") {
    return { start: startExpression, end: startExpression };
  }
  const startEventId = `effect-start:${target.id}:${effectIndex}:${partIndex ?? ""}`;
  const start = materializeEvent(
    startExpression,
    startEventId,
    undefined,
    facts,
    target.path,
  );
  if (start.state === "no-natural-completion") {
    return { start, end: start };
  }
  if (effect.repeatCount === "forever") {
    return { start, end: { state: "no-natural-completion" } };
  }
  const rounds = effect.repeatCount ?? 1;
  const durationMs = Math.round(effect.duration) * rounds;
  if (
    start.state === "known" &&
    !Number.isSafeInteger(start.timeMs + durationMs)
  ) {
    fail(
      "invalid-motion-time",
      "Effect completion moment must be a safe integer in milliseconds",
      target.path,
    );
  }
  const dependency: MotionStartDependency = {
    targetId: target.id,
    effectIndex,
    ...(partIndex === undefined ? {} : { partIndex }),
    point: "end",
    offsetMs: 0,
  };
  const endEventId = dependencyEventId(dependency);
  const existingEndTime = facts.exactTime(endEventId);
  if (existingEndTime !== undefined && !Number.isSafeInteger(existingEndTime)) {
    return {
      start,
      end: eventTime(endEventId, "known", existingEndTime, dependency),
    };
  }
  // Current spTgt retains single behavior for entire object; by-word still has this positive lower bound even without subsets.
  const minimumDuration = durationMs;
  if (expansion.type === "iterate") {
    facts.addLowerBound(startEventId, endEventId, minimumDuration, target.path);
  } else {
    // Finite behavior preserves relative time difference after full repetitions without requiring absolute numeric start.
    facts.addLowerBound(startEventId, endEventId, durationMs, target.path);
    facts.addExactOffset(startEventId, endEventId, durationMs, target.path);
  }
  if (expansion.type !== "iterate" && start.state === "known") {
    const endMs = start.timeMs + durationMs;
    facts.setExactTime(endEventId, endMs, target.path);
    return {
      start,
      end: eventTime(endEventId, "known", endMs, dependency),
    };
  }
  return {
    start,
    end: eventTime(endEventId, "unknown", undefined, dependency),
  };
}

function motionScope(value: MotionScopeDeclaration): boolean {
  return (
    typeof value.identity === "string" &&
    typeof value.mode === "string" &&
    typeof value.delay === "number"
  );
}

function motionPlan(
  contributions: readonly CapabilityContribution[],
): MotionPlan {
  const stepRootHostIdentity = "\u0000motion-step-root";
  const scopes = new Map<string, ScopeState>();
  const targets = new Map<number, TargetState>();
  const roots: ScopeMember[] = [];
  const effectExpansions = new Map<string, EffectExpansion>();
  const getExpansion = (target: TargetState, effectIndex: number) => {
    const key = `${target.id}:${effectIndex}`;
    const existing = effectExpansions.get(key);
    if (existing !== undefined) return existing;
    const effect = target.value.effects[effectIndex];
    if (effect === undefined) {
      fail("invalid-motion", "Motion effect index does not exist", target.path);
    }
    const expansion = expansionFor(target.value.target, effect, target.path);
    effectExpansions.set(key, expansion);
    return expansion;
  };

  for (const contribution of contributions) {
    if (contribution.capability !== "motion") continue;
    if (isScopeValue(contribution.value)) {
      const declaration = contribution.value.scope;
      if (!motionScope(declaration)) {
        fail(
          "invalid-scope",
          "Choreography scope declaration is invalid",
          contribution.path,
        );
      }
      if (scopes.has(declaration.identity)) {
        fail(
          "duplicate-scope",
          "Same choreography scope identity cannot be declared twice",
          contribution.path,
        );
      }
      scopes.set(declaration.identity, {
        declaration,
        path: contribution.path,
        order: contribution.order,
        members: [],
      });
      continue;
    }
    if (!isMotionValue(contribution.value)) {
      fail(
        "invalid-motion",
        `motion annotation is invalid at ${contribution.path}`,
        contribution.path,
      );
    }
    if (contribution.targetId === undefined) {
      fail(
        "invalid-motion",
        "Object motion lacks native object identity",
        contribution.path,
      );
    }
    const trigger = stepInfo(contribution.value.step);
    if (targets.has(contribution.targetId)) {
      fail(
        "invalid-motion",
        "Same native object repeatedly submitted motion declarations",
        contribution.path,
      );
    }
    targets.set(contribution.targetId, {
      id: contribution.targetId,
      path: contribution.path,
      order: contribution.order,
      value: contribution.value,
      stepKey: trigger.key,
      trigger: trigger.trigger,
      ...(trigger.number === undefined ? {} : { number: trigger.number }),
    });
  }

  for (const target of targets.values()) {
    const granularities = new Set(
      target.value.effects
        .map(({ by }) => by)
        .filter((by): by is MotionGranularity => typeof by === "string"),
    );
    const textGranularities = [...granularities].filter(
      (by) => by === "paragraph" || by === "word" || by === "letter",
    );
    const chartGranularities = [...granularities].filter(
      (by) => by === "series" || by === "category",
    );
    const incompatible =
      target.value.target.kind === "chart"
        ? chartGranularities.length > 1
        : textGranularities.length > 1;
    if (incompatible) {
      fail(
        "invalid-motion-granularity",
        "Parallel motions on the same native object cannot mix incompatible build granularities",
        target.path,
      );
    }
  }

  for (const scope of scopes.values()) {
    const parentId = scope.declaration.parent;
    if (parentId === undefined) continue;
    const parent = scopes.get(parentId);
    if (parent === undefined) {
      fail(
        "missing-scope",
        "Parent instance of choreography scope does not exist",
        scope.path,
      );
    }
    parent.members.push({
      kind: "scope",
      id: scope.declaration.identity,
      order: scope.order,
    });
  }

  for (const scope of scopes.values()) {
    if (scope.declaration.parent === undefined) {
      roots.push({
        kind: "scope",
        id: scope.declaration.identity,
        order: scope.order,
      });
    }
  }

  for (const target of targets.values()) {
    let parent: string | undefined;
    for (const declaration of target.value.scopes) {
      if (!motionScope(declaration)) {
        fail(
          "invalid-scope",
          "Object contains invalid choreography scope path",
          target.path,
        );
      }
      const scope = scopes.get(declaration.identity);
      if (scope === undefined) {
        fail(
          "missing-scope",
          "Choreography scope instance referenced by object does not exist",
          target.path,
        );
      }
      if (scope.declaration.parent !== parent) {
        fail(
          "invalid-scope",
          "Object's choreography scope path does not match host declaration",
          target.path,
        );
      }
      parent = declaration.identity;
    }
    const member: ScopeMember = {
      kind: "object",
      id: target.id,
      order: target.order,
    };
    if (parent === undefined) roots.push(member);
    else scopes.get(parent)?.members.push(member);
  }

  roots.sort((left, right) => left.order - right.order);
  for (const scope of scopes.values()) {
    scope.members.sort((left, right) => left.order - right.order);
  }

  const queryDeclarationIndex = new Map<
    string,
    { readonly parent?: string; readonly members: readonly ScopeMember[] }
  >();
  for (const scope of scopes.values()) {
    queryDeclarationIndex.set(scope.declaration.identity, {
      ...(scope.declaration.parent === undefined
        ? {}
        : { parent: scope.declaration.parent }),
      members: scope.members,
    });
  }

  const activeSteps = new Map<string, Set<string>>();
  const collectActiveSteps = (identity: string): Set<string> => {
    const cached = activeSteps.get(identity);
    if (cached !== undefined) return cached;
    const result = new Set<string>();
    activeSteps.set(identity, result);
    const scope = scopes.get(identity);
    if (scope === undefined) return result;
    for (const member of scope.members) {
      if (member.kind === "object") {
        const target = targets.get(member.id);
        if (target !== undefined) result.add(target.stepKey);
      } else {
        for (const key of collectActiveSteps(member.id)) result.add(key);
      }
    }
    return result;
  };
  const scopeSteps = (scope: ScopeState): readonly string[] => {
    const active = [...collectActiveSteps(scope.declaration.identity)];
    return active.length > 0 ? active : [stepKey(scope.declaration.step)];
  };
  const activeMemberSteps = (member: ScopeMember): readonly string[] =>
    member.kind === "object"
      ? [targets.get(member.id)?.stepKey ?? "slide-entry"]
      : [...collectActiveSteps(member.id)];
  const memberExistsInStep = (member: ScopeMember, key: string): boolean => {
    const memberSteps = activeMemberSteps(member);
    if (memberSteps.length > 0) return memberSteps.includes(key);
    if (member.kind !== "scope") return false;
    const child = scopes.get(member.id);
    return child !== undefined && stepKey(child.declaration.step) === key;
  };
  const membersInStep = (
    scope: ScopeState,
    key: string,
  ): readonly ScopeMember[] =>
    scope.members.filter((member) => memberExistsInStep(member, key));
  const memberIsActive = (member: ScopeMember, key: string): boolean =>
    activeMemberSteps(member).includes(key);

  const cuePublishers = new Map<
    string,
    {
      readonly scope: string;
      readonly point: "start" | "end";
      readonly path: string;
    }
  >();
  for (const scope of scopes.values()) {
    const keys = scopeSteps(scope);
    if (
      (scope.declaration.publishStart !== undefined ||
        scope.declaration.publishEnd !== undefined) &&
      keys.length !== 1
    ) {
      fail(
        "cross-step-motion",
        "Scope spanning multiple click steps cannot publish a single moment",
        scope.path,
      );
    }
    for (const [publisher, point] of [
      [scope.declaration.publishStart, "start"],
      [scope.declaration.publishEnd, "end"],
    ] as const) {
      if (publisher === undefined) continue;
      const key = momentKey(publisher);
      if (cuePublishers.has(key)) {
        fail(
          "duplicate-cue",
          "Same named moment cannot be published twice",
          scope.path,
        );
      }
      cuePublishers.set(key, {
        scope: scope.declaration.identity,
        point,
        path: scope.path,
      });
    }
  }

  const scopeObservations: MotionScopeObservation[] = [];
  const stepStates = new Map<string, StepState>();
  const registerStep = (
    key: string,
    trigger: "auto" | "click",
    number: number | undefined,
    order: number,
  ) => {
    const existing = stepStates.get(key);
    if (existing === undefined) {
      stepStates.set(key, {
        key,
        trigger,
        ...(number === undefined ? {} : { number }),
        order,
      });
    } else {
      existing.order = Math.min(existing.order, order);
    }
  };
  for (const target of targets.values()) {
    registerStep(target.stepKey, target.trigger, target.number, target.order);
  }
  for (const scope of scopes.values()) {
    for (const key of scopeSteps(scope)) {
      const isExplicit = key.startsWith("explicit:");
      const number = isExplicit
        ? Number(key.slice("explicit:".length))
        : undefined;
      registerStep(
        key,
        key === "slide-entry" ? "auto" : "click",
        number,
        scope.order,
      );
    }
  }

  const allScheduledEffects = new Map<string, ScheduledEffect[]>();
  const allStepNodes = new Map<string, MotionPlanNode[]>();
  const querySourcesByStep = new Map<
    string,
    Omit<MotionQueryStepSource, "effects"> & {
      readonly effectSources: Map<string, MotionQueryEffectSource>;
    }
  >();
  const allStepConditionDefinitions = new Map<
    string,
    readonly MotionPlanConditionPlacement[]
  >();
  const addStepEffect = (effect: ScheduledEffect) => {
    const items = allScheduledEffects.get(effect.stepKey) ?? [];
    items.push(effect);
    allScheduledEffects.set(effect.stepKey, items);
  };

  for (const step of stepStates.values()) {
    const key = step.key;
    const facts = new MotionTimingFacts();
    const originEventId = `step-origin:${key}`;
    facts.setExactTime(originEventId, 0, "");
    const origin = eventTime(originEventId, "known", 0, undefined, {
      kind: "parent-start",
      delayMs: 0,
    });
    const starts = new Map<string, TimePoint>();
    const ends = new Map<string, TimePoint>();
    // Logical start preserves source representation; TimePoint still uses named event identity for internal completion resolution.
    const scopeStartExpressions = new Map<string, TimeExpression>();
    const scopeCompletionMembers = new Map<string, readonly TimePoint[]>();
    const scopeStartPlans = new Map<string, MotionPlanStart>();
    const conditionJoinsById = new Map<
      string,
      Extract<MotionPlanStart, { readonly kind: "condition-join" }>
    >();
    const conditionPlacementsByHost = new Map<
      string,
      MotionPlanConditionPlacement[]
    >();
    const conditionStartsByEventId = new Map<
      string,
      {
        readonly id: string;
        readonly candidates: readonly TimeCandidate[];
        readonly delayLayers: readonly number[];
      }
    >();
    const resolvingStarts = new Set<string>();
    const resolvingEnds = new Set<string>();
    const scheduledObjects = new Set<number>();
    const queryScopes: MotionQueryScopeSource[] = [];
    const queryEffectSources = new Map<string, MotionQueryEffectSource>();
    // Query completes scopes along structural parent chain; each source still binds to local clock of current step.
    const queryScopeIdentities = new Set<string>();
    for (const scope of scopes.values()) {
      if (!scopeSteps(scope).includes(key)) continue;
      let identity: string | undefined = scope.declaration.identity;
      while (identity !== undefined && !queryScopeIdentities.has(identity)) {
        queryScopeIdentities.add(identity);
        identity = queryDeclarationIndex.get(identity)?.parent;
      }
    }
    // Scope edges come from declaration closure; object effects still filter by active members of this step.
    const queryDeclarationMembers = (
      scope: ScopeState,
    ): readonly ScopeMember[] =>
      queryDeclarationIndex
        .get(scope.declaration.identity)
        ?.members.filter((member) =>
          member.kind === "scope"
            ? queryScopeIdentities.has(member.id)
            : memberIsActive(member, key),
        ) ?? [];
    const projectedNodeHosts = new Map<string, string>();
    const projectedEndHosts = new Map<string, string>();
    const preservedScopeStarts = new Set<string>();
    let projectionChanged = false;

    const directMemberForScope = (
      host: ScopeState,
      identity: string,
    ): number | undefined => {
      let current = scopes.get(identity);
      if (current === undefined || current.declaration.parent === undefined) {
        return undefined;
      }
      while (current.declaration.parent !== host.declaration.identity) {
        const parentId: string | undefined = current.declaration.parent;
        if (parentId === undefined) return undefined;
        const parent = scopes.get(parentId);
        if (parent === undefined || parent.declaration.parent === undefined) {
          return undefined;
        }
        current = parent;
      }
      return host.members.findIndex(
        (member) =>
          member.kind === "scope" && member.id === current.declaration.identity,
      );
    };

    const directMemberForDependency = (
      host: ScopeState,
      dependency: MotionStartDependency,
    ): number | undefined => {
      if ("scopeId" in dependency) {
        const source = [...scopes.values()].find(
          (scope) =>
            scopeId(scope.declaration.identity, scopes) === dependency.scopeId,
        );
        return source === undefined
          ? undefined
          : directMemberForScope(host, source.declaration.identity);
      }
      const source = targets.get(dependency.targetId);
      if (source === undefined) return undefined;
      const containingScope = source.value.scopes.at(-1)?.identity;
      if (containingScope === host.declaration.identity) {
        return host.members.findIndex(
          (member) => member.kind === "object" && member.id === source.id,
        );
      }
      return containingScope === undefined
        ? undefined
        : directMemberForScope(host, containingScope);
    };

    const rootMemberForScope = (identity: string): ScopeMember | undefined => {
      let current = scopes.get(identity);
      while (current?.declaration.parent !== undefined) {
        current = scopes.get(current.declaration.parent);
      }
      if (current === undefined || !scopeSteps(current).includes(key)) {
        return undefined;
      }
      return roots.find(
        (member) =>
          member.kind === "scope" &&
          member.id === current?.declaration.identity,
      );
    };

    const rootMemberForTarget = (
      target: TargetState,
    ): ScopeMember | undefined => {
      const containingScope = target.value.scopes.at(-1)?.identity;
      if (containingScope !== undefined) {
        return rootMemberForScope(containingScope);
      }
      return target.stepKey === key
        ? roots.find(
            (member) => member.kind === "object" && member.id === target.id,
          )
        : undefined;
    };

    const directRootMemberForDependency = (
      dependency: MotionStartDependency,
    ): ScopeMember | undefined => {
      if ("scopeId" in dependency) {
        const source = scopeForPlanId(dependency.scopeId);
        return source === undefined
          ? undefined
          : rootMemberForScope(source.declaration.identity);
      }
      const source = targets.get(dependency.targetId);
      return source === undefined ? undefined : rootMemberForTarget(source);
    };

    const directRootMemberForConsumer = (
      consumer: ScopeMember,
    ): ScopeMember | undefined => {
      if (consumer.kind === "scope") {
        return rootMemberForScope(consumer.id);
      }
      const target = targets.get(consumer.id);
      return target === undefined ? undefined : rootMemberForTarget(target);
    };

    const memberNodeId = (member: ScopeMember): string =>
      member.kind === "scope"
        ? `scope:${scopeId(member.id, scopes)}`
        : `target:${member.id}`;

    const directConsumerAtHost = (
      candidateHost: ScopeState,
      actualParent: ScopeState,
      consumer: ScopeMember,
    ): ScopeMember | undefined => {
      if (
        candidateHost.declaration.identity === actualParent.declaration.identity
      ) {
        return consumer;
      }
      let current =
        consumer.kind === "scope" ? scopes.get(consumer.id) : actualParent;
      while (
        current !== undefined &&
        current.declaration.parent !== candidateHost.declaration.identity
      ) {
        const parentId = current.declaration.parent;
        if (parentId === undefined) return undefined;
        current = scopes.get(parentId);
      }
      return current === undefined
        ? undefined
        : {
            kind: "scope",
            id: current.declaration.identity,
            order: current.order,
          };
    };

    const strictlyAfterParentStart = (
      candidates: readonly TimeCandidate[],
      parentStart: TimePoint | TimeExpression,
    ): boolean => {
      if (parentStart.state === "no-natural-completion") return false;
      if (parentStart.state === "known") {
        return candidates.some((candidate) => {
          const earliestTime = facts.earliestTime(
            candidate.eventId,
            originEventId,
          );
          return (
            earliestTime !== undefined &&
            earliestTime + candidate.offsetMs > parentStart.timeMs
          );
        });
      }
      if (parentStart.candidates.length === 0) return false;
      // Both sides are candidate maximums; find a completion candidate strictly later than each parent candidate.
      return parentStart.candidates.every((parentCandidate) =>
        candidates.some((candidate) => {
          const lowerBound = facts.lowerBound(
            parentCandidate.eventId,
            candidate.eventId,
          );
          return (
            lowerBound !== undefined &&
            lowerBound + candidate.offsetMs - parentCandidate.offsetMs > 0
          );
        }),
      );
    };

    const resolveNativeStart = (
      expression: TimeExpression,
      parentStart: TimePoint,
      path: string,
      eventId: string,
      hostScope?: ScopeState,
      consumer?: ScopeMember,
    ): MotionPlanStart => {
      if (expression.state === "no-natural-completion") {
        fail(
          "no-natural-completion",
          "Content without natural completion cannot start subsequent effects",
          path,
        );
      }
      if (parentStart.state === "no-natural-completion") {
        fail(
          "no-natural-completion",
          "Parent scope without natural completion cannot start subsequent effects",
          path,
        );
      }
      if (expression.state === "known") {
        if (parentStart.state !== "known") {
          fail(
            "unknown-motion-start",
            "Numeric start after unknown parent start has no verifiable merge relation",
            path,
          );
        }
        const delayMs = expression.timeMs - parentStart.timeMs;
        if (delayMs < 0) {
          fail(
            "unknown-motion-start",
            "Child start is earlier than its containing scope",
            path,
          );
        }
        const lowered = addMotionPlanStartDelay(
          { kind: "parent-start", delayMs: 0 },
          delayMs,
        );
        if (lowered === undefined) {
          fail(
            "invalid-motion-time",
            "Delay layer of known start cannot be preserved as safe integer in milliseconds",
            path,
          );
        }
        return lowered;
      }

      const parentStartDelay = commonCandidateOffset(
        expression.candidates,
        parentStart.candidates,
      );
      if (parentStartDelay !== undefined && parentStartDelay >= 0) {
        // Reuse parent start only when every original event increases by the same exact offset.
        const lowered = addMotionPlanStartDelay(
          { kind: "parent-start", delayMs: 0 },
          parentStartDelay,
        );
        if (lowered !== undefined) return lowered;
      }
      for (const reference of expression.references) {
        const canonical = conditionStartsByEventId.get(reference.eventId);
        if (canonical === undefined) continue;
        const delayMs = commonCandidateOffset(
          expression.candidates,
          canonical.candidates,
        );
        if (delayMs === undefined || delayMs !== reference.offsetMs) continue;
        const canonicalDelay = canonical.delayLayers.reduce(
          (total, delay) => total + delay,
          0,
        );
        if (!Number.isSafeInteger(canonicalDelay)) continue;
        // Offset belongs to later alias; cannot restore join END that occurred before real parent activation of consumer.
        const joinEndCandidates = canonical.candidates.map((candidate) => ({
          ...candidate,
          // Recorded candidate describes start of alias scope.
          offsetMs: candidate.offsetMs - canonicalDelay,
        }));
        const joinEndIsPending = joinEndCandidates.some((candidate) => {
          if (parentStart.state === "known") {
            const earliestTime = facts.earliestTime(
              candidate.eventId,
              originEventId,
            );
            return (
              earliestTime !== undefined &&
              earliestTime + candidate.offsetMs > parentStart.timeMs
            );
          }
          return (
            parentStart.candidates.length > 0 &&
            parentStart.candidates.every((parentCandidate) => {
              const lowerBound = facts.lowerBound(
                parentCandidate.eventId,
                candidate.eventId,
              );
              return (
                lowerBound !== undefined &&
                lowerBound + candidate.offsetMs - parentCandidate.offsetMs > 0
              );
            })
          );
        });
        if (!joinEndIsPending) {
          // When canonical fast path is unsafe, preserve full candidate expression for general host projection.
          continue;
        }
        let conditionEndStart: MotionPlanStart = {
          kind: "condition-end",
          id: canonical.id,
          delayMs: 0,
        };
        for (const delay of canonical.delayLayers) {
          const next = addMotionPlanStartDelay(conditionEndStart, delay);
          if (next === undefined) break;
          conditionEndStart = next;
        }
        const lowered = addMotionPlanStartDelay(conditionEndStart, delayMs);
        if (lowered !== undefined) return lowered;
      }

      if (expression.candidates.length !== 1) {
        const startCandidates = expression.candidates.flatMap((candidate) =>
          expandStartCandidate(candidate),
        );
        const unsupported = (reason: string): never =>
          fail(
            "unknown-motion-start",
            `Multiple incomparable unknown starts do not satisfy verified condition collection prerequisites: ${reason}`,
            path,
          );
        if (consumer === undefined) {
          return unsupported("Requires real consumer member in the same step");
        }
        const candidateHosts: (ScopeState | undefined)[] = [];
        let currentHost: ScopeState | undefined = hostScope;
        while (currentHost !== undefined) {
          candidateHosts.push(currentHost);
          const parentId: string | undefined = currentHost.declaration.parent;
          currentHost =
            parentId === undefined ? undefined : scopes.get(parentId);
        }
        candidateHosts.push(undefined);
        let lastReason = "No real host can define all sources beforehand";
        for (const host of candidateHosts) {
          const isStepRoot = host === undefined;
          const hostIdentity =
            host?.declaration.identity ?? stepRootHostIdentity;
          const rootMembers = roots.filter((member) =>
            memberExistsInStep(member, key),
          );
          const consumerBranch = isStepRoot
            ? directRootMemberForConsumer(consumer)
            : hostScope === undefined
              ? undefined
              : directConsumerAtHost(host, hostScope, consumer);
          if (consumerBranch === undefined) {
            lastReason =
              "Cannot locate actual position of consumer member in candidate host";
            continue;
          }
          const hostMembers = isStepRoot ? rootMembers : host.members;
          const consumerIndex = hostMembers.findIndex(
            (member) =>
              member.kind === consumerBranch.kind &&
              member.id === consumerBranch.id,
          );
          if (consumerIndex < 0) {
            lastReason = "Cannot locate actual consumer member";
            continue;
          }
          const hostStart = isStepRoot ? origin : resolveScopeStart(host);
          if (hostStart.state === "no-natural-completion") {
            lastReason = "Candidate host has no subscribable start";
            continue;
          }
          const gates: MotionPlanStartGate[] = [];
          const subscribedSources: TimeCandidate[] = [];
          let validHost = true;
          let requiresProjection = false;
          const appendGateStart = (
            start: MotionPlanStart,
            sourceDependency?: MotionStartDependency,
          ): boolean => {
            if (start.kind === "delay-layer") {
              gates.push({
                kind: "plan-start",
                start,
                ...(sourceDependency === undefined ? {} : { sourceDependency }),
              });
              return true;
            }
            if (start.kind === "parent-start") {
              gates.push({
                kind: "parent-start",
                delayMs: start.delayMs,
                ...(sourceDependency === undefined ? {} : { sourceDependency }),
              });
              return true;
            }
            if (start.kind === "end") {
              gates.push({ kind: "end", dependency: start.dependency });
              return true;
            }
            if (start.kind === "event") {
              gates.push({ kind: "begin", dependency: start.dependency });
              return true;
            }
            return false;
          };
          for (const candidate of startCandidates) {
            let dependency = candidate.dependency;
            if (
              candidate.eventId === parentStart.eventId &&
              !isStepRoot &&
              hostScope !== undefined &&
              host.declaration.identity !== hostScope.declaration.identity
            ) {
              // Parent start keeps independent event identity, avoiding merging multi-layer legal delays into a single field.
              dependency = {
                scopeId: scopeId(hostScope.declaration.identity, scopes),
                point: "start",
                offsetMs: candidate.offsetMs,
              };
            }
            if (
              !Number.isSafeInteger(candidate.offsetMs) ||
              candidate.offsetMs < 0
            ) {
              lastReason = "Completion offset exceeds safe integer range";
              validHost = false;
              break;
            }
            if (candidate.eventId === hostStart.eventId) {
              const lowered = addMotionPlanStartDelay(
                { kind: "parent-start", delayMs: 0 },
                candidate.offsetMs,
              );
              if (lowered === undefined || !appendGateStart(lowered)) {
                lastReason =
                  "Parent start candidate cannot be preserved as safe integer delay layer";
                validHost = false;
                break;
              }
              continue;
            }
            const exactHostOffset = facts.exactOffset(
              hostStart.eventId,
              candidate.eventId,
            );
            if (exactHostOffset !== undefined) {
              const delayMs = exactHostOffset + candidate.offsetMs;
              if (!Number.isSafeInteger(delayMs)) {
                lastReason =
                  "Exact host cumulative offset exceeds safe integer range";
                validHost = false;
                break;
              }
              if (delayMs >= 0) {
                // Sources that already occurred can be offset by sufficiently large candidate delay, then folded into host delay.
                const lowered = addMotionPlanStartDelay(
                  { kind: "parent-start", delayMs: 0 },
                  delayMs,
                );
                const sourceDependencyAtCandidate =
                  dependency === undefined
                    ? undefined
                    : dependencyAtEvent(dependency, candidate.offsetMs);
                if (
                  lowered === undefined ||
                  !appendGateStart(lowered, sourceDependencyAtCandidate)
                ) {
                  lastReason =
                    "Exact host cumulative offset cannot be preserved as safe integer delay layer";
                  validHost = false;
                  break;
                }
                continue;
              }
            }
            if (exactHostOffset !== undefined && exactHostOffset < 0) {
              lastReason =
                "Candidate source is already earlier than host activation and cannot fold into a legal offset";
              validHost = false;
              break;
            }
            if (dependency === undefined) {
              const sourceStart = scopeForStartEvent(candidate.eventId);
              if (sourceStart !== undefined) {
                dependency = {
                  scopeId: scopeId(sourceStart.declaration.identity, scopes),
                  point: "start",
                  offsetMs: candidate.offsetMs,
                };
              }
            }
            if (
              dependency === undefined ||
              !dependencyMatchesEvent(dependency, candidate.eventId)
            ) {
              lastReason =
                "Sources outside candidate host start have no real subscribable events";
              validHost = false;
              break;
            }
            const sourceIndex = isStepRoot
              ? (() => {
                  const sourceBranch =
                    directRootMemberForDependency(dependency);
                  return sourceBranch === undefined
                    ? -1
                    : hostMembers.findIndex(
                        (member) =>
                          member.kind === sourceBranch.kind &&
                          member.id === sourceBranch.id,
                      );
                })()
              : directMemberForDependency(host, dependency);
            if (sourceIndex === consumerIndex) {
              requiresProjection = true;
            }
            if ("scopeId" in dependency && dependency.point === "end") {
              const sourceScope = scopeForPlanId(dependency.scopeId);
              if (isStepRoot) {
                if (
                  sourceScope !== undefined &&
                  hostScope !== undefined &&
                  isAncestorScope(sourceScope, hostScope)
                ) {
                  lastReason =
                    "Overall completion of source would wait for its inner host to finish";
                  validHost = false;
                  break;
                }
              } else {
                let current = host;
                while (sourceScope !== undefined) {
                  if (
                    current.declaration.identity ===
                    sourceScope.declaration.identity
                  ) {
                    lastReason =
                      "Overall completion of source would wait for its inner host to finish";
                    validHost = false;
                    break;
                  }
                  const parentId = current.declaration.parent;
                  if (parentId === undefined) break;
                  const parent = scopes.get(parentId);
                  if (parent === undefined) break;
                  current = parent;
                }
                if (!validHost) break;
              }
            }
            const sourceAfterHost =
              facts.lowerBound(hostStart.eventId, candidate.eventId) ??
              (hostStart.state === "known"
                ? (facts.earliestTime(candidate.eventId, originEventId) ??
                    Number.NEGATIVE_INFINITY) - hostStart.timeMs
                : undefined);
            if (sourceAfterHost === undefined || sourceAfterHost <= 0) {
              lastReason =
                dependency.point === "end"
                  ? "Cannot prove real source end is later than candidate host activation"
                  : "Cannot prove real start event is later than candidate host activation";
              validHost = false;
              break;
            }
            const lowered = addMotionPlanStartDelay(
              dependency.point === "start"
                ? {
                    kind: "event",
                    dependency: dependencyAtEvent(dependency),
                  }
                : {
                    kind: "end",
                    dependency: dependencyAtEvent(dependency),
                  },
              candidate.offsetMs,
            );
            if (lowered === undefined || !appendGateStart(lowered)) {
              lastReason =
                "Completion dependency cannot be preserved as safe integer delay layer";
              validHost = false;
              break;
            }
            subscribedSources.push({ ...candidate, offsetMs: 0 });
            if (dependency.point === "start") {
              const sourceScope = scopeForPlanId(dependency.scopeId);
              if (sourceScope !== undefined) {
                preservedScopeStarts.add(sourceScope.declaration.identity);
              }
            }
          }
          if (!validHost) continue;
          if (
            !isStepRoot &&
            hostScope !== undefined &&
            host.declaration.identity !== hostScope.declaration.identity &&
            !strictlyAfterParentStart(startCandidates, parentStart)
          ) {
            // When parent scope is not yet activated, condition and consumer must be moved together to an earlier host.
            requiresProjection = true;
          }
          if (isStepRoot && hostScope !== undefined) {
            requiresProjection = true;
          }
          const consumerActivation =
            requiresProjection &&
            (isStepRoot ||
              (hostScope !== undefined &&
                host.declaration.identity !== hostScope.declaration.identity))
              ? isStepRoot
                ? hostStart
                : memberStart(host, consumerBranch)
              : requiresProjection
                ? hostStart
                : parentStart;
          if (
            requiresProjection &&
            subscribedSources.length > 0 &&
            !strictlyAfterParentStart(subscribedSources, consumerActivation)
          ) {
            lastReason =
              "Cannot prove condition collection completion is later than consumer member parent scope activation: real source occurred before actual condition host activation";
            continue;
          }
          if (!strictlyAfterParentStart(startCandidates, consumerActivation)) {
            lastReason =
              "Cannot prove condition collection completion is later than consumer member parent scope activation";
            continue;
          }
          const projectedConsumerId = memberNodeId(consumer);
          if (requiresProjection) {
            const previousHost = projectedNodeHosts.get(projectedConsumerId);
            if (previousHost !== undefined && previousHost !== hostIdentity) {
              lastReason =
                "Same consumer node requires multiple physical early hosts";
              continue;
            }
            projectedNodeHosts.set(projectedConsumerId, hostIdentity);
            if (isStepRoot) {
              if (hostScope !== undefined) {
                markProjectedAncestorsAtStepRoot(hostScope);
              }
            } else if (hostScope !== undefined) {
              markProjectedAncestors(hostScope, host);
            }
          }
          const join: Extract<
            MotionPlanStart,
            { readonly kind: "condition-join" }
          > = {
            kind: "condition-join",
            id: eventId,
            hostScopeId: isStepRoot
              ? stepRootHostIdentity
              : scopeId(host.declaration.identity, scopes),
            consumerParentScopeId:
              requiresProjection || hostScope === undefined
                ? hostIdentity
                : scopeId(hostScope.declaration.identity, scopes),
            beforeNodeId: memberNodeId(
              requiresProjection ? consumer : consumerBranch,
            ),
            gates,
          };
          conditionJoinsById.set(join.id, join);
          return join;
        }
        return unsupported(lastReason);
      }
      const [candidate] = expression.candidates;
      if (candidate === undefined) {
        fail(
          "unknown-motion-start",
          "Unknown start has no single native completion event to reference",
          path,
        );
      }

      if (candidate.eventId === parentStart.eventId) {
        const lowered = addMotionPlanStartDelay(
          { kind: "parent-start", delayMs: 0 },
          candidate.offsetMs,
        );
        if (lowered !== undefined) return lowered;
      }
      if (
        parentStart.nativeStart?.kind === "end" &&
        dependencyEventId(parentStart.nativeStart.dependency) ===
          candidate.eventId
      ) {
        const delayMs =
          candidate.offsetMs - parentStart.nativeStart.dependency.offsetMs;
        if (delayMs >= 0) {
          const lowered = addMotionPlanStartDelay(
            { kind: "parent-start", delayMs: 0 },
            delayMs,
          );
          if (lowered !== undefined) return lowered;
        }
      }
      const sourceStart = scopeForStartEvent(candidate.eventId);
      if (sourceStart !== undefined && candidate.dependency === undefined) {
        const dependency: MotionStartDependency = {
          scopeId: scopeId(sourceStart.declaration.identity, scopes),
          point: "start",
          offsetMs: candidate.offsetMs,
        };
        const eventAfterParent = facts.lowerBound(
          parentStart.eventId,
          candidate.eventId,
        );
        if (eventAfterParent !== undefined && eventAfterParent > 0) {
          preservedScopeStarts.add(sourceStart.declaration.identity);
          const lowered = addMotionPlanStartDelay(
            { kind: "event", dependency: { ...dependency, offsetMs: 0 } },
            candidate.offsetMs,
          );
          if (lowered !== undefined) return lowered;
        }
      }

      // Preserve real completion event still later than parent start; switch to exact offset only when it already preceded parent start.
      const eventAfterParent =
        parentStart.eventId === undefined
          ? undefined
          : facts.lowerBound(parentStart.eventId, candidate.eventId);
      const eventAfterKnownParent =
        parentStart.state === "known"
          ? facts.earliestTime(candidate.eventId, originEventId)
          : undefined;
      const knownParentTime =
        parentStart.state === "known" ? parentStart.timeMs : undefined;
      if (
        candidate.dependency !== undefined &&
        ((eventAfterParent !== undefined && eventAfterParent > 0) ||
          (eventAfterKnownParent !== undefined &&
            knownParentTime !== undefined &&
            eventAfterKnownParent > knownParentTime))
      ) {
        const lowered = addMotionPlanStartDelay(
          candidate.dependency.point === "start"
            ? {
                kind: "event",
                dependency: { ...candidate.dependency, offsetMs: 0 },
              }
            : {
                kind: "end",
                dependency: { ...candidate.dependency, offsetMs: 0 },
              },
          candidate.offsetMs,
        );
        if (lowered !== undefined) return lowered;
      }
      const exactParentOffset = facts.exactOffset(
        candidate.eventId,
        parentStart.eventId,
      );
      if (
        exactParentOffset !== undefined &&
        candidate.offsetMs >= exactParentOffset
      ) {
        const lowered = addMotionPlanStartDelay(
          { kind: "parent-start", delayMs: 0 },
          candidate.offsetMs - exactParentOffset,
        );
        if (lowered !== undefined) return lowered;
      }
      fail(
        "unknown-motion-start",
        "Unknown start has no provable parent activation or completion event source",
        path,
      );
    };

    const reportedDependency = (
      nativeStart: MotionPlanStart,
      parentStart: TimePoint,
    ): MotionStartDependency | undefined => {
      if (nativeStart.kind === "delay-layer") {
        const inherited = reportedDependency(nativeStart.start, parentStart);
        return inherited === undefined
          ? undefined
          : dependencyAtEvent(
              inherited,
              inherited.offsetMs + nativeStart.delayMs,
            );
      }
      if (nativeStart.kind === "end" || nativeStart.kind === "event") {
        return nativeStart.dependency;
      }
      if (nativeStart.kind !== "parent-start") return undefined;
      if (parentStart.state === "no-natural-completion") return undefined;
      const inherited = parentStart.reportedStartAfter;
      return inherited === undefined
        ? undefined
        : dependencyAtEvent(
            inherited,
            inherited.offsetMs + nativeStart.delayMs,
          );
    };

    const reportedDependencies = (
      nativeStart: MotionPlanStart,
      parentStart: TimePoint,
    ): readonly MotionStartDependency[] | undefined => {
      if (nativeStart.kind === "delay-layer") {
        const inherited = reportedDependencies(nativeStart.start, parentStart);
        return inherited === undefined
          ? undefined
          : Object.freeze(
              inherited.map((dependency) =>
                dependencyAtEvent(
                  dependency,
                  dependency.offsetMs + nativeStart.delayMs,
                ),
              ),
            );
      }
      const expandStartDependency = (
        dependency: MotionStartDependency,
        seen: ReadonlySet<string> = new Set(),
      ): readonly MotionStartDependency[] => {
        if (!("scopeId" in dependency) || dependency.point !== "start") {
          return [dependency];
        }
        const source = scopeForPlanId(dependency.scopeId);
        if (source === undefined || seen.has(source.declaration.identity)) {
          return [dependency];
        }
        const original = scopeStartExpressions.get(source.declaration.identity);
        if (
          original === undefined ||
          original.state === "no-natural-completion"
        ) {
          return [dependency];
        }
        const nextSeen = new Set(seen);
        nextSeen.add(source.declaration.identity);
        return original.candidates.flatMap((candidate) => {
          let sourceDependency = candidate.dependency;
          if (sourceDependency === undefined) {
            const sourceStart = scopeForStartEvent(candidate.eventId);
            if (sourceStart !== undefined) {
              sourceDependency = {
                scopeId: scopeId(sourceStart.declaration.identity, scopes),
                point: "start",
                offsetMs: 0,
              };
            }
          }
          if (sourceDependency === undefined) return [dependency];
          return expandStartDependency(
            dependencyAtEvent(
              sourceDependency,
              candidate.offsetMs + dependency.offsetMs,
            ),
            nextSeen,
          );
        });
      };
      const reportedGateDependencies = (
        gate: MotionPlanStartGate,
        hostScopeId: string,
      ): readonly MotionStartDependency[] => {
        if (gate.kind === "end") return [gate.dependency];
        if (gate.kind === "begin") {
          return expandStartDependency(gate.dependency);
        }
        if (gate.kind === "parent-start") {
          return gate.sourceDependency === undefined
            ? [
                {
                  // Native numeric gate is relative to real start of collection host; query also preserves this source.
                  scopeId: hostScopeId,
                  point: "start",
                  offsetMs: gate.delayMs,
                },
              ]
            : expandStartDependency(gate.sourceDependency);
        }
        if (gate.sourceDependency !== undefined) {
          return expandStartDependency(gate.sourceDependency);
        }
        const { base, delays } = motionPlanStartLayers(gate.start);
        if (base.kind === "parent-start") {
          return [
            {
              scopeId: hostScopeId,
              point: "start",
              offsetMs:
                base.delayMs +
                delays.reduce((total, delayMs) => total + delayMs, 0),
            },
          ];
        }
        return reportedDependencies(gate.start, parentStart) ?? [];
      };
      if (nativeStart.kind === "end" || nativeStart.kind === "event") {
        return Object.freeze(expandStartDependency(nativeStart.dependency));
      }
      if (nativeStart.kind === "condition-join") {
        const dependencies = nativeStart.gates.flatMap((gate) =>
          reportedGateDependencies(gate, nativeStart.hostScopeId),
        );
        return Object.freeze(dependencies);
      }
      if (nativeStart.kind === "condition-end") {
        const canonical = conditionJoinsById.get(nativeStart.id);
        if (canonical === undefined) {
          throw new Error(
            "MotionPlan condition alias lacks original condition collection",
          );
        }
        const dependencies = canonical.gates.flatMap((gate) =>
          reportedGateDependencies(gate, canonical.hostScopeId),
        );
        return Object.freeze(
          dependencies.map((dependency) =>
            dependencyAtEvent(
              dependency,
              dependency.offsetMs + nativeStart.delayMs,
            ),
          ),
        );
      }
      if (nativeStart.kind !== "parent-start") return undefined;
      if (parentStart.state === "no-natural-completion") return undefined;
      const inherited = parentStart.reportedStartAfterAll;
      if (inherited === undefined) return undefined;
      return Object.freeze(
        inherited.map((dependency) =>
          dependencyAtEvent(
            dependency,
            dependency.offsetMs + nativeStart.delayMs,
          ),
        ),
      );
    };

    const pointExpression = (point: TimePoint): TimeExpression =>
      asExpression(point);

    const scopeStartEventId = (scope: ScopeState): string =>
      `scope-start:${scope.declaration.identity}`;

    const stepKeysOf = (scope: ScopeState): readonly string[] =>
      scopeSteps(scope);
    const ensureSingleStep = (scope: ScopeState): void => {
      if (stepKeysOf(scope).length !== 1) {
        fail(
          "cross-step-motion",
          "Moment of a single choreography scope cannot be determined across click steps",
          scope.path,
        );
      }
    };

    const resolveMoment = (
      reference: MotionMoment,
      fromPath: string,
    ): TimePoint | TimeExpression => {
      const owner = scopes.get(reference.scope);
      if (owner === undefined) {
        fail(
          "missing-cue",
          "Choreography scope corresponding to moment handle does not exist",
          fromPath,
        );
      }
      ensureSingleStep(owner);
      if (scopeSteps(owner)[0] !== key) {
        fail(
          "cross-step-motion",
          "Moment dependencies cannot cross different click steps",
          fromPath,
        );
      }
      let sourceScope = owner;
      let point: "start" | "end" = reference.point === "end" ? "end" : "start";
      if (reference.point === "named") {
        const publisher = cuePublishers.get(momentKey(reference));
        if (publisher === undefined) {
          fail(
            "missing-cue",
            `No component published named moment ${reference.name ?? ""}`,
            fromPath,
          );
        }
        const publishedScope = scopes.get(publisher.scope);
        if (publishedScope === undefined) {
          fail(
            "missing-cue",
            "Publishing scope for named moment does not exist",
            publisher.path,
          );
        }
        ensureSingleStep(publishedScope);
        if (scopeSteps(publishedScope)[0] !== key) {
          fail(
            "cross-step-motion",
            "Named moment dependencies cannot cross different click steps",
            fromPath,
          );
        }
        sourceScope = publishedScope;
        point = publisher.point;
      }
      if (point === "start") {
        const start = resolveScopeStart(sourceScope);
        const expression = asExpression(start);
        const original = scopeStartExpressions.get(
          sourceScope.declaration.identity,
        );
        if (activeChildren(sourceScope).length === 0) {
          // Empty scope can be referenced as real timeline node, preserving its START/END alias layer.
          preservedScopeStarts.add(sourceScope.declaration.identity);
          return expression;
        }
        return expression.state === "no-natural-completion" ||
          original?.state !== "unknown"
          ? expression
          : { ...expression, candidates: original.candidates };
      }
      return resolveScopeEnd(sourceScope);
    };

    const resolveCueTime = (
      timing: MotionTiming,
      fromPath: string,
    ): TimeExpression => {
      const { moment, offsetMs } = timingParts(timing);
      return add(resolveMoment(moment, fromPath), offsetMs);
    };

    const activeChildren = (scope: ScopeState): readonly ScopeMember[] =>
      membersInStep(scope, key).filter((member) => memberIsActive(member, key));

    const directChildren = (scope: ScopeState): readonly ScopeMember[] =>
      membersInStep(scope, key);

    const memberEnd = (parent: ScopeState, member: ScopeMember): TimePoint => {
      if (member.kind === "scope") {
        const child = scopes.get(member.id);
        if (child === undefined) {
          fail(
            "missing-scope",
            "Member declaration of choreography scope does not exist",
            parent.path,
          );
        }
        return resolveScopeEnd(child);
      }
      const target = targets.get(member.id);
      if (target === undefined) {
        fail(
          "invalid-motion",
          "Object member declaration of choreography scope does not exist",
          parent.path,
        );
      }
      const base = memberStart(parent, member);
      const effectEnds = target.value.effects.map((effect, effectIndex) =>
        effectEnd(
          base,
          target,
          effect,
          effectIndex,
          getExpansion(target, effectIndex),
          facts,
        ),
      );
      const animationNodeCount = target.value.effects.reduce(
        (count, _effect, effectIndex) => {
          const expansion = getExpansion(target, effectIndex);
          return count + (expansion.type === "parts" ? expansion.count : 1);
        },
        0,
      );
      if (animationNodeCount > 1) {
        return completion(
          materializeEvent(
            base,
            `target-start:${target.id}`,
            undefined,
            facts,
            target.path,
          ),
          effectEnds,
          {
            targetId: target.id,
            group: "effects",
            point: "end",
            offsetMs: 0,
          },
          facts,
          target.path,
        );
      }
      return (
        effectEnds[0] ?? {
          state: "no-natural-completion",
        }
      );
    };

    const memberStart = (
      parent: ScopeState,
      member: ScopeMember,
    ): TimeExpression => {
      const parentStart = resolveScopeStart(parent);
      const active = activeChildren(parent);
      const activeIndex = active.findIndex(
        (candidate) =>
          candidate.kind === member.kind && candidate.id === member.id,
      );
      if (parent.declaration.mode === "parallel") {
        return pointExpression(parentStart);
      }
      if (parent.declaration.mode === "stagger") {
        if (activeIndex < 0) return pointExpression(parentStart);
        return add(parentStart, activeIndex * (parent.declaration.each ?? 0));
      }
      let cursor = pointExpression(parentStart);
      const direct = directChildren(parent);
      const directIndex = direct.findIndex(
        (candidate) =>
          candidate.kind === member.kind && candidate.id === member.id,
      );
      for (const previous of direct.slice(0, directIndex)) {
        if (!memberIsActive(previous, key)) continue;
        const ended = memberEnd(parent, previous);
        if (
          ended.state === "no-natural-completion" &&
          direct
            .slice(directIndex + 1)
            .some((candidate) => memberIsActive(candidate, key))
        ) {
          return { state: "no-natural-completion" };
        }
        cursor = pointExpression(ended);
      }
      return cursor;
    };

    const resolveScopeStart = (scope: ScopeState): TimePoint => {
      const cached = starts.get(scope.declaration.identity);
      if (cached !== undefined) return cached;
      if (resolvingStarts.has(scope.declaration.identity)) {
        fail(
          "motion-cycle",
          "Choreography moment dependencies contain a cycle",
          scope.path,
        );
      }
      if (scope.declaration.delay > maximumTimeNodeValue) {
        fail(
          "invalid-motion-time",
          "Single scope delay field exceeds ST_TLTime range",
          scope.path,
        );
      }
      resolvingStarts.add(scope.declaration.identity);
      let parentStart = origin;
      let parentScope: ScopeState | undefined;
      let base = pointExpression(origin);
      if (scope.declaration.parent !== undefined) {
        const parent = scopes.get(scope.declaration.parent);
        if (parent === undefined) {
          fail(
            "missing-scope",
            "Parent instance of choreography scope does not exist",
            scope.path,
          );
        }
        parentScope = parent;
        parentStart = resolveScopeStart(parent);
        base = memberStart(parent, {
          kind: "scope",
          id: scope.declaration.identity,
          order: scope.order,
        });
      }
      if (scope.declaration.startAt !== undefined) {
        const fromCue = resolveCueTime(scope.declaration.startAt, scope.path);
        base = maximum([base, fromCue], facts);
      }
      const startExpression = add(base, scope.declaration.delay);
      if (startExpression.state === "no-natural-completion") {
        fail(
          "no-natural-completion",
          "Content without natural completion cannot start subsequent effects",
          scope.path,
        );
      }
      if (parentStart.state === "no-natural-completion") {
        fail(
          "no-natural-completion",
          "Parent scope without natural completion cannot start subsequent effects",
          scope.path,
        );
      }
      const nativeStartArguments = [
        parentStart,
        scope.path,
        scopeStartEventId(scope),
        parentScope,
        parentScope === undefined
          ? undefined
          : {
              kind: "scope" as const,
              id: scope.declaration.identity,
              order: scope.order,
            },
      ] as const;
      const nativeFieldWouldOverflow =
        scope.declaration.delay > 0 &&
        (startExpression.state === "known"
          ? parentStart.state === "known" &&
            startExpression.timeMs - parentStart.timeMs > maximumTimeNodeValue
          : startExpression.candidates.some(
              (candidate) => candidate.offsetMs > maximumTimeNodeValue,
            ));
      let nativeStart: MotionPlanStart | undefined;
      if (nativeFieldWouldOverflow) {
        // Only when accumulated value exceeds single native field, restore encoding from declared start and delay layers.
        const resolvedBaseStart = resolveNativeStart(
          base,
          ...nativeStartArguments,
        );
        nativeStart = addMotionPlanStartDelay(
          resolvedBaseStart,
          scope.declaration.delay,
        );
      } else {
        nativeStart = resolveNativeStart(
          startExpression,
          ...nativeStartArguments,
        );
      }
      if (nativeStart === undefined) {
        fail(
          "invalid-motion-time",
          "Delay layer of scope start cannot be preserved as safe integer in milliseconds",
          scope.path,
        );
      }
      // Alias reads original expression; TimePoint preserves multi-source candidates only, avoiding rewriting internal event identities.
      const start = materializeEvent(
        startExpression,
        scopeStartEventId(scope),
        undefined,
        facts,
        scope.path,
        nativeStart,
        reportedDependency(nativeStart, parentStart),
        reportedDependencies(nativeStart, parentStart),
        startExpression.candidates.length > 1
          ? startExpression.candidates
          : undefined,
      );
      if (nativeStart.kind === "parent-start") {
        // Candidates may expand along aliases to ancestor sources; still record actual parent start threshold into facts graph.
        facts.addLowerBound(
          parentStart.eventId,
          scopeStartEventId(scope),
          nativeStart.delayMs,
          scope.path,
        );
        facts.addExactOffset(
          parentStart.eventId,
          scopeStartEventId(scope),
          nativeStart.delayMs,
          scope.path,
        );
      } else if (nativeStart.kind === "condition-end") {
        facts.addLowerBound(
          nativeStart.id,
          scopeStartEventId(scope),
          nativeStart.delayMs,
          scope.path,
        );
        facts.addExactOffset(
          nativeStart.id,
          scopeStartEventId(scope),
          nativeStart.delayMs,
          scope.path,
        );
      } else if (nativeStart.kind === "event") {
        const sourceEventId = logicalEventId(nativeStart.dependency);
        facts.addLowerBound(
          sourceEventId,
          scopeStartEventId(scope),
          nativeStart.dependency.offsetMs,
          scope.path,
        );
        facts.addExactOffset(
          sourceEventId,
          scopeStartEventId(scope),
          nativeStart.dependency.offsetMs,
          scope.path,
        );
      }
      const { base: nativeStartBase, delays: nativeStartDelays } =
        motionPlanStartLayers(nativeStart);
      const localDelayMs = nativeStartDelays.reduce(
        (total, delayMs) => total + delayMs,
        0,
      );
      if (!Number.isSafeInteger(localDelayMs)) {
        fail(
          "invalid-motion-time",
          "Accumulated delay of scope start must be a safe integer in milliseconds",
          scope.path,
        );
      }
      if (nativeStartBase.kind === "condition-join") {
        conditionStartsByEventId.set(scopeStartEventId(scope), {
          id: nativeStartBase.id,
          candidates: startExpression.candidates,
          delayLayers: nativeStartDelays,
        });
      } else if (nativeStartBase.kind === "condition-end") {
        conditionStartsByEventId.set(scopeStartEventId(scope), {
          id: nativeStartBase.id,
          candidates: startExpression.candidates,
          delayLayers: [nativeStartBase.delayMs, ...nativeStartDelays],
        });
      } else if (nativeStartBase.kind === "parent-start") {
        const inherited = conditionStartsByEventId.get(parentStart.eventId);
        if (inherited !== undefined) {
          conditionStartsByEventId.set(scopeStartEventId(scope), {
            id: inherited.id,
            candidates: startExpression.candidates,
            delayLayers: [
              ...inherited.delayLayers,
              ...(nativeStartBase.delayMs === 0
                ? []
                : [nativeStartBase.delayMs]),
              ...nativeStartDelays,
            ],
          });
        }
      }
      scopeStartExpressions.set(scope.declaration.identity, startExpression);
      scopeStartPlans.set(scope.declaration.identity, nativeStart);
      starts.set(scope.declaration.identity, start);
      resolvingStarts.delete(scope.declaration.identity);
      return start;
    };

    const resolveScopeEnd = (scope: ScopeState): TimePoint => {
      const cached = ends.get(scope.declaration.identity);
      if (cached !== undefined) return cached;
      if (resolvingEnds.has(scope.declaration.identity)) {
        fail(
          "motion-cycle",
          "Choreography moment dependencies contain a cycle",
          scope.path,
        );
      }
      resolvingEnds.add(scope.declaration.identity);
      const start = resolveScopeStart(scope);
      // Inactive scopes can still provide their own boundary moments, but do not extend overall completion of parent scope.
      const members = activeChildren(scope);
      const memberEnds = members.map((member) => memberEnd(scope, member));
      scopeCompletionMembers.set(scope.declaration.identity, memberEnds);
      const end = completion(
        start,
        memberEnds,
        {
          scopeId: scopeId(scope.declaration.identity, scopes),
          point: "end",
          offsetMs: 0,
        },
        facts,
        scope.path,
      );
      ends.set(scope.declaration.identity, end);
      resolvingEnds.delete(scope.declaration.identity);
      return end;
    };

    for (const scope of scopes.values()) {
      if (!scopeSteps(scope).includes(key)) continue;
      if (
        (scope.declaration.mode === "sequence" ||
          scope.declaration.mode === "stagger") &&
        collectActiveSteps(scope.declaration.identity).size > 1
      ) {
        fail(
          "cross-step-motion",
          "Sequence and stagger scopes cannot choreograph different click steps",
          scope.path,
        );
      }
      if (
        scope.declaration.startAt !== undefined &&
        collectActiveSteps(scope.declaration.identity).size > 1
      ) {
        fail(
          "cross-step-motion",
          "Scope spanning multiple click steps cannot start from a single moment",
          scope.path,
        );
      }
      if (!scopeSteps(scope).includes(key)) continue;
      const start = resolveScopeStart(scope);
      const end = resolveScopeEnd(scope);
      const parent =
        scope.declaration.parent === undefined
          ? undefined
          : scopes.get(scope.declaration.parent);
      const id = scopeId(scope.declaration.identity, scopes);
      const parentId =
        parent === undefined
          ? undefined
          : scopeId(parent.declaration.identity, scopes);
      const observedMembers = directChildren(scope).map((member) =>
        member.kind === "scope"
          ? { kind: "scope" as const, id: scopeId(member.id, scopes) }
          : { kind: "object" as const, id: member.id },
      );
      scopeObservations.push({
        id,
        ...(parentId === undefined ? {} : { parentId }),
        mode: scope.declaration.mode,
        ...(scope.declaration.each === undefined
          ? {}
          : { eachMs: Math.round(scope.declaration.each) }),
        delayMs: Math.round(scope.declaration.delay),
        ...(start.state === "known" ? { startMs: start.timeMs } : {}),
        ...(start.state === "unknown" && start.reportedStartAfter !== undefined
          ? { startAfter: start.reportedStartAfter }
          : {}),
        ...(start.state === "unknown" &&
        start.reportedStartAfterAll !== undefined
          ? { startAfterAll: start.reportedStartAfterAll }
          : {}),
        ...(end.state === "known" ? { endMs: end.timeMs } : {}),
        completion: end.state,
        members: observedMembers,
      });
      queryScopes.push({
        identity: scope.declaration.identity,
        id,
        ...(scope.declaration.parent === undefined
          ? {}
          : { parentIdentity: scope.declaration.parent }),
        mode: scope.declaration.mode,
        ...(scope.declaration.each === undefined
          ? {}
          : { eachMs: Math.round(scope.declaration.each) }),
        delayMs: Math.round(scope.declaration.delay),
        members: queryDeclarationMembers(scope).map((member) =>
          member.kind === "scope"
            ? { kind: "scope" as const, identity: member.id }
            : { kind: "object" as const, targetId: member.id },
        ),
        start: queryMomentSource(start),
        end: queryMomentSource(end),
        declaration: scope.declaration,
      });
    }

    const targetsForStep = [...targets.values()]
      .filter((target) => target.stepKey === key)
      .sort((left, right) => left.order - right.order);
    for (const target of targetsForStep) {
      if (scheduledObjects.has(target.id)) continue;
      const containingScope = target.value.scopes.at(-1)?.identity;
      const parent =
        containingScope === undefined ? undefined : scopes.get(containingScope);
      const parentStart =
        parent === undefined ? origin : resolveScopeStart(parent);
      const base =
        parent === undefined
          ? pointExpression(origin)
          : memberStart(parent, {
              kind: "object",
              id: target.id,
              order: target.order,
            });
      for (const [index, effect] of target.value.effects.entries()) {
        const expansion = getExpansion(target, index);
        const entries =
          expansion.type === "parts"
            ? expansion.ranges.map((range, partIndex) => ({
                range,
                partIndex,
              }))
            : [expansion.type === "iterate" ? { iterate: expansion } : {}];
        for (const entry of entries) {
          const partIndex = "partIndex" in entry ? entry.partIndex : undefined;
          const start = add(
            base,
            Math.round(effect.delay) +
              (partIndex === undefined
                ? 0
                : partIndex * Math.round(effect.duration)),
          );
          if (start.state === "no-natural-completion") {
            fail(
              "no-natural-completion",
              "Content without natural completion cannot start subsequent effects",
              target.path,
            );
          }
          const nativeStart = resolveNativeStart(
            start,
            parentStart,
            target.path,
            `effect-start:${target.id}:${index}:${partIndex ?? ""}`,
            parent,
            {
              kind: "object",
              id: target.id,
              order: target.order,
            },
          );
          const startAfter =
            start.state === "unknown"
              ? reportedDependency(nativeStart, parentStart)
              : undefined;
          const startAfterAll =
            start.state === "unknown"
              ? reportedDependencies(nativeStart, parentStart)
              : undefined;
          const timing = effectTimingAt(
            start,
            target,
            effect,
            index,
            partIndex,
            expansion,
            facts,
          );
          const knownStart =
            timing.start.state === "known" ? timing.start.timeMs : undefined;
          const endMs =
            timing.end.state === "known" ? timing.end.timeMs : undefined;
          const from =
            typeof effect.from === "number"
              ? Math.round(effect.from * 60000) / 60000
              : effect.from;
          const by =
            typeof effect.by === "number"
              ? Math.round(effect.by * 60000) / 60000
              : effect.by;
          const easing = normalizedEasing(effect.easing);
          const scheduledEffect: ScheduledEffect = {
            target: {
              id: target.id,
              name: target.value.target.name,
              ...(target.value.target.authorId === undefined
                ? {}
                : { authorId: target.value.target.authorId }),
              ...("range" in entry ? { range: entry.range } : {}),
            },
            kind: effect.kind,
            preset: effect.preset,
            durationMs: Math.round(effect.duration),
            ...(effect.repeatCount === undefined
              ? {}
              : { repeatCount: effect.repeatCount }),
            delayMs: Math.round(effect.delay),
            ...(by === undefined ? {} : { by }),
            ...(expansion.type === "iterate"
              ? { intervalMs: expansion.intervalMs }
              : {}),
            ...(from === undefined ? {} : { from }),
            ...(easing === undefined ? {} : { easing }),
            ...(knownStart === undefined ? {} : { startMs: knownStart }),
            ...(effect.preset === "breathe" || effect.preset === "fillColor"
              ? {
                  fromColor: effect.fromColor,
                  toColor: effect.toColor,
                }
              : {}),
            ...(timing.end.state === "no-natural-completion"
              ? { completion: "no-natural-completion" as const }
              : endMs === undefined
                ? { completion: "unknown" as const }
                : { endMs }),
            ...(startAfter === undefined ? {} : { startAfter }),
            ...(startAfterAll === undefined ? {} : { startAfterAll }),
            stepKey: key,
            order: target.order + index,
            effectIndex: index,
            ...(partIndex === undefined ? {} : { partIndex }),
            targetTextBody: target.value.target.textBody,
            targetFilled: target.value.target.filled,
            ...(target.value.target.box === undefined
              ? {}
              : { targetBox: target.value.target.box }),
            ...(target.value.target.rotationDegrees === undefined
              ? {}
              : { targetRotationDegrees: target.value.target.rotationDegrees }),
            targetHasBackground: target.value.target.hasBackground ?? false,
            ...(target.value.target.paragraphCount === undefined
              ? {}
              : { targetParagraphCount: target.value.target.paragraphCount }),
            ...(target.value.target.members === undefined
              ? {}
              : { targetMembers: target.value.target.members }),
            ...(expansion.type === "iterate"
              ? {
                  iterate: {
                    type: expansion.by === "word" ? "wd" : "lt",
                    intervalMs: expansion.intervalMs,
                  },
                }
              : {}),
          };
          addStepEffect(scheduledEffect);
          queryEffectSources.set(motionEffectKey(target.id, index, partIndex), {
            observation: scheduledEffect,
            targetId: target.id,
            effectIndex: index,
            ...(partIndex === undefined ? {} : { partIndex }),
            start: queryMomentSource(timing.start),
            end: queryMomentSource(timing.end),
            scopeIdentities: target.value.scopes.map(
              ({ identity }) => identity,
            ),
          });
        }
      }
      scheduledObjects.add(target.id);
    }

    const targetNode = (
      target: TargetState,
      parent?: ScopeState,
    ): MotionPlanNode | undefined => {
      const effects = (allScheduledEffects.get(key) ?? []).filter(
        (effect) => effect.target.id === target.id,
      );
      if (effects.length === 0) return undefined;
      const parentStart =
        parent === undefined ? origin : resolveScopeStart(parent);
      const start =
        parent === undefined
          ? pointExpression(origin)
          : memberStart(parent, {
              kind: "object",
              id: target.id,
              order: target.order,
            });
      const nativeStart = resolveNativeStart(
        start,
        parentStart,
        target.path,
        `target-start:${target.id}`,
        parent,
        parent === undefined
          ? undefined
          : {
              kind: "object",
              id: target.id,
              order: target.order,
            },
      );
      recordStartCondition(nativeStart);
      return {
        kind: "target",
        targetId: target.id,
        nativeStart: consumerStart(nativeStart),
        effects,
      };
    };

    function scopeForStartEvent(eventId: string): ScopeState | undefined {
      return [...scopes.values()].find(
        (scope) => `scope-start:${scope.declaration.identity}` === eventId,
      );
    }

    function scopeForPlanId(id: string): ScopeState | undefined {
      return [...scopes.values()].find(
        (scope) => scopeId(scope.declaration.identity, scopes) === id,
      );
    }

    function isAncestorScope(
      ancestor: ScopeState,
      descendant: ScopeState,
    ): boolean {
      let current: ScopeState | undefined = descendant;
      while (current !== undefined) {
        if (current.declaration.identity === ancestor.declaration.identity) {
          return true;
        }
        const parentId: string | undefined = current.declaration.parent;
        current = parentId === undefined ? undefined : scopes.get(parentId);
      }
      return false;
    }

    function hasActiveMembers(scope: ScopeState): boolean {
      return membersInStep(scope, key).some((member) =>
        memberIsActive(member, key),
      );
    }

    function expandStartCandidate(
      candidate: TimeCandidate,
      seen: ReadonlySet<string> = new Set(),
    ): readonly TimeCandidate[] {
      const source = scopeForStartEvent(candidate.eventId);
      if (source === undefined || seen.has(source.declaration.identity)) {
        return [candidate];
      }
      if (!hasActiveMembers(source)) {
        // Empty scope END equals START; preserve it as real timeline node, avoiding folding legal delay layers.
        preservedScopeStarts.add(source.declaration.identity);
        return [candidate];
      }
      const original = scopeStartExpressions.get(source.declaration.identity);
      if (
        original === undefined ||
        original.state === "no-natural-completion"
      ) {
        return [candidate];
      }
      const nextSeen = new Set(seen);
      nextSeen.add(source.declaration.identity);
      const expanded: TimeCandidate[] = [];
      for (const sourceCandidate of original.candidates) {
        const offsetMs = sourceCandidate.offsetMs + candidate.offsetMs;
        if (
          !Number.isSafeInteger(offsetMs) ||
          offsetMs < 0 ||
          offsetMs > maximumTimeNodeValue
        ) {
          if (hasActiveMembers(source)) return [candidate];
          fail(
            "invalid-motion-time",
            "Synthesized offset of empty scope start cannot be written to single ST_TLTime field",
            source.path,
          );
        }
        expanded.push(
          ...expandStartCandidate({ ...sourceCandidate, offsetMs }, nextSeen),
        );
      }
      return expanded.length === 0 ? [candidate] : expanded;
    }

    function dependencyMatchesEvent(
      dependency: MotionStartDependency,
      eventId: string,
    ): boolean {
      if ("scopeId" in dependency && dependency.point === "start") {
        const source = scopeForPlanId(dependency.scopeId);
        return (
          source !== undefined &&
          `scope-start:${source.declaration.identity}` === eventId
        );
      }
      return dependencyEventId(dependency) === eventId;
    }

    function logicalEventId(dependency: MotionStartDependency): string {
      if ("scopeId" in dependency && dependency.point === "start") {
        const source = scopeForPlanId(dependency.scopeId);
        if (source !== undefined) {
          return `scope-start:${source.declaration.identity}`;
        }
      }
      return dependencyEventId(dependency);
    }

    function markProjectedAncestors(
      actualParent: ScopeState,
      host: ScopeState,
    ): void {
      let projectionHost = host;
      const existingHostId = projectedEndHosts.get(
        actualParent.declaration.identity,
      );
      if (existingHostId === stepRootHostIdentity) return;
      const existingHost =
        existingHostId === undefined ? undefined : scopes.get(existingHostId);
      if (existingHost !== undefined && isAncestorScope(existingHost, host)) {
        projectionHost = existingHost;
      } else if (
        existingHost !== undefined &&
        !isAncestorScope(host, existingHost)
      ) {
        fail(
          "motion-plan-conflict",
          "Same logical completion cannot be projected to an unrelated physical host",
          actualParent.path,
        );
      }
      let current: ScopeState | undefined = actualParent;
      while (
        current !== undefined &&
        current.declaration.identity !== projectionHost.declaration.identity
      ) {
        projectedEndHosts.set(
          current.declaration.identity,
          projectionHost.declaration.identity,
        );
        const parentId: string | undefined = current.declaration.parent;
        current = parentId === undefined ? undefined : scopes.get(parentId);
      }
      if (current === undefined) {
        fail(
          "motion-plan-conflict",
          "Projection host is not an ancestor of consumer node",
          actualParent.path,
        );
      }
    }

    function markProjectedAncestorsAtStepRoot(actualParent: ScopeState): void {
      let current: ScopeState | undefined = actualParent;
      while (current !== undefined) {
        projectedEndHosts.set(
          current.declaration.identity,
          stepRootHostIdentity,
        );
        const parentId: string | undefined = current.declaration.parent;
        current = parentId === undefined ? undefined : scopes.get(parentId);
      }
    }

    const recordConditionPlacement = (
      join: Extract<MotionPlanStart, { readonly kind: "condition-join" }>,
    ): void => {
      const projectedHost = projectedNodeHosts.get(join.beforeNodeId);
      const placement =
        projectedHost === undefined
          ? join
          : {
              ...join,
              hostScopeId:
                projectedHost === stepRootHostIdentity
                  ? stepRootHostIdentity
                  : scopeId(projectedHost, scopes),
              consumerParentScopeId:
                projectedHost === stepRootHostIdentity
                  ? stepRootHostIdentity
                  : scopeId(projectedHost, scopes),
            };
      const placements =
        conditionPlacementsByHost.get(placement.hostScopeId) ?? [];
      const existingIndex = placements.findIndex(
        ({ join: existing }) => existing.id === placement.id,
      );
      if (existingIndex >= 0) {
        placements[existingIndex] = { join: placement };
        return;
      }
      placements.push({ join: placement });
      conditionPlacementsByHost.set(placement.hostScopeId, placements);
    };

    const recordStartCondition = (start: MotionPlanStart): void => {
      const { base } = motionPlanStartLayers(start);
      if (base.kind === "condition-join") {
        recordConditionPlacement(base);
      } else if (base.kind === "condition-end") {
        const join = conditionJoinsById.get(base.id);
        if (join !== undefined) recordConditionPlacement(join);
      }
    };

    const consumerStart = (start: MotionPlanStart): MotionPlanStart => {
      if (start.kind === "condition-join") {
        return { kind: "condition-end", id: start.id, delayMs: 0 };
      }
      if (start.kind === "delay-layer") {
        return { ...start, start: consumerStart(start.start) };
      }
      return start;
    };

    const conditionEvent = (id: string) => `condition:${id}`;
    const planNodeId = (node: MotionPlanNode): string =>
      node.kind === "scope"
        ? `scope:${node.id}`
        : node.kind === "completion-collector"
          ? `completion:${node.scopeId}`
          : `target:${node.targetId}`;
    const nodePath = (node: MotionPlanNode): string | undefined =>
      node.kind === "scope"
        ? [...scopes.values()].find(
            (scope) => scopeId(scope.declaration.identity, scopes) === node.id,
          )?.path
        : node.kind === "completion-collector"
          ? [...scopes.values()].find(
              (scope) =>
                scopeId(scope.declaration.identity, scopes) === node.scopeId,
            )?.path
          : targets.get(node.targetId)?.path;
    const sortNativeRoots = (
      rootNodes: readonly MotionPlanNode[],
      conditionDefinitionsBefore: readonly MotionPlanConditionPlacement[] = [],
      parentScope?: ScopeState,
    ): MotionPlanNode[] => {
      const rootDefinitions = new Map<string, number>();
      const rootReferences = rootNodes.map(
        (): { readonly id: string; readonly node: MotionPlanNode }[] => [],
      );
      const collectStartDependency = (
        start: MotionPlanStart,
        rootIndex: number,
        node: MotionPlanNode,
      ): void => {
        const collectDependency = (id: string) =>
          rootReferences[rootIndex]?.push({ id, node });
        if (start.kind === "delay-layer") {
          collectStartDependency(start.start, rootIndex, node);
        } else if (start.kind === "end" || start.kind === "event") {
          collectDependency(eventSource(start.dependency));
        } else if (start.kind === "condition-end") {
          collectDependency(conditionEvent(start.id));
        } else if (start.kind === "condition-join") {
          rootDefinitions.set(conditionEvent(start.id), rootIndex);
          for (const gate of start.gates) {
            collectStartDependency(planStartForGate(gate), rootIndex, node);
          }
        }
      };
      const collectRootDependencies = (
        node: MotionPlanNode,
        rootIndex: number,
      ): void => {
        if (node.kind === "scope") {
          rootDefinitions.set(`scope-start:${node.id}`, rootIndex);
          if (node.splitCompletion !== true) {
            rootDefinitions.set(`scope:${node.id}`, rootIndex);
          }
        } else if (node.kind === "completion-collector") {
          rootDefinitions.set(`scope:${node.scopeId}`, rootIndex);
          for (const gate of node.gates) {
            collectStartDependency(planStartForGate(gate), rootIndex, node);
          }
        } else {
          if (node.effects.length > 1) {
            rootDefinitions.set(`target:${node.targetId}`, rootIndex);
          }
          for (const effect of node.effects) {
            rootDefinitions.set(
              eventSource({
                targetId: effect.target.id,
                effectIndex: effect.effectIndex,
                ...(effect.partIndex === undefined
                  ? {}
                  : { partIndex: effect.partIndex }),
                point: "end",
                offsetMs: 0,
              }),
              rootIndex,
            );
          }
        }
        collectStartDependency(node.nativeStart, rootIndex, node);
        if (node.kind === "scope") {
          for (const placement of node.conditionDefinitionsBefore ?? []) {
            rootDefinitions.set(conditionEvent(placement.join.id), rootIndex);
            for (const gate of placement.join.gates) {
              collectStartDependency(planStartForGate(gate), rootIndex, node);
            }
          }
          for (const child of node.children) {
            collectRootDependencies(child, rootIndex);
          }
        }
      };
      rootNodes.forEach((node, index) => collectRootDependencies(node, index));

      for (const placement of conditionDefinitionsBefore) {
        const consumerIndex = rootNodes.findIndex(
          (node) => planNodeId(node) === placement.join.beforeNodeId,
        );
        if (consumerIndex < 0) continue;
        rootDefinitions.set(conditionEvent(placement.join.id), consumerIndex);
        for (const gate of placement.join.gates) {
          const consumer = rootNodes[consumerIndex];
          if (consumer !== undefined) {
            collectStartDependency(
              planStartForGate(gate),
              consumerIndex,
              consumer,
            );
          }
        }
      }

      // Adjust only serialization order of the same group of native timeline nodes; dependency conditions remain unchanged.
      const nextRootNodes = rootNodes.map(() => new Set<number>());
      const rootIndegrees = rootNodes.map(() => 0);
      for (const [consumerIndex, references] of rootReferences.entries()) {
        for (const reference of references) {
          const definitionIndex = rootDefinitions.get(reference.id);
          // Source may belong to outer root node of the same step; handled by outer sort and final order check.
          if (definitionIndex === undefined) continue;
          if (
            definitionIndex !== consumerIndex &&
            !nextRootNodes[definitionIndex]?.has(consumerIndex)
          ) {
            nextRootNodes[definitionIndex]?.add(consumerIndex);
            rootIndegrees[consumerIndex] =
              (rootIndegrees[consumerIndex] ?? 0) + 1;
          }
        }
      }
      const readyRootNodes = rootIndegrees.flatMap((indegree, index) =>
        indegree === 0 ? [index] : [],
      );
      const rootOrder: number[] = [];
      while (readyRootNodes.length > 0) {
        readyRootNodes.sort((left, right) => left - right);
        const current = readyRootNodes.shift();
        if (current === undefined) continue;
        rootOrder.push(current);
        for (const following of nextRootNodes[current] ?? []) {
          const remaining = (rootIndegrees[following] ?? 0) - 1;
          rootIndegrees[following] = remaining;
          if (remaining === 0) readyRootNodes.push(following);
        }
      }
      if (rootOrder.length !== rootNodes.length) {
        const reachableFrom = (start: number): Set<number> => {
          const reached = new Set<number>();
          const pending = [...(nextRootNodes[start] ?? [])];
          while (pending.length > 0) {
            const current = pending.pop();
            if (current === undefined || reached.has(current)) continue;
            reached.add(current);
            pending.push(...(nextRootNodes[current] ?? []));
          }
          return reached;
        };
        const reachable = rootNodes.map((_, index) => reachableFrom(index));
        const cycleRootIndices = new Set<number>();
        for (let left = 0; left < rootNodes.length; left += 1) {
          for (let right = left + 1; right < rootNodes.length; right += 1) {
            if (reachable[left]?.has(right) && reachable[right]?.has(left)) {
              cycleRootIndices.add(left);
              cycleRootIndices.add(right);
            }
          }
        }
        const placementsIn = (
          node: MotionPlanNode,
        ): MotionPlanConditionPlacement[] =>
          node.kind === "scope"
            ? [
                ...(node.conditionDefinitionsBefore ?? []),
                ...node.children.flatMap(placementsIn),
              ]
            : [];
        const parentForConsumer = (
          consumerId: string,
        ): ScopeState | undefined => {
          if (consumerId.startsWith("scope:")) {
            const consumerScope = scopeForPlanId(
              consumerId.slice("scope:".length),
            );
            const parentId = consumerScope?.declaration.parent;
            return parentId === undefined ? undefined : scopes.get(parentId);
          }
          if (consumerId.startsWith("target:")) {
            const target = targets.get(
              Number(consumerId.slice("target:".length)),
            );
            const parentId = target?.value.scopes.at(-1)?.identity;
            return parentId === undefined ? undefined : scopes.get(parentId);
          }
          return undefined;
        };
        if (parentScope !== undefined) {
          for (const rootIndex of cycleRootIndices) {
            const root = rootNodes[rootIndex];
            if (root === undefined) continue;
            for (const { join } of placementsIn(root)) {
              const hasCrossingSource = join.gates.some((gate) => {
                return gateSourceDependencies(gate).some((dependency) => {
                  const sourceRoot = rootDefinitions.get(
                    eventSource(dependency),
                  );
                  return (
                    sourceRoot !== undefined &&
                    sourceRoot !== rootIndex &&
                    cycleRootIndices.has(sourceRoot)
                  );
                });
              });
              if (!hasCrossingSource) continue;
              const consumerId = join.beforeNodeId;
              const actualParent = parentForConsumer(consumerId);
              if (
                actualParent === undefined ||
                !isAncestorScope(parentScope, actualParent)
              ) {
                continue;
              }
              const existingHost = projectedNodeHosts.get(consumerId);
              if (existingHost === parentScope.declaration.identity) continue;
              projectedNodeHosts.set(
                consumerId,
                parentScope.declaration.identity,
              );
              markProjectedAncestors(actualParent, parentScope);
              projectionChanged = true;
            }
          }
        }
        if (projectionChanged) return [...rootNodes];
        fail(
          "unknown-motion-start",
          "Node cycles that cannot be defined by real timeline still exist after recursive projection",
          parentScope?.path,
        );
      }
      return rootOrder.length === rootNodes.length
        ? rootOrder.flatMap((index) => {
            const node = rootNodes[index];
            return node === undefined ? [] : [node];
          })
        : [...rootNodes];
    };

    const projectedNodesByHost = new Map<string, MotionPlanNode[]>();

    const collectorForProjectedScope = (
      scope: ScopeState,
      hostStart: TimePoint,
    ): MotionPlanNode | undefined => {
      const end = ends.get(scope.declaration.identity);
      const memberEnds = scopeCompletionMembers.get(scope.declaration.identity);
      if (
        end === undefined ||
        end.state === "no-natural-completion" ||
        memberEnds === undefined ||
        memberEnds.length !== activeChildren(scope).length ||
        hostStart.state === "no-natural-completion"
      ) {
        return undefined;
      }
      const gates = memberEnds.map((memberEnd): MotionPlanStartGate => {
        if (memberEnd.state === "no-natural-completion") {
          fail(
            "motion-plan-conflict",
            "Finite scope completion collection contains member without natural completion",
            scope.path,
          );
        }
        let dependency = memberEnd.dependency;
        if (dependency === undefined) {
          const startSource = scopeForStartEvent(memberEnd.eventId);
          if (startSource !== undefined) {
            dependency = {
              scopeId: scopeId(startSource.declaration.identity, scopes),
              point: "start",
              offsetMs: 0,
            };
          }
        }
        const exactOffset =
          facts.exactOffset(hostStart.eventId, memberEnd.eventId) ??
          (hostStart.state === "known" && memberEnd.state === "known"
            ? memberEnd.timeMs - hostStart.timeMs
            : undefined);
        if (exactOffset !== undefined && exactOffset >= 0) {
          if (exactOffset <= maximumTimeNodeValue) {
            return {
              kind: "parent-start",
              delayMs: exactOffset,
              ...(dependency === undefined
                ? {}
                : { sourceDependency: dependency }),
            };
          }
          if (dependency === undefined) {
            fail(
              "invalid-motion-time",
              "Exact member start of completion collection cannot be preserved as a legal delay layer",
              scope.path,
            );
          }
        }
        if (
          dependency === undefined ||
          !strictlyAfterParentStart(memberEnd.candidates, hostStart)
        ) {
          fail(
            "unknown-motion-start",
            "Cannot prove real member END is later than early host activation and can be collected entirely",
            scope.path,
          );
        }
        return {
          kind: dependency.point === "start" ? "begin" : "end",
          dependency,
        };
      });
      return {
        kind: "completion-collector",
        scopeId: scopeId(scope.declaration.identity, scopes),
        // Each gate references complete END of original direct member; physical host END does not enter this collection.
        nativeStart: { kind: "parent-start", delayMs: 0 },
        gates,
      };
    };

    const scopeNode = (scope: ScopeState): MotionPlanNode[] => {
      const children = activeChildren(scope).flatMap((member) => {
        if (member.kind === "scope") {
          const childScope = scopes.get(member.id);
          if (childScope === undefined) {
            fail(
              "missing-scope",
              "Member declaration of choreography scope does not exist",
              scope.path,
            );
          }
          return scopeNode(childScope);
        }
        const target = targets.get(member.id);
        if (target === undefined) {
          fail(
            "invalid-motion",
            "Object member declaration of choreography scope does not exist",
            scope.path,
          );
        }
        const node = targetNode(target, scope);
        if (node === undefined) return [];
        const projectedHost = projectedNodeHosts.get(`target:${target.id}`);
        if (
          projectedHost !== undefined &&
          projectedHost !== scope.declaration.identity
        ) {
          const lifted = projectedNodesByHost.get(projectedHost) ?? [];
          lifted.push(node);
          projectedNodesByHost.set(projectedHost, lifted);
          return [];
        }
        return [node];
      });
      const identity = scope.declaration.identity;
      const hostIdentity = projectedNodeHosts.get(
        `scope:${scopeId(identity, scopes)}`,
      );
      if (hostIdentity !== undefined && hostIdentity !== identity) {
        const lifted = projectedNodesByHost.get(hostIdentity) ?? [];
        const node = createScopeNode(scope, children);
        lifted.push(node);
        projectedNodesByHost.set(hostIdentity, lifted);
        return [];
      }

      const hostLifted = projectedNodesByHost.get(identity) ?? [];
      const hostScopes = [...projectedEndHosts.entries()]
        .filter(([, host]) => host === identity)
        .map(([scopeIdentity]) => scopes.get(scopeIdentity))
        .filter((candidate): candidate is ScopeState => candidate !== undefined)
        .sort((left, right) => left.order - right.order);
      for (const projectedScope of hostScopes) {
        const collector = collectorForProjectedScope(
          projectedScope,
          resolveScopeStart(scope),
        );
        if (collector !== undefined) hostLifted.push(collector);
      }
      const allChildren = [...children, ...hostLifted];
      projectedNodesByHost.delete(identity);
      const preserveStart = preservedScopeStarts.has(identity);
      if (allChildren.length === 0 && !preserveStart) return [];
      const node = createScopeNode(scope, allChildren, preserveStart);
      return [node];
    };

    const createScopeNode = (
      scope: ScopeState,
      children: readonly MotionPlanNode[],
      preserveStart = preservedScopeStarts.has(scope.declaration.identity),
    ): MotionPlanNode => {
      const nativeStart = scopeStartPlans.get(scope.declaration.identity);
      if (nativeStart === undefined) {
        fail("motion-plan-conflict", "Scope lacks resolved start", scope.path);
      }
      recordStartCondition(nativeStart);
      const id = scopeId(scope.declaration.identity, scopes);
      const conditionDefinitionsBefore =
        conditionPlacementsByHost.get(id) ?? [];
      return {
        kind: "scope",
        id,
        nativeStart: consumerStart(nativeStart),
        children: sortNativeRoots(children, conditionDefinitionsBefore, scope),
        ...(projectedEndHosts.has(scope.declaration.identity)
          ? { splitCompletion: true as const }
          : {}),
        ...(preserveStart ? { preserveStart: true as const } : {}),
        ...(conditionDefinitionsBefore.length === 0
          ? {}
          : { conditionDefinitionsBefore }),
      };
    };

    let rootNodes: MotionPlanNode[] = [];
    let nodes: MotionPlanNode[] = [];
    let rootConditionDefinitionsBefore: MotionPlanConditionPlacement[] = [];
    while (true) {
      projectionChanged = false;
      conditionPlacementsByHost.clear();
      projectedNodesByHost.clear();
      rootNodes = roots.flatMap((member) => {
        if (member.kind === "scope") {
          const scope = scopes.get(member.id);
          if (scope === undefined || !scopeSteps(scope).includes(key))
            return [];
          return scopeNode(scope);
        }
        const target = targets.get(member.id);
        if (target === undefined || target.stepKey !== key) return [];
        const node = targetNode(target);
        return node === undefined ? [] : [node];
      });
      if (projectionChanged) continue;
      rootNodes.push(...(projectedNodesByHost.get(stepRootHostIdentity) ?? []));
      projectedNodesByHost.delete(stepRootHostIdentity);
      const projectedRootScopes = [...projectedEndHosts.entries()]
        .filter(([, host]) => host === stepRootHostIdentity)
        .map(([scopeIdentity]) => scopes.get(scopeIdentity))
        .filter(
          (candidate): candidate is ScopeState =>
            candidate !== undefined && scopeSteps(candidate).includes(key),
        )
        .sort((left, right) => left.order - right.order);
      for (const projectedScope of projectedRootScopes) {
        const collector = collectorForProjectedScope(projectedScope, origin);
        if (collector !== undefined) rootNodes.push(collector);
      }
      rootConditionDefinitionsBefore =
        conditionPlacementsByHost.get(stepRootHostIdentity) ?? [];
      nodes = sortNativeRoots(rootNodes, rootConditionDefinitionsBefore);
      if (!projectionChanged) break;
    }
    const definedEvents = new Set<string>();
    const definedConditions = new Set<string>();
    const verifyEventReference = (id: string, node: MotionPlanNode): void => {
      if (!definedEvents.has(id)) {
        fail(
          "unknown-motion-start",
          `Native completion source ${id} must be defined before its condition reference`,
          nodePath(node),
        );
      }
    };
    const verifyConditionJoins = (
      start: MotionPlanStart,
      node: MotionPlanNode,
    ): void => {
      if (start.kind === "delay-layer") {
        verifyConditionJoins(start.start, node);
      } else if (start.kind === "condition-join") {
        for (const gate of start.gates) {
          verifyGate(gate, node);
        }
        definedConditions.add(conditionEvent(start.id));
      }
    };
    const verifyStartDependencies = (
      start: MotionPlanStart,
      node: MotionPlanNode,
    ): void => {
      if (start.kind === "delay-layer") {
        verifyStartDependencies(start.start, node);
      } else if (start.kind === "end" || start.kind === "event") {
        verifyEventReference(eventSource(start.dependency), node);
      } else if (
        start.kind === "condition-end" &&
        !definedConditions.has(conditionEvent(start.id))
      ) {
        fail(
          "unknown-motion-start",
          "Condition collection must be defined before its consumers",
          nodePath(node),
        );
      }
    };
    const verifyGate = (
      gate: MotionPlanStartGate,
      node: MotionPlanNode,
    ): void => {
      const start = planStartForGate(gate);
      verifyConditionJoins(start, node);
      verifyStartDependencies(start, node);
    };
    const verifyNodeOrder = (node: MotionPlanNode): void => {
      verifyConditionJoins(node.nativeStart, node);
      if (node.kind === "completion-collector") {
        for (const gate of node.gates) {
          verifyGate(gate, node);
        }
        const eventId = `scope:${node.scopeId}`;
        if (definedEvents.has(eventId)) {
          fail(
            "motion-plan-conflict",
            "Same scope repeatedly defined completion event",
            nodePath(node),
          );
        }
        definedEvents.add(eventId);
      } else if (node.kind === "scope") {
        definedEvents.add(`scope-start:${node.id}`);
        if (node.splitCompletion !== true) {
          definedEvents.add(`scope:${node.id}`);
        }
      } else {
        if (node.effects.length > 1) {
          definedEvents.add(`target:${node.targetId}`);
        }
        for (const effect of node.effects) {
          definedEvents.add(
            eventSource({
              targetId: effect.target.id,
              effectIndex: effect.effectIndex,
              ...(effect.partIndex === undefined
                ? {}
                : { partIndex: effect.partIndex }),
              point: "end",
              offsetMs: 0,
            }),
          );
        }
      }
      verifyStartDependencies(node.nativeStart, node);
      if (node.kind === "scope") {
        const placements = node.conditionDefinitionsBefore ?? [];
        const remaining = new Set(placements);
        for (const child of node.children) {
          const childId = planNodeId(child);
          for (const placement of placements) {
            if (placement.join.beforeNodeId !== childId) continue;
            for (const gate of placement.join.gates) {
              verifyGate(gate, node);
            }
            if (definedConditions.has(conditionEvent(placement.join.id))) {
              fail(
                "motion-plan-conflict",
                "Same step repeatedly defined condition collection",
                nodePath(node),
              );
            }
            definedConditions.add(conditionEvent(placement.join.id));
            remaining.delete(placement);
          }
          verifyNodeOrder(child);
        }
        if (remaining.size > 0) {
          fail(
            "motion-plan-conflict",
            "Condition collection host lacks actual insertion position",
            nodePath(node),
          );
        }
      }
    };
    const remainingRootPlacements = new Set(rootConditionDefinitionsBefore);
    for (const node of nodes) {
      const nodeId = planNodeId(node);
      for (const placement of rootConditionDefinitionsBefore) {
        if (placement.join.beforeNodeId !== nodeId) continue;
        for (const gate of placement.join.gates) {
          verifyGate(gate, node);
        }
        if (definedConditions.has(conditionEvent(placement.join.id))) {
          fail(
            "motion-plan-conflict",
            "Same step repeatedly defined condition collection",
            nodePath(node),
          );
        }
        definedConditions.add(conditionEvent(placement.join.id));
        remainingRootPlacements.delete(placement);
      }
      verifyNodeOrder(node);
    }
    if (remainingRootPlacements.size > 0) {
      fail(
        "motion-plan-conflict",
        "Root with-group lacks actual consumer node for condition collection",
      );
    }
    // Resolves boundaries of missing ancestors only after native plan is complete; does not participate in schedule generation or writing.
    const queryScopesByIdentity = new Map(
      queryScopes.map((scope) => [scope.identity, scope] as const),
    );
    for (const scope of scopes.values()) {
      const identity = scope.declaration.identity;
      if (
        !queryScopeIdentities.has(identity) ||
        queryScopesByIdentity.has(identity)
      ) {
        continue;
      }
      const start = resolveScopeStart(scope);
      const end = resolveScopeEnd(scope);
      queryScopesByIdentity.set(identity, {
        identity,
        id: scopeId(identity, scopes),
        ...(scope.declaration.parent === undefined
          ? {}
          : { parentIdentity: scope.declaration.parent }),
        mode: scope.declaration.mode,
        ...(scope.declaration.each === undefined
          ? {}
          : { eachMs: Math.round(scope.declaration.each) }),
        delayMs: Math.round(scope.declaration.delay),
        members: queryDeclarationMembers(scope).map((member) =>
          member.kind === "scope"
            ? { kind: "scope" as const, identity: member.id }
            : { kind: "object" as const, targetId: member.id },
        ),
        start: queryMomentSource(start),
        end: queryMomentSource(end),
        declaration: scope.declaration,
      });
    }
    queryScopes.length = 0;
    for (const scope of scopes.values()) {
      const queryScope = queryScopesByIdentity.get(scope.declaration.identity);
      if (queryScope !== undefined) queryScopes.push(queryScope);
    }
    allStepNodes.set(key, nodes);
    const queryCues: MotionQueryCueSource[] = [];
    for (const [cueKey, publisher] of cuePublishers) {
      const publisherScope = scopes.get(publisher.scope);
      if (
        publisherScope === undefined ||
        !scopeSteps(publisherScope).includes(key)
      ) {
        continue;
      }
      const declaration =
        publisher.point === "start"
          ? publisherScope.declaration.publishStart
          : publisherScope.declaration.publishEnd;
      if (declaration?.point !== "named" || declaration.name === undefined) {
        continue;
      }
      queryCues.push({
        key: cueKey,
        name: declaration.name,
        scopeIdentity: publisher.scope,
        point: publisher.point,
      });
    }
    querySourcesByStep.set(key, {
      key,
      trigger: step.trigger,
      ...(step.number === undefined ? {} : { number: step.number }),
      facts: {
        exactOffset: (from, to) => facts.exactOffset(from, to),
        lowerBound: (from, to) => facts.lowerBound(from, to),
      },
      effectSources: queryEffectSources,
      scopes: queryScopes,
      cues: queryCues,
    });
    allStepConditionDefinitions.set(key, rootConditionDefinitionsBefore);
  }

  const plans = [...stepStates.values()]
    .sort((left, right) => {
      if (left.trigger !== right.trigger)
        return left.trigger === "auto" ? -1 : 1;
      return left.order - right.order;
    })
    .map((step) => {
      const effects = (allScheduledEffects.get(step.key) ?? [])
        .slice()
        .sort(
          (left, right) =>
            (left.startMs ?? Number.MAX_SAFE_INTEGER) -
              (right.startMs ?? Number.MAX_SAFE_INTEGER) ||
            left.target.id - right.target.id ||
            left.order - right.order,
        );
      const conditionDefinitionsBefore =
        allStepConditionDefinitions.get(step.key) ?? [];
      return {
        key: step.key,
        trigger: step.trigger,
        ...(step.number === undefined ? {} : { number: step.number }),
        order: step.order,
        effects,
        nodes: allStepNodes.get(step.key) ?? [],
        ...(conditionDefinitionsBefore.length === 0
          ? {}
          : { conditionDefinitionsBefore }),
      } satisfies MotionPlanStep;
    });

  const querySources: MotionQueryStepSource[] = plans.map((step) => {
    const source = querySourcesByStep.get(step.key);
    if (source === undefined)
      throw new Error("MotionPlan query step is missing");
    const effects = step.effects.map((effect) => {
      const effectSource = source.effectSources.get(
        motionEffectKey(effect.target.id, effect.effectIndex, effect.partIndex),
      );
      if (effectSource === undefined) {
        throw new Error("MotionPlan query effect is missing");
      }
      return effectSource;
    });
    return {
      key: source.key,
      trigger: source.trigger,
      ...(source.number === undefined ? {} : { number: source.number }),
      facts: source.facts,
      effects,
      scopes: source.scopes,
      cues: source.cues,
    };
  });

  const scopesObservation = scopeObservations.sort((left, right) => {
    const leftScope = [...scopes.values()].find(
      (scope) => scopeId(scope.declaration.identity, scopes) === left.id,
    );
    const rightScope = [...scopes.values()].find(
      (scope) => scopeId(scope.declaration.identity, scopes) === right.id,
    );
    return (leftScope?.order ?? 0) - (rightScope?.order ?? 0);
  });
  const observation: MotionObservation = Object.freeze({
    steps: Object.freeze(
      plans.map(({ effects, trigger, number }) => ({
        trigger,
        ...(number === undefined ? {} : { number }),
        effects: Object.freeze(
          effects.map(
            ({
              stepKey: _stepKey,
              order: _order,
              targetTextBody: _targetTextBody,
              targetFilled: _targetFilled,
              targetBox: _targetBox,
              targetRotationDegrees: _targetRotationDegrees,
              targetHasBackground: _targetHasBackground,
              targetParagraphCount: _targetParagraphCount,
              targetMembers: _targetMembers,
              effectIndex: _effectIndex,
              partIndex: _partIndex,
              iterate: _iterate,
              ...effect
            }) => Object.freeze(effect),
          ),
        ),
      })),
    ),
    ...(scopesObservation.length === 0
      ? {}
      : { scopes: Object.freeze(scopesObservation) }),
  });

  const query = createMotionQuery(querySources);
  return { steps: plans, observation, query };
}

function queryMomentSource(time: TimePoint): MotionQueryMomentSource {
  if (time.state === "no-natural-completion") {
    return { time: { state: "no-natural-completion" } };
  }
  return {
    eventId: time.eventId,
    time:
      time.state === "known"
        ? { state: "known", ms: time.timeMs }
        : { state: "unknown" },
  };
}

function motionEffectKey(
  targetId: number,
  effectIndex: number,
  partIndex: number | undefined,
): string {
  return `${targetId}:${effectIndex}:${partIndex ?? ""}`;
}

function scopeId(
  identity: string,
  scopes: ReadonlyMap<string, ScopeState>,
): string {
  const ordered = [...scopes.values()].sort(
    (left, right) => left.order - right.order,
  );
  const index = ordered.findIndex(
    (scope) => scope.declaration.identity === identity,
  );
  return `scope-${index + 1}`;
}

function normalizedEasing(
  easing: Effect["easing"],
): Effect["easing"] | undefined {
  if (easing === undefined) return undefined;
  const accelerate =
    easing.accelerate === undefined
      ? undefined
      : Math.round(easing.accelerate * 100000) / 100000;
  const decelerate =
    easing.decelerate === undefined
      ? undefined
      : Math.round(easing.decelerate * 100000) / 100000;
  if (
    (accelerate === undefined || accelerate === 0) &&
    (decelerate === undefined || decelerate === 0)
  ) {
    return undefined;
  }
  return Object.freeze({
    ...(accelerate === undefined || accelerate === 0 ? {} : { accelerate }),
    ...(decelerate === undefined || decelerate === 0 ? {} : { decelerate }),
  });
}

export function compileMotionPlan(
  contributions: readonly CapabilityContribution[],
): MotionPlan {
  return motionPlan(contributions);
}
