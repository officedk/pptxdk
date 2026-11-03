import { CapabilityError, type CapabilityContribution } from "@react-pptx/core";

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
} from "./motion-api.ts";

type StepDeclaration =
  | { readonly mode: "automatic"; readonly identity: string }
  | { readonly mode: "explicit"; readonly number: number };

type MotionTarget = {
  readonly kind: string;
  readonly name: string;
  readonly testId?: string;
  readonly textBody: boolean;
  readonly hasTextCharacters?: boolean;
  readonly filled: boolean;
  readonly paragraphCount?: number;
  readonly chartSeriesCount?: number;
  readonly chartCategoryCount?: number;
};

type MotionValue = {
  readonly target: MotionTarget;
  readonly step?: StepDeclaration;
  readonly scopes: readonly MotionScopeDeclaration[];
  readonly effects: readonly Effect[];
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
};

/** 真实时序包含树；scope 和多效果对象节点会成为原生并行容器。 */
export type MotionPlanNode =
  | {
      readonly kind: "scope";
      readonly id: string;
      readonly nativeStart: MotionPlanStart;
      readonly children: readonly MotionPlanNode[];
    }
  | {
      readonly kind: "target";
      readonly targetId: number;
      readonly nativeStart: MotionPlanStart;
      readonly effects: readonly ScheduledEffect[];
    };

/** 已求解的起点形式；PPTX 写出只负责序列化，不再判断因果。 */
export type MotionPlanStart =
  | { readonly kind: "parent-start"; readonly delayMs: number }
  | { readonly kind: "end"; readonly dependency: MotionStartDependency };

/** The single slide-local result consumed by both the writer and build report. */
export type MotionPlan = {
  readonly steps: readonly MotionPlanStep[];
  readonly observation: MotionObservation;
};

type TimeCandidate = {
  readonly eventId: string;
  readonly offsetMs: number;
  readonly dependency?: MotionStartDependency;
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
    }
  | {
      readonly state: "unknown";
      readonly candidates: readonly TimeCandidate[];
      readonly eventId: string;
      readonly dependency?: MotionStartDependency;
      readonly nativeStart?: MotionPlanStart;
      readonly reportedStartAfter?: MotionStartDependency;
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
    }
  | { readonly state: "unknown"; readonly candidates: readonly TimeCandidate[] }
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
  if ("scopeId" in event) return `scope:${event.scopeId}`;
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
): TimePoint {
  const candidate: TimeCandidate = {
    eventId,
    offsetMs: 0,
    ...(dependency === undefined
      ? {}
      : { dependency: dependencyAtEvent(dependency) }),
  };
  if (state === "known") {
    if (timeMs === undefined) throw new Error("已知时刻缺少数值");
    return {
      state,
      timeMs,
      eventId,
      candidates: [candidate],
      ...(dependency === undefined ? {} : { dependency }),
      ...(nativeStart === undefined ? {} : { nativeStart }),
      ...(reportedStartAfter === undefined ? {} : { reportedStartAfter }),
    };
  }
  return {
    state,
    eventId,
    candidates: [candidate],
    ...(dependency === undefined ? {} : { dependency }),
    ...(nativeStart === undefined ? {} : { nativeStart }),
    ...(reportedStartAfter === undefined ? {} : { reportedStartAfter }),
  };
}

class MotionTimingFacts {
  private readonly successors = new Map<string, Map<string, number>>();
  // 精确时差可正反推导；反向关系不进入必要前置图的环检查。
  private readonly exactRelations = new Map<string, Map<string, number>>();
  private readonly exactTimes = new Map<string, number>();

  setExactTime(eventId: string, timeMs: number, path: string): void {
    const existing = this.exactTimes.get(eventId);
    if (existing !== undefined && existing !== timeMs) {
      fail("motion-plan-conflict", "同一编排事件得到冲突的数值时刻", path);
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
      fail("motion-plan-conflict", "精确编排时差与已知时刻冲突", path);
    }
    const existing = this.exactRelations.get(before)?.get(after);
    if (existing !== undefined && existing !== offsetMs) {
      fail("motion-plan-conflict", "同一编排事件得到冲突的精确时差", path);
    }
    const existingReverse = this.exactRelations.get(after)?.get(before);
    if (existingReverse !== undefined && existingReverse !== -offsetMs) {
      fail("motion-plan-conflict", "同一编排事件得到冲突的精确时差", path);
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
      fail("motion-plan-conflict", "精确编排时差与已知时刻冲突");
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
        fail("motion-cycle", "编排时刻依赖存在循环", path);
      }
      return;
    }
    if (this.hasPath(after, before)) {
      fail("motion-cycle", "编排时刻依赖存在循环", path);
    }
    const beforeTime = this.exactTime(before);
    const afterTime = this.exactTime(after);
    if (
      beforeTime !== undefined &&
      afterTime !== undefined &&
      afterTime - beforeTime < minimumMs
    ) {
      fail("motion-plan-conflict", "编排时差事实与已知时刻冲突", path);
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
      fail("motion-plan-conflict", "编排时差事实与精确时差冲突");
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
            fail("motion-plan-conflict", "同一编排时刻存在冲突的精确关系");
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
  return time.state === "known"
    ? {
        state: "known",
        timeMs: time.timeMs,
        candidates: time.candidates,
      }
    : { state: "unknown", candidates: time.candidates };
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
  return time.state === "known"
    ? { state: "known", timeMs: time.timeMs + offsetMs, candidates }
    : { state: "unknown", candidates };
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
  return { state: "unknown", candidates: reduced };
}

function materializeEvent(
  expression: TimeExpression,
  eventId: string,
  dependency: MotionStartDependency | undefined,
  facts: MotionTimingFacts,
  path: string,
  nativeStart?: MotionPlanStart,
  reportedStartAfter?: MotionStartDependency,
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
    );
  }
  return eventTime(
    eventId,
    "unknown",
    undefined,
    dependency,
    nativeStart,
    reportedStartAfter,
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
      // 唯一胜出的真实完成事件与整体结束保持精确等价，绝对时刻仍可未知。
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
  if (effect.by === undefined) return { type: "single", count: 1 };
  if (effect.by === "word" || effect.by === "letter") {
    if (!target.textBody || !target.paragraphCount) {
      fail(
        "invalid-motion-granularity",
        `by: "${effect.by}" 需要含文字的原生对象`,
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
        'by: "paragraph" 需要含段落的原生文字对象',
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
      `by: "${effect.by}" 需要含数据的原生图表`,
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
  effect: Effect,
  effectIndex: number,
  expansion: EffectExpansion,
  facts: MotionTimingFacts,
): TimePoint {
  const partIndex =
    expansion.type === "parts" ? expansion.count - 1 : undefined;
  const dependency: MotionStartDependency = {
    targetId: target.id,
    effectIndex,
    ...(partIndex === undefined ? {} : { partIndex }),
    point: "end",
    offsetMs: 0,
  };
  const durationMs = Math.round(effect.duration);
  const startOffsetMs =
    Math.round(effect.delay) +
    (partIndex === undefined ? 0 : partIndex * durationMs);
  const startExpression = add(base, startOffsetMs);
  if (startExpression.state === "no-natural-completion") {
    return startExpression;
  }
  const startEventId = `effect-start:${target.id}:${effectIndex}:${partIndex ?? ""}`;
  const start = materializeEvent(
    startExpression,
    startEventId,
    undefined,
    facts,
    target.path,
  );
  if (start.state === "no-natural-completion") return start;

  const endEventId = dependencyEventId(dependency);
  const minimumDuration =
    expansion.type === "iterate" &&
    expansion.by === "letter" &&
    target.value.target.hasTextCharacters === true
      ? durationMs
      : expansion.type === "iterate"
        ? 0
        : durationMs;
  if (expansion.type === "iterate") {
    facts.addLowerBound(startEventId, endEventId, minimumDuration, target.path);
  } else {
    // 有限行为保留相对时差，不要求效果起点已有绝对数值。
    facts.addLowerBound(startEventId, endEventId, durationMs, target.path);
    facts.addExactOffset(startEventId, endEventId, durationMs, target.path);
  }
  if (expansion.type !== "iterate" && start.state === "known") {
    const endMs = start.timeMs + durationMs;
    facts.setExactTime(endEventId, endMs, target.path);
    return eventTime(endEventId, "known", endMs, dependency);
  }
  return eventTime(endEventId, "unknown", undefined, dependency);
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
      fail("invalid-motion", "动效效果索引不存在", target.path);
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
        fail("invalid-scope", "编排作用域声明无效", contribution.path);
      }
      if (scopes.has(declaration.identity)) {
        fail(
          "duplicate-scope",
          "同一编排作用域身份不能声明两次",
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
        `motion 注解在 ${contribution.path} 中无效`,
        contribution.path,
      );
    }
    if (contribution.targetId === undefined) {
      fail("invalid-motion", "对象动效缺少原生对象身份", contribution.path);
    }
    const trigger = stepInfo(contribution.value.step);
    if (targets.has(contribution.targetId)) {
      fail("invalid-motion", "同一原生对象重复提交动效声明", contribution.path);
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
        .filter((by): by is MotionGranularity => by !== undefined),
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
        "同一原生对象的并行动效不能混用不兼容的构建粒度",
        target.path,
      );
    }
  }

  for (const scope of scopes.values()) {
    const parentId = scope.declaration.parent;
    if (parentId === undefined) continue;
    const parent = scopes.get(parentId);
    if (parent === undefined) {
      fail("missing-scope", "编排作用域的父实例不存在", scope.path);
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
        fail("invalid-scope", "对象包含无效编排作用域路径", target.path);
      }
      const scope = scopes.get(declaration.identity);
      if (scope === undefined) {
        fail("missing-scope", "对象引用的编排作用域实例不存在", target.path);
      }
      if (scope.declaration.parent !== parent) {
        fail(
          "invalid-scope",
          "对象的编排作用域路径与宿主声明不一致",
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
        "跨多个点击步的作用域不能发布单一时刻",
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
        fail("duplicate-cue", "同一命名时刻不能发布两次", scope.path);
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
    const scopeStartPlans = new Map<string, MotionPlanStart>();
    const resolvingStarts = new Set<string>();
    const resolvingEnds = new Set<string>();
    const scheduledObjects = new Set<number>();

    const resolveNativeStart = (
      expression: TimeExpression,
      parentStart: TimePoint,
      path: string,
    ): MotionPlanStart => {
      if (expression.state === "no-natural-completion") {
        fail("no-natural-completion", "无自然完成的内容不能启动后继效果", path);
      }
      if (parentStart.state === "no-natural-completion") {
        fail(
          "no-natural-completion",
          "无自然完成的父作用域不能启动后继效果",
          path,
        );
      }
      if (expression.state === "known") {
        if (parentStart.state !== "known") {
          fail(
            "unknown-motion-start",
            "未知父起点后的数值起点没有可验证的合并关系",
            path,
          );
        }
        const delayMs = expression.timeMs - parentStart.timeMs;
        if (delayMs < 0) {
          fail("unknown-motion-start", "子项起点早于其包含作用域", path);
        }
        return { kind: "parent-start", delayMs };
      }

      if (expression.candidates.length !== 1) {
        fail(
          "unknown-motion-start",
          "多个不可比较的未知起点没有已核验的原生合并形式",
          path,
        );
      }
      const [candidate] = expression.candidates;
      if (candidate === undefined) {
        fail(
          "unknown-motion-start",
          "未知起点没有可引用的单个原生完成事件",
          path,
        );
      }

      if (candidate.eventId === parentStart.eventId) {
        return { kind: "parent-start", delayMs: candidate.offsetMs };
      }
      if (
        parentStart.nativeStart?.kind === "end" &&
        dependencyEventId(parentStart.nativeStart.dependency) ===
          candidate.eventId
      ) {
        const delayMs =
          candidate.offsetMs - parentStart.nativeStart.dependency.offsetMs;
        if (delayMs >= 0) return { kind: "parent-start", delayMs };
      }

      // 保留仍晚于父启动的真实完成事件；仅在它已先于父启动时改用精确偏移。
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
        ((eventAfterParent !== undefined && eventAfterParent >= 0) ||
          (eventAfterKnownParent !== undefined &&
            knownParentTime !== undefined &&
            eventAfterKnownParent >= knownParentTime))
      ) {
        return {
          kind: "end",
          dependency: dependencyAtEvent(
            candidate.dependency,
            candidate.offsetMs,
          ),
        };
      }
      const exactParentOffset = facts.exactOffset(
        candidate.eventId,
        parentStart.eventId,
      );
      if (
        exactParentOffset !== undefined &&
        candidate.offsetMs >= exactParentOffset
      ) {
        return {
          kind: "parent-start",
          delayMs: candidate.offsetMs - exactParentOffset,
        };
      }
      fail(
        "unknown-motion-start",
        "未知起点没有可证明的父级激活或完成事件来源",
        path,
      );
    };

    const reportedDependency = (
      nativeStart: MotionPlanStart,
      parentStart: TimePoint,
    ): MotionStartDependency | undefined => {
      if (nativeStart.kind === "end") return nativeStart.dependency;
      if (parentStart.state === "no-natural-completion") return undefined;
      const inherited = parentStart.reportedStartAfter;
      return inherited === undefined
        ? undefined
        : dependencyAtEvent(
            inherited,
            inherited.offsetMs + nativeStart.delayMs,
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
          "单一编排作用域的时刻不能跨点击步确定",
          scope.path,
        );
      }
    };

    const resolveMoment = (
      reference: MotionMoment,
      fromPath: string,
    ): TimePoint => {
      const owner = scopes.get(reference.scope);
      if (owner === undefined) {
        fail("missing-cue", "时刻句柄对应的编排作用域不存在", fromPath);
      }
      ensureSingleStep(owner);
      if (scopeSteps(owner)[0] !== key) {
        fail("cross-step-motion", "时刻依赖不能跨不同点击步", fromPath);
      }
      let sourceScope = owner;
      let point: "start" | "end" = reference.point === "end" ? "end" : "start";
      if (reference.point === "named") {
        const publisher = cuePublishers.get(momentKey(reference));
        if (publisher === undefined) {
          fail(
            "missing-cue",
            `没有组件发布命名时刻 ${reference.name ?? ""}`,
            fromPath,
          );
        }
        const publishedScope = scopes.get(publisher.scope);
        if (publishedScope === undefined) {
          fail("missing-cue", "命名时刻的发布作用域不存在", publisher.path);
        }
        ensureSingleStep(publishedScope);
        if (scopeSteps(publishedScope)[0] !== key) {
          fail("cross-step-motion", "命名时刻依赖不能跨不同点击步", fromPath);
        }
        sourceScope = publishedScope;
        point = publisher.point;
      }
      return point === "start"
        ? resolveScopeStart(sourceScope)
        : resolveScopeEnd(sourceScope);
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
          fail("missing-scope", "编排作用域的成员声明不存在", parent.path);
        }
        return resolveScopeEnd(child);
      }
      const target = targets.get(member.id);
      if (target === undefined) {
        fail("invalid-motion", "编排作用域的对象成员声明不存在", parent.path);
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
        fail("motion-cycle", "编排时刻依赖存在循环", scope.path);
      }
      resolvingStarts.add(scope.declaration.identity);
      let parentStart = origin;
      let base = pointExpression(origin);
      if (scope.declaration.parent !== undefined) {
        const parent = scopes.get(scope.declaration.parent);
        if (parent === undefined) {
          fail("missing-scope", "编排作用域的父实例不存在", scope.path);
        }
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
      const nativeStart = resolveNativeStart(
        startExpression,
        parentStart,
        scope.path,
      );
      const start = materializeEvent(
        startExpression,
        scopeStartEventId(scope),
        undefined,
        facts,
        scope.path,
        nativeStart,
        reportedDependency(nativeStart, parentStart),
      );
      scopeStartPlans.set(scope.declaration.identity, nativeStart);
      starts.set(scope.declaration.identity, start);
      resolvingStarts.delete(scope.declaration.identity);
      return start;
    };

    const resolveScopeEnd = (scope: ScopeState): TimePoint => {
      const cached = ends.get(scope.declaration.identity);
      if (cached !== undefined) return cached;
      if (resolvingEnds.has(scope.declaration.identity)) {
        fail("motion-cycle", "编排时刻依赖存在循环", scope.path);
      }
      resolvingEnds.add(scope.declaration.identity);
      const start = resolveScopeStart(scope);
      // 非活动作用域仍可提供自身边界时刻，但不延长父作用域的整体完成。
      const members = activeChildren(scope);
      const memberEnds = members.map((member) => memberEnd(scope, member));
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
          "顺序和交错作用域不能编排不同点击步",
          scope.path,
        );
      }
      if (
        scope.declaration.startAt !== undefined &&
        collectActiveSteps(scope.declaration.identity).size > 1
      ) {
        fail(
          "cross-step-motion",
          "跨多个点击步的作用域不能从单一时刻启动",
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
        ...(end.state === "known" ? { endMs: end.timeMs } : {}),
        completion: end.state,
        members: observedMembers,
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
              "无自然完成的内容不能启动后继效果",
              target.path,
            );
          }
          const nativeStart = resolveNativeStart(
            start,
            parentStart,
            target.path,
          );
          const startAfter =
            start.state === "unknown"
              ? reportedDependency(nativeStart, parentStart)
              : undefined;
          const knownStart = start.state === "known" ? start.timeMs : undefined;
          const endMs =
            expansion.type === "iterate" || knownStart === undefined
              ? undefined
              : knownStart + Math.round(effect.duration);
          const from =
            typeof effect.from === "number"
              ? Math.round(effect.from * 60000) / 60000
              : effect.from;
          const easing = normalizedEasing(effect.easing);
          addStepEffect({
            target: {
              id: target.id,
              name: target.value.target.name,
              ...(target.value.target.testId === undefined
                ? {}
                : { testId: target.value.target.testId }),
              ...("range" in entry ? { range: entry.range } : {}),
            },
            kind: effect.kind,
            preset: effect.preset,
            durationMs: Math.round(effect.duration),
            delayMs: Math.round(effect.delay),
            ...(effect.by === undefined ? {} : { by: effect.by }),
            ...(expansion.type === "iterate"
              ? { intervalMs: expansion.intervalMs }
              : {}),
            ...(from === undefined ? {} : { from }),
            ...(easing === undefined ? {} : { easing }),
            ...(knownStart === undefined ? {} : { startMs: knownStart }),
            ...(endMs === undefined
              ? { completion: "unknown" as const }
              : { endMs }),
            ...(startAfter === undefined ? {} : { startAfter }),
            stepKey: key,
            order: target.order + index,
            effectIndex: index,
            ...(partIndex === undefined ? {} : { partIndex }),
            targetTextBody: target.value.target.textBody,
            targetFilled: target.value.target.filled,
            ...(expansion.type === "iterate"
              ? {
                  iterate: {
                    type: expansion.by === "word" ? "wd" : "lt",
                    intervalMs: expansion.intervalMs,
                  },
                }
              : {}),
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
      return {
        kind: "target",
        targetId: target.id,
        nativeStart: resolveNativeStart(start, parentStart, target.path),
        effects,
      };
    };

    const scopeNode = (scope: ScopeState): MotionPlanNode | undefined => {
      const children = activeChildren(scope).flatMap((member) => {
        if (member.kind === "scope") {
          const childScope = scopes.get(member.id);
          if (childScope === undefined) {
            fail("missing-scope", "编排作用域的成员声明不存在", scope.path);
          }
          const node = scopeNode(childScope);
          return node === undefined ? [] : [node];
        }
        const target = targets.get(member.id);
        if (target === undefined) {
          fail("invalid-motion", "编排作用域的对象成员声明不存在", scope.path);
        }
        const node = targetNode(target, scope);
        return node === undefined ? [] : [node];
      });
      if (children.length === 0) return undefined;
      const nativeStart = scopeStartPlans.get(scope.declaration.identity);
      if (nativeStart === undefined) {
        fail("motion-plan-conflict", "作用域缺少已求解的起点", scope.path);
      }
      return {
        kind: "scope",
        id: scopeId(scope.declaration.identity, scopes),
        nativeStart,
        children,
      };
    };

    const nodes = roots.flatMap((member) => {
      if (member.kind === "scope") {
        const scope = scopes.get(member.id);
        if (scope === undefined || !scopeSteps(scope).includes(key)) return [];
        const node = scopeNode(scope);
        return node === undefined ? [] : [node];
      }
      const target = targets.get(member.id);
      if (target === undefined || target.stepKey !== key) return [];
      const node = targetNode(target);
      return node === undefined ? [] : [node];
    });
    allStepNodes.set(key, nodes);
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
      return {
        key: step.key,
        trigger: step.trigger,
        ...(step.number === undefined ? {} : { number: step.number }),
        order: step.order,
        effects,
        nodes: allStepNodes.get(step.key) ?? [],
      } satisfies MotionPlanStep;
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

  return { steps: plans, observation };
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
