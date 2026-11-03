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
      readonly startMs?: number;
      readonly startAfter?: MotionStartDependency;
      readonly children: readonly MotionPlanNode[];
    }
  | {
      readonly kind: "target";
      readonly targetId: number;
      readonly startMs?: number;
      readonly startAfter?: MotionStartDependency;
      readonly effects: readonly ScheduledEffect[];
    };

/** The single slide-local result consumed by both the writer and build report. */
export type MotionPlan = {
  readonly steps: readonly MotionPlanStep[];
  readonly observation: MotionObservation;
};

type Completion =
  | { readonly state: "known"; readonly timeMs: number }
  | { readonly state: "unknown"; readonly event?: MotionStartDependency }
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

function known(timeMs: number): TimePoint {
  return { state: "known", timeMs };
}

function eventSource(event: MotionStartDependency): string {
  if ("scopeId" in event) return `scope:${event.scopeId}`;
  if ("group" in event) return `target:${event.targetId}`;
  return `effect:${event.targetId}:${event.effectIndex}:${event.partIndex ?? ""}`;
}

function sameEvent(
  left: MotionStartDependency,
  right: MotionStartDependency,
): boolean {
  return eventSource(left) === eventSource(right);
}

function add(time: TimePoint, offsetMs: number): TimePoint {
  if (time.state === "known") return known(time.timeMs + offsetMs);
  if (time.state === "unknown" && time.event !== undefined) {
    return {
      state: "unknown",
      event: { ...time.event, offsetMs: time.event.offsetMs + offsetMs },
    };
  }
  return time;
}

function maximum(times: readonly TimePoint[]): TimePoint {
  if (times.some(({ state }) => state === "no-natural-completion")) {
    return { state: "no-natural-completion" };
  }
  const unknownTimes = times.filter(({ state }) => state === "unknown");
  if (unknownTimes.length > 0) {
    const knownTimes = times.filter(
      (time): time is { readonly state: "known"; readonly timeMs: number } =>
        time.state === "known",
    );
    const onlyZeroBoundary = knownTimes.every(({ timeMs }) => timeMs === 0);
    const [firstUnknown] = unknownTimes;
    const events = unknownTimes.map((time) =>
      time.state === "unknown" ? time.event : undefined,
    );
    const sharedEvent =
      events.length > 0 &&
      events.every(
        (event): event is MotionStartDependency =>
          event !== undefined &&
          firstUnknown?.state === "unknown" &&
          firstUnknown.event !== undefined &&
          sameEvent(event, firstUnknown.event),
      );
    if (sharedEvent && onlyZeroBoundary) {
      const event = events.reduce<MotionStartDependency | undefined>(
        (latest, current) =>
          latest === undefined || current.offsetMs > latest.offsetMs
            ? current
            : latest,
        undefined,
      );
      return { state: "unknown", ...(event === undefined ? {} : { event }) };
    }
    return {
      state: "unknown",
      ...(unknownTimes.length === 1 &&
      onlyZeroBoundary &&
      firstUnknown?.state === "unknown" &&
      firstUnknown.event !== undefined
        ? { event: firstUnknown.event }
        : {}),
    };
  }
  return known(
    Math.max(
      0,
      ...times.map((time) => (time.state === "known" ? time.timeMs : 0)),
    ),
  );
}

function completion(
  start: TimePoint,
  children: readonly TimePoint[],
  event: MotionStartDependency,
): TimePoint {
  if (
    start.state === "no-natural-completion" ||
    children.some(({ state }) => state === "no-natural-completion")
  ) {
    return { state: "no-natural-completion" };
  }
  if (children.length === 0) return start;
  const end = maximum([start, ...children]);
  return end.state === "unknown" ? { state: "unknown", event } : end;
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
  base: TimePoint,
  target: TargetState,
  effect: Effect,
  effectIndex: number,
  expansion: EffectExpansion,
): TimePoint {
  if (expansion.type === "iterate") {
    // 原生逐词/逐字节点会自然结束，但本计划不从单次行为时长估算整段时长。
    return {
      state: "unknown",
      event: {
        targetId: target.id,
        effectIndex,
        point: "end",
        offsetMs: 0,
      },
    };
  }
  const offset =
    Math.round(effect.delay) + expansion.count * Math.round(effect.duration);
  const end = add(base, offset);
  if (end.state === "known") return end;
  if (end.state === "no-natural-completion") return end;
  const partIndex =
    expansion.type === "parts" ? expansion.count - 1 : undefined;
  return {
    state: "unknown",
    event: {
      targetId: target.id,
      effectIndex,
      ...(partIndex === undefined ? {} : { partIndex }),
      point: "end",
      offsetMs: 0,
    },
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
    const starts = new Map<string, TimePoint>();
    const ends = new Map<string, TimePoint>();
    const resolvingStarts = new Set<string>();
    const resolvingEnds = new Set<string>();
    const scheduledObjects = new Set<number>();

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
    ): TimePoint => {
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
        ),
      );
      const animationNodeCount = target.value.effects.reduce(
        (count, _effect, effectIndex) => {
          const expansion = getExpansion(target, effectIndex);
          return count + (expansion.type === "parts" ? expansion.count : 1);
        },
        0,
      );
      const end =
        effectEnds.length === 1 ? (effectEnds[0] ?? base) : maximum(effectEnds);
      return end.state === "unknown" && animationNodeCount > 1
        ? {
            state: "unknown",
            event: {
              targetId: target.id,
              group: "effects",
              point: "end",
              offsetMs: 0,
            },
          }
        : end;
    };

    const memberStart = (
      parent: ScopeState,
      member: ScopeMember,
    ): TimePoint => {
      const parentStart = resolveScopeStart(parent);
      const active = activeChildren(parent);
      const activeIndex = active.findIndex(
        (candidate) =>
          candidate.kind === member.kind && candidate.id === member.id,
      );
      if (parent.declaration.mode === "parallel") return parentStart;
      if (parent.declaration.mode === "stagger") {
        if (activeIndex < 0) return parentStart;
        return add(parentStart, activeIndex * (parent.declaration.each ?? 0));
      }
      let cursor = parentStart;
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
        cursor = ended;
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
      let base: TimePoint = known(0);
      if (scope.declaration.parent !== undefined) {
        const parent = scopes.get(scope.declaration.parent);
        if (parent === undefined) {
          fail("missing-scope", "编排作用域的父实例不存在", scope.path);
        }
        base = memberStart(parent, {
          kind: "scope",
          id: scope.declaration.identity,
          order: scope.order,
        });
      }
      if (scope.declaration.startAt !== undefined) {
        const fromCue = resolveCueTime(scope.declaration.startAt, scope.path);
        base = maximum([base, fromCue]);
      }
      const start = add(base, scope.declaration.delay);
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
      const end = completion(start, memberEnds, {
        scopeId: scopeId(scope.declaration.identity, scopes),
        point: "end",
        offsetMs: 0,
      });
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
        ...(start.state === "unknown" && start.event !== undefined
          ? { startAfter: start.event }
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
      const base =
        containingScope === undefined
          ? known(0)
          : memberStart(scopes.get(containingScope)!, {
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
          if (start.state === "unknown" && start.event === undefined) {
            fail(
              "unknown-motion-start",
              "未知起点没有可引用的单个原生完成事件",
              target.path,
            );
          }
          const knownStart = start.state === "known" ? start.timeMs : undefined;
          const end =
            expansion.type === "iterate"
              ? { state: "unknown" as const }
              : knownStart === undefined
                ? { state: "unknown" as const }
                : known(knownStart + Math.round(effect.duration));
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
            ...(end.state === "known"
              ? { endMs: end.timeMs }
              : { completion: "unknown" as const }),
            ...(start.state === "unknown" && start.event !== undefined
              ? { startAfter: start.event }
              : {}),
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
      const start =
        parent === undefined
          ? known(0)
          : memberStart(parent, {
              kind: "object",
              id: target.id,
              order: target.order,
            });
      return {
        kind: "target",
        targetId: target.id,
        ...(start.state === "known" ? { startMs: start.timeMs } : {}),
        ...(start.state === "unknown" && start.event !== undefined
          ? { startAfter: start.event }
          : {}),
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
      const start = resolveScopeStart(scope);
      return {
        kind: "scope",
        id: scopeId(scope.declaration.identity, scopes),
        ...(start.state === "known" ? { startMs: start.timeMs } : {}),
        ...(start.state === "unknown" && start.event !== undefined
          ? { startAfter: start.event }
          : {}),
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
