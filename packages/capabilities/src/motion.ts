import {
  CapabilityError,
  xmlEscape,
  type CapabilityInput,
  type SlideCapabilityModule,
  type TransparentCapabilityModule,
} from "@react-pptx/core";

import type {
  MotionDirection,
  MotionEffectObservation,
  MotionObservation,
  MotionStartDependency,
} from "./motion-api.ts";
import { maximumTimeNodeValue, motionScopeAnnotation } from "./motion-api.ts";
import {
  composeMotionPlanStarts,
  compileMotionPlan,
  type MotionPlanConditionPlacement,
  type MotionPlanNode,
  type MotionPlanStart,
  type MotionPlanStartGate,
  type MotionPlanStep,
} from "./motion-plan.ts";

export {
  after,
  breathe,
  Box,
  MotionScope,
  Step,
  fade,
  fly,
  floatUp,
  repeat,
  spin,
  spinIn,
  stretch,
  useMotionScope,
  wipe,
  zoom,
} from "./motion-api.ts";
export type {
  BoxProps,
  Easing,
  Effect,
  EffectOptions,
  EffectPreset,
  BreatheEffect,
  BreatheOptions,
  EmphasisEffect,
  EmphasisEffectDeclaration,
  EntranceEffect,
  EntranceEffectDeclaration,
  MotionCompletionState,
  MotionDirection,
  MotionGranularity,
  MotionEffectObservation,
  MotionMoment,
  MotionMomentOffset,
  MotionObservation,
  MotionScopeHandle,
  MotionScopeMode,
  MotionScopeObservation,
  MotionScopeProps,
  MotionProps,
  MotionStartDependency,
  MotionTargetRange,
  MotionTiming,
  MotionStepObservation,
  RepeatCount,
  RepeatedEffect,
  StepProps,
} from "./motion-api.ts";

type XmlAttributes = Record<string, string | number | boolean | undefined>;

function invalidMotionTime(field: string): never {
  throw new CapabilityError(
    "invalid-motion-time",
    `${field} 必须是精确的非负整数毫秒，且实际写入时不得超过 ST_TLTime 范围`,
  );
}

// 累计总时长可由 repeat 表达；只有实际投影到 ST_TLTime 的字段受其上限约束。
function validateKnownTime(value: number | undefined, field: string): void {
  if (value !== undefined && (!Number.isSafeInteger(value) || value < 0)) {
    invalidMotionTime(field);
  }
}

function timeNodeValue(value: number, field: string): number {
  validateKnownTime(value, field);
  if (value > maximumTimeNodeValue) invalidMotionTime(field);
  return value;
}

function validateKnownMotionTimes(observation: MotionObservation): void {
  for (const scope of observation.scopes ?? []) {
    validateKnownTime(scope.delayMs, "编排作用域延迟");
    validateKnownTime(scope.eachMs, "交错间隔");
    validateKnownTime(scope.startMs, "编排作用域起点");
    validateKnownTime(scope.endMs, "编排作用域完成时刻");
    validateKnownTime(scope.startAfter?.offsetMs, "作用域完成依赖偏移");
    for (const dependency of scope.startAfterAll ?? []) {
      validateKnownTime(dependency.offsetMs, "作用域起点约束偏移");
    }
  }
  for (const step of observation.steps) {
    for (const effect of step.effects) {
      validateKnownTime(effect.durationMs, "动画时长");
      validateKnownTime(effect.delayMs, "动画延迟");
      validateKnownTime(effect.intervalMs, "动画迭代间隔");
      validateKnownTime(effect.startMs, "动画起点");
      validateKnownTime(effect.endMs, "动画完成时刻");
      validateKnownTime(effect.startAfter?.offsetMs, "动画完成依赖偏移");
      for (const dependency of effect.startAfterAll ?? []) {
        validateKnownTime(dependency.offsetMs, "动画起点约束偏移");
      }
    }
  }
}

function element(
  name: string,
  attributes: XmlAttributes | undefined,
  ...children: readonly (string | undefined)[]
): string {
  const attrs = Object.entries(attributes ?? {})
    .filter(([, value]) => value !== undefined)
    .map(([key, value]) => {
      const normalized =
        typeof value === "boolean" ? (value ? "1" : "0") : value;
      return ` ${key}="${xmlEscape(String(normalized))}"`;
    })
    .join("");
  return `<${name}${attrs}>${children.filter(Boolean).join("")}</${name}>`;
}

class TimingIds {
  #next = 1;

  take(): number {
    return this.#next++;
  }
}

type TimingEffect = MotionEffectObservation & {
  readonly effectIndex: number;
  readonly partIndex?: number;
  readonly targetTextBody: boolean;
  readonly targetFilled: boolean;
};

function timingTarget(effect: TimingEffect): string {
  const range = effect.target.range;
  return element(
    "p:tgtEl",
    undefined,
    element(
      "p:spTgt",
      { spid: effect.target.id },
      range?.type === "paragraph"
        ? element(
            "p:txEl",
            undefined,
            element("p:pRg", { st: range.start, end: range.end }),
          )
        : undefined,
      range?.type === "chart"
        ? element(
            "p:graphicEl",
            undefined,
            element("a:chart", {
              seriesIdx: range.seriesIndex,
              categoryIdx: range.categoryIndex,
              bldStep: range.step,
            }),
          )
        : undefined,
    ),
  );
}

function attributeNames(value: string): string {
  return element(
    "p:attrNameLst",
    undefined,
    element("p:attrName", undefined, value),
  );
}

function direction(effect: MotionEffectObservation): MotionDirection {
  return typeof effect.from === "string" ? effect.from : "bottom";
}

function behaviorIds(effect: TimingEffect, ids: TimingIds): readonly string[] {
  const target = () => timingTarget(effect);
  const duration = effect.durationMs;
  validateKnownTime(duration, "动画时长");
  if (effect.preset === "breathe") {
    timeNodeValue(Math.round(duration / 2), "呼吸半周期");
  } else {
    timeNodeValue(duration, "动画行为时长");
  }
  if (effect.repeatCount !== undefined && effect.repeatCount !== "forever") {
    timeNodeValue(effect.repeatCount * 1000, "动画重复次数");
  }
  if (effect.intervalMs !== undefined) {
    timeNodeValue(effect.intervalMs, "动画迭代间隔");
  }
  const show = () =>
    element(
      "p:set",
      undefined,
      element(
        "p:cBhvr",
        undefined,
        element(
          "p:cTn",
          { id: ids.take(), dur: 1, fill: "hold" },
          element("p:stCondLst", undefined, element("p:cond", { delay: 0 })),
        ),
        target(),
        attributeNames("style.visibility"),
      ),
      element("p:to", undefined, element("p:strVal", { val: "visible" })),
    );
  const filter = (name: string) =>
    element(
      "p:animEffect",
      { transition: "in", filter: name },
      element(
        "p:cBhvr",
        undefined,
        element("p:cTn", { id: ids.take(), dur: duration }),
        target(),
      ),
    );
  const value = (item: string | number) =>
    typeof item === "number"
      ? element("p:fltVal", { val: item })
      : element("p:strVal", { val: item });
  const motion = (
    name: string,
    from: string | number,
    to: string | number,
    additive: boolean,
  ) =>
    element(
      "p:anim",
      { calcmode: "lin", valueType: "num" },
      element(
        "p:cBhvr",
        { additive: additive ? "base" : undefined },
        element("p:cTn", { id: ids.take(), dur: duration, fill: "hold" }),
        target(),
        attributeNames(name),
      ),
      element(
        "p:tavLst",
        undefined,
        element("p:tav", { tm: 0 }, element("p:val", undefined, value(from))),
        element(
          "p:tav",
          { tm: 100000 },
          element("p:val", undefined, value(to)),
        ),
      ),
    );

  switch (effect.preset) {
    case "fade":
      return [show(), filter("fade")];
    case "wipe": {
      const filters: Record<MotionDirection, string> = {
        top: "wipe(up)",
        right: "wipe(right)",
        bottom: "wipe(down)",
        left: "wipe(left)",
      };
      return [show(), filter(filters[direction(effect)])];
    }
    case "zoom":
      return [
        show(),
        motion("ppt_w", 0, "#ppt_w", false),
        motion("ppt_h", 0, "#ppt_h", false),
        filter("fade"),
      ];
    case "fly": {
      const flyFrom: Record<MotionDirection, readonly [string, string]> = {
        left: ["0-#ppt_w/2", "#ppt_y"],
        right: ["1+#ppt_w/2", "#ppt_y"],
        top: ["#ppt_x", "0-#ppt_h/2"],
        bottom: ["#ppt_x", "1+#ppt_h/2"],
      };
      const [x, y] = flyFrom[direction(effect)];
      return [
        show(),
        motion("ppt_x", x, "#ppt_x", true),
        motion("ppt_y", y, "#ppt_y", true),
      ];
    }
    case "spinIn":
      return [
        show(),
        motion("ppt_w", 0, "#ppt_w", false),
        motion("ppt_h", 0, "#ppt_h", false),
        motion("style.rotation", 360, 0, false),
        filter("fade"),
      ];
    case "stretch":
      return [
        show(),
        motion("ppt_w", 0, "#ppt_w", false),
        motion("ppt_h", "#ppt_h", "#ppt_h", false),
      ];
    case "floatUp":
      return [
        show(),
        filter("fade"),
        motion("ppt_x", "#ppt_x", "#ppt_x", false),
        motion("ppt_y", "#ppt_y+.1", "#ppt_y", false),
      ];
    case "spin":
      return [
        element(
          "p:animRot",
          {
            from: Math.round(
              (typeof effect.from === "number" ? effect.from : 360) * 60000,
            ),
            to: 0,
          },
          element(
            "p:cBhvr",
            undefined,
            element("p:cTn", { id: ids.take(), dur: duration, fill: "hold" }),
            target(),
            attributeNames("r"),
          ),
        ),
      ];
    case "breathe":
      return [
        element(
          "p:animClr",
          { clrSpc: "rgb" },
          element(
            "p:cBhvr",
            undefined,
            element("p:cTn", {
              id: ids.take(),
              dur: Math.round(duration / 2),
              fill: "hold",
            }),
            target(),
            attributeNames("fillcolor"),
          ),
          element(
            "p:from",
            undefined,
            element("a:srgbClr", {
              val: effect.fromColor?.slice(1),
            }),
          ),
          element(
            "p:to",
            undefined,
            element("a:srgbClr", {
              val: effect.toColor?.slice(1),
            }),
          ),
        ),
      ];
  }
}

function effectTree(
  effect: TimingEffect,
  nodeType: string,
  startCondition: string,
  ids: TimingIds,
  effectId: number,
  buildGroupId?: number,
): string {
  if (effect.startAfter !== undefined) {
    validateKnownTime(effect.startAfter.offsetMs, "完成依赖偏移");
  }
  const preset = {
    fade: { id: 10, class: "entr" },
    wipe: { id: 22, class: "entr" },
    zoom: { id: 53, class: "entr" },
    fly: { id: 2, class: "entr" },
    floatUp: { id: 42, class: "entr" },
    spinIn: { id: 49, class: "entr" },
    stretch: { id: 17, class: "entr" },
    spin: { id: 8, class: "emph" },
    breathe: { id: 0, class: "emph" },
  }[effect.preset];
  const subtype =
    effect.preset === "zoom"
      ? 16
      : effect.preset === "stretch"
        ? 10
        : effect.preset === "wipe" || effect.preset === "fly"
          ? { top: 1, right: 2, bottom: 4, left: 8 }[direction(effect)]
          : 0;
  const easing = effect.easing;
  const properties: XmlAttributes = {
    id: effectId,
    presetID: preset.id,
    presetClass: preset.class,
    presetSubtype: subtype,
    repeatCount:
      effect.repeatCount === "forever"
        ? "indefinite"
        : effect.repeatCount === undefined
          ? undefined
          : effect.repeatCount * 1000,
    autoRev: effect.preset === "breathe" ? true : undefined,
    accel:
      easing?.accelerate === undefined || easing.accelerate === 0
        ? undefined
        : Math.round(easing.accelerate * 100000),
    decel:
      easing?.decelerate === undefined || easing.decelerate === 0
        ? undefined
        : Math.round(easing.decelerate * 100000),
    fill: "hold",
    grpId: effect.targetTextBody ? 0 : buildGroupId,
    nodeType,
  };
  return element(
    "p:par",
    undefined,
    element(
      "p:cTn",
      properties,
      element("p:stCondLst", undefined, startCondition),
      effect.by !== "word" && effect.by !== "letter"
        ? undefined
        : element(
            "p:iterate",
            { type: effect.by === "word" ? "wd" : "lt" },
            // 单次 dur 只描述当前词或字，不能伪装成整段自然完成时刻。
            element("p:tmAbs", { val: effect.intervalMs }),
          ),
      element("p:childTnLst", undefined, ...behaviorIds(effect, ids)),
    ),
  );
}

function effectKey(effect: TimingEffect): string {
  return `${effect.target.id}:${effect.effectIndex}:${effect.partIndex ?? ""}`;
}

function startAfterDelay(
  start: MotionPlanStart,
  delayMs: number,
): MotionPlanStart {
  if (start.kind === "parent-start") {
    return { kind: start.kind, delayMs: start.delayMs + delayMs };
  }
  if (start.kind === "end") {
    return {
      kind: start.kind,
      dependency: {
        ...start.dependency,
        offsetMs: start.dependency.offsetMs + delayMs,
      },
    };
  }
  if (start.kind === "event") {
    return {
      kind: start.kind,
      dependency: {
        ...start.dependency,
        offsetMs: start.dependency.offsetMs + delayMs,
      },
    };
  }
  if (start.kind === "condition-join") {
    return { kind: "condition-end", id: start.id, delayMs };
  }
  return { ...start, delayMs: start.delayMs + delayMs };
}

// 仅在合并结果仍可写入单个 ST_TLTime 字段时折叠延迟层。
function addNativeStartDelay(
  start: MotionPlanStart,
  delayMs: number,
): MotionPlanStart | undefined {
  const delayed = startAfterDelay(start, delayMs);
  const value =
    delayed.kind === "parent-start"
      ? delayed.delayMs
      : delayed.kind === "end" || delayed.kind === "event"
        ? delayed.dependency.offsetMs
        : delayed.kind === "condition-end"
          ? delayed.delayMs
          : undefined;
  return Number.isSafeInteger(value) &&
    value !== undefined &&
    value >= 0 &&
    value <= maximumTimeNodeValue
    ? delayed
    : undefined;
}

function dependencyTimingId(
  dependency: MotionStartDependency,
  stepKey: string,
  effectNodeIds: ReadonlyMap<string, number>,
  scopeNodeIds: ReadonlyMap<string, number>,
  targetGroupIds: ReadonlyMap<string, number>,
): number | undefined {
  if ("scopeId" in dependency) {
    return scopeNodeIds.get(
      `${stepKey}:${dependency.point}:${dependency.scopeId}`,
    );
  }
  if ("group" in dependency) {
    return targetGroupIds.get(`${stepKey}:${dependency.targetId}`);
  }
  return effectNodeIds.get(
    `${dependency.targetId}:${dependency.effectIndex}:${dependency.partIndex ?? ""}`,
  );
}

function nativeStartCondition(
  start: MotionPlanStart,
  stepKey: string,
  effectNodeIds: ReadonlyMap<string, number>,
  scopeNodeIds: ReadonlyMap<string, number>,
  targetGroupIds: ReadonlyMap<string, number>,
  conditionNodeIds: ReadonlyMap<string, number>,
): string {
  if (start.kind === "parent-start") {
    return element("p:cond", {
      delay: timeNodeValue(start.delayMs, "动画起点延迟"),
    });
  }
  if (start.kind === "condition-join") {
    throw new Error("MotionPlan 条件集合必须先于其消费者写出");
  }
  if (start.kind === "condition-end") {
    const conditionId = conditionNodeIds.get(`${stepKey}:${start.id}`);
    if (conditionId === undefined) {
      throw new Error("MotionPlan 条件集合缺少原生时间节点身份");
    }
    return element(
      "p:cond",
      {
        evt: "end",
        delay: timeNodeValue(start.delayMs, "条件集合完成偏移"),
      },
      element("p:tn", { val: conditionId }),
    );
  }
  const dependencyId = dependencyTimingId(
    start.dependency,
    stepKey,
    effectNodeIds,
    scopeNodeIds,
    targetGroupIds,
  );
  if (dependencyId === undefined) {
    throw new Error("MotionPlan 完成依赖没有对应的原生时序节点");
  }
  return element(
    "p:cond",
    {
      evt:
        start.kind === "event" || start.dependency.point === "start"
          ? "begin"
          : "end",
      delay: timeNodeValue(start.dependency.offsetMs, "完成依赖偏移"),
    },
    element("p:tn", { val: dependencyId }),
  );
}

function endSync(): string {
  // 容器等全部真实子节点自然结束；不从这些子项估算聚合数值时长。
  return element(
    "p:endSync",
    { evt: "end", delay: 0 },
    element("p:rtn", { val: "all" }),
  );
}

function timingContainer(
  id: number,
  startCondition: string,
  children: readonly string[],
  nodeType?: string,
): string {
  const isEmpty = children.length === 0;
  return element(
    "p:par",
    undefined,
    element(
      "p:cTn",
      { id, ...(isEmpty ? { dur: 0 } : {}), fill: "hold", nodeType },
      element("p:stCondLst", undefined, startCondition),
      isEmpty ? undefined : endSync(),
      // 空作用域的 END 与其未知 START 相同；省略空列表以保持合法 XML。
      isEmpty ? undefined : element("p:childTnLst", undefined, ...children),
    ),
  );
}

function conditionJoinTree(
  start: Extract<MotionPlanStart, { readonly kind: "condition-join" }>,
  stepKey: string,
  ids: TimingIds,
  effectNodeIds: ReadonlyMap<string, number>,
  scopeNodeIds: ReadonlyMap<string, number>,
  targetGroupIds: ReadonlyMap<string, number>,
  conditionNodeIds: ReadonlyMap<string, number>,
): string {
  const id = conditionNodeIds.get(`${stepKey}:${start.id}`);
  if (id === undefined) {
    throw new Error("MotionPlan 条件集合缺少原生时间节点身份");
  }
  return endCollectorTree(
    id,
    element("p:cond", { delay: 0 }),
    start.gates,
    stepKey,
    ids,
    effectNodeIds,
    scopeNodeIds,
    targetGroupIds,
    conditionNodeIds,
  );
}

function gateStart(gate: MotionPlanStartGate): MotionPlanStart {
  switch (gate.kind) {
    case "parent-start":
      return { kind: "parent-start", delayMs: gate.delayMs };
    case "end":
      return { kind: "end", dependency: gate.dependency };
    case "begin":
      return { kind: "event", dependency: gate.dependency };
  }
}

function endCollectorTree(
  id: number,
  startCondition: string,
  gates: readonly MotionPlanStartGate[],
  stepKey: string,
  ids: TimingIds,
  effectNodeIds: ReadonlyMap<string, number>,
  scopeNodeIds: ReadonlyMap<string, number>,
  targetGroupIds: ReadonlyMap<string, number>,
  conditionNodeIds: ReadonlyMap<string, number>,
  nodeType?: string,
): string {
  const leaves = gates.map((gate) =>
    element(
      "p:par",
      undefined,
      element(
        "p:cTn",
        { id: ids.take(), dur: 0, fill: "hold" },
        element(
          "p:stCondLst",
          undefined,
          nativeStartCondition(
            gateStart(gate),
            stepKey,
            effectNodeIds,
            scopeNodeIds,
            targetGroupIds,
            conditionNodeIds,
          ),
        ),
      ),
    ),
  );
  return timingContainer(id, startCondition, leaves, nodeType);
}

function mainSequenceWithGroup(
  id: number,
  children: readonly string[],
): string {
  // mainSeq 的标准 step→with-group→入口节点层级供原生导入识别步首角色。
  // 分组从步首立即开始，不复制计划起点，也不创建 Shape 或完成来源。
  return element(
    "p:par",
    undefined,
    element(
      "p:cTn",
      { id, fill: "hold" },
      element("p:stCondLst", undefined, element("p:cond", { delay: 0 })),
      element("p:childTnLst", undefined, ...children),
    ),
  );
}

function timingXml(steps: readonly MotionPlanStep[]): string {
  const ids = new TimingIds();
  const rootId = ids.take();
  const sequenceId = ids.take();
  const effectNodeIds = new Map<string, number>();
  const scopeNodeIds = new Map<string, number>();
  const targetGroupIds = new Map<string, number>();
  const conditionNodeIds = new Map<string, number>();
  const referencedScopeNodes = new Set<string>();
  const referencedScopeStartNodes = new Set<string>();
  const referencedTargetGroups = new Set<string>();
  const keepDependencyContainer = (
    dependency: MotionStartDependency,
    stepKey: string,
  ): void => {
    if ("scopeId" in dependency) {
      const key = `${stepKey}:${dependency.point}:${dependency.scopeId}`;
      if (dependency.point === "start") {
        referencedScopeStartNodes.add(key);
      } else {
        referencedScopeNodes.add(key);
      }
    } else if ("group" in dependency) {
      referencedTargetGroups.add(`${stepKey}:${dependency.targetId}`);
    }
  };

  // 只为真实 end 依赖保留聚合节点；无消费者的容器可把起点投影给成员。
  const collectReferencedContainers = (
    nodes: readonly MotionPlanNode[],
    stepKey: string,
  ): void => {
    for (const node of nodes) {
      if (
        node.nativeStart.kind === "end" ||
        node.nativeStart.kind === "event"
      ) {
        keepDependencyContainer(node.nativeStart.dependency, stepKey);
      } else if (node.nativeStart.kind === "condition-join") {
        referencedScopeStartNodes.add(
          `${stepKey}:start:${node.nativeStart.hostScopeId}`,
        );
        referencedScopeStartNodes.add(
          `${stepKey}:start:${node.nativeStart.consumerParentScopeId}`,
        );
        for (const gate of node.nativeStart.gates) {
          if (gate.kind === "end" || gate.kind === "begin") {
            keepDependencyContainer(gate.dependency, stepKey);
          }
        }
      }
      if (node.kind === "completion-collector") {
        for (const gate of node.gates) {
          if (gate.kind === "end" || gate.kind === "begin") {
            keepDependencyContainer(gate.dependency, stepKey);
          }
        }
      }
      if (node.kind === "scope") {
        for (const placement of node.conditionDefinitionsBefore ?? []) {
          referencedScopeStartNodes.add(
            `${stepKey}:start:${placement.join.hostScopeId}`,
          );
          referencedScopeStartNodes.add(
            `${stepKey}:start:${placement.join.consumerParentScopeId}`,
          );
          for (const gate of placement.join.gates) {
            if (gate.kind === "end" || gate.kind === "begin") {
              keepDependencyContainer(gate.dependency, stepKey);
            }
          }
        }
        collectReferencedContainers(node.children, stepKey);
      }
    }
  };
  for (const step of steps) {
    collectReferencedContainers(step.nodes, step.key);
    for (const placement of step.conditionDefinitionsBefore ?? []) {
      for (const gate of placement.join.gates) {
        if (gate.kind === "end" || gate.kind === "begin") {
          keepDependencyContainer(gate.dependency, step.key);
        }
      }
    }
  }

  const registerNodeIds = (
    nodes: readonly MotionPlanNode[],
    stepKey: string,
  ): void => {
    for (const node of nodes) {
      if (node.nativeStart.kind === "condition-join") {
        conditionNodeIds.set(`${stepKey}:${node.nativeStart.id}`, ids.take());
      }
      if (node.kind === "scope") {
        for (const placement of node.conditionDefinitionsBefore ?? []) {
          conditionNodeIds.set(`${stepKey}:${placement.join.id}`, ids.take());
        }
        const scopeTimingId = ids.take();
        scopeNodeIds.set(`${stepKey}:start:${node.id}`, scopeTimingId);
        if (node.splitCompletion !== true) {
          scopeNodeIds.set(`${stepKey}:end:${node.id}`, scopeTimingId);
        }
        registerNodeIds(node.children, stepKey);
        continue;
      }
      if (node.kind === "completion-collector") {
        scopeNodeIds.set(`${stepKey}:end:${node.scopeId}`, ids.take());
        continue;
      }
      if (node.effects.length > 1) {
        targetGroupIds.set(`${stepKey}:${node.targetId}`, ids.take());
      }
      for (const effect of node.effects) {
        effectNodeIds.set(effectKey(effect), ids.take());
      }
    }
  };
  for (const step of steps) {
    for (const placement of step.conditionDefinitionsBefore ?? []) {
      conditionNodeIds.set(`${step.key}:${placement.join.id}`, ids.take());
    }
    registerNodeIds(step.nodes, step.key);
  }
  const chartBuildGroupIds = new Map<number, number>();
  for (const step of steps) {
    for (const effect of step.effects) {
      if (
        effect.target.range?.type === "chart" &&
        !chartBuildGroupIds.has(effect.target.id)
      ) {
        // bldGraphic 的 grpId 必须关联该图表实际存在的时间节点。
        const id = effectNodeIds.get(effectKey(effect));
        if (id !== undefined) chartBuildGroupIds.set(effect.target.id, id);
      }
    }
  }
  const stepGroups = steps.map((step) => {
    const stepId = ids.take();
    const trigger =
      step.trigger === "click"
        ? element("p:cond", { delay: "indefinite" })
        : `${element("p:cond", { delay: "indefinite" })}${element(
            "p:cond",
            { evt: "onBegin", delay: 0 },
            element("p:tn", { val: sequenceId }),
          )}`;
    const childDelay = (effect: TimingEffect): number =>
      timeNodeValue(
        effect.delayMs + (effect.partIndex ?? 0) * effect.durationMs,
        "动画对象成员延迟",
      );
    const renderNodes = (
      nodes: readonly MotionPlanNode[],
      inheritedStart: MotionPlanStart | undefined,
      isTopLevel: boolean,
      firstRootNodeType?: string,
      conditionDefinitionsBefore: readonly MotionPlanConditionPlacement[] = [],
    ): string[] => {
      const rendered: string[] = [];
      let pendingRootNodeType = firstRootNodeType;
      for (const [index, node] of nodes.entries()) {
        const nodeId =
          node.kind === "scope"
            ? `scope:${node.id}`
            : node.kind === "completion-collector"
              ? `completion:${node.scopeId}`
              : `target:${node.targetId}`;
        for (const placement of conditionDefinitionsBefore) {
          if (placement.join.beforeNodeId !== nodeId) continue;
          rendered.push(
            conditionJoinTree(
              placement.join,
              step.key,
              ids,
              effectNodeIds,
              scopeNodeIds,
              targetGroupIds,
              conditionNodeIds,
            ),
          );
        }
        const defaultNodeType =
          index === 0 && isTopLevel && step.trigger === "click"
            ? "clickEffect"
            : index === 0
              ? "afterEffect"
              : "withEffect";
        const nodeType = defaultNodeType;
        let effectiveStart = composeMotionPlanStarts(
          inheritedStart,
          node.nativeStart,
        );
        if (effectiveStart === undefined) {
          throw new Error("原生时序投影需要保留延迟分层");
        }
        if (effectiveStart.kind === "condition-join") {
          rendered.push(
            conditionJoinTree(
              effectiveStart,
              step.key,
              ids,
              effectNodeIds,
              scopeNodeIds,
              targetGroupIds,
              conditionNodeIds,
            ),
          );
          effectiveStart = {
            kind: "condition-end",
            id: effectiveStart.id,
            delayMs: 0,
          };
        }
        if (node.kind === "scope") {
          const scopeId = scopeNodeIds.get(`${step.key}:start:${node.id}`);
          if (scopeId === undefined) {
            throw new Error("MotionPlan 作用域缺少原生时间节点身份");
          }
          const needsCompletionNode =
            node.splitCompletion !== true &&
            referencedScopeNodes.has(`${step.key}:end:${node.id}`);
          const needsStartNode =
            node.preserveStart === true ||
            referencedScopeStartNodes.has(`${step.key}:start:${node.id}`);
          const childrenCanLift = node.children.every((child) => {
            const childStart = composeMotionPlanStarts(
              effectiveStart,
              child.nativeStart,
            );
            if (childStart === undefined) return false;
            if (
              child.kind === "scope" ||
              child.kind === "completion-collector" ||
              child.effects.length > 1
            ) {
              return true;
            }
            const [effect] = child.effects;
            return (
              effect !== undefined &&
              addNativeStartDelay(childStart, childDelay(effect)) !== undefined
            );
          });
          if (needsCompletionNode || needsStartNode || !childrenCanLift) {
            const condition = nativeStartCondition(
              effectiveStart,
              step.key,
              effectNodeIds,
              scopeNodeIds,
              targetGroupIds,
              conditionNodeIds,
            );
            const children = renderNodes(
              node.children,
              undefined,
              false,
              undefined,
              node.conditionDefinitionsBefore,
            );
            rendered.push(
              timingContainer(
                scopeId,
                condition,
                children,
                pendingRootNodeType,
              ),
            );
          } else {
            rendered.push(
              ...renderNodes(
                node.children,
                effectiveStart,
                false,
                pendingRootNodeType,
                node.conditionDefinitionsBefore,
              ),
            );
          }
          if (rendered.length > 0) pendingRootNodeType = undefined;
          continue;
        }

        if (node.kind === "completion-collector") {
          const collectorId = scopeNodeIds.get(
            `${step.key}:end:${node.scopeId}`,
          );
          if (collectorId === undefined) {
            throw new Error("MotionPlan 完成集合缺少原生时间节点身份");
          }
          const condition = nativeStartCondition(
            effectiveStart,
            step.key,
            effectNodeIds,
            scopeNodeIds,
            targetGroupIds,
            conditionNodeIds,
          );
          rendered.push(
            endCollectorTree(
              collectorId,
              condition,
              node.gates,
              step.key,
              ids,
              effectNodeIds,
              scopeNodeIds,
              targetGroupIds,
              conditionNodeIds,
              pendingRootNodeType,
            ),
          );
          pendingRootNodeType = undefined;
          continue;
        }

        const condition = nativeStartCondition(
          effectiveStart,
          step.key,
          effectNodeIds,
          scopeNodeIds,
          targetGroupIds,
          conditionNodeIds,
        );
        const targetGroupKey = `${step.key}:${node.targetId}`;
        const effectStarts = node.effects.map((effect) =>
          addNativeStartDelay(effectiveStart, childDelay(effect)),
        );
        const keepTargetGroup =
          node.effects.length > 1 &&
          (referencedTargetGroups.has(targetGroupKey) ||
            effectStarts.some((start) => start === undefined));
        if (keepTargetGroup) {
          const groupId = targetGroupIds.get(targetGroupKey);
          if (groupId === undefined) {
            throw new Error("MotionPlan 对象成员缺少原生聚合节点身份");
          }
          const children = node.effects.map((effect, effectIndex) => {
            const effectId = effectNodeIds.get(effectKey(effect));
            if (effectId === undefined) {
              throw new Error("MotionPlan 效果缺少原生时间节点身份");
            }
            const buildGroupId =
              effect.target.range?.type === "chart"
                ? chartBuildGroupIds.get(effect.target.id)
                : undefined;
            return effectTree(
              effect,
              effectIndex === 0 ? "afterEffect" : "withEffect",
              element("p:cond", { delay: childDelay(effect) }),
              ids,
              effectId,
              buildGroupId,
            );
          });
          rendered.push(
            timingContainer(groupId, condition, children, pendingRootNodeType),
          );
          pendingRootNodeType = undefined;
          continue;
        }

        for (const [effectIndex, effect] of node.effects.entries()) {
          const effectId = effectNodeIds.get(effectKey(effect));
          if (effectId === undefined) {
            throw new Error("MotionPlan 效果缺少原生时间节点身份");
          }
          const localStart = effectStarts[effectIndex];
          if (localStart === undefined) {
            throw new Error("原生时序投影丢失可编码的动画延迟分层");
          }
          const effectCondition = nativeStartCondition(
            localStart,
            step.key,
            effectNodeIds,
            scopeNodeIds,
            targetGroupIds,
            conditionNodeIds,
          );
          const buildGroupId =
            effect.target.range?.type === "chart"
              ? chartBuildGroupIds.get(effect.target.id)
              : undefined;
          rendered.push(
            effectTree(
              effect,
              pendingRootNodeType ??
                (effectIndex === 0 ? nodeType : "withEffect"),
              effectCondition,
              ids,
              effectId,
              buildGroupId,
            ),
          );
          pendingRootNodeType = undefined;
        }
      }
      return rendered;
    };
    // nodeType 仅标记真实步首节点，应用的时刻条件仍完全来自计划。
    const stepChildren = renderNodes(
      step.nodes,
      undefined,
      true,
      step.trigger === "click" ? "clickEffect" : "afterEffect",
      step.conditionDefinitionsBefore,
    );
    const stepTimingGroups =
      stepChildren.length === 0
        ? []
        : [mainSequenceWithGroup(ids.take(), stepChildren)];
    return element(
      "p:par",
      undefined,
      element(
        "p:cTn",
        { id: stepId, fill: "hold" },
        element("p:stCondLst", undefined, trigger),
        element("p:childTnLst", undefined, ...stepTimingGroups),
      ),
    );
  });
  const slideTarget = element(
    "p:tgtEl",
    undefined,
    element("p:sldTgt", undefined),
  );
  const buildTargets = new Map<number, TimingEffect>();
  for (const step of steps) {
    for (const effect of step.effects) {
      const timed = effect as TimingEffect;
      if (timed.targetTextBody || timed.target.range?.type === "chart") {
        const previous = buildTargets.get(timed.target.id);
        // 原生构建粒度属于整个文字对象，整体效果不能覆盖逐段构建要求。
        buildTargets.set(
          timed.target.id,
          timed.targetTextBody &&
            (timed.by === "paragraph" || previous?.by === "paragraph")
            ? { ...timed, by: "paragraph" }
            : timed,
        );
      }
    }
  }
  return element(
    "p:timing",
    undefined,
    element(
      "p:tnLst",
      undefined,
      element(
        "p:par",
        undefined,
        element(
          "p:cTn",
          {
            id: rootId,
            dur: "indefinite",
            restart: "never",
            nodeType: "tmRoot",
          },
          element(
            "p:childTnLst",
            undefined,
            element(
              "p:seq",
              { concurrent: true, nextAc: "seek" },
              element(
                "p:cTn",
                { id: sequenceId, dur: "indefinite", nodeType: "mainSeq" },
                element("p:childTnLst", undefined, ...stepGroups),
              ),
              element(
                "p:prevCondLst",
                undefined,
                element("p:cond", { evt: "onPrev", delay: 0 }, slideTarget),
              ),
              element(
                "p:nextCondLst",
                undefined,
                element("p:cond", { evt: "onNext", delay: 0 }, slideTarget),
              ),
            ),
          ),
        ),
      ),
    ),
    buildTargets.size === 0
      ? undefined
      : element(
          "p:bldLst",
          undefined,
          ...[...buildTargets.values()].map((effect) => {
            if (effect.target.range?.type === "chart") {
              const by = effect.target.range.step;
              return element(
                "p:bldGraphic",
                {
                  spid: effect.target.id,
                  grpId: chartBuildGroupIds.get(effect.target.id),
                },
                by === "series" || by === "category"
                  ? element(
                      "p:bldSub",
                      undefined,
                      element("a:bldChart", { bld: by }),
                    )
                  : element("p:bldAsOne", undefined),
              );
            }
            return element("p:bldP", {
              spid: effect.target.id,
              grpId: 0,
              build: effect.by === "paragraph" ? "p" : undefined,
              animBg: effect.targetFilled ? true : undefined,
            });
          }),
        ),
  );
}

/** 页面级动画模块写入对象时序，并返回仅描述声明与可确定相对时间的观察。 */
export const motionCapability: SlideCapabilityModule = Object.freeze({
  declaration: "motion",
  slideProp: "__motion",
  slot: "timing",
  always: true,
  transparentDeclarations: [
    {
      declaration: "rp:motion-scope",
      compileDeclaration(input) {
        return [motionScopeAnnotation(input.props)];
      },
    } satisfies TransparentCapabilityModule,
  ],
  compileSlide(_input: CapabilityInput, context) {
    const plan = compileMotionPlan(context.contributions);
    validateKnownMotionTimes(plan.observation);
    const hasEffects = plan.steps.some((step) => step.effects.length > 0);
    if (!hasEffects) {
      return {
        xml: "",
        extensions: [],
        ...(plan.observation.scopes === undefined
          ? {}
          : { observation: plan.observation }),
      };
    }
    return {
      xml: timingXml(plan.steps),
      extensions: [],
      observation: plan.observation satisfies MotionObservation,
    };
  },
});
