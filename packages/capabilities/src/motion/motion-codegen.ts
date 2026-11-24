import {
  defineObjectCodegen,
  expr,
  type CodegenContext,
  type CodegenExpr,
} from "@pptxdk/codegen";
import { CommonNode } from "@pptxdk/core/document";

import {
  breatheOptionsSchema,
  effectDelaySchema,
  effectOptionsSchema,
  emphasisEffectSchema,
  entranceEffectSchema,
  fillColorOptionsSchema,
  motionDirectionSchema,
  motionMomentOffsetSchema,
  motionMomentSchema,
  motionScopeHandleSchema,
  motionScopeModeSchema,
  motionStepAuthorPropsSchema,
  ownerMotionSchema,
  repeatCountSchema,
  spinOptionsSchema,
  type Easing,
  type Effect,
  type MotionMoment,
  type MotionMomentOffset,
} from "./motion-schemas.ts";

const stepAllowedPropKeys = new Set(["n", "stepIdentity"]);
const scopeAllowedPropKeys = new Set([
  "scope",
  "mode",
  "each",
  "delay",
  "startAt",
  "publishStart",
  "publishEnd",
  "motion",
]);

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function subPropPath(ctx: CodegenContext, segment: string): string {
  return ctx.path.includes("@") ? `${ctx.path}.${segment}` : `@${segment}`;
}

function emitContainerChildren(
  node: CommonNode,
  ctx: CodegenContext,
  ownerLabel: string,
): readonly string[] {
  const children: string[] = [];
  let childIndex = 0;
  for (const child of node.children) {
    if (child instanceof CommonNode) {
      children.push(ctx.emitChildNode(child, childIndex));
      childIndex += 1;
    } else if (
      "text" in child &&
      typeof child.text === "string" &&
      child.text.trim() !== ""
    ) {
      ctx.fail(
        "invalid-document",
        `${ownerLabel} cannot directly contain text content`,
      );
    }
  }
  return children;
}

function serializeMotionMomentValue(
  value: unknown,
  ctx: CodegenContext,
  propSegment: string,
): CodegenExpr {
  const parsed = motionMomentSchema.safeParse(value);
  if (!parsed.success) {
    ctx.fail(
      "unexpressible-value",
      "Expected a valid MotionMoment handle from a motion scope",
      subPropPath(ctx, propSegment),
    );
  }
  const moment: MotionMoment = parsed.data;
  const scopeVar = ctx.allocateMotionScope(moment.scope);
  if (moment.point === "start") {
    return expr(`${scopeVar}.start`);
  }
  if (moment.point === "end") {
    return expr(`${scopeVar}.end`);
  }
  return expr(`${scopeVar}.cue(${JSON.stringify(moment.name)})`);
}

function serializeMotionTimingValue(
  value: unknown,
  ctx: CodegenContext,
  propSegment: string,
): CodegenExpr {
  if (isRecord(value) && value["type"] === "motion-offset") {
    const rawOffset = value["offsetMs"];
    if (typeof rawOffset === "number" && !Number.isFinite(rawOffset)) {
      ctx.fail(
        "unexpressible-value",
        `Non-finite number ${String(rawOffset)} cannot be expressed in TypeScript`,
        subPropPath(ctx, `${propSegment}.offsetMs`),
      );
    }
    const parsed = motionMomentOffsetSchema.safeParse(value);
    if (!parsed.success) {
      ctx.fail(
        "unexpressible-value",
        "Expected a valid MotionMomentOffset created by after(moment, offsetMs)",
        subPropPath(ctx, propSegment),
      );
    }
    const offset: MotionMomentOffset = parsed.data;
    const afterFn = ctx.useImport("after");
    const momentExpr = serializeMotionMomentValue(
      offset.moment,
      ctx,
      `${propSegment}.moment`,
    );
    const offsetExpr = ctx.emitValueExpr(
      offset.offsetMs,
      subPropPath(ctx, `${propSegment}.offsetMs`),
    );
    return expr(`${afterFn}(${momentExpr.code}, ${offsetExpr})`);
  }
  return serializeMotionMomentValue(value, ctx, propSegment);
}

function collectStepProps(
  node: CommonNode,
  ctx: CodegenContext,
): readonly (readonly [string, unknown])[] {
  for (const [key, value] of Object.entries(node.props)) {
    if (value === undefined) continue;
    if (!stepAllowedPropKeys.has(key)) {
      ctx.fail(
        "invalid-document",
        `Step declaration does not support property ${JSON.stringify(key)}`,
        `@${key}`,
      );
    }
  }
  const rawN = node.props["n"];
  if (rawN === undefined) {
    return [];
  }
  if (typeof rawN === "number" && !Number.isFinite(rawN)) {
    ctx.fail(
      "unexpressible-value",
      `Non-finite number ${String(rawN)} cannot be expressed in TypeScript`,
      "@n",
    );
  }
  const parsed = motionStepAuthorPropsSchema.safeParse({ n: rawN });
  if (!parsed.success || parsed.data.n === undefined) {
    ctx.fail("unexpressible-value", "Step.n must be a positive integer", "@n");
  }
  return [["n", parsed.data.n]];
}

function collectScopeProps(
  node: CommonNode,
  ctx: CodegenContext,
): {
  readonly scopeVar: string;
  readonly optionEntries: readonly (readonly [string, unknown])[];
} {
  for (const [key, value] of Object.entries(node.props)) {
    if (value === undefined) continue;
    if (!scopeAllowedPropKeys.has(key)) {
      ctx.fail(
        "invalid-document",
        `MotionScope declaration does not support property ${JSON.stringify(key)}`,
        `@${key}`,
      );
    }
  }

  const parsedScope = motionScopeHandleSchema.safeParse(node.props["scope"]);
  if (!parsedScope.success) {
    ctx.fail(
      "unexpressible-value",
      "MotionScope requires a handle created by createMotionScopeHandle() or useMotionScope()",
      "@scope",
    );
  }
  const scopeVar = ctx.allocateMotionScope(parsedScope.data.identity);

  const optionEntries: [string, unknown][] = [];
  const rawMode = node.props["mode"];
  if (rawMode !== undefined) {
    const parsedMode = motionScopeModeSchema.safeParse(rawMode);
    if (!parsedMode.success) {
      ctx.fail(
        "unexpressible-value",
        'MotionScope.mode must be "parallel", "sequence", or "stagger"',
        "@mode",
      );
    }
    optionEntries.push(["mode", parsedMode.data]);
  }

  const rawEach = node.props["each"];
  if (rawEach !== undefined) {
    if (typeof rawEach === "number" && !Number.isFinite(rawEach)) {
      ctx.fail(
        "unexpressible-value",
        `Non-finite number ${String(rawEach)} cannot be expressed in TypeScript`,
        "@each",
      );
    }
    if (
      typeof rawEach !== "number" ||
      !Number.isSafeInteger(rawEach) ||
      rawEach < 0
    ) {
      ctx.fail(
        "unexpressible-value",
        "MotionScope.each must be a non-negative integer",
        "@each",
      );
    }
    optionEntries.push(["each", rawEach]);
  }

  const rawDelay = node.props["delay"];
  if (rawDelay !== undefined) {
    if (typeof rawDelay === "number" && !Number.isFinite(rawDelay)) {
      ctx.fail(
        "unexpressible-value",
        `Non-finite number ${String(rawDelay)} cannot be expressed in TypeScript`,
        "@delay",
      );
    }
    const parsedDelay = effectDelaySchema.safeParse(rawDelay);
    if (!parsedDelay.success) {
      ctx.fail(
        "unexpressible-value",
        "MotionScope.delay must be a finite non-negative integer in milliseconds",
        "@delay",
      );
    }
    optionEntries.push(["delay", parsedDelay.data]);
  }

  const rawStartAt = node.props["startAt"];
  if (rawStartAt !== undefined) {
    optionEntries.push([
      "startAt",
      serializeMotionTimingValue(rawStartAt, ctx, "startAt"),
    ]);
  }

  const rawPublishStart = node.props["publishStart"];
  if (rawPublishStart !== undefined) {
    optionEntries.push([
      "publishStart",
      serializeMotionMomentValue(rawPublishStart, ctx, "publishStart"),
    ]);
  }

  const rawPublishEnd = node.props["publishEnd"];
  if (rawPublishEnd !== undefined) {
    optionEntries.push([
      "publishEnd",
      serializeMotionMomentValue(rawPublishEnd, ctx, "publishEnd"),
    ]);
  }

  const rawMotion = node.props["motion"];
  if (rawMotion !== undefined) {
    if (typeof rawMotion !== "boolean") {
      ctx.fail(
        "unexpressible-value",
        "MotionScope.motion must be a boolean",
        "@motion",
      );
    }
    optionEntries.push(["motion", rawMotion]);
  }

  return { scopeVar, optionEntries };
}

function failAtSubPath(
  ctx: CodegenContext,
  message: string,
  subPath?: string,
): never {
  return ctx.fail("unexpressible-value", message, subPath);
}

function assertFiniteNumbersInRecord(
  value: unknown,
  ctx: CodegenContext,
  basePath: string,
): void {
  if (typeof value === "number") {
    if (!Number.isFinite(value)) {
      failAtSubPath(
        ctx,
        `Non-finite number ${String(value)} cannot be expressed in TypeScript`,
        basePath,
      );
    }
    return;
  }
  if (Array.isArray(value)) {
    for (const [index, item] of value.entries()) {
      assertFiniteNumbersInRecord(item, ctx, `${basePath}[${index}]`);
    }
    return;
  }
  if (!isRecord(value)) return;
  for (const [key, entryValue] of Object.entries(value)) {
    assertFiniteNumbersInRecord(entryValue, ctx, `${basePath}.${key}`);
  }
}

function cleanEasing(
  easing: Easing | undefined,
): Readonly<Record<string, number>> | undefined {
  if (easing === undefined) return undefined;
  const cleaned: Record<string, number> = {};
  if (easing.accelerate !== undefined) {
    cleaned["accelerate"] = easing.accelerate;
  }
  if (easing.decelerate !== undefined) {
    cleaned["decelerate"] = easing.decelerate;
  }
  return Object.keys(cleaned).length === 0 ? undefined : cleaned;
}

function serializeEffectCall(
  rawEffect: unknown,
  expectedKind: Effect["kind"] | undefined,
  ctx: CodegenContext,
  effectPath: string,
): CodegenExpr {
  if (!isRecord(rawEffect) || rawEffect["type"] !== "effect") {
    failAtSubPath(
      ctx,
      "Expected a valid animation effect declaration",
      effectPath,
    );
  }
  assertFiniteNumbersInRecord(rawEffect, ctx, effectPath);

  let effect: Effect;
  if (expectedKind === "entrance") {
    const parsed = entranceEffectSchema.safeParse(rawEffect);
    if (!parsed.success) {
      failAtSubPath(
        ctx,
        parsed.error.issues[0]?.message ?? "Invalid entrance animation effect",
        effectPath,
      );
    }
    effect = parsed.data;
  } else if (expectedKind === "emphasis") {
    const parsed = emphasisEffectSchema.safeParse(rawEffect);
    if (!parsed.success) {
      failAtSubPath(
        ctx,
        parsed.error.issues[0]?.message ?? "Invalid emphasis animation effect",
        effectPath,
      );
    }
    effect = parsed.data;
  } else {
    const parsed =
      rawEffect["kind"] === "entrance"
        ? entranceEffectSchema.safeParse(rawEffect)
        : emphasisEffectSchema.safeParse(rawEffect);
    if (!parsed.success) {
      failAtSubPath(
        ctx,
        parsed.error.issues[0]?.message ?? "Invalid animation effect",
        effectPath,
      );
    }
    effect = parsed.data;
  }
  const easing = cleanEasing(effect.easing);

  if (effect.preset === "breathe") {
    if (
      Object.getOwnPropertySymbols(rawEffect).length > 0 ||
      effect.from !== undefined ||
      effect.by !== undefined ||
      effect.intervalMs !== undefined
    ) {
      failAtSubPath(
        ctx,
        "breathe() contains unsupported parameters",
        effectPath,
      );
    }
    const candidate = {
      period: effect.duration,
      from: effect.fromColor,
      to: effect.toColor,
      ...(effect.delay === 0 ? {} : { delay: effect.delay }),
      ...(easing === undefined ? {} : { easing }),
    };
    const parsed = breatheOptionsSchema.safeParse(candidate);
    if (!parsed.success) {
      failAtSubPath(
        ctx,
        parsed.error.issues[0]?.message ?? "Invalid breathe() options",
        effectPath,
      );
    }
    const fn = ctx.useImport("breathe");
    const entries: (readonly [string, unknown])[] = [
      ["period", parsed.data.period],
      ["from", parsed.data.from],
      ["to", parsed.data.to],
      ...(parsed.data.delay === undefined || parsed.data.delay === 0
        ? []
        : [["delay", parsed.data.delay] as const]),
      ...(parsed.data.easing === undefined
        ? []
        : [["easing", parsed.data.easing] as const]),
    ];
    return expr(`${fn}(${ctx.formatProps(entries)})`);
  }

  if (effect.preset === "fillColor") {
    if (
      effect.from !== undefined ||
      effect.by !== undefined ||
      effect.intervalMs !== undefined
    ) {
      failAtSubPath(
        ctx,
        "fillColor() contains unsupported parameters",
        effectPath,
      );
    }
    const candidate = {
      from: effect.fromColor,
      to: effect.toColor,
      ...(effect.duration === 500 ? {} : { duration: effect.duration }),
      ...(effect.delay === 0 ? {} : { delay: effect.delay }),
      ...(easing === undefined ? {} : { easing }),
    };
    const parsed = fillColorOptionsSchema.safeParse(candidate);
    if (!parsed.success) {
      failAtSubPath(
        ctx,
        parsed.error.issues[0]?.message ?? "Invalid fillColor() options",
        effectPath,
      );
    }
    const fn = ctx.useImport("fillColor");
    const entries: (readonly [string, unknown])[] = [
      ["from", parsed.data.from],
      ["to", parsed.data.to],
      ...(parsed.data.duration === undefined
        ? []
        : [["duration", parsed.data.duration] as const]),
      ...(parsed.data.delay === undefined || parsed.data.delay === 0
        ? []
        : [["delay", parsed.data.delay] as const]),
      ...(parsed.data.easing === undefined
        ? []
        : [["easing", parsed.data.easing] as const]),
    ];
    return expr(`${fn}(${ctx.formatProps(entries)})`);
  }

  if (effect.fromColor !== undefined || effect.toColor !== undefined) {
    failAtSubPath(
      ctx,
      `${effect.preset}() does not accept color parameters`,
      effectPath,
    );
  }

  if (effect.preset === "spin") {
    const omitDefaultFrom =
      typeof effect.by !== "number" && effect.from === 360;
    const candidate = {
      ...(effect.from === undefined || omitDefaultFrom
        ? {}
        : { from: effect.from }),
      ...(effect.duration === 1000 ? {} : { duration: effect.duration }),
      ...(effect.delay === 0 ? {} : { delay: effect.delay }),
      ...(effect.by === undefined ? {} : { by: effect.by }),
      ...(effect.intervalMs === undefined
        ? {}
        : { intervalMs: effect.intervalMs }),
      ...(easing === undefined ? {} : { easing }),
    };
    const parsed = spinOptionsSchema.safeParse(candidate);
    if (!parsed.success) {
      failAtSubPath(
        ctx,
        parsed.error.issues[0]?.message ?? "Invalid spin() options",
        effectPath,
      );
    }
    const fn = ctx.useImport("spin");
    const entries: [string, unknown][] = [];
    if (parsed.data.from !== undefined)
      entries.push(["from", parsed.data.from]);
    if (parsed.data.by !== undefined) entries.push(["by", parsed.data.by]);
    if (parsed.data.duration !== undefined) {
      entries.push(["duration", parsed.data.duration]);
    }
    if (parsed.data.delay !== undefined && parsed.data.delay !== 0) {
      entries.push(["delay", parsed.data.delay]);
    }
    if (parsed.data.intervalMs !== undefined) {
      entries.push(["intervalMs", parsed.data.intervalMs]);
    }
    if (parsed.data.easing !== undefined) {
      entries.push(["easing", parsed.data.easing]);
    }
    return expr(
      entries.length === 0 ? `${fn}()` : `${fn}(${ctx.formatProps(entries)})`,
    );
  }

  // Remaining 8 entrance presets
  const acceptsDirection = effect.preset === "wipe" || effect.preset === "fly";
  if (!acceptsDirection && effect.from !== undefined) {
    failAtSubPath(
      ctx,
      `${effect.preset}() does not accept from parameter`,
      `${effectPath}.from`,
    );
  }
  if (
    acceptsDirection &&
    effect.from !== undefined &&
    !motionDirectionSchema.safeParse(effect.from).success
  ) {
    failAtSubPath(
      ctx,
      `${effect.preset}() from must be a valid direction`,
      `${effectPath}.from`,
    );
  }
  const baseCandidate = {
    ...(effect.duration === 500 ? {} : { duration: effect.duration }),
    ...(effect.delay === 0 ? {} : { delay: effect.delay }),
    ...(effect.by === undefined ? {} : { by: effect.by }),
    ...(effect.intervalMs === undefined
      ? {}
      : { intervalMs: effect.intervalMs }),
    ...(easing === undefined ? {} : { easing }),
  };
  const parsedBase = effectOptionsSchema.safeParse(baseCandidate);
  if (!parsedBase.success) {
    failAtSubPath(
      ctx,
      parsedBase.error.issues[0]?.message ??
        `Invalid ${effect.preset}() options`,
      effectPath,
    );
  }

  const fn = ctx.useImport(effect.preset);
  const entries: [string, unknown][] = [];
  if (
    acceptsDirection &&
    effect.from !== undefined &&
    effect.from !== "bottom"
  ) {
    entries.push(["from", effect.from]);
  }
  if (parsedBase.data.duration !== undefined) {
    entries.push(["duration", parsedBase.data.duration]);
  }
  if (parsedBase.data.delay !== undefined && parsedBase.data.delay !== 0) {
    entries.push(["delay", parsedBase.data.delay]);
  }
  if (parsedBase.data.by !== undefined) {
    entries.push(["by", parsedBase.data.by]);
  }
  if (parsedBase.data.intervalMs !== undefined) {
    entries.push(["intervalMs", parsedBase.data.intervalMs]);
  }
  if (parsedBase.data.easing !== undefined) {
    entries.push(["easing", parsedBase.data.easing]);
  }
  return expr(
    entries.length === 0 ? `${fn}()` : `${fn}(${ctx.formatProps(entries)})`,
  );
}

function serializeEffectDeclarationExpr(
  rawDeclaration: unknown,
  expectedKind: Effect["kind"] | undefined,
  ctx: CodegenContext,
  declPath: string,
): CodegenExpr {
  if (
    isRecord(rawDeclaration) &&
    rawDeclaration["type"] === "repeated-effect"
  ) {
    const unknownKey = Object.keys(rawDeclaration).find(
      (key) => key !== "type" && key !== "effect" && key !== "count",
    );
    if (unknownKey !== undefined) {
      failAtSubPath(
        ctx,
        `repeat() contains unsupported field ${JSON.stringify(unknownKey)}`,
        `${declPath}.${unknownKey}`,
      );
    }
    const rawCount = rawDeclaration["count"];
    if (typeof rawCount === "number" && !Number.isFinite(rawCount)) {
      failAtSubPath(
        ctx,
        `Non-finite number ${String(rawCount)} cannot be expressed in TypeScript`,
        `${declPath}.count`,
      );
    }
    const parsedCount = repeatCountSchema.safeParse(rawCount);
    if (!parsedCount.success) {
      failAtSubPath(
        ctx,
        parsedCount.error.issues[0]?.message ?? "Invalid repeat count",
        `${declPath}.count`,
      );
    }
    const innerExpr = serializeEffectCall(
      rawDeclaration["effect"],
      expectedKind,
      ctx,
      `${declPath}.effect`,
    );
    const repeatFn = ctx.useImport("repeat");
    return expr(
      ctx.formatVanillaCall(repeatFn, [
        innerExpr.code,
        ctx.emitValueExpr(parsedCount.data, `${declPath}.count`),
      ]),
    );
  }
  return serializeEffectCall(rawDeclaration, expectedKind, ctx, declPath);
}

function serializeOwnerMotion(
  value: unknown,
  ctx: CodegenContext,
): CodegenExpr {
  if (ctx.path.startsWith("deck/slide[")) {
    ctx.fail(
      "unsupported-slide-property",
      "Slide does not support owner motion property",
    );
  }
  if (!isRecord(value)) {
    failAtSubPath(ctx, "OwnerMotion must be an object", ctx.path);
  }
  assertFiniteNumbersInRecord(value, ctx, ctx.path);

  const rawSteps = value["steps"];
  if (!Array.isArray(rawSteps)) {
    failAtSubPath(
      ctx,
      "OwnerMotion.steps must be an array",
      `${ctx.path}.steps`,
    );
  }

  const parsedOwner = ownerMotionSchema.safeParse(value);
  if (!parsedOwner.success) {
    const issue = parsedOwner.error.issues[0];
    const issuePath =
      issue !== undefined && issue.path.length > 0
        ? `${ctx.path}.${issue.path.join(".")}`
        : ctx.path;
    failAtSubPath(
      ctx,
      issue?.message ?? "OwnerMotion declaration is invalid",
      issuePath,
    );
  }

  const serializedSteps = parsedOwner.data.steps.map((step, stepIdx) => ({
    trigger: step.trigger,
    effects: step.effects.map((item, effectIdx) => {
      const effectExpr = serializeEffectDeclarationExpr(
        item.effect,
        undefined,
        ctx,
        `${ctx.path}.steps[${stepIdx}].effects[${effectIdx}].effect`,
      );
      return {
        target: item.target,
        effect: effectExpr,
        ...(item.range === undefined ? {} : { range: item.range }),
      };
    }),
  }));

  return expr(ctx.formatProps([["steps", serializedSteps]]));
}

function serializeMotionValue(
  value: unknown,
  ctx: CodegenContext,
): CodegenExpr | undefined {
  if (!isRecord(value)) return undefined;
  const type = value["type"];
  if (type === "motion-moment") {
    return serializeMotionMomentValue(value, ctx, "moment");
  }
  if (type === "motion-offset") {
    return serializeMotionTimingValue(value, ctx, "timing");
  }
  if (type === "effect" || type === "repeated-effect") {
    return serializeEffectDeclarationExpr(value, undefined, ctx, ctx.path);
  }
  return undefined;
}

/**
 * Source codegen contribution for `motion-step`, `motion-scope`, object motion properties
 * (`enter`, `emphasis`, `motion`), and owner-level `OwnerMotion`.
 */
export const motionCodegen = defineObjectCodegen({
  declarations: {
    "motion-step": {
      emitVanilla(node, ctx) {
        const callee = ctx.useImport("step");
        const propsEntries = collectStepProps(node, ctx);
        const args =
          propsEntries.length === 0 ? [] : [ctx.formatProps(propsEntries)];
        const children = emitContainerChildren(node, ctx, "Step");
        return ctx.formatVanillaCall(callee, args, children);
      },
      emitReact(node, ctx) {
        const tag = ctx.useImport("Step");
        const propsEntries = collectStepProps(node, ctx);
        const children = emitContainerChildren(node, ctx, "Step");
        return ctx.formatJsxElement(tag, propsEntries, children);
      },
    },
    "motion-scope": {
      emitVanilla(node, ctx) {
        const callee = ctx.useImport("motionScope");
        const { scopeVar, optionEntries } = collectScopeProps(node, ctx);
        const args =
          optionEntries.length === 0
            ? [scopeVar]
            : [scopeVar, ctx.formatProps(optionEntries)];
        const children = emitContainerChildren(node, ctx, "MotionScope");
        return ctx.formatVanillaCall(callee, args, children);
      },
      emitReact(node, ctx) {
        const tag = ctx.useImport("MotionScope");
        const { scopeVar, optionEntries } = collectScopeProps(node, ctx);
        const children = emitContainerChildren(node, ctx, "MotionScope");
        return ctx.formatJsxElement(
          tag,
          [["scope", expr(scopeVar)], ...optionEntries],
          children,
        );
      },
    },
  },
  objectProps: {
    enter(value, ctx) {
      if (Array.isArray(value)) {
        const items = value.map((item, index) =>
          serializeEffectDeclarationExpr(
            item,
            "entrance",
            ctx,
            `${ctx.path}[${index}]`,
          ),
        );
        return expr(ctx.emitValueExpr(items));
      }
      return serializeEffectDeclarationExpr(value, "entrance", ctx, ctx.path);
    },
    emphasis(value, ctx) {
      return serializeEffectDeclarationExpr(value, "emphasis", ctx, ctx.path);
    },
    motion(value, ctx) {
      if (typeof value !== "boolean") {
        ctx.fail(
          "unexpressible-value",
          "Object motion property must be a boolean",
        );
      }
      return expr(value ? "true" : "false");
    },
  },
  ownerProps: {
    motion: serializeOwnerMotion,
  },
  valueSerializers: [serializeMotionValue],
});
