import {
  defineObjectCodegen,
  type CodegenContext,
  type CodegenExpr,
} from "@pptxdk/codegen";
import {
  CapabilityError,
  readInternalDeclarationProps,
} from "@pptxdk/core/capability";
import { CommonNode } from "@pptxdk/core/document";
import type { z } from "zod";

import { motionPropKeySet } from "../motion/motion-schemas.ts";
import {
  chartAllowedPropKeys,
  chartAxesSchema,
  chartDisplayBlanksAsSchema,
  chartLegendSchema,
  chartPlotAreaSchema,
  chartStyleSchema,
  chartTitleSchema,
  parseChartGroupBySchema,
} from "./chart-schemas.ts";

const chartAllowedPropKeySet = new Set(chartAllowedPropKeys);
const objectMotionPropKeys = motionPropKeySet;

function isPlainRecord(value: unknown): value is Record<string, unknown> {
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    return false;
  }
  const proto = Object.getPrototypeOf(value);
  return proto === Object.prototype || proto === null;
}

function validatePropWithSchema(
  schema: z.ZodType,
  value: unknown,
  propKey: string,
  ctx: CodegenContext,
): void {
  if (value === undefined) return;
  const parsed = schema.safeParse(value);
  if (!parsed.success) {
    const issue = parsed.error.issues[0];
    ctx.fail(
      "unexpressible-value",
      issue?.message ?? `Chart.${propKey} cannot be expressed`,
      `@${propKey}`,
    );
  }
}

function validateChartStructure(
  props: Readonly<Record<string, unknown>>,
  ctx: CodegenContext,
): void {
  const rawGroups = props["groups"];
  if (!Array.isArray(rawGroups) || rawGroups.length === 0) {
    ctx.fail(
      "unexpressible-value",
      "Chart.groups must be a non-empty array",
      "@groups",
    );
  }
  for (const [index, rawGroup] of rawGroups.entries()) {
    try {
      parseChartGroupBySchema(rawGroup, `Chart.groups[${index}]`);
    } catch (error) {
      const message =
        error instanceof CapabilityError
          ? error.message
          : `Chart.groups[${index}] is invalid`;
      ctx.fail("unexpressible-value", message, `@groups[${index}]`);
    }
  }

  const nativePh = props["nativePlaceholder"];
  if (
    isPlainRecord(nativePh) &&
    Object.hasOwn(nativePh, "sz") &&
    nativePh["sz"] !== undefined
  ) {
    ctx.fail(
      "unexpressible-value",
      "Chart.nativePlaceholder.sz is no longer supported; use size",
      "@nativePlaceholder.sz",
    );
  }
  validatePropWithSchema(chartAxesSchema, props["axes"], "axes", ctx);
  validatePropWithSchema(chartTitleSchema, props["title"], "title", ctx);
  validatePropWithSchema(
    chartDisplayBlanksAsSchema,
    props["displayBlanksAs"],
    "displayBlanksAs",
    ctx,
  );
  validatePropWithSchema(chartLegendSchema, props["legend"], "legend", ctx);
  validatePropWithSchema(
    chartPlotAreaSchema,
    props["plotArea"],
    "plotArea",
    ctx,
  );
  validatePropWithSchema(chartStyleSchema, props["style"], "style", ctx);
}

function collectChartProps(
  node: CommonNode,
  ctx: CodegenContext,
): readonly (readonly [string, unknown])[] {
  if (node.children.length > 0) {
    ctx.fail("invalid-document", "Chart declaration cannot have children");
  }

  const props = readInternalDeclarationProps(node);
  const extra: Readonly<Record<string, CodegenExpr>> =
    ctx.emitObjectExtraProps(props);
  const entries: [string, unknown][] = [];

  for (const [key, value] of Object.entries(props)) {
    if (value === undefined) continue;
    if (!chartAllowedPropKeySet.has(key)) {
      ctx.fail(
        "invalid-document",
        `chart declaration does not support property ${JSON.stringify(key)}`,
        `@${key}`,
      );
    }
    if (objectMotionPropKeys.has(key)) {
      if (Object.hasOwn(extra, key)) {
        entries.push([key, extra[key]]);
      } else {
        ctx.fail(
          "unexpressible-value",
          `Object property ${JSON.stringify(key)} requires motion codegen support`,
          `@${key}`,
        );
      }
      continue;
    }
    entries.push([key, value]);
  }

  return entries;
}

function emitChartVanilla(node: CommonNode, ctx: CodegenContext): string {
  const callee = ctx.useImport("chart");
  const propsEntries = collectChartProps(node, ctx);
  const formattedProps = ctx.formatProps(propsEntries);
  validateChartStructure(readInternalDeclarationProps(node), ctx);
  return ctx.formatVanillaCall(callee, [formattedProps]);
}

function emitChartReact(node: CommonNode, ctx: CodegenContext): string {
  const tag = ctx.useImport("Chart");
  const propsEntries = collectChartProps(node, ctx);
  const jsx = ctx.formatJsxElement(tag, propsEntries, []);
  validateChartStructure(readInternalDeclarationProps(node), ctx);
  return jsx;
}

/** Source codegen contribution for `chart` (`chart(props)` / `<Chart {...props} />`) declarations. */
export const chartCodegen = defineObjectCodegen({
  declarations: {
    chart: {
      emitVanilla: emitChartVanilla,
      emitReact: emitChartReact,
    },
  },
});
