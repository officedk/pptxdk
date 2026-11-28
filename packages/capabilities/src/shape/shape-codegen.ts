import {
  defineObjectCodegen,
  expr,
  type CodegenContext,
  type CodegenExpr,
} from "@pptxdk/codegen";
import type { DeclarationText } from "@pptxdk/core/build";
import { readInternalDeclarationProps } from "@pptxdk/core/capability";
import { CommonNode } from "@pptxdk/core/document";

import { serializeImageDescriptorValue } from "../image/image-codegen.ts";
import { motionPropKeySet } from "../motion/motion-schemas.ts";
import {
  shapeAllowedPropKeyList,
  shapeStyleSchema,
} from "./shape-declaration.ts";

const objectMotionPropKeys = motionPropKeySet;

const shapeAllowedPropKeys = new Set(shapeAllowedPropKeyList);

const shapeAllowedChildKinds = new Set([
  "paragraph",
  "span",
  "slide-number",
  "br",
  "motion-step",
]);

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function isPlainRecord(value: unknown): value is Record<string, unknown> {
  if (!isRecord(value)) return false;
  const proto = Object.getPrototypeOf(value);
  return proto === Object.prototype || proto === null;
}

function singleTextLeaf(node: CommonNode): string | undefined {
  if (node.children.length !== 1) return undefined;
  const only = node.children[0];
  if (only === undefined || only instanceof CommonNode) return undefined;
  const leaf = only as DeclarationText;
  return leaf.type === "#text" ? leaf.text : undefined;
}

function assertAllowedShapeChildren(
  node: CommonNode,
  ctx: CodegenContext,
): void {
  let childIndex = 0;
  for (const child of node.children) {
    if (!(child instanceof CommonNode)) continue;
    const label =
      child.id === undefined
        ? `${child.type}[${childIndex + 1}]`
        : `${child.type}[${child.id}]`;
    if (!shapeAllowedChildKinds.has(child.type)) {
      ctx.fail(
        "invalid-document",
        `shape declaration cannot contain ${JSON.stringify(child.type)} child`,
        `${ctx.path}/${label}`,
      );
    }
    childIndex += 1;
  }
}

function formatSvgNumber(
  value: unknown,
  ctx: CodegenContext,
  propPath: string,
): string {
  if (typeof value !== "number" || !Number.isFinite(value)) {
    ctx.fail(
      "unexpressible-value",
      `Non-finite geometry coordinate ${String(value)} cannot be expressed in source`,
      propPath,
    );
  }
  return Object.is(value, -0) ? "0" : String(value);
}

function commandsToSvgPath(
  commands: readonly unknown[],
  ctx: CodegenContext,
  basePath: string,
): string | undefined {
  if (commands.length === 0) return undefined;
  const segments: string[] = [];
  for (let index = 0; index < commands.length; index += 1) {
    const command = commands[index];
    if (!isRecord(command) || typeof command["op"] !== "string") {
      return undefined;
    }
    const cmdPath = `${basePath}.commands[${index}]`;
    switch (command["op"]) {
      case "M":
      case "L": {
        if (index === 0 && command["op"] !== "M") return undefined;
        const x = formatSvgNumber(command["x"], ctx, `${cmdPath}.x`);
        const y = formatSvgNumber(command["y"], ctx, `${cmdPath}.y`);
        segments.push(`${command["op"]} ${x} ${y}`);
        break;
      }
      case "C": {
        if (index === 0) return undefined;
        const x1 = formatSvgNumber(command["x1"], ctx, `${cmdPath}.x1`);
        const y1 = formatSvgNumber(command["y1"], ctx, `${cmdPath}.y1`);
        const x2 = formatSvgNumber(command["x2"], ctx, `${cmdPath}.x2`);
        const y2 = formatSvgNumber(command["y2"], ctx, `${cmdPath}.y2`);
        const x = formatSvgNumber(command["x"], ctx, `${cmdPath}.x`);
        const y = formatSvgNumber(command["y"], ctx, `${cmdPath}.y`);
        segments.push(`C ${x1} ${y1} ${x2} ${y2} ${x} ${y}`);
        break;
      }
      case "Q": {
        if (index === 0) return undefined;
        const x1 = formatSvgNumber(command["x1"], ctx, `${cmdPath}.x1`);
        const y1 = formatSvgNumber(command["y1"], ctx, `${cmdPath}.y1`);
        const x = formatSvgNumber(command["x"], ctx, `${cmdPath}.x`);
        const y = formatSvgNumber(command["y"], ctx, `${cmdPath}.y`);
        segments.push(`Q ${x1} ${y1} ${x} ${y}`);
        break;
      }
      case "Z":
        if (index === 0) return undefined;
        segments.push("Z");
        break;
      default:
        return undefined;
    }
  }
  return segments.join(" ");
}

function trySerializeCustomAsPath(
  value: Record<string, unknown>,
  ctx: CodegenContext,
): CodegenExpr | undefined {
  const width = value["width"];
  const height = value["height"];
  const paths = value["paths"];
  if (
    typeof width !== "number" ||
    !Number.isFinite(width) ||
    width <= 0 ||
    typeof height !== "number" ||
    !Number.isFinite(height) ||
    height <= 0 ||
    !Array.isArray(paths) ||
    paths.length === 0
  ) {
    return undefined;
  }

  const firstPath = paths[0];
  if (!isRecord(firstPath)) return undefined;
  const uniformFill = firstPath["fill"];
  const uniformStroke = firstPath["stroke"];

  const svgStrings: string[] = [];
  for (let index = 0; index < paths.length; index += 1) {
    const entry = paths[index];
    if (!isRecord(entry) || !Array.isArray(entry["commands"])) {
      return undefined;
    }
    if (entry["fill"] !== uniformFill || entry["stroke"] !== uniformStroke) {
      return undefined;
    }
    const svg = commandsToSvgPath(
      entry["commands"],
      ctx,
      `@geometry.paths[${index}]`,
    );
    if (svg === undefined) return undefined;
    svgStrings.push(svg);
  }

  const pathFn = ctx.useImport("path");
  const dataExpr =
    svgStrings.length === 1
      ? JSON.stringify(svgStrings[0])
      : ctx.emitValueExpr(svgStrings, "@geometry.paths");
  const optionsObj: Record<string, unknown> = {
    viewBox: [width, height],
    ...(uniformFill === undefined ? {} : { fill: uniformFill }),
    ...(uniformStroke === undefined ? {} : { stroke: uniformStroke }),
  };
  return expr(
    `${pathFn}(${dataExpr}, ${ctx.emitValueExpr(optionsObj, "@geometry")})`,
  );
}

export function serializeGeometryProp(
  value: unknown,
  ctx: CodegenContext,
): unknown {
  if (typeof value === "string") return value;
  if (!isRecord(value)) {
    return expr(ctx.emitValueExpr(value, "@geometry"));
  }
  if (value["type"] === "preset" && typeof value["name"] === "string") {
    const presetFn = ctx.useImport("preset");
    const adjust = value["adjust"];
    if (adjust !== undefined) {
      if (!isRecord(adjust)) {
        ctx.fail(
          "unexpressible-value",
          "Preset geometry adjust must be a plain record",
          "@geometry.adjust",
        );
      }
      const definedAdjustEntries = Object.entries(adjust).filter(
        ([, entryValue]) => entryValue !== undefined,
      );
      if (definedAdjustEntries.length > 0) {
        return expr(
          `${presetFn}(${JSON.stringify(value["name"])}, ${ctx.emitValueExpr(adjust, "@geometry.adjust")})`,
        );
      }
    }
    return expr(`${presetFn}(${JSON.stringify(value["name"])})`);
  }
  if (value["type"] === "custom") {
    const asPath = trySerializeCustomAsPath(value, ctx);
    if (asPath !== undefined) return asPath;
  }
  return expr(ctx.emitValueExpr(value, "@geometry"));
}

function serializeGradientValue(
  value: unknown,
  ctx: CodegenContext,
): CodegenExpr | undefined {
  if (!isRecord(value)) return undefined;
  const type = value["type"];
  if (type === "linear-gradient") {
    const angle = value["angle"];
    const stops = value["stops"];
    if (typeof angle !== "number" || !Array.isArray(stops)) return undefined;
    const fn = ctx.useImport("linearGradient");
    const args = [
      ctx.emitValueExpr(angle, "@angle"),
      ...stops.map((stop, idx) => ctx.emitValueExpr(stop, `@stops[${idx}]`)),
    ];
    return expr(`${fn}(${args.join(", ")})`);
  }
  if (type === "radial-gradient") {
    const stops = value["stops"];
    if (!Array.isArray(stops)) return undefined;
    const fn = ctx.useImport("radialGradient");
    const args = stops.map((stop, idx) =>
      ctx.emitValueExpr(stop, `@stops[${idx}]`),
    );
    return expr(`${fn}(${args.join(", ")})`);
  }
  return undefined;
}

function serializeShapeStyleProp(
  style: unknown,
  ctx: CodegenContext,
): CodegenExpr {
  if (!isPlainRecord(style)) {
    ctx.fail(
      "unexpressible-value",
      "Shape.style must be a plain object",
      "@style",
    );
  }
  const parsed = shapeStyleSchema.safeParse(style);
  if (!parsed.success) {
    const issue = parsed.error.issues[0];
    const issuePath =
      issue !== undefined && issue.path.length > 0
        ? `@style.${issue.path.join(".")}`
        : "@style";
    ctx.fail(
      "unexpressible-value",
      issue?.message ?? "Shape.style cannot be expressed",
      issuePath,
    );
  }
  const entries: [string, unknown][] = [];
  for (const [key, entryValue] of Object.entries(style)) {
    if (entryValue === undefined) continue;
    if (key === "background") {
      const serializedImage = serializeImageDescriptorValue(entryValue, ctx);
      entries.push([key, serializedImage ?? entryValue]);
      continue;
    }
    entries.push([key, entryValue]);
  }
  return expr(ctx.formatProps(entries));
}

function collectShapeProps(
  node: CommonNode,
  ctx: CodegenContext,
): readonly (readonly [string, unknown])[] {
  const props = readInternalDeclarationProps(node);
  const extra = ctx.emitObjectExtraProps(props);
  const entries: [string, unknown][] = [];
  for (const [key, value] of Object.entries(props)) {
    if (value === undefined) continue;
    if (!shapeAllowedPropKeys.has(key)) {
      ctx.fail(
        "invalid-document",
        `shape declaration does not support property ${JSON.stringify(key)}`,
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
    if (key === "geometry") {
      entries.push([key, serializeGeometryProp(value, ctx)]);
      continue;
    }
    if (key === "style") {
      entries.push([key, serializeShapeStyleProp(value, ctx)]);
      continue;
    }
    entries.push([key, value]);
  }
  return entries;
}

function emitShapeVanilla(node: CommonNode, ctx: CodegenContext): string {
  assertAllowedShapeChildren(node, ctx);
  const callee = ctx.useImport("shape");
  const propsEntries = collectShapeProps(node, ctx);
  const leafText = singleTextLeaf(node);
  if (leafText !== undefined) {
    const args =
      propsEntries.length === 0
        ? [JSON.stringify(leafText)]
        : [JSON.stringify(leafText), ctx.formatProps(propsEntries)];
    return ctx.formatVanillaCall(callee, args);
  }
  if (node.children.length === 0) {
    const args =
      propsEntries.length === 0 ? [] : [ctx.formatProps(propsEntries)];
    return ctx.formatVanillaCall(callee, args);
  }
  const children = ctx.emitRichTextChildren(node.children);
  const contentArg = ctx.emitValueExpr(children.map((code) => expr(code)));
  const args =
    propsEntries.length === 0
      ? [contentArg]
      : [contentArg, ctx.formatProps(propsEntries)];
  return ctx.formatVanillaCall(callee, args);
}

function emitShapeReact(node: CommonNode, ctx: CodegenContext): string {
  assertAllowedShapeChildren(node, ctx);
  const tag = ctx.useImport("Shape");
  const propsEntries = collectShapeProps(node, ctx);
  const children =
    node.children.length === 0 ? [] : ctx.emitRichTextChildren(node.children);
  return ctx.formatJsxElement(tag, propsEntries, children);
}

/** Source codegen contribution for `shape` declarations and DrawingML gradient values. */
export const shapeCodegen = defineObjectCodegen({
  declarations: {
    shape: {
      emitVanilla: emitShapeVanilla,
      emitReact: emitShapeReact,
    },
  },
  valueSerializers: [serializeGradientValue],
});
