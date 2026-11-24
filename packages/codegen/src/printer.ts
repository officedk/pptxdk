import { isSchemeColorName } from "@pptxdk/core";

import { isCodegenExpr, type CodegenContext } from "./types.ts";

const identifierKeyPattern = /^[A-Za-z_$][A-Za-z0-9_$]*$/u;
const multilineThreshold = 72;

export function indentBlock(text: string, spaces = 2): string {
  const prefix = " ".repeat(spaces);
  return text
    .split("\n")
    .map((line) => (line === "" ? "" : `${prefix}${line}`))
    .join("\n");
}

function isPlainRecord(value: unknown): value is Record<string, unknown> {
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    return false;
  }
  const proto = Object.getPrototypeOf(value);
  return proto === Object.prototype || proto === null;
}

function formatPropertyKey(key: string): string {
  return identifierKeyPattern.test(key) ? key : JSON.stringify(key);
}

function formatNumberLiteral(value: number): string {
  return Object.is(value, -0) ? "-0" : String(value);
}

function tryFormatThemeShade(
  record: Record<string, unknown>,
  ctx: CodegenContext,
): string | undefined {
  const keys = Object.keys(record).filter((key) => record[key] !== undefined);
  if (keys.length < 2 || keys.length > 3) return undefined;
  const slot = record["slot"];
  const lumMod = record["lumMod"];
  const lumOff = record["lumOff"];
  if (
    typeof slot !== "string" ||
    !isSchemeColorName(slot) ||
    typeof lumMod !== "number" ||
    !Number.isFinite(lumMod)
  ) {
    return undefined;
  }
  for (const key of keys) {
    if (key !== "slot" && key !== "lumMod" && key !== "lumOff") {
      return undefined;
    }
  }
  if (
    typeof lumOff === "number" &&
    Number.isFinite(lumOff) &&
    Math.abs(lumMod + lumOff - 1) < 1e-12
  ) {
    return `${ctx.useImport("lighter")}(${JSON.stringify(slot)}, ${formatNumberLiteral(lumOff)})`;
  }
  if (lumOff === undefined) {
    const amount = Number((1 - lumMod).toFixed(12));
    if (1 - amount === lumMod) {
      return `${ctx.useImport("darker")}(${JSON.stringify(slot)}, ${formatNumberLiteral(amount)})`;
    }
  }
  return undefined;
}

export type ValueHooks = {
  readonly serializeSpecial: (
    value: unknown,
    propPath: string,
  ) => string | undefined;
};

/**
 * Formats an arbitrary author value as a TypeScript expression, throwing `CodegenError`
 * via `ctx.fail("unexpressible-value", ...)` on non-finite numbers or unserializable objects.
 */
export function formatTsValue(
  value: unknown,
  ctx: CodegenContext,
  hooks: ValueHooks,
  propPath = "",
): string {
  if (isCodegenExpr(value)) {
    return value.code;
  }
  if (value === null) return "null";
  if (value === undefined) return "undefined";
  if (typeof value === "boolean") return value ? "true" : "false";
  if (typeof value === "string") return JSON.stringify(value);
  if (typeof value === "number") {
    if (!Number.isFinite(value)) {
      ctx.fail(
        "unexpressible-value",
        `Non-finite number ${String(value)} cannot be expressed in TypeScript`,
        propPath,
      );
    }
    return formatNumberLiteral(value);
  }
  if (typeof value === "bigint" || typeof value === "symbol") {
    ctx.fail(
      "unexpressible-value",
      `Unsupported ${typeof value} value cannot be expressed in TypeScript`,
      propPath,
    );
  }
  if (typeof value === "function") {
    ctx.fail(
      "unexpressible-value",
      "Function value cannot be expressed in generated source",
      propPath,
    );
  }

  const special = hooks.serializeSpecial(value, propPath);
  if (special !== undefined) {
    return special;
  }

  if (value instanceof URL) {
    return `new URL(${JSON.stringify(value.href)})`;
  }

  if (Array.isArray(value)) {
    if (value.length === 0) return "[]";
    const items = value.map((item, index) =>
      formatTsValue(item, ctx, hooks, `${propPath}[${index}]`),
    );
    const singleLine = `[${items.join(", ")}]`;
    if (!singleLine.includes("\n") && singleLine.length <= multilineThreshold) {
      return singleLine;
    }
    return `[\n${indentBlock(items.join(",\n"))},\n]`;
  }

  if (!isPlainRecord(value)) {
    ctx.fail(
      "unexpressible-value",
      `Unsupported non-plain object value cannot be expressed in generated source`,
      propPath,
    );
  }

  const shadeExpr = tryFormatThemeShade(value, ctx);
  if (shadeExpr !== undefined) {
    return shadeExpr;
  }

  const entries = Object.entries(value).filter(
    ([, entryValue]) => entryValue !== undefined,
  );
  if (entries.length === 0) return "{}";
  const formattedEntries = entries.map(([key, entryValue]) => {
    const nextPath = propPath === "" ? `@${key}` : `${propPath}.${key}`;
    return `${formatPropertyKey(key)}: ${formatTsValue(entryValue, ctx, hooks, nextPath)}`;
  });
  const singleLine = `{ ${formattedEntries.join(", ")} }`;
  if (!singleLine.includes("\n") && singleLine.length <= multilineThreshold) {
    return singleLine;
  }
  return `{\n${indentBlock(formattedEntries.join(",\n"))},\n}`;
}

export function formatPropsObject(
  entries: readonly (readonly [string, unknown])[],
  ctx: CodegenContext,
  hooks: ValueHooks,
  basePath = "",
): string {
  const defined = entries.filter(([, value]) => value !== undefined);
  if (defined.length === 0) return "{}";
  const formatted = defined.map(([key, value]) => {
    const nextPath = basePath === "" ? `@${key}` : `${basePath}.${key}`;
    return `${formatPropertyKey(key)}: ${formatTsValue(value, ctx, hooks, nextPath)}`;
  });
  const singleLine = `{ ${formatted.join(", ")} }`;
  if (!singleLine.includes("\n") && singleLine.length <= multilineThreshold) {
    return singleLine;
  }
  return `{\n${indentBlock(formatted.join(",\n"))},\n}`;
}

const bareJsxUnsafeChars = /[&{}<>\r\n\t]/u;

/**
 * Formats a `#text` declaration leaf for JSX children without losing leading/trailing spaces,
 * repeated spaces, line breaks, or JSX-reserved characters.
 */
export function formatJsxTextChild(text: string): string {
  if (
    text.length > 0 &&
    text.trim() === text &&
    !bareJsxUnsafeChars.test(text) &&
    !/\s{2,}/u.test(text)
  ) {
    return text;
  }
  return `{${JSON.stringify(text)}}`;
}

function isSimpleJsxStringAttr(value: string): boolean {
  if (/["\r\n\t]/u.test(value)) return false;
  for (const ch of value) {
    const code = ch.codePointAt(0);
    if (code !== undefined && (code < 0x20 || code === 0x7f)) {
      return false;
    }
  }
  return true;
}

export function formatJsxAttributes(
  entries: readonly (readonly [string, unknown])[],
  ctx: CodegenContext,
  hooks: ValueHooks,
  basePath = "",
): readonly string[] {
  const attrs: string[] = [];
  for (const [key, value] of entries) {
    if (value === undefined) continue;
    const nextPath = basePath === "" ? `@${key}` : `${basePath}.${key}`;
    if (isCodegenExpr(value)) {
      attrs.push(`${key}={${value.code}}`);
      continue;
    }
    if (value === true) {
      attrs.push(key);
      continue;
    }
    if (typeof value === "string" && isSimpleJsxStringAttr(value)) {
      attrs.push(`${key}=${JSON.stringify(value)}`);
      continue;
    }
    attrs.push(`${key}={${formatTsValue(value, ctx, hooks, nextPath)}}`);
  }
  return attrs;
}

export function formatJsxElement(
  tag: string,
  entries: readonly (readonly [string, unknown])[],
  children: readonly string[],
  ctx: CodegenContext,
  hooks: ValueHooks,
  basePath = "",
): string {
  const attrs = formatJsxAttributes(entries, ctx, hooks, basePath);
  const inlineAttrs = attrs.join(" ");
  const hasMultilineAttr = attrs.some((attr) => attr.includes("\n"));
  const openTag =
    attrs.length === 0
      ? `<${tag}`
      : !hasMultilineAttr && tag.length + inlineAttrs.length + 2 <= 80
        ? `<${tag} ${inlineAttrs}`
        : `<${tag}\n${indentBlock(attrs.join("\n"))}\n`;

  if (children.length === 0) {
    return openTag.endsWith("\n") ? `${openTag}/>` : `${openTag} />`;
  }

  const head = openTag.endsWith("\n") ? `${openTag}>` : `${openTag}>`;
  if (
    children.length === 1 &&
    !children[0]!.includes("\n") &&
    !head.includes("\n") &&
    head.length + children[0]!.length + tag.length + 4 <= 80
  ) {
    return `${head}${children[0]}</${tag}>`;
  }
  return `${head}\n${indentBlock(children.join("\n"))}\n</${tag}>`;
}

export function formatVanillaCall(
  callee: string,
  args: readonly string[],
  children: readonly string[],
): string {
  const callArgs = args.join(", ");
  const hasMultilineArg = args.some((arg) => arg.includes("\n"));
  const baseCall =
    args.length === 0
      ? `${callee}()`
      : !hasMultilineArg && callee.length + callArgs.length + 2 <= 80
        ? `${callee}(${callArgs})`
        : `${callee}(\n${indentBlock(args.join(",\n"))},\n)`;

  if (children.length === 0) {
    return baseCall;
  }
  const singleLineChildren = children.join(", ");
  if (
    !baseCall.includes("\n") &&
    !singleLineChildren.includes("\n") &&
    baseCall.length + singleLineChildren.length + 10 <= 80
  ) {
    return `${baseCall}.append(${singleLineChildren})`;
  }
  return `${baseCall}.append(\n${indentBlock(children.join(",\n"))},\n)`;
}
