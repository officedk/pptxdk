import type { CapabilityInput } from "@pptxdk/core/capability";

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

export function layoutStyleValue(
  input:
    | CapabilityInput
    | { readonly props: Readonly<Record<string, unknown>> },
  name: "width" | "height",
): number | "auto" | `${number}%` | undefined {
  const value = input.props["style"];
  if (!isRecord(value)) return undefined;
  const raw = value[name];
  return typeof raw === "number" ||
    raw === "auto" ||
    (typeof raw === "string" && raw.endsWith("%"))
    ? (raw as number | "auto" | `${number}%`)
    : undefined;
}
