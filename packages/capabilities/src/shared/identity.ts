import {
  CapabilityError,
  isMorphKey,
  parseLinkTarget,
  type LinkTarget,
  type MorphKey,
  type NativeLink,
} from "@pptxdk/core/capability";
import { xmlEscape } from "@pptxdk/emit";

export { parseLinkTarget };
export type { LinkTarget };

/** Validates an optional non-empty object name parameter. */
export function parseOptionalName(
  value: unknown,
  path?: string,
): string | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== "string" || value.trim() === "") {
    throw new CapabilityError(
      "invalid-parameter",
      "name must be a non-empty string",
      path,
    );
  }
  return value;
}

/** Validates an optional Morph handle created by createMorphScope().handle(name). */
export function parseMorphKey(
  value: unknown,
  path?: string,
): MorphKey | undefined {
  if (value === undefined) return undefined;
  if (!isMorphKey(value)) {
    throw new CapabilityError(
      "invalid-morph-key",
      "morphKey must be created by createMorphScope().handle(name)",
      path,
    );
  }
  return value;
}

/** Native object capabilities share the same click link representation, with relationships and actions pre-resolved by core. */
export function nativeLinkXml(link: NativeLink | undefined): string {
  if (link === undefined) return "";
  const attributes = [
    ...(link.relationshipId === undefined
      ? []
      : [`r:id="${xmlEscape(link.relationshipId)}"`]),
    ...(link.action === undefined
      ? []
      : [`action="${xmlEscape(link.action)}"`]),
  ];
  return `<a:hlinkClick ${attributes.join(" ")}/>`;
}
