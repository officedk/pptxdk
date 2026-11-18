import {
  CapabilityError,
  xmlEscape,
  type LinkTarget,
  type NativeLink,
} from "@react-pptx/core/capability";

/** Validates link target declarations once during capability bind. */
export function parseLinkTarget(
  value: unknown,
  path?: string,
): LinkTarget | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    throw new CapabilityError(
      "invalid-link",
      "link must specify id, url, or home action",
      path,
    );
  }
  const entries = Object.entries(value).filter(
    ([, target]) => target !== undefined,
  );
  if (entries.length !== 1) {
    throw new CapabilityError(
      "invalid-link",
      "link must specify exactly one target",
      path,
    );
  }
  const [kind, target] = entries[0] ?? [];
  if (kind === "id" && typeof target === "string") {
    return { id: target };
  }
  if (kind === "url" && typeof target === "string") {
    return { url: target };
  }
  if (kind === "action" && target === "home") {
    return { action: "home" };
  }
  throw new CapabilityError(
    "invalid-link",
    "link supports { id }, { url }, or { action: 'home' }",
    path,
  );
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
