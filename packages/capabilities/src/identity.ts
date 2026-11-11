import { xmlEscape, type NativeLink } from "@react-pptx/core/capability";

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
