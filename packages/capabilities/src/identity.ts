import { xmlEscape, type NativeLink } from "@react-pptx/core";

/** 原生对象能力共用同一种点击字段表达，关系与动作由 core 预先解析。 */
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
