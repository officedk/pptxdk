import type { Font } from "fontkit";

export function fontFamilyKey(value: string): string {
  return value.trim().replace(/\s+/gu, " ").toLocaleLowerCase("en-US");
}

/** 非 RIBBI 字重可保留旧式字族，同时通过 name ID 16 声明统一的优先字族。 */
export function matchesFontFamily(font: Font, family: string): boolean {
  const preferredFamily = font.getName("preferredFamily", "en");
  return [font.familyName, preferredFamily].some(
    (candidate) =>
      candidate !== null && fontFamilyKey(candidate) === fontFamilyKey(family),
  );
}
