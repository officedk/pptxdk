import type { Font } from "fontkit";

export function fontFamilyKey(value: string): string {
  return value.trim().replace(/\s+/gu, " ").toLocaleLowerCase("en-US");
}

/** Non-RIBBI weights can retain legacy families while declaring a unified preferred family via name ID 16. */
export function matchesFontFamily(font: Font, family: string): boolean {
  const preferredFamily = font.getName("preferredFamily", "en");
  return [font.familyName, preferredFamily].some(
    (candidate) =>
      candidate !== null && fontFamilyKey(candidate) === fontFamilyKey(family),
  );
}
