import { defineSlideLayout } from "@officedk/pptxdk/react";

import { colors } from "../theme.ts";

function headings(ink: string, secondary: string) {
  return {
    chapter: {
      type: "body" as const,
      idx: 1,
      x: 64,
      y: 30,
      width: 700,
      height: 26,
      style: { fontSize: 14, color: secondary, letterSpacing: 1 },
    },
    title: {
      type: "title" as const,
      x: 64,
      y: 70,
      width: 1110,
      height: 70,
      style: {
        fontFamily: "display",
        fontSize: 43,
        color: ink,
        lineHeight: 54,
        textWrap: "nowrap" as const,
      },
    },
    source: {
      type: "body" as const,
      idx: 2,
      x: 64,
      y: 676,
      width: 1060,
      height: 22,
      style: { fontSize: 11, color: secondary, textWrap: "nowrap" as const },
    },
  };
}

export const coverLayout = defineSlideLayout({
  id: "cover",
  name: "封面",
  type: "title",
  background: colors.indigo,
  showMasterShapes: false,
  headerFooter: { slideNumber: false },
  placeholders: {
    title: {
      type: "ctrTitle",
      x: 68,
      y: 154,
      width: 650,
      height: 206,
      style: {
        fontFamily: "display",
        fontSize: 76,
        lineHeight: 98,
        color: colors.paperBright,
      },
    },
    subtitle: {
      type: "subTitle",
      idx: 1,
      x: 72,
      y: 390,
      width: 540,
      height: 48,
      style: { fontSize: 27, color: colors.sun },
    },
  },
});

export const pageNumber = {
  idx: 4,
  x: 1130,
  y: 52,
  width: 66,
  height: 34,
  style: {
    fontFamily: "display",
    fontSize: 22,
    color: colors.muted,
    align: "right" as const,
  },
};

export const contentLayout = defineSlideLayout({
  id: "content",
  name: "浅色正文",
  type: "obj",
  background: colors.paper,
  headerFooter: { slideNumber: true },
  footerPlaceholders: { slideNumber: pageNumber },
  placeholders: {
    ...headings(colors.ink, colors.muted),
    content: { type: "obj", idx: 3, x: 64, y: 150, width: 1152, height: 510 },
  },
});

export const summaryLayout = defineSlideLayout({
  id: "summary",
  name: "深色总结",
  type: "cust",
  background: colors.indigo,
  headerFooter: { slideNumber: true },
  footerPlaceholders: {
    slideNumber: {
      ...pageNumber,
      style: { ...pageNumber.style, color: colors.mist },
    },
  },
  placeholders: headings(colors.paperBright, colors.mist),
});
