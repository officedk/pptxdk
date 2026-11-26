import { defineSlideLayout } from "@officedk/pptxdk/react";

import { colors } from "../theme.ts";

function headings(ink: string, secondary: string) {
  return {
    chapter: {
      type: "body" as const,
      idx: 1,
      style: {
        position: "absolute" as const,
        left: 64,
        top: 30,
        width: 700,
        height: 26,
        fontSize: 14,
        color: secondary,
        letterSpacing: 1,
      },
    },
    title: {
      type: "title" as const,
      style: {
        position: "absolute" as const,
        left: 64,
        top: 70,
        width: 1110,
        height: 70,
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
      style: {
        position: "absolute" as const,
        left: 64,
        top: 676,
        width: 1060,
        height: 22,
        fontSize: 11,
        color: secondary,
        textWrap: "nowrap" as const,
      },
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
      style: {
        position: "absolute",
        left: 68,
        top: 154,
        width: 650,
        height: 206,
        fontFamily: "display",
        fontSize: 76,
        lineHeight: 98,
        color: colors.paperBright,
      },
    },
    subtitle: {
      type: "subTitle",
      idx: 1,
      style: {
        position: "absolute",
        left: 72,
        top: 390,
        width: 540,
        height: 48,
        fontSize: 27,
        color: colors.sun,
      },
    },
  },
});

export const pageNumber = {
  idx: 4,
  style: {
    position: "absolute" as const,
    left: 1130,
    top: 52,
    width: 66,
    height: 34,
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
    content: {
      type: "obj",
      idx: 3,
      style: {
        position: "absolute",
        left: 64,
        top: 150,
        width: 1152,
        height: 510,
      },
    },
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
