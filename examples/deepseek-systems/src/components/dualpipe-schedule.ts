import {
  box,
  shape,
  text,
  type BoxNode,
  type CommonNode,
  type ThemeColor,
} from "@officedk/pptxdk/vanilla";

import { colors } from "../theme.ts";

type ScheduleTone = "compute" | "comm" | "copper" | "steel";

interface ScheduleBlock {
  readonly label: string;
  readonly tone: ScheduleTone;
  readonly width: number;
}

interface ScheduleLane {
  readonly title: string;
  readonly labelTop: number;
  readonly blockTop: number;
  readonly blocks: readonly ScheduleBlock[];
}

/** Props for positioning the DualPipe schedule container relative to its enclosing panel origin. */
export interface DualPipeScheduleProps {
  readonly x: number;
  readonly y: number;
  readonly id?: string;
}

const PANEL_WIDTH = 628;
const PANEL_HEIGHT = 252;
const LABEL_LEFT = 20;
const LABEL_WIDTH = 124;
const BLOCK_START_X = 148;
const BLOCK_GAP = 6;
const BLOCK_HEIGHT = 34;

const toneFill: Record<ScheduleTone, ThemeColor> = {
  compute: colors.compute,
  comm: colors.comm,
  copper: colors.copper,
  steel: colors.steel,
};

const scheduleLanes: readonly ScheduleLane[] = [
  {
    title: "正向数据流 (Rank 0→P)",
    labelTop: 72,
    blockTop: 64,
    blocks: [
      { label: "F0", tone: "compute", width: 50 },
      { label: "F1", tone: "compute", width: 50 },
      { label: "F&B (双向重叠)", tone: "copper", width: 124 },
      { label: "F&B (双向重叠)", tone: "copper", width: 124 },
      { label: "B1", tone: "comm", width: 52 },
      { label: "W1", tone: "steel", width: 48 },
    ],
  },
  {
    title: "镜像数据流 (Rank P→0)",
    labelTop: 118,
    blockTop: 110,
    blocks: [
      { label: "F'0", tone: "compute", width: 50 },
      { label: "F&B (双向重叠)", tone: "copper", width: 124 },
      { label: "F&B (双向重叠)", tone: "copper", width: 124 },
      { label: "B'0", tone: "comm", width: 52 },
      { label: "B'1", tone: "comm", width: 52 },
      { label: "W'0", tone: "steel", width: 46 },
    ],
  },
];

function buildLaneNodes(lane: ScheduleLane): CommonNode[] {
  const labelNode = text(lane.title, {
    style: {
      position: "absolute",
      left: LABEL_LEFT,
      top: lane.labelTop,
      width: LABEL_WIDTH,
      padding: 0,
      fontSize: 11,
      color: colors.steel,
      textWrap: "nowrap",
    },
  });

  let cursor = BLOCK_START_X;
  const blockNodes = lane.blocks.map((block) => {
    const left = cursor;
    cursor += block.width + BLOCK_GAP;
    return shape(block.label, {
      style: {
        position: "absolute",
        left,
        top: lane.blockTop,
        width: block.width,
        height: BLOCK_HEIGHT,
        background: toneFill[block.tone],
        padding: [0, 2],
        fontSize: 10.5,
        color: colors.surface,
        align: "center",
        verticalAlign: "middle",
      },
    });
  });

  return [labelNode, ...blockNodes];
}

/**
 * Builds the DualPipe bidirectional microbatch schedule diagram inside a transparent layout
 * box so horizontal block cursor math lives outside React rendering.
 */
export function createDualPipeSchedule({
  x,
  y,
  id = "dualpipe-schedule",
}: DualPipeScheduleProps): BoxNode {
  // A style-free box emits no native PPTX shape while offsetting child coordinates by (x, y).
  return box({
    id,
    style: {
      position: "absolute",
      left: x,
      top: y,
      width: PANEL_WIDTH,
      height: PANEL_HEIGHT,
    },
  }).append(...scheduleLanes.flatMap(buildLaneNodes));
}
