import type { CapabilityModule } from "@pptxdk/core/capability";

import { boxCapability } from "./box/box.ts";
import { chartCapability } from "./chart/chart.ts";
import { groupCapability } from "./group/group.ts";
import { imageCapability } from "./image/image.ts";
import {
  footerPlaceholderCapability,
  mastersCapability,
} from "./masters/masters.ts";
import { motionCapability } from "./motion/motion.ts";
import { notesCapability } from "./notes/notes.ts";
import { placeholderCapability } from "./placeholder/placeholder.ts";
import { shapeCapability } from "./shape/shape.ts";
import { slideBackgroundCapability } from "./slide-background/slide-background.ts";
import { tableCapability } from "./table/table.ts";
import { textCapability } from "./text/text.ts";
import { transitionCapability } from "./transition/transition.ts";

const baselineCapabilities = [
  boxCapability,
  chartCapability,
  footerPlaceholderCapability,
  groupCapability,
  imageCapability,
  mastersCapability,
  motionCapability,
  notesCapability,
  placeholderCapability,
  shapeCapability,
  slideBackgroundCapability,
  tableCapability,
  textCapability,
  transitionCapability,
] as const satisfies readonly CapabilityModule[];

/** Baseline capabilities enabled by the public document build; factories never register capabilities. */
export const documentCapabilities: readonly [
  typeof boxCapability,
  typeof chartCapability,
  typeof footerPlaceholderCapability,
  typeof groupCapability,
  typeof imageCapability,
  typeof mastersCapability,
  typeof motionCapability,
  typeof notesCapability,
  typeof placeholderCapability,
  typeof shapeCapability,
  typeof slideBackgroundCapability,
  typeof tableCapability,
  typeof textCapability,
  typeof transitionCapability,
] = Object.freeze(baselineCapabilities);
