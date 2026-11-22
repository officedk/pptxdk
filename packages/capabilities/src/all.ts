import type { RegisteredCapabilityModule } from "@pptxdk/core/capability";

import { boxCapability } from "./box/box.ts";
import { chartCapability } from "./chart/chart.ts";
import { groupCapability } from "./group/group.ts";
import { imageCapability } from "./image/image.ts";
import { motionCapability } from "./motion/motion.ts";
import { notesCapability } from "./notes/notes.ts";
import { shapeCapability } from "./shape/shape.ts";
import { slideBackgroundCapability } from "./slide-background/slide-background.ts";
import { tableCapability } from "./table/table.ts";
import { textCapability } from "./text/text.ts";
import { transitionCapability } from "./transition/transition.ts";

const baselineCapabilities = [
  boxCapability,
  chartCapability,
  groupCapability,
  imageCapability,
  motionCapability,
  notesCapability,
  shapeCapability,
  slideBackgroundCapability,
  tableCapability,
  textCapability,
  transitionCapability,
] as const satisfies readonly RegisteredCapabilityModule[];

/** Baseline capabilities enabled by the public document build; factories never register capabilities. */
export const documentCapabilities: typeof baselineCapabilities =
  Object.freeze(baselineCapabilities);
