import { hostNodeCreator } from "@pptxdk/core/host";
import type { ReactNode } from "react";

import { motionElement } from "../motion/motion-elements.ts";
import { createChartNode, type ChartProps } from "./chart.ts";

/** Declares an object compiled into a native chart by the official capability module. */
export function Chart(props: ChartProps): ReactNode {
  return motionElement("chart", {
    ...props,
    [hostNodeCreator]: createChartNode,
  });
}
