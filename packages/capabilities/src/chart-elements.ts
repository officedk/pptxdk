import { hostNodeCreator } from "@react-pptx/core/host";
import type { ReactNode } from "react";

import { createChartNode, type ChartProps } from "./chart.ts";
import { motionElement } from "./motion-elements.ts";

/** Declares an object compiled into a native chart by the official capability module. */
export function Chart(props: ChartProps): ReactNode {
  return motionElement("rp:chart", {
    ...props,
    [hostNodeCreator]: createChartNode,
  });
}
