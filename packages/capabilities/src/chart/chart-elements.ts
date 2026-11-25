import { hostNodeCreator } from "@pptxdk/core/host";

import { motionElement } from "../motion/motion-elements.ts";
import type { TypedElement } from "../shared/typed-element.ts";
import { createChartNode, type ChartProps } from "./chart-declaration.ts";

export type { ChartProps } from "./chart-declaration.ts";
export type { TypedElement } from "../shared/typed-element.ts";

/** Declares an object compiled into a native chart by the official capability module. */
export function Chart(props: ChartProps): TypedElement<"chart"> {
  return motionElement("chart", {
    ...props,
    [hostNodeCreator]: createChartNode,
  }) as TypedElement<"chart">;
}
