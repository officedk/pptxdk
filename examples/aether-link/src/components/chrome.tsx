import {
  Notes,
  P,
  Slide,
  fromVanilla,
  type JsxPlaceholders,
  type PlaceholderMap,
  type SlideLayoutSource,
  type ThemeOverride,
  type Transition,
} from "@officedk/pptxdk/react";
import type { ReactNode } from "react";

import {
  EDITORIAL_BODY_STYLE,
  PALETTE,
  editorialLightLayout,
  jostFont,
} from "../theme.ts";
import {
  createBeamDivergenceDiagram,
  createMetricTile,
  createTakeawayStrip,
  createTelemetryPanel,
} from "./vanilla.ts";

// React component bridges generated from the zero-React Vanilla library.
export const MetricTile = fromVanilla(createMetricTile);
export const TakeawayStrip = fromVanilla(createTakeawayStrip);
export const TelemetryPanel = fromVanilla(createTelemetryPanel);
export const BeamDivergenceDiagram = fromVanilla(createBeamDivergenceDiagram);

/** Shared light editorial slide shell with layout placeholders and optional speaker notes. */
export function EditorialSlide(props: {
  readonly id: string;
  readonly layout?: SlideLayoutSource | undefined;
  readonly kicker: string;
  readonly title: string;
  readonly notes?: string | undefined;
  readonly transition?: Transition | undefined;
  readonly themeOverride?: ThemeOverride | undefined;
  readonly extraPlaceholders?: JsxPlaceholders<PlaceholderMap> | undefined;
  readonly children?: ReactNode;
}) {
  const layout = props.layout ?? editorialLightLayout;
  const placeholders: JsxPlaceholders<PlaceholderMap> = {
    kicker: props.kicker,
    title: props.title,
    ...props.extraPlaceholders,
  };

  return (
    <Slide
      id={props.id}
      layout={layout}
      placeholders={placeholders}
      dateTime={true}
      footer={true}
      slideNumber={true}
      {...(props.transition === undefined
        ? {}
        : { transition: props.transition })}
      {...(props.themeOverride === undefined
        ? {}
        : { themeOverride: props.themeOverride })}
      style={EDITORIAL_BODY_STYLE}
    >
      {props.children}
      {props.notes ? (
        <Notes
          style={{ fontFamily: jostFont, fontSize: 12, color: PALETTE.ink }}
        >
          <P>{props.notes}</P>
        </Notes>
      ) : null}
    </Slide>
  );
}
