import { hostNodeCreator } from "@pptxdk/core/host";
import { createElement, useState, type ReactNode } from "react";

import {
  createMotionScopeHandle,
  createMotionScopeNode,
  createMotionStepIdentity,
  createMotionStepNode,
  type MotionProps,
  type MotionScopeHandle,
  type MotionScopeProps,
  type StepProps,
} from "./motion-api.ts";

export { Box, type BoxProps } from "../box/box-elements.ts";

/**
 * Keeps one pure scope identity for the lifetime of this component instance.
 *
 * The React hook only stores the plain handle; scope parent chains, click steps, and inheritance are
 * derived by the build from the document structure, so moving content does not carry React state.
 */
export function useMotionScope(): MotionScopeHandle {
  const [handle] = useState(createMotionScopeHandle);
  return handle;
}

/** Explicit orchestration scope only declares timing relationships without creating layout or native objects. */
export function MotionScope(props: MotionScopeProps): ReactNode {
  const { children, ...declaration } = props;
  return createElement(
    "motion-scope",
    { ...declaration, [hostNodeCreator]: createMotionScopeNode },
    children,
  );
}

/** A single click triggers a group of objects; declarations with the same number share a click step. */
export function Step(props: StepProps): ReactNode {
  const { children, ...declaration } = props;
  const [stepIdentity] = useState(createMotionStepIdentity);
  return createElement(
    "motion-step",
    { ...declaration, stepIdentity, [hostNodeCreator]: createMotionStepNode },
    children,
  );
}

/**
 * Passes local motion props to a native object host.
 *
 * Object capabilities emit only local declarations; inherited scope, step, default entrance, and
 * disabled ancestors are resolved by the formal build from the current document.
 */
export function motionElement(
  type: string,
  props: MotionProps & {
    readonly children?: ReactNode;
    readonly [hostNodeCreator]?: unknown;
  },
): ReactNode {
  return createElement(type, props, props.children);
}
