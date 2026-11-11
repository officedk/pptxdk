import {
  createContext,
  createElement,
  useContext,
  useId,
  type ReactNode,
} from "react";

import type {
  MotionContextValue,
  MotionMoment,
  MotionMomentOffset,
  MotionProps,
  MotionScopeDeclaration,
  MotionScopeHandle,
  MotionScopeProps,
  MotionTiming,
  StepDeclaration,
  StepProps,
} from "./motion-api.ts";

export { Box, type BoxProps } from "./box-elements.ts";

export const motionContext = createContext<MotionContextValue>({
  enabled: true,
  scopes: [],
});

function moment(
  scope: string,
  point: MotionMoment["point"],
  name?: string,
): MotionMoment {
  return Object.freeze({
    type: "motion-moment",
    scope,
    point,
    ...(name === undefined ? {} : { name }),
  });
}

/** Creates explicit orchestration scope and publishable named moments for current React component instance. */
export function useMotionScope(): MotionScopeHandle {
  const identity = useId();
  const cue = (name: string): MotionMoment => moment(identity, "named", name);
  return Object.freeze({
    type: "motion-scope-handle",
    identity,
    start: moment(identity, "start"),
    end: moment(identity, "end"),
    cue,
  });
}

/** Explicit orchestration scope only declares timing relationships without creating layout or native objects. */
export function MotionScope(props: MotionScopeProps): ReactNode {
  const { children, ...publicProps } = props;
  const {
    scope,
    mode = "parallel",
    each,
    delay = 0,
    startAt,
    publishStart,
    publishEnd,
    motion,
  } = publicProps;
  const parent = useContext(motionContext);
  const parentScope = parent.scopes.at(-1);
  const declaration: MotionScopeDeclaration = Object.freeze({
    identity: scope.identity,
    ...(parentScope === undefined ? {} : { parent: parentScope.identity }),
    ...(parent.step === undefined ? {} : { step: parent.step }),
    mode,
    delay,
    ...(each === undefined ? {} : { each }),
    ...(startAt === undefined ? {} : { startAt }),
    ...(publishStart === undefined ? {} : { publishStart }),
    ...(publishEnd === undefined ? {} : { publishEnd }),
  });
  const context = Object.freeze({
    ...parent,
    enabled: parent.enabled && motion !== false,
    scopes: Object.freeze([...parent.scopes, declaration]),
  });
  return createElement(
    motionContext.Provider,
    { value: context },
    createElement(
      "rp:motion-scope",
      { ...publicProps, __motionScope: declaration },
      children,
    ),
  );
}

/** A single click triggers a group of objects; declarations with the same number share a click step. */
export function Step(props: StepProps): ReactNode {
  const { children, ...publicProps } = props;
  const { n } = publicProps;
  const parent = useContext(motionContext);
  const identity = useId();
  const step: StepDeclaration =
    n === undefined
      ? { mode: "automatic", identity }
      : { mode: "explicit", number: n };
  return createElement(
    motionContext.Provider,
    { value: { ...parent, step } },
    createElement(
      "rp:motion-step",
      { ...publicProps, __motionStep: step },
      children,
    ),
  );
}

/** Passes current Step and disabled status to the formal native object host node. */
export function motionElement(
  type: string,
  props: MotionProps & { readonly children?: ReactNode },
): ReactNode {
  const parent = useContext(motionContext);
  const enabled = parent.enabled && props.motion !== false;
  const context = Object.freeze({ ...parent, enabled });
  const host = createElement(
    type,
    {
      ...props,
      __motionContext: context,
    },
    props.children,
  );
  if (enabled === parent.enabled) return host;
  return createElement(motionContext.Provider, { value: context }, host);
}

export type { MotionMoment, MotionMomentOffset, MotionTiming };
