import { createContext, type ReactNode } from "react";
import Reconciler from "react-reconciler";
import {
  DefaultEventPriority,
  LegacyRoot,
} from "react-reconciler/constants.js";

import {
  HostNode,
  type HostChild,
  type HostContainer,
  type HostText,
  withoutChildren,
} from "#/host/nodes.ts";

type Props = Record<string, unknown>;

let currentUpdatePriority = 0;

function insertChild(
  children: HostChild[],
  child: HostChild,
  before?: HostChild,
): void {
  const existing = children.indexOf(child);
  if (existing !== -1) children.splice(existing, 1);
  const index = before === undefined ? -1 : children.indexOf(before);
  if (index === -1) children.push(child);
  else children.splice(index, 0, child);
}

function removeChild(children: HostChild[], child: HostChild): void {
  const index = children.indexOf(child);
  if (index !== -1) children.splice(index, 1);
}

const hostTransitionContext = createContext<null>(null);

const reconciler = Reconciler<
  string,
  Props,
  HostContainer,
  HostNode,
  HostText,
  never,
  never,
  never,
  never,
  HostNode | HostText,
  Record<string, never>,
  never,
  ReturnType<typeof setTimeout>,
  -1,
  null,
  null,
  null,
  never,
  never,
  never
>({
  supportsMutation: true,
  supportsPersistence: false,
  supportsHydration: false,
  isPrimaryRenderer: false,
  rendererVersion: "0.0.0",
  rendererPackageName: "@react-pptx/core",
  extraDevToolsConfig: null,
  noTimeout: -1,
  scheduleTimeout: setTimeout,
  cancelTimeout: clearTimeout,
  NotPendingTransition: null,
  HostTransitionContext: hostTransitionContext,

  createInstance(type, props, _root, _context, internalHandle) {
    return new HostNode(type, withoutChildren(props), internalHandle);
  },
  createTextInstance(text) {
    return { type: "#text", text };
  },
  appendInitialChild(parent, child) {
    insertChild(parent.children, child);
  },
  appendChild(parent, child) {
    insertChild(parent.children, child);
  },
  appendChildToContainer(container, child) {
    insertChild(container.children, child);
  },
  insertBefore(parent, child, before) {
    insertChild(parent.children, child, before);
  },
  insertInContainerBefore(container, child, before) {
    insertChild(container.children, child, before);
  },
  removeChild(parent, child) {
    removeChild(parent.children, child);
  },
  removeChildFromContainer(container, child) {
    removeChild(container.children, child);
  },
  commitUpdate(instance, _type, _previous, next, internalHandle) {
    instance.props = withoutChildren(next);
    instance.fiber = internalHandle;
  },
  commitTextUpdate(textInstance, _previous, next) {
    textInstance.text = next;
  },
  resetTextContent() {},
  clearContainer(container) {
    container.children.length = 0;
  },
  finalizeInitialChildren: () => false,
  shouldSetTextContent: () => false,
  getRootHostContext: () => ({}),
  getChildHostContext: (parent) => parent,
  getPublicInstance: () => null as never,
  prepareForCommit: () => null,
  resetAfterCommit() {},
  preparePortalMount() {},
  getInstanceFromNode: () => null,
  beforeActiveInstanceBlur() {},
  afterActiveInstanceBlur() {},
  prepareScopeUpdate() {},
  getInstanceFromScope: () => null,
  detachDeletedInstance() {},
  bindToConsole: (methodName, args) =>
    Function.prototype.bind.call(
      console[methodName as "log"],
      console,
      ...(args as unknown[]),
    ) as () => void,
  setCurrentUpdatePriority(priority) {
    currentUpdatePriority = priority;
  },
  getCurrentUpdatePriority: () => currentUpdatePriority,
  resolveUpdatePriority: () =>
    currentUpdatePriority === 0 ? DefaultEventPriority : currentUpdatePriority,
  resetFormInstance() {},
  requestPostPaintCallback() {},
  shouldAttemptEagerTransition: () => false,
  trackSchedulerEvent() {},
  resolveEventType: () => null,
  resolveEventTimeStamp: () => -1.1,
  maySuspendCommit: () => false,
  maySuspendCommitOnUpdate: () => false,
  maySuspendCommitInSyncRender: () => false,
  preloadInstance: () => true,
  startSuspendingCommit: () => null,
  suspendInstance() {},
  suspendOnActiveViewTransition() {},
  waitForCommitToBeReady: () => null,
  getSuspendedCommitReason: () => null,
});

/** 为一次构建创建专属 React root，并在序列化完成后卸载。 */
export function withHostTree<T>(
  element: ReactNode,
  consume: (container: HostContainer) => T,
): T {
  const container: HostContainer = { children: [] };
  let failure: { error: unknown } | undefined;
  const record = (error: unknown): void => {
    failure ??= { error };
  };
  const root = reconciler.createContainer(
    container,
    LegacyRoot,
    null,
    false,
    null,
    "react-pptx",
    record,
    record,
    () => {},
    () => {},
    null,
  );

  try {
    reconciler.updateContainerSync(element, root, null, null);
    reconciler.flushSyncWork();
    reconciler.flushPassiveEffects();
    if (failure !== undefined) throw failure.error;
    return consume(container);
  } finally {
    reconciler.updateContainerSync(null, root, null, null);
    reconciler.flushSyncWork();
    reconciler.flushPassiveEffects();
  }
}
