import { createContext, type ReactNode } from "react";
import Reconciler from "react-reconciler";
import {
  DefaultEventPriority,
  ConcurrentRoot,
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

type FiberNode = {
  readonly tag: number;
  readonly memoizedState: unknown;
  readonly child: FiberNode | null;
  readonly sibling: FiberNode | null;
};

type FiberRoot = {
  current: FiberNode;
  readonly pendingLanes: number;
};

const suspenseComponentTag = 13;

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
    return { type: "#text", text, hidden: false };
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
  hideInstance(instance) {
    instance.hidden = true;
  },
  hideTextInstance(textInstance) {
    textInstance.hidden = true;
  },
  unhideInstance(instance) {
    instance.hidden = false;
  },
  unhideTextInstance(textInstance) {
    textInstance.hidden = false;
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

function containsSuspendedBoundary(fiber: FiberNode | null): boolean {
  for (let current = fiber; current !== null; current = current.sibling) {
    if (
      current.tag === suspenseComponentTag &&
      current.memoizedState !== null
    ) {
      return true;
    }
    if (containsSuspendedBoundary(current.child)) return true;
  }
  return false;
}

/** 为一次构建创建专属 React root，等待已声明内容并在序列化后卸载。 */
export async function withHostTree<T>(
  element: ReactNode,
  consume: (container: HostContainer) => T,
): Promise<T> {
  const container: HostContainer = { children: [] };
  let failure: { error: unknown } | undefined;
  let queueCompletionCheck = (): void => {};
  const root = reconciler.createContainer(
    container,
    ConcurrentRoot,
    null,
    false,
    null,
    "react-pptx",
    (error) => {
      failure ??= { error };
      queueCompletionCheck();
    },
    () => {},
    () => {},
    () => {},
    null,
  );
  const fiberRoot = root as unknown as FiberRoot;
  const currentDescriptor = Object.getOwnPropertyDescriptor(root, "current");
  let current = fiberRoot.current;
  let currentAdapterInstalled = false;

  try {
    if (
      currentDescriptor?.configurable !== true ||
      currentDescriptor.writable !== true
    ) {
      throw new Error("当前 React root 不支持已验证的提交观察适配");
    }

    await new Promise<void>((resolve, reject) => {
      let settled = false;
      let revision = 0;

      const finish = (failed: boolean, error?: unknown): void => {
        if (settled) return;
        settled = true;
        if (failed) reject(error);
        else if (failure !== undefined) reject(failure.error);
        else resolve();
      };

      queueCompletionCheck = (): void => {
        if (settled) return;
        const observedRevision = ++revision;
        queueMicrotask(() => {
          if (settled || observedRevision !== revision) return;
          try {
            reconciler.flushPassiveEffects();
            if (settled || observedRevision !== revision) return;
            if (failure !== undefined) {
              finish(true, failure.error);
              return;
            }
            if (
              !containsSuspendedBoundary(fiberRoot.current) &&
              fiberRoot.pendingLanes === 0
            ) {
              finish(false);
            }
          } catch (error) {
            finish(true, error);
          }
        });
      };

      Object.defineProperty(root, "current", {
        enumerable: currentDescriptor.enumerable === true,
        configurable: true,
        get: () => current,
        set(next: FiberNode) {
          current = next;
          try {
            queueCompletionCheck();
          } catch (error) {
            failure ??= { error };
          }
        },
      });
      currentAdapterInstalled = true;

      reconciler.updateContainer(element, root, null, null);
    });

    return consume(container);
  } finally {
    try {
      reconciler.updateContainerSync(null, root, null, null);
      reconciler.flushSyncWork();
      reconciler.flushPassiveEffects();
    } finally {
      if (currentAdapterInstalled && currentDescriptor !== undefined) {
        Object.defineProperty(root, "current", {
          ...currentDescriptor,
          value: current,
        });
      }
    }
  }
}
