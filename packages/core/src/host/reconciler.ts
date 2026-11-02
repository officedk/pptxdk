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
  signal?: AbortSignal,
): Promise<T> {
  if (signal?.aborted) throw signal.reason;

  const container: HostContainer = { children: [] };
  type Outcome = "running" | "completed" | "failed" | "cancelled";
  let outcome: Outcome = "running";
  let outcomeError: unknown;
  let cleanupError: { error: unknown } | undefined;
  let cleaning = false;
  let resolveCompletion = (): void => {};
  let rejectCompletion = (_error: unknown): void => {};
  const completion = new Promise<void>((resolve, reject) => {
    resolveCompletion = resolve;
    rejectCompletion = reject;
  });

  const finish = (next: Exclude<Outcome, "running">, error?: unknown) => {
    // 首个终态请求锁定结果；已认领的完成不会被后来到达的 abort 回滚。
    if (outcome !== "running") return;
    outcome = next;
    outcomeError = error;
    if (next === "completed") resolveCompletion();
    else rejectCompletion(error);
  };

  let queueCompletionCheck = (): void => {};
  const root = reconciler.createContainer(
    container,
    ConcurrentRoot,
    null,
    false,
    null,
    "react-pptx",
    (error) => {
      if (cleaning) {
        cleanupError ??= { error };
      } else if (outcome === "running") {
        finish("failed", error);
      }
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
  let abortListenerInstalled = false;
  let revision = 0;
  let operationError: unknown;
  let operationFailed = false;
  let result: T | undefined;
  let hasResult = false;

  const currentOutcome = (): Outcome => outcome;
  const onAbort = (): void => finish("cancelled", signal?.reason);

  try {
    if (
      currentDescriptor?.configurable !== true ||
      currentDescriptor.writable !== true
    ) {
      throw new Error("当前 React root 不支持已验证的提交观察适配");
    }

    queueCompletionCheck = (): void => {
      if (outcome !== "running") return;
      try {
        const observedRevision = ++revision;
        queueMicrotask(() => {
          if (outcome !== "running" || observedRevision !== revision) return;
          try {
            reconciler.flushPassiveEffects();
            if (outcome !== "running" || observedRevision !== revision) return;
            if (
              !containsSuspendedBoundary(fiberRoot.current) &&
              fiberRoot.pendingLanes === 0
            ) {
              finish("completed");
            }
          } catch (error) {
            finish("failed", error);
          }
        });
      } catch (error) {
        finish("failed", error);
      }
    };

    Object.defineProperty(root, "current", {
      enumerable: currentDescriptor.enumerable === true,
      configurable: true,
      get: () => current,
      set(next: FiberNode) {
        current = next;
        queueCompletionCheck();
      },
    });
    currentAdapterInstalled = true;

    if (signal !== undefined) {
      signal.addEventListener("abort", onAbort, { once: true });
      abortListenerInstalled = true;
      if (signal.aborted) onAbort();
    }
    if (outcome === "running") {
      reconciler.updateContainer(element, root, null, null);
    }

    try {
      await completion;
      if (currentOutcome() === "completed") {
        result = consume(container);
        hasResult = true;
      }
    } catch (error) {
      operationFailed = true;
      operationError = error;
    }
  } catch (error) {
    operationFailed = true;
    operationError = error;
  } finally {
    cleaning = true;
    const attemptCleanup = (cleanup: () => void): void => {
      try {
        cleanup();
      } catch (error) {
        cleanupError ??= { error };
      }
    };

    try {
      attemptCleanup(() => {
        reconciler.updateContainerSync(null, root, null, null);
      });
      attemptCleanup(() => reconciler.flushSyncWork());
      attemptCleanup(() => reconciler.flushPassiveEffects());
    } finally {
      if (currentAdapterInstalled && currentDescriptor !== undefined) {
        attemptCleanup(() => {
          Object.defineProperty(root, "current", {
            ...currentDescriptor,
            value: current,
          });
        });
      }
      if (abortListenerInstalled && signal !== undefined) {
        attemptCleanup(() => signal.removeEventListener("abort", onAbort));
      }
    }
  }

  const finalOutcome = currentOutcome();
  if (operationFailed) throw operationError;
  if (finalOutcome === "failed" || finalOutcome === "cancelled") {
    throw outcomeError;
  }
  if (cleanupError !== undefined) throw cleanupError.error;
  if (finalOutcome !== "completed" || !hasResult) {
    throw new Error("React 构建未产生完成结果");
  }
  return result as T;
}
