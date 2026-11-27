import {
  BuildError,
  toBuildDiagnostic,
  type BuildDiagnostic,
} from "@pptxdk/core";
import { CommonContainer, CommonNode } from "@pptxdk/core/document";
import type { HostTextNode } from "@pptxdk/core/host";
import { createContext, type ReactNode } from "react";
import Reconciler from "react-reconciler";
import {
  ConcurrentRoot,
  DefaultEventPriority,
} from "react-reconciler/constants.js";

import {
  appendInitialHostChild,
  attachHostChild,
  commitMountHostInstance,
  createHostContainer,
  createHostInstance,
  createHostText,
  detachHostChild,
  type AdapterChild,
  type HostContainer,
  type HostInstance,
  insertRootChild,
  isMountHostRecord,
  lockHandoffTree,
  releaseContainerMounts,
  removeRootChild,
  updateHostInstance,
  updateHostText,
} from "#/document-adapter.ts";

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

const hostTransitionContext = createContext<null>(null);

const reconciler = Reconciler<
  string,
  Props,
  HostContainer,
  HostInstance,
  HostTextNode,
  never,
  never,
  never,
  never,
  AdapterChild,
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
  rendererPackageName: "@pptxdk/react",
  extraDevToolsConfig: null,
  noTimeout: -1,
  scheduleTimeout: setTimeout,
  cancelTimeout: clearTimeout,
  NotPendingTransition: null,
  HostTransitionContext: hostTransitionContext,

  createInstance(type, props, rootContainer) {
    return createHostInstance(type, props, rootContainer);
  },
  createTextInstance(text, rootContainer) {
    return createHostText(text, rootContainer);
  },
  appendInitialChild(parent, child) {
    appendInitialHostChild(parent, child);
  },
  appendChild(parent, child) {
    attachHostChild(parent, child);
  },
  appendChildToContainer(container, child) {
    insertRootChild(container, child);
  },
  insertBefore(parent, child, before) {
    attachHostChild(parent, child, before);
  },
  insertInContainerBefore(container, child, before) {
    insertRootChild(container, child, before);
  },
  removeChild(parent, child) {
    detachHostChild(parent, child);
  },
  removeChildFromContainer(container, child) {
    removeRootChild(container, child);
  },
  commitUpdate(instance, _type, _previous, next) {
    updateHostInstance(instance, next);
  },
  commitTextUpdate(textInstance, _previous, next) {
    updateHostText(textInstance, next);
  },
  hideInstance() {},
  hideTextInstance() {},
  unhideInstance() {},
  unhideTextInstance() {},
  resetTextContent() {},
  clearContainer(container) {
    container.children.length = 0;
    container.hostOrder.length = 0;
  },
  finalizeInitialChildren: (instance) => isMountHostRecord(instance),
  commitMount(instance) {
    commitMountHostInstance(instance);
  },
  shouldSetTextContent: () => false,
  getRootHostContext: () => ({}),
  getChildHostContext: (parent) => parent,
  // oxlint-disable-next-line typescript/no-unsafe-type-assertion
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
    // oxlint-disable-next-line typescript/no-unsafe-type-assertion
    Function.prototype.bind.call(
      // oxlint-disable-next-line typescript/no-unsafe-type-assertion
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

function diagnostic(
  path: string,
  code: string,
  message: string,
): BuildDiagnostic {
  return toBuildDiagnostic(path, code, message);
}

function isDeckNode(
  node: CommonNode | undefined,
): node is CommonContainer<"deck"> {
  return node instanceof CommonContainer && node.type === "deck";
}

/**
 * Selects the single completed Deck root before renderer teardown.
 *
 * Root-level checks mirror the formal declaration contract so an invalid React tree fails expansion
 * with the same diagnostics instead of silently handing off a partial document.
 */
function completedDocument(container: HostContainer): CommonContainer<"deck"> {
  const nodes = container.children.filter(
    (child): child is CommonNode => child instanceof CommonNode,
  );
  if (nodes.length !== 1 || !isDeckNode(nodes[0])) {
    throw new BuildError([
      diagnostic(
        "deck",
        "document-root",
        "Build content must have exactly one Deck root component",
      ),
    ]);
  }
  if (
    container.children.some(
      (child) => !("props" in child) && child.text.trim() !== "",
    )
  ) {
    throw new BuildError([
      diagnostic(
        "deck",
        "unexpected-text",
        "No text allowed outside the Deck root component",
      ),
    ]);
  }
  return nodes[0];
}

/**
 * Runs one dedicated React root to completion and hands off content selected from its final commit.
 *
 * React creates and updates the same `CommonNode` instances the caller receives. On completion the
 * candidate is closed to further adapter edits before renderer teardown, then the root is unmounted
 * and its listeners/adapter state are cleaned up. The candidate is returned only when both expansion
 * and cleanup succeed; unhandled errors, cancellation, and cleanup failures reject the expansion.
 */
async function expandReactRoot<Result>(
  element: ReactNode,
  signal: AbortSignal | undefined,
  select: (container: HostContainer) => Result,
): Promise<Result> {
  if (signal?.aborted) throw signal.reason;

  const container = createHostContainer();
  type Outcome = "running" | "completed" | "failed" | "cancelled";
  let outcome: Outcome = "running";
  let outcomeError: unknown;
  let cleanupError: { error: unknown } | undefined;
  let cleaning = false;
  // oxlint-disable-next-line unicorn/consistent-function-scoping
  let resolveCompletion = (): void => {};
  // oxlint-disable-next-line unicorn/consistent-function-scoping
  let rejectCompletion = (_error: unknown): void => {};
  const completion = new Promise<void>((resolve, reject) => {
    resolveCompletion = resolve;
    rejectCompletion = reject;
  });

  const finish = (next: Exclude<Outcome, "running">, error?: unknown) => {
    // The first terminal state request locks the outcome; a claimed completion is not rolled back by a later abort.
    if (outcome !== "running") return;
    outcome = next;
    outcomeError = error;
    if (next === "completed") resolveCompletion();
    else rejectCompletion(error);
  };

  // oxlint-disable-next-line unicorn/consistent-function-scoping
  let queueCompletionCheck = (): void => {};
  const root = reconciler.createContainer(
    container,
    ConcurrentRoot,
    null,
    false,
    null,
    "pptxdk",
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
  // oxlint-disable-next-line typescript/no-unsafe-type-assertion
  const fiberRoot = root as unknown as FiberRoot;
  const currentDescriptor = Object.getOwnPropertyDescriptor(root, "current");
  let current = fiberRoot.current;
  let currentAdapterInstalled = false;
  let abortListenerInstalled = false;
  let revision = 0;
  let operationError: unknown;
  let operationFailed = false;
  let candidate: Result | undefined;
  let rootError: BuildError | undefined;

  const currentOutcome = (): Outcome => outcome;
  const onAbort = (): void => finish("cancelled", signal?.reason);

  try {
    if (
      currentDescriptor?.configurable !== true ||
      currentDescriptor.writable !== true
    ) {
      throw new Error(
        "Current React root does not support verified commit observation adaptation",
      );
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
        try {
          candidate = select(container);
        } catch (error) {
          if (error instanceof BuildError) rootError = error;
          else throw error;
        }
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
      attemptCleanup(() => releaseContainerMounts(container));
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
  if (rootError !== undefined) throw rootError;
  if (finalOutcome !== "completed" || candidate === undefined) {
    throw new Error("React expansion did not produce handoff content");
  }
  return candidate;
}

/**
 * Runs a dedicated React root to completion and hands off the completed common document.
 *
 * The document is returned only when both expansion and cleanup succeed; unhandled errors,
 * cancellation, and cleanup failures reject the expansion.
 */
export async function expandReactDocument(
  element: ReactNode,
  signal?: AbortSignal,
): Promise<CommonContainer<"deck">> {
  return expandReactRoot(element, signal, (container) => {
    const candidate = completedDocument(container);
    lockHandoffTree(container);
    return candidate;
  });
}

function completedFragment(container: HostContainer): readonly CommonNode[] {
  const nodes = container.children.filter(
    (child): child is CommonNode => child instanceof CommonNode,
  );
  if (
    container.children.some(
      (child) => "text" in child && child.text.trim() !== "",
    )
  ) {
    throw new BuildError([
      diagnostic(
        "fragment",
        "unexpected-text",
        "No text allowed outside an editable node in a JSX fragment",
      ),
    ]);
  }
  const frozen = Object.freeze(nodes);
  lockHandoffTree(container);
  return frozen;
}

/**
 * Runs a dedicated React root to completion and hands off top-level editable nodes from a fragment.
 *
 * The same renderer, completion wait, cleanup, and handoff lock as `expandReactDocument` apply; the
 * fragment does not need a Deck root because it is staged ownerless until the caller appends it.
 */
export async function expandReactFragment(
  element: ReactNode,
  signal?: AbortSignal,
): Promise<readonly CommonNode[]> {
  return expandReactRoot(element, signal, completedFragment);
}
