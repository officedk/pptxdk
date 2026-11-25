import assert from "node:assert/strict";
import { getEventListeners } from "node:events";
import { createServer, type ServerResponse } from "node:http";

import { BuildError, build, emit, resolve } from "@officedk/pptxdk";
import { Deck, expandDocument, Slide, Text } from "@officedk/pptxdk/react";
import { validatePptx } from "@pptxdk/validation";
import {
  Component,
  createElement,
  Fragment,
  Suspense,
  use,
  useEffect,
  useLayoutEffect,
  useState,
  type ReactNode,
} from "react";
import { describe, it } from "vitest";

function deck(children: ReactNode) {
  return createElement(
    Deck,
    { width: 960, height: 540 },
    createElement(Slide, { id: "async" }, children),
  );
}

function Read({
  promise,
  onRead,
  id,
}: {
  promise: Promise<string>;
  onRead?: () => void;
  id?: string;
}) {
  const text = use(promise);
  useEffect(() => {
    onRead?.();
  }, [onRead, text]);
  return createElement(
    Text,
    {
      ...(id === undefined ? {} : { id }),
      x: 0,
      y: 0,
      width: 240,
      height: 48,
    },
    text,
  );
}

class ErrorBoundary extends Component<
  { children: ReactNode },
  { failed: boolean }
> {
  override state = { failed: false };

  static getDerivedStateFromError() {
    return { failed: true };
  }

  override render() {
    return this.state.failed
      ? createElement(
          Text,
          { x: 0, y: 0, width: 240, height: 48 },
          "边界业务回退",
        )
      : this.props.children;
  }
}

describe("React asynchronous content in public build", () => {
  it("loads React element implementation corresponding to NODE_ENV", () => {
    const element = createElement("react-mode-probe");

    // React 19 development elements carry _store; production elements do not. Check the actual module branch rather than the environment variable itself.
    assert.equal(
      Object.hasOwn(element, "_store"),
      process.env.NODE_ENV !== "production",
    );
  });

  it("does not execute author components or leave AbortSignal listeners on pre-cancellation", async () => {
    const controller = new AbortController();
    const reason = new Error("caller cancelled");
    controller.abort(reason);
    let rendered = 0;

    function AuthorContent() {
      rendered += 1;
      return createElement(
        Text,
        { x: 0, y: 0, width: 240, height: 48 },
        "不应执行",
      );
    }

    await assert.rejects(
      build(deck(createElement(AuthorContent)), { signal: controller.signal }),
      (error: unknown) => error === reason,
    );

    assert.equal(rendered, 0);
    assert.equal(getEventListeners(controller.signal, "abort").length, 0);
  });

  it("rejects build, unmounts author components, and removes listeners when cancelled while suspended", async () => {
    const controller = new AbortController();
    const reason = new Error("aborted waiting for async content");
    const never = new Promise<string>(() => {});
    const started = Promise.withResolvers<void>();
    let cleaned = 0;

    function PendingContent() {
      useEffect(() => {
        started.resolve();
        return () => {
          cleaned += 1;
        };
      }, []);
      return createElement(
        ErrorBoundary,
        null,
        createElement(
          Suspense,
          {
            fallback: createElement(
              Text,
              { x: 0, y: 0, width: 240, height: 48 },
              "加载占位",
            ),
          },
          createElement(Read, { promise: never }),
        ),
      );
    }

    const pendingBuild = build(deck(createElement(PendingContent)), {
      signal: controller.signal,
    });
    const rejected = assert.rejects(
      pendingBuild,
      (error: unknown) => error === reason,
    );
    await started.promise;
    assert.equal(getEventListeners(controller.signal, "abort").length, 1);

    controller.abort(reason);
    await rejected;

    assert.equal(cleaned, 1);
    assert.equal(getEventListeners(controller.signal, "abort").length, 0);
  });

  it("does not re-enter commit phase when cancelled inside LayoutEffect", async () => {
    const controller = new AbortController();
    const reason = new Error("cancelled inside LayoutEffect");
    let insideLayoutEffect = false;
    let cleanupRanInsideLayoutEffect = false;

    function CancelDuringLayoutEffect() {
      useLayoutEffect(() => {
        insideLayoutEffect = true;
        controller.abort(reason);
        insideLayoutEffect = false;
        return () => {
          cleanupRanInsideLayoutEffect = insideLayoutEffect;
        };
      }, []);
      return createElement(
        Text,
        { x: 0, y: 0, width: 240, height: 48 },
        "取消前的提交内容",
      );
    }

    await assert.rejects(
      build(deck(createElement(CancelDuringLayoutEffect)), {
        signal: controller.signal,
      }),
      (error: unknown) => error === reason,
    );

    assert.equal(cleanupRanInsideLayoutEffect, false);
    assert.equal(getEventListeners(controller.signal, "abort").length, 0);
  });

  it("prioritizes cancellation over unmount cleanup exceptions, still completing cleanup and removing listeners", async () => {
    const controller = new AbortController();
    const reason = new Error("caller cancellation ends build with priority");
    const cleanupFailure = new Error("cleanup failure during cancellation");
    const never = new Promise<string>(() => {});
    const started = Promise.withResolvers<void>();
    let cleaned = 0;

    function PendingCleanupFailure() {
      useEffect(() => {
        started.resolve();
        return () => {
          cleaned += 1;
          throw cleanupFailure;
        };
      }, []);
      return createElement(
        Suspense,
        { fallback: null },
        createElement(Read, { promise: never }),
      );
    }

    const pendingBuild = build(deck(createElement(PendingCleanupFailure)), {
      signal: controller.signal,
    });
    const rejected = assert.rejects(
      pendingBuild,
      (error: unknown) => error === reason,
    );
    await started.promise;
    controller.abort(reason);
    await rejected;

    assert.equal(cleaned, 1);
    assert.equal(getEventListeners(controller.signal, "abort").length, 0);
    const recovered = await build(
      deck(
        createElement(
          Text,
          { x: 0, y: 0, width: 240, height: 48 },
          "取消后恢复",
        ),
      ),
    );
    const file = await validatePptx(recovered.pptx);
    assert.deepEqual(file.errors, []);
    assert.equal(file.slides[0]?.objects[0]?.text?.text, "取消后恢复");
  });

  it("isolates cancellation across concurrent builds without late Promises affecting another real artifact", async () => {
    const cancelled = new AbortController();
    const continuing = new AbortController();
    const cancelReason = new Error("cancel only the first build");
    const late = Promise.withResolvers<string>();
    const otherContent = Promise.withResolvers<string>();
    const cancelledStarted = Promise.withResolvers<void>();
    const continuingStarted = Promise.withResolvers<void>();
    let cleaned = 0;
    let lateCallbacks = 0;

    function CancelledContent() {
      const [text, setText] = useState("原始内容");
      useEffect(() => {
        cancelledStarted.resolve();
        void late.promise.then((value) => {
          lateCallbacks += 1;
          setText(value);
        });
        return () => {
          cleaned += 1;
        };
      }, []);
      return createElement(
        Fragment,
        null,
        createElement(Text, { x: 0, y: 0, width: 240, height: 48 }, text),
        createElement(
          Suspense,
          {
            fallback: createElement(
              Text,
              { x: 0, y: 52, width: 240, height: 48 },
              "等待",
            ),
          },
          createElement(Read, { promise: late.promise }),
        ),
      );
    }

    function ContinuingContent() {
      useEffect(() => {
        continuingStarted.resolve();
      }, []);
      return createElement(
        Suspense,
        { fallback: null },
        createElement(Read, { promise: otherContent.promise }),
      );
    }

    const cancelledBuild = build(deck(createElement(CancelledContent)), {
      signal: cancelled.signal,
    });
    const continuingBuild = build(deck(createElement(ContinuingContent)), {
      signal: continuing.signal,
    });
    const cancelledRejection = assert.rejects(
      cancelledBuild,
      (error: unknown) => error === cancelReason,
    );
    await Promise.all([cancelledStarted.promise, continuingStarted.promise]);

    cancelled.abort(cancelReason);
    await cancelledRejection;
    late.resolve("取消后到达的更新");
    await late.promise;
    otherContent.resolve("另一构建的最终内容");

    const result = await continuingBuild;
    const file = await validatePptx(result.pptx);

    assert.equal(cleaned, 1);
    assert.equal(lateCallbacks, 1);
    assert.equal(getEventListeners(cancelled.signal, "abort").length, 0);
    assert.equal(getEventListeners(continuing.signal, "abort").length, 0);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(
      result.report.slides[0]?.objects.map(({ text }) => text),
      ["另一构建的最终内容"],
    );
    assert.equal(file.slides[0]?.objects[0]?.text?.text, "另一构建的最终内容");
  });

  it("preserves same build result when completion precedes cancellation", async () => {
    const controller = new AbortController();
    const result = await build(
      deck(
        createElement(
          Text,
          {
            id: "completed-before-abort",
            x: 0,
            y: 0,
            width: 240,
            height: 48,
          },
          "完成内容",
        ),
      ),
      { signal: controller.signal },
    );
    const report = result.report;
    const pptx = result.pptx.slice();

    controller.abort(new Error("late cancellation"));
    const file = await validatePptx(result.pptx);

    assert.equal(getEventListeners(controller.signal, "abort").length, 0);
    assert.deepEqual(result.report, report);
    assert.deepEqual(result.pptx, pptx);
    assert.deepEqual(file.errors, []);
    assert.equal(file.slides[0]?.objects[0]?.text?.text, "完成内容");
  });

  it("keeps the expanded document when layout cleanup aborts the expansion signal, while resolve/build honor their own signals", async () => {
    const expandController = new AbortController();
    const expandReason = new Error("cleanup abort after expansion completion");
    let cleanupRan = false;

    function AbortsDuringLayoutCleanup() {
      useLayoutEffect(
        () => () => {
          cleanupRan = true;
          expandController.abort(expandReason);
        },
        [],
      );
      return createElement(
        Text,
        {
          id: "cleanup-abort",
          x: 0,
          y: 0,
          width: 240,
          height: 48,
        },
        "完成内容",
      );
    }

    const expanded = await expandDocument(
      deck(createElement(AbortsDuringLayoutCleanup)),
      { signal: expandController.signal },
    );
    assert.equal(cleanupRan, true);
    assert.equal(getEventListeners(expandController.signal, "abort").length, 0);

    // Sharing the already-aborted expansion signal with resolve() or build() cancels the resolve stage.
    assert.throws(
      () => resolve(expanded, { signal: expandController.signal }),
      (error: unknown) => error === expandReason,
    );
    await assert.rejects(
      () =>
        build(deck(createElement(AbortsDuringLayoutCleanup)), {
          signal: expandController.signal,
        }),
      (error: unknown) => error === expandReason,
    );

    // A separate signal for resolve/emit succeeds when expansion itself succeeded, and cancels independently when aborted.
    const resolveController = new AbortController();
    const resolution = resolve(expanded, { signal: resolveController.signal });
    const emitted = emit(resolution, { signal: resolveController.signal });
    const file = await validatePptx(emitted.pptx);
    assert.deepEqual(file.errors, []);
    assert.equal(file.slides[0]?.objects[0]?.text?.text, "完成内容");

    const abortedResolveController = new AbortController();
    const resolveReason = new Error("resolve cancelled independently");
    abortedResolveController.abort(resolveReason);
    assert.throws(
      () => resolve(expanded, { signal: abortedResolveController.signal }),
      (error: unknown) => error === resolveReason,
    );
  });

  it("awaits stable Promises inside Suspense and generates final content", async () => {
    const started = Promise.withResolvers<void>();
    let resolveContent: (value: string) => void = () => {};
    const content = new Promise<string>((resolve) => {
      resolveContent = resolve;
    });

    function SuspendedContent() {
      useEffect(() => {
        started.resolve();
      }, []);
      return createElement(
        Suspense,
        {
          fallback: createElement(
            Text,
            { x: 0, y: 0, width: 240, height: 48 },
            "加载中",
          ),
        },
        createElement(Read, { promise: content }),
      );
    }

    const pendingBuild = build(deck(createElement(SuspendedContent)));
    let settled = false;
    void pendingBuild.then(
      () => {
        settled = true;
      },
      () => {
        settled = true;
      },
    );

    await started.promise;
    assert.equal(settled, false);

    resolveContent("最终内容");
    const result = await pendingBuild;
    const file = await validatePptx(result.pptx);

    assert.deepEqual(file.errors, []);
    assert.deepEqual(
      result.report.slides[0]?.objects.map(({ text }) => text),
      ["最终内容"],
    );
    assert.equal(file.slides[0]?.objects[0]?.text?.text, "最终内容");
  });

  it("mixes synchronous and entry-prefetched content, awaiting final commits of nested and empty fallbacks", async () => {
    const outer = Promise.withResolvers<string>();
    const inner = Promise.withResolvers<string>();
    const started = Promise.withResolvers<void>();
    const outerRead = Promise.withResolvers<void>();
    const innerRead = Promise.withResolvers<void>();
    const prefetched = await Promise.resolve("入口预取");

    function MixedContent() {
      useEffect(() => {
        started.resolve();
      }, []);
      return createElement(
        Fragment,
        null,
        createElement(
          Text,
          { id: "sync", x: 0, y: 0, width: 240, height: 48 },
          "同步内容",
        ),
        createElement(
          Text,
          { id: "prefetched", x: 0, y: 52, width: 240, height: 48 },
          prefetched,
        ),
        createElement(
          Suspense,
          {
            fallback: createElement(
              Text,
              { x: 0, y: 104, width: 240, height: 48 },
              "外层等待",
            ),
          },
          createElement(Read, {
            promise: outer.promise,
            onRead: () => outerRead.resolve(),
          }),
        ),
        createElement(
          Suspense,
          {
            fallback: createElement(
              Text,
              { x: 0, y: 156, width: 240, height: 48 },
              "嵌套等待",
            ),
          },
          createElement(
            Suspense,
            { fallback: null },
            createElement(Read, {
              promise: inner.promise,
              onRead: () => innerRead.resolve(),
            }),
          ),
        ),
      );
    }

    const pendingBuild = build(deck(createElement(MixedContent)));
    let settled = false;
    void pendingBuild.then(
      () => {
        settled = true;
      },
      () => {
        settled = true;
      },
    );

    await started.promise;
    assert.equal(settled, false);
    outer.resolve("外层最终内容");
    await outerRead.promise;
    assert.equal(settled, false);
    inner.resolve("嵌套最终内容");
    await innerRead.promise;

    const result = await pendingBuild;
    const file = await validatePptx(result.pptx);

    assert.deepEqual(file.errors, []);
    assert.deepEqual(
      result.report.slides[0]?.objects.map(({ text }) => text),
      ["同步内容", "入口预取", "外层最终内容", "嵌套最终内容"],
    );
    assert.deepEqual(
      file.slides[0]?.objects.map(({ text }) => text?.text),
      ["同步内容", "入口预取", "外层最终内容", "嵌套最终内容"],
    );
  });

  it("awaits sequential synchronous Effect updates even when state commits produce no host changes", async () => {
    const updates: number[] = [];
    const stableHost = createElement(
      Text,
      { id: "stable-host", x: 0, y: 0, width: 240, height: 48 },
      "宿主文字保持不变",
    );

    function EffectUpdates() {
      const [value, setValue] = useState(0);
      useEffect(() => {
        updates.push(value);
        if (value < 2) setValue(value + 1);
      }, [value]);
      return stableHost;
    }

    const result = await build(deck(createElement(EffectUpdates)));
    const file = await validatePptx(result.pptx);

    assert.deepEqual(updates, [0, 1, 2]);
    assert.deepEqual(file.errors, []);
    assert.equal(result.report.slides[0]?.objects[0]?.text, "宿主文字保持不变");
    assert.equal(file.slides[0]?.objects[0]?.text?.text, "宿主文字保持不变");
  });

  it("generates fallbacks from author catches and error boundaries without producing business warnings", async () => {
    const caughtResource = Promise.withResolvers<string>();
    const boundaryResource = Promise.withResolvers<string>();
    const abandonedResource = new Promise<string>(() => {});
    const started = Promise.withResolvers<void>();
    const recovered = caughtResource.promise.catch(() => "catch 业务回退");
    const warnings: string[] = [];
    const originalError = console.error;

    function HandledErrors() {
      useEffect(() => {
        started.resolve();
      }, []);
      return createElement(
        Fragment,
        null,
        createElement(
          Suspense,
          { fallback: null },
          createElement(Read, { promise: recovered }),
        ),
        createElement(
          ErrorBoundary,
          null,
          createElement(
            Fragment,
            null,
            createElement(
              Suspense,
              { fallback: null },
              createElement(Read, { promise: abandonedResource }),
            ),
            createElement(
              Suspense,
              { fallback: null },
              createElement(Read, { promise: boundaryResource.promise }),
            ),
          ),
        ),
      );
    }

    console.error = (...values: unknown[]) => {
      warnings.push(values.map(String).join(" "));
    };
    try {
      const pendingBuild = build(deck(createElement(HandledErrors)));
      await started.promise;
      caughtResource.reject(new Error("catch recovered"));
      boundaryResource.reject(new Error("boundary recovered"));

      const result = await pendingBuild;
      const file = await validatePptx(result.pptx);

      assert.deepEqual(file.errors, []);
      assert.deepEqual(
        result.report.diagnostics.map(({ code }) => code),
        ["font-measurement-unavailable", "font-measurement-unavailable"],
      );
      assert.deepEqual(warnings, []);
      assert.deepEqual(
        result.report.slides[0]?.objects.map(({ text }) => text),
        ["catch 业务回退", "边界业务回退"],
      );
      assert.deepEqual(
        file.slides[0]?.objects.map(({ text }) => text?.text),
        ["catch 业务回退", "边界业务回退"],
      );
    } finally {
      console.error = originalError;
    }
  });

  it("fails build on unhandled direct Effect exception and unmounts other committed Effects", async () => {
    const controller = new AbortController();
    const failure = new Error("direct Effect failed");
    let cleaned = 0;

    function DirectEffectFailure() {
      useEffect(() => {
        return () => {
          cleaned += 1;
        };
      }, []);
      useEffect(() => {
        throw failure;
      }, []);
      return createElement(
        Text,
        { x: 0, y: 0, width: 240, height: 48 },
        "不能导出的已提交内容",
      );
    }

    await assert.rejects(
      build(deck(createElement(DirectEffectFailure)), {
        signal: controller.signal,
      }),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "react-render");
        assert.match(error.message, /direct Effect failed/u);
        return true;
      },
    );

    assert.equal(cleaned, 1);
    assert.equal(getEventListeners(controller.signal, "abort").length, 0);
    const recovered = await build(
      deck(
        createElement(Text, { x: 0, y: 0, width: 240, height: 48 }, "恢复构建"),
      ),
    );
    const file = await validatePptx(recovered.pptx);
    assert.deepEqual(file.errors, []);
    assert.equal(file.slides[0]?.objects[0]?.text?.text, "恢复构建");
  });

  it("does not return successful artifacts when passive Effect cleanup throws during normal content unmount", async () => {
    const controller = new AbortController();
    const cleanupFailure = new Error("passive cleanup failure");

    function CleanupFailure() {
      useEffect(() => {
        return () => {
          throw cleanupFailure;
        };
      }, []);
      return createElement(
        Text,
        { x: 0, y: 0, width: 240, height: 48 },
        "清理失败时不能返回",
      );
    }

    await assert.rejects(
      build(deck(createElement(CleanupFailure)), {
        signal: controller.signal,
      }),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "react-render");
        assert.match(error.message, /passive cleanup failure/u);
        return true;
      },
    );
    assert.equal(getEventListeners(controller.signal, "abort").length, 0);

    const recovered = await build(
      deck(
        createElement(
          Text,
          { x: 0, y: 0, width: 240, height: 48 },
          "清理后恢复",
        ),
      ),
    );
    const file = await validatePptx(recovered.pptx);
    assert.deepEqual(file.errors, []);
    assert.equal(file.slides[0]?.objects[0]?.text?.text, "清理后恢复");
  });

  it("still rejects and removes cancellation listeners when LayoutEffect cleanup throws", async () => {
    const controller = new AbortController();
    const cleanupFailure = new Error("layout cleanup failure");

    function LayoutCleanupFailure() {
      useLayoutEffect(() => {
        return () => {
          throw cleanupFailure;
        };
      }, []);
      return createElement(
        Text,
        { x: 0, y: 0, width: 240, height: 48 },
        "LayoutEffect 内容",
      );
    }

    await assert.rejects(
      build(deck(createElement(LayoutCleanupFailure)), {
        signal: controller.signal,
      }),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "react-render");
        assert.match(error.message, /layout cleanup failure/u);
        return true;
      },
    );
    assert.equal(getEventListeners(controller.signal, "abort").length, 0);
  });

  it("fails build when unhandled Promise errors propagate to the build entry", async () => {
    const resource = Promise.withResolvers<string>();
    const started = Promise.withResolvers<void>();
    const failure = new Error("unhandled data error");
    let cleaned = 0;

    function UnhandledResource() {
      useEffect(() => {
        started.resolve();
        return () => {
          cleaned += 1;
        };
      }, []);
      return createElement(
        Suspense,
        { fallback: null },
        createElement(Read, { promise: resource.promise }),
      );
    }

    const pendingBuild = build(deck(createElement(UnhandledResource)));
    await started.promise;
    const rejected = assert.rejects(pendingBuild, (error: unknown) => {
      assert.ok(error instanceof BuildError);
      assert.equal(error.diagnostics[0]?.code, "react-render");
      assert.match(error.message, /unhandled data error/u);
      return true;
    });
    resource.reject(failure);
    await rejected;
    assert.equal(cleaned, 1);
  });

  it("does not await undeclared background Effects, keeping report and PPTX unchanged after completion", async () => {
    const late = Promise.withResolvers<string>();
    const started = Promise.withResolvers<void>();

    function BackgroundUpdate() {
      const [text, setText] = useState("已提交内容");
      useEffect(() => {
        started.resolve();
        void late.promise.then(setText);
      }, []);
      return createElement(
        Text,
        { id: "background", x: 0, y: 0, width: 240, height: 48 },
        text,
      );
    }

    const pendingBuild = build(deck(createElement(BackgroundUpdate)));
    await started.promise;
    const result = await pendingBuild;
    const reportBefore = result.report;
    const pptxBefore = result.pptx.slice();

    late.resolve("构建完成后的后台更新");
    await late.promise;
    const file = await validatePptx(result.pptx);

    assert.deepEqual(result.report, reportBefore);
    assert.deepEqual(result.pptx, pptxBefore);
    assert.deepEqual(file.errors, []);
    assert.equal(result.report.slides[0]?.objects[0]?.text, "已提交内容");
    assert.equal(file.slides[0]?.objects[0]?.text?.text, "已提交内容");
  });

  it("isolates controlled HTTP data inside components across instances and concurrent builds", async () => {
    const responses = new Map<string, ServerResponse>();
    const requested = Promise.withResolvers<void>();
    const firstInstanceRead = Promise.withResolvers<void>();
    const secondInstanceRead = Promise.withResolvers<void>();
    const otherBuildRead = Promise.withResolvers<void>();
    const readSignals = new Map([
      ["/first", firstInstanceRead],
      ["/second", secondInstanceRead],
      ["/other-build", otherBuildRead],
    ]);
    const server = createServer((request, response) => {
      const path = request.url;
      if (path === undefined) {
        response.destroy(new Error("request has no pathname"));
        return;
      }
      responses.set(path, response);
      if (responses.size === 3) requested.resolve();
    });

    function SalesTable({ source, id }: { source: string; id: string }) {
      const [promise, setPromise] = useState<Promise<string> | null>(null);
      useEffect(() => {
        const read = fetch(source)
          .then((response) => response.text())
          .then((content) => {
            readSignals.get(new URL(source).pathname)?.resolve();
            return content;
          });
        setPromise(read);
      }, [source]);
      return promise === null
        ? null
        : createElement(
            Suspense,
            { fallback: null },
            createElement(Read, { promise, id }),
          );
    }

    try {
      await new Promise<void>((resolve, reject) => {
        server.once("error", reject);
        server.listen(0, "127.0.0.1", resolve);
      });
      const address = server.address();
      assert.ok(address && typeof address !== "string");
      const origin = `http://127.0.0.1:${address.port}`;
      const firstBuild = build(
        deck(
          createElement(
            Fragment,
            null,
            createElement(SalesTable, {
              source: `${origin}/first`,
              id: "first-instance",
            }),
            createElement(SalesTable, {
              source: `${origin}/second`,
              id: "second-instance",
            }),
          ),
        ),
      );
      const otherBuild = build(
        deck(
          createElement(SalesTable, {
            source: `${origin}/other-build`,
            id: "other-build",
          }),
        ),
      );
      let firstStatus = "pending";
      let otherStatus = "pending";
      void firstBuild.then(
        () => {
          firstStatus = "done";
        },
        () => {
          firstStatus = "failed";
        },
      );
      void otherBuild.then(
        () => {
          otherStatus = "done";
        },
        () => {
          otherStatus = "failed";
        },
      );

      await requested.promise;
      assert.equal(firstStatus, "pending");
      assert.equal(otherStatus, "pending");

      const otherResponse = responses.get("/other-build");
      assert.ok(otherResponse);
      otherResponse.end("另一构建的数据");
      await otherBuildRead.promise;
      const otherResult = await otherBuild;
      const otherFile = await validatePptx(otherResult.pptx);
      assert.deepEqual(otherFile.errors, []);
      assert.equal(
        otherResult.report.slides[0]?.objects[0]?.text,
        "另一构建的数据",
      );
      assert.equal(
        otherFile.slides[0]?.objects[0]?.text?.text,
        "另一构建的数据",
      );
      assert.equal(firstStatus, "pending");

      const secondResponse = responses.get("/second");
      assert.ok(secondResponse);
      secondResponse.end("第二个组件实例的数据");
      await secondInstanceRead.promise;
      assert.equal(firstStatus, "pending");

      const firstResponse = responses.get("/first");
      assert.ok(firstResponse);
      firstResponse.end("第一个组件实例的数据");
      await firstInstanceRead.promise;
      const firstResult = await firstBuild;
      const firstFile = await validatePptx(firstResult.pptx);

      assert.deepEqual(firstFile.errors, []);
      assert.deepEqual(
        firstResult.report.slides[0]?.objects.map(({ authorId, text }) => ({
          authorId,
          text,
        })),
        [
          { authorId: "first-instance", text: "第一个组件实例的数据" },
          { authorId: "second-instance", text: "第二个组件实例的数据" },
        ],
      );
      assert.deepEqual(
        firstFile.slides[0]?.objects.map(({ name, text }) => ({
          name,
          text: text?.text,
        })),
        [
          { name: "Text 1", text: "第一个组件实例的数据" },
          { name: "Text 2", text: "第二个组件实例的数据" },
        ],
      );
      assert.equal(otherStatus, "done");
      assert.equal(firstStatus, "done");
    } finally {
      server.closeAllConnections();
      if (server.listening) {
        await new Promise<void>((resolve, reject) => {
          server.close((error) => {
            if (error !== undefined) reject(error);
            else resolve();
          });
        });
      }
    }
  }, 30_000);
});
