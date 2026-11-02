import assert from "node:assert/strict";
import { createServer, type ServerResponse } from "node:http";

import { BuildError, build, Deck, Slide, Text } from "@react-pptx/react-pptx";
import { validatePptx } from "@react-pptx/validation";
import {
  Component,
  createElement,
  Fragment,
  Suspense,
  use,
  useEffect,
  useState,
  type ReactNode,
} from "react";
import { describe, it } from "vitest";

function deck(children: ReactNode) {
  return createElement(
    Deck,
    { width: 960, height: 540 },
    createElement(Slide, { slideId: "async" }, children),
  );
}

function Read({
  promise,
  onRead,
  testId,
}: {
  promise: Promise<string>;
  onRead?: () => void;
  testId?: string;
}) {
  const text = use(promise);
  useEffect(() => {
    onRead?.();
  }, [onRead, text]);
  return createElement(
    Text,
    {
      ...(testId === undefined ? {} : { testId }),
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

describe("公开构建中的 React 异步内容", () => {
  it("等待 Suspense 中的稳定 Promise 并生成最终内容", async () => {
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

  it("混合同步与入口预取内容，并等待嵌套和空 fallback 的最终提交", async () => {
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
          { testId: "sync", x: 0, y: 0, width: 240, height: 48 },
          "同步内容",
        ),
        createElement(
          Text,
          { testId: "prefetched", x: 0, y: 52, width: 240, height: 48 },
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

  it("等待同步 Effect 连续更新，即使状态提交没有宿主变化", async () => {
    const updates: number[] = [];
    const stableHost = createElement(
      Text,
      { testId: "stable-host", x: 0, y: 0, width: 240, height: 48 },
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

  it("作者 catch 和错误边界的回退正常生成且不产生业务警告", async () => {
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
      caughtResource.reject(new Error("catch 已恢复"));
      boundaryResource.reject(new Error("边界已恢复"));

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

  it("未处理并传播到构建入口的 Promise 错误使构建失败", async () => {
    const resource = Promise.withResolvers<string>();
    const started = Promise.withResolvers<void>();
    const failure = new Error("未处理的数据错误");
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
      assert.match(error.message, /未处理的数据错误/u);
      return true;
    });
    resource.reject(failure);
    await rejected;
    assert.equal(cleaned, 1);
  });

  it("不等待未声明的后台 Effect，完成后的报告与 PPTX 保持不变", async () => {
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
        { testId: "background", x: 0, y: 0, width: 240, height: 48 },
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

  it("组件内受控 HTTP 数据按实例和并发构建隔离", async () => {
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
        response.destroy(new Error("请求没有路径"));
        return;
      }
      responses.set(path, response);
      if (responses.size === 3) requested.resolve();
    });

    function SalesTable({
      source,
      testId,
    }: {
      source: string;
      testId: string;
    }) {
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
            createElement(Read, { promise, testId }),
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
              testId: "first-instance",
            }),
            createElement(SalesTable, {
              source: `${origin}/second`,
              testId: "second-instance",
            }),
          ),
        ),
      );
      const otherBuild = build(
        deck(
          createElement(SalesTable, {
            source: `${origin}/other-build`,
            testId: "other-build",
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
        firstResult.report.slides[0]?.objects.map(({ testId, text }) => ({
          testId,
          text,
        })),
        [
          { testId: "first-instance", text: "第一个组件实例的数据" },
          { testId: "second-instance", text: "第二个组件实例的数据" },
        ],
      );
      assert.deepEqual(
        firstFile.slides[0]?.objects.map(({ name, text }) => ({
          name,
          text: text?.text,
        })),
        [
          { name: "first-instance", text: "第一个组件实例的数据" },
          { name: "second-instance", text: "第二个组件实例的数据" },
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
