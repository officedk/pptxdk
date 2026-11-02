const state = globalThis.__reactPptxSourceLoadState;
if (state === undefined) throw new Error("缺少源码加载测试状态");
state.concurrentStarted = true;
await new Promise<void>((resolve) => {
  state.allowConcurrentImport = resolve;
});

const deck = await import("./source-after-cancellation.tsx");
export default deck.default;
