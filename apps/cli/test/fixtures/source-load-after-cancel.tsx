const state = globalThis.__reactPptxSourceLoadState;
if (state === undefined) throw new Error("Missing source load test state");
state.concurrentStarted = true;
await new Promise<void>((resolve) => {
  state.allowConcurrentImport = resolve;
});

const deck = await import("./source-after-cancellation.tsx");
export default deck.default;
