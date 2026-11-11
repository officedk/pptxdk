import { Deck, Slide } from "@react-pptx/react-pptx";

declare global {
  var __reactPptxSourceLoadState:
    | {
        readonly controller: AbortController;
        started: boolean;
        concurrentStarted: boolean;
        allowConcurrentImport: (() => void) | undefined;
        rejectPending: ((reason: Error) => void) | undefined;
      }
    | undefined;
}

const state = globalThis.__reactPptxSourceLoadState;
if (state === undefined) throw new Error("Missing source load test state");
state.started = true;
await new Promise<void>((_resolve, reject) => {
  state.rejectPending = (reason) => reject(reason);
});

export default function PendingSourceDeck() {
  return (
    <Deck>
      <Slide id="pending-source" />
    </Deck>
  );
}
