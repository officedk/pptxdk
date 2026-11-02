import { Deck, Slide } from "@react-pptx/react-pptx";
import { Suspense, use, useEffect } from "react";

declare global {
  var __reactPptxCancellation:
    | {
        readonly controller: AbortController;
        readonly reason: Error;
      }
    | undefined;
}

const pending = new Promise<string>(() => {});

function PendingContent() {
  return use(pending);
}

function NotifyMount() {
  useEffect(() => {
    const cancellation = globalThis.__reactPptxCancellation;
    if (cancellation !== undefined) {
      cancellation.controller.abort(cancellation.reason);
    }
    process.stdout.write("__react_pptx_cancel_ready__\n");
  }, []);
  return null;
}

export default function CancelledDeck() {
  return (
    <Deck>
      <Slide slideId="cancelled">
        <NotifyMount />
        <Suspense fallback={null}>
          <PendingContent />
        </Suspense>
      </Slide>
    </Deck>
  );
}
