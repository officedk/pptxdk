import { Deck, Slide } from "@officedk/pptxdk/react";
import { Suspense, use, useEffect } from "react";

declare global {
  var __pptxdkCancellation:
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
    const cancellation = globalThis.__pptxdkCancellation;
    if (cancellation !== undefined) {
      cancellation.controller.abort(cancellation.reason);
    }
    process.stdout.write("__pptxdk_cancel_ready__\n");
  }, []);
  return null;
}

export default function CancelledDeck() {
  return (
    <Deck>
      <Slide id="cancelled">
        <NotifyMount />
        <Suspense fallback={null}>
          <PendingContent />
        </Suspense>
      </Slide>
    </Deck>
  );
}
