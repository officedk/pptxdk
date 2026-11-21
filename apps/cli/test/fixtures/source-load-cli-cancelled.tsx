import { Deck, Slide } from "@react-pptx/react-pptx";

await new Promise<void>((_resolve, reject) => {
  const keepAlive = setInterval(() => {}, 1000);
  process.once("SIGINT", () => {
    clearInterval(keepAlive);
    reject(new Error("Author source preparation cooperatively cancelled"));
  });
  process.stdout.write("__pptxdk_source_load_cancel_ready__\n");
});

export default function SourceLoadCancelledDeck() {
  return (
    <Deck>
      <Slide id="cancelled-source-load" />
    </Deck>
  );
}
