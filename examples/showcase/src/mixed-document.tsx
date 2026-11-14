import { expandDocument } from "@react-pptx/react-pptx";

import { ShowcaseDeck } from "./deck.tsx";
import { createComparisonPage } from "./pages/comparison-document.ts";

/** Expands the JSX deck and replaces its comparison page through public editable document operations. */
export async function createShowcaseDocument() {
  const document = await expandDocument(<ShowcaseDeck />);
  const original = document.query.get("comparison").node;
  const replacement = createComparisonPage();

  replacement.moveTo(document, original);
  original.remove();

  return document;
}
