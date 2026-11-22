import type { ReactElement } from "react";

declare const elementKindBrand: unique symbol;

/**
 * Invariant JSX element brand distinguishing native capability categories (e.g. text vs image)
 * so `<Slide placeholders={{ ... }} />` rejects mismatched element kinds at compile time.
 */
export interface TypedElement<Kind extends string> extends ReactElement {
  readonly [elementKindBrand]: (kind: Kind) => Kind;
}
