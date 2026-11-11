/**
 * Plain authoring declaration tree consumed by the formal build.
 *
 * React components and command authoring produce editable common nodes that expose this structure
 * once; the formal build reads it without executing React lifecycle or depending on React at
 * runtime. Build-time parsing and layout remain derived structures of the current declaration, not
 * another long-lived authoring model.
 */
export type DeclarationText = {
  readonly type: "#text";
  readonly text: string;
};

export type DeclarationNode = {
  readonly type: string;
  readonly props: Readonly<Record<string, unknown>>;
  readonly children: readonly DeclarationChild[];
};

export type DeclarationChild = DeclarationNode | DeclarationText;

export type DeclarationContainer = {
  readonly children: readonly DeclarationChild[];
};
