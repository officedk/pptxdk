export type HostText = {
  readonly type: "#text";
  text: string;
  hidden: boolean;
};

export class HostNode {
  readonly type: string;
  props: Readonly<Record<string, unknown>>;
  readonly children: HostChild[] = [];
  fiber: unknown;
  hidden = false;

  constructor(
    type: string,
    props: Readonly<Record<string, unknown>>,
    fiber: unknown,
  ) {
    this.type = type;
    this.props = props;
    this.fiber = fiber;
  }
}

export type HostChild = HostNode | HostText;
export type HostContainer = { readonly children: HostChild[] };

export function withoutChildren(
  props: Readonly<Record<string, unknown>>,
): Readonly<Record<string, unknown>> {
  if (!("children" in props)) return props;
  const { children: _children, ...rest } = props;
  return rest;
}
