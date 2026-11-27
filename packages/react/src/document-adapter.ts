import { BuildError, toBuildDiagnostic } from "@pptxdk/core";
import { CommonContainer, CommonNode } from "@pptxdk/core/document";
import {
  hostNodeCreator,
  HostTextNode,
  insertDeclarationChild,
  removeDeclarationChild,
  replaceDeclarationProps,
  setDeclarationText,
  type DeclarationChild,
  type DocumentNodeCreator,
} from "@pptxdk/core/host";

import { hostMountMarker, hostMountPlaceholder } from "#/mount.ts";

export type AdapterChild = DeclarationChild | MountHostRecord;
export type HostInstance = CommonNode | MountHostRecord;
export type HostParent = CommonNode | MountHostRecord | HostContainer;

/** Renderer-side root during a single React expansion; it does not own a second document tree. */
export type HostContainer = {
  children: DeclarationChild[];
  hostOrder: AdapterChild[];
  mountedRecords: Set<MountHostRecord>;
  locked: boolean;
};

type AppliedPlaceholder = {
  readonly node: CommonNode;
  readonly hasPrevious: boolean;
  readonly previousValue: unknown;
};

/**
 * In-memory adapter record for a `<Mount>` element; never enters the editable `CommonNode` tree.
 */
export type MountHostRecord = {
  readonly kind: "mount";
  readonly container: HostContainer;
  adoptedNodes: CommonNode[];
  slot: CommonContainer | undefined;
  placeholder: string | undefined;
  appliedPlaceholder: AppliedPlaceholder | undefined;
  hostParent: HostParent | undefined;
  committed: boolean;
  slotChildren: AdapterChild[];
};

const nodeOwnerContainer = new WeakMap<object, HostContainer>();
const activeMounts = new WeakMap<CommonNode, MountHostRecord>();
const hostChildrenByParent = new WeakMap<CommonNode, AdapterChild[]>();

export function createHostContainer(): HostContainer {
  return {
    children: [],
    hostOrder: [],
    mountedRecords: new Set(),
    locked: false,
  };
}

export function isMountHostRecord(value: unknown): value is MountHostRecord {
  return (
    typeof value === "object" &&
    value !== null &&
    "kind" in value &&
    value.kind === "mount" &&
    "adoptedNodes" in value
  );
}

function isHostContainer(value: HostParent): value is HostContainer {
  return "hostOrder" in value && "mountedRecords" in value;
}

function isLocked(target: AdapterChild | HostContainer): boolean {
  if (isMountHostRecord(target)) return target.container.locked;
  if ("hostOrder" in target && "mountedRecords" in target) return target.locked;
  return nodeOwnerContainer.get(target)?.locked === true;
}

function mountError(message: string): BuildError {
  return new BuildError([
    toBuildDiagnostic("mount", "invalid-structure", message),
  ]);
}

function isDescendantOf(candidate: CommonNode, ancestor: CommonNode): boolean {
  for (
    let current = candidate.parent;
    current !== undefined;
    current = current.parent
  ) {
    if (current === ancestor) return true;
  }
  return false;
}

function parseMountDeclaration(props: Readonly<Record<PropertyKey, unknown>>): {
  readonly nodes: CommonNode[];
  readonly slot: CommonContainer | undefined;
  readonly placeholder: string | undefined;
} {
  const rawNode = props["node"];
  const rawSlot = props["slot"];
  const rawPlaceholder = props[hostMountPlaceholder];
  const placeholder =
    typeof rawPlaceholder === "string" ? rawPlaceholder : undefined;

  if (rawSlot !== undefined) {
    if (!(rawNode instanceof CommonNode)) {
      throw mountError(
        "Mount with a slot requires node to be a single CommonNode",
      );
    }
    if (!(rawSlot instanceof CommonContainer)) {
      throw mountError("Mount slot must be a CommonContainer");
    }
    if (rawSlot !== rawNode && !isDescendantOf(rawSlot, rawNode)) {
      throw mountError(
        "Mount slot must be the adopted node itself or one of its descendant containers",
      );
    }
    return { nodes: [rawNode], slot: rawSlot, placeholder };
  }

  if (rawNode === null || rawNode === undefined) {
    return { nodes: [], slot: undefined, placeholder };
  }
  if (rawNode instanceof CommonNode) {
    return { nodes: [rawNode], slot: undefined, placeholder };
  }
  if (Array.isArray(rawNode)) {
    const nodes: CommonNode[] = [];
    const seen = new Set<CommonNode>();
    for (const item of rawNode) {
      if (!(item instanceof CommonNode)) {
        throw mountError(
          "Mount node array must contain only CommonNode instances",
        );
      }
      if (seen.has(item)) {
        throw mountError(
          "Mount cannot adopt the same CommonNode instance more than once",
        );
      }
      seen.add(item);
      nodes.push(item);
    }
    return { nodes, slot: undefined, placeholder };
  }

  throw mountError(
    "Mount node must be a CommonNode, a readonly CommonNode array, null, or undefined",
  );
}

function withoutHostMetadata(
  props: Readonly<Record<string, unknown>>,
): Readonly<Record<string, unknown>> {
  const {
    children: _children,
    [hostNodeCreator]: _create,
    ...declaration
  } = props;
  return declaration;
}

function isDocumentNodeCreator(value: unknown): value is DocumentNodeCreator {
  return typeof value === "function";
}

function hostChildrenOf(parent: CommonNode): AdapterChild[] {
  let list = hostChildrenByParent.get(parent);
  if (list === undefined) {
    list = [...parent.children];
    hostChildrenByParent.set(parent, list);
  }
  return list;
}

function orderListOf(parent: HostParent): AdapterChild[] {
  if (isHostContainer(parent)) return parent.hostOrder;
  if (isMountHostRecord(parent)) return parent.slotChildren;
  return hostChildrenOf(parent);
}

function realSlotContainer(record: MountHostRecord): CommonContainer {
  if (record.slot === undefined) {
    throw mountError("Mount cannot accept children without a slot container");
  }
  return record.slot;
}

function realParentNode(parent: CommonNode | MountHostRecord): CommonNode {
  return isMountHostRecord(parent) ? realSlotContainer(parent) : parent;
}

function realChildrenOf(parent: HostParent): readonly DeclarationChild[] {
  if (isHostContainer(parent)) return parent.children;
  return realParentNode(parent).children;
}

function firstLiveChildOf(
  entry: AdapterChild,
  liveChildren: readonly DeclarationChild[],
): DeclarationChild | undefined {
  if (!isMountHostRecord(entry)) {
    return liveChildren.includes(entry) ? entry : undefined;
  }
  if (!entry.committed) return undefined;
  for (const adopted of entry.adoptedNodes) {
    if (liveChildren.includes(adopted)) return adopted;
  }
  return undefined;
}

function resolveAnchorAfter(
  order: readonly AdapterChild[],
  startIndex: number,
  liveChildren: readonly DeclarationChild[],
): DeclarationChild | undefined {
  for (let i = startIndex; i < order.length; i += 1) {
    const candidate = order[i];
    if (candidate === undefined) continue;
    const live = firstLiveChildOf(candidate, liveChildren);
    if (live !== undefined) return live;
  }
  return undefined;
}

function recordOrderInsert(
  order: AdapterChild[],
  child: AdapterChild,
  before: AdapterChild | undefined,
): void {
  const existing = order.indexOf(child);
  if (existing !== -1) order.splice(existing, 1);
  const index = before === undefined ? -1 : order.indexOf(before);
  if (index === -1) order.push(child);
  else order.splice(index, 0, child);
}

function sameNodeList(
  a: readonly CommonNode[],
  b: readonly CommonNode[],
): boolean {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i += 1) {
    if (!Object.is(a[i], b[i])) return false;
  }
  return true;
}

function assertCanAdopt(
  nodes: readonly CommonNode[],
  record: MountHostRecord,
): void {
  for (const node of nodes) {
    const existingMount = activeMounts.get(node);
    if (existingMount !== undefined && existingMount !== record) {
      throw mountError(
        `CommonNode (${node.kind}) is already mounted by another active Mount`,
      );
    }
    if (node.parent !== undefined && existingMount !== record) {
      throw mountError(
        `CommonNode (${node.kind}) is already attached to a parent container`,
      );
    }
  }
}

function spliceRootChildren(
  container: HostContainer,
  items: readonly DeclarationChild[],
  before: DeclarationChild | undefined,
): void {
  for (const item of items) {
    const current = container.children.indexOf(item);
    if (current !== -1) container.children.splice(current, 1);
    const index =
      before === undefined ? -1 : container.children.indexOf(before);
    if (index === -1) container.children.push(item);
    else container.children.splice(index, 0, item);
  }
}

function removeFromRootChildren(
  container: HostContainer,
  items: readonly DeclarationChild[],
): void {
  for (const item of items) {
    const index = container.children.indexOf(item);
    if (index !== -1) container.children.splice(index, 1);
  }
}

function insertAdoptedIntoParent(
  parent: HostParent,
  nodes: readonly CommonNode[],
  before: DeclarationChild | undefined,
  record: MountHostRecord,
): void {
  if (nodes.length === 0) return;
  if (isHostContainer(parent)) {
    spliceRootChildren(parent, nodes, before);
    for (const node of nodes) activeMounts.set(node, record);
    return;
  }

  const targetNode = realParentNode(parent);
  const inserted: CommonNode[] = [];
  try {
    for (const node of nodes) {
      insertDeclarationChild(targetNode, node, before);
      inserted.push(node);
    }
  } catch (error) {
    for (const node of inserted) {
      removeDeclarationChild(targetNode, node);
    }
    throw mountError(error instanceof Error ? error.message : String(error));
  }
  for (const node of inserted) activeMounts.set(node, record);
}

function detachAdoptedFromParent(
  parent: HostParent,
  nodes: readonly CommonNode[],
  record: MountHostRecord,
  locked: boolean,
): void {
  for (const node of nodes) {
    if (activeMounts.get(node) === record) activeMounts.delete(node);
  }
  if (isHostContainer(parent)) {
    removeFromRootChildren(parent, nodes);
    return;
  }
  if (locked) return;
  const targetNode = realParentNode(parent);
  for (const node of nodes) {
    removeDeclarationChild(targetNode, node);
  }
}

function applyPlaceholderBinding(record: MountHostRecord): void {
  if (record.placeholder === undefined || record.adoptedNodes.length !== 1) {
    return;
  }
  const node = record.adoptedNodes[0]!;
  if (
    record.appliedPlaceholder !== undefined &&
    record.appliedPlaceholder.node === node
  ) {
    node.set({ placeholder: record.placeholder });
    return;
  }
  const currentProps = node.props;
  const hasPrevious = Object.prototype.hasOwnProperty.call(
    currentProps,
    "placeholder",
  );
  record.appliedPlaceholder = {
    node,
    hasPrevious,
    previousValue: currentProps["placeholder"],
  };
  node.set({ placeholder: record.placeholder });
}

function restorePlaceholderBinding(record: MountHostRecord): void {
  const applied = record.appliedPlaceholder;
  if (applied === undefined) return;
  record.appliedPlaceholder = undefined;
  applied.node.set({
    placeholder: applied.hasPrevious ? applied.previousValue : undefined,
  });
}

function commitMountIntoParent(
  parent: HostParent,
  record: MountHostRecord,
  anchor: DeclarationChild | undefined,
): void {
  const isReposition = record.committed && record.hostParent === parent;
  if (!isReposition) {
    assertCanAdopt(record.adoptedNodes, record);
  }
  record.hostParent = parent;
  insertAdoptedIntoParent(parent, record.adoptedNodes, anchor, record);
  try {
    if (!isReposition) {
      applyPlaceholderBinding(record);
    }
    record.committed = true;
    record.container.mountedRecords.add(record);
    if (!isReposition && record.slot !== undefined) {
      attachSlotChildren(record);
    }
  } catch (error) {
    if (!isReposition) {
      restorePlaceholderBinding(record);
      detachAdoptedFromParent(parent, record.adoptedNodes, record, false);
      record.committed = false;
      record.container.mountedRecords.delete(record);
    }
    throw error;
  }
}

function attachSlotChildren(record: MountHostRecord): void {
  const slot = record.slot;
  if (slot === undefined) return;
  for (const child of record.slotChildren) {
    if (isMountHostRecord(child)) {
      if (!child.committed) commitMountHostInstance(child);
      else
        insertAdoptedIntoParent(record, child.adoptedNodes, undefined, child);
    } else {
      insertDeclarationChild(slot, child, undefined);
    }
  }
}

function detachSlotChildren(record: MountHostRecord, locked: boolean): void {
  const slot = record.slot;
  if (slot === undefined) return;
  for (const child of record.slotChildren) {
    if (isMountHostRecord(child)) {
      detachAdoptedFromParent(record, child.adoptedNodes, child, locked);
    } else if (!locked) {
      removeDeclarationChild(slot, child);
    }
  }
}

function releaseMountSubtree(entry: AdapterChild, locked: boolean): void {
  if (isMountHostRecord(entry)) {
    for (const child of entry.slotChildren) {
      releaseMountSubtree(child, locked);
    }
    if (!locked && entry.slot !== undefined) {
      for (const child of entry.slotChildren) {
        if (!isMountHostRecord(child)) {
          removeDeclarationChild(entry.slot, child);
        }
      }
    }
    if (entry.hostParent !== undefined && entry.committed) {
      detachAdoptedFromParent(
        entry.hostParent,
        entry.adoptedNodes,
        entry,
        locked,
      );
    } else {
      for (const node of entry.adoptedNodes) {
        if (activeMounts.get(node) === entry) activeMounts.delete(node);
      }
    }
    entry.container.mountedRecords.delete(entry);
    if (!locked) {
      restorePlaceholderBinding(entry);
      entry.committed = false;
      entry.hostParent = undefined;
    }
    return;
  }

  if (entry instanceof CommonNode) {
    const children = hostChildrenByParent.get(entry);
    if (children !== undefined) {
      for (const child of children) {
        releaseMountSubtree(child, locked);
      }
      if (!locked) hostChildrenByParent.delete(entry);
    }
  }
}

export function createHostInstance(
  type: string,
  props: Readonly<Record<PropertyKey, unknown>>,
  container: HostContainer,
): HostInstance {
  if (props[hostMountMarker] === true) {
    const { nodes, slot, placeholder } = parseMountDeclaration(props);
    return {
      kind: "mount",
      container,
      adoptedNodes: nodes,
      slot,
      placeholder,
      appliedPlaceholder: undefined,
      hostParent: undefined,
      committed: false,
      slotChildren: [],
    };
  }

  const create = props[hostNodeCreator];
  const declaration = withoutHostMetadata(props);
  const instance = isDocumentNodeCreator(create)
    ? create(declaration)
    : new CommonNode(type, type, declaration);
  nodeOwnerContainer.set(instance, container);
  return instance;
}

export function createHostText(
  text: string,
  container: HostContainer,
): HostTextNode {
  const node = new HostTextNode(text);
  nodeOwnerContainer.set(node, container);
  return node;
}

export function updateHostText(text: HostTextNode, next: string): void {
  if (isLocked(text)) return;
  setDeclarationText(text, next);
}

function updateMountInstance(
  record: MountHostRecord,
  nextProps: Readonly<Record<PropertyKey, unknown>>,
): void {
  if (isLocked(record)) return;
  const {
    nodes: nextNodes,
    slot: nextSlot,
    placeholder: nextPlaceholder,
  } = parseMountDeclaration(nextProps);
  const nodesUnchanged = sameNodeList(record.adoptedNodes, nextNodes);
  const slotUnchanged = record.slot === nextSlot;
  const placeholderUnchanged = record.placeholder === nextPlaceholder;
  if (nodesUnchanged && slotUnchanged && placeholderUnchanged) return;

  assertCanAdopt(nextNodes, record);

  const prevNodes = record.adoptedNodes;
  const prevSlot = record.slot;

  if (!slotUnchanged && prevSlot !== undefined && record.committed) {
    detachSlotChildren(record, false);
  }

  if (!nodesUnchanged && record.hostParent !== undefined && record.committed) {
    const order = orderListOf(record.hostParent);
    const recordIndex = order.indexOf(record);
    const anchor =
      recordIndex === -1
        ? undefined
        : resolveAnchorAfter(
            order,
            recordIndex + 1,
            realChildrenOf(record.hostParent),
          );
    restorePlaceholderBinding(record);
    detachAdoptedFromParent(record.hostParent, prevNodes, record, false);
    try {
      insertAdoptedIntoParent(record.hostParent, nextNodes, anchor, record);
    } catch (error) {
      insertAdoptedIntoParent(record.hostParent, prevNodes, anchor, record);
      applyPlaceholderBinding(record);
      if (!slotUnchanged && prevSlot !== undefined) {
        attachSlotChildren(record);
      }
      throw error;
    }
  }

  record.adoptedNodes = nextNodes;
  record.slot = nextSlot;
  record.placeholder = nextPlaceholder;

  if (record.committed) {
    if (nextPlaceholder === undefined) {
      restorePlaceholderBinding(record);
    } else {
      applyPlaceholderBinding(record);
    }
  }

  if (!slotUnchanged && nextSlot !== undefined && record.committed) {
    attachSlotChildren(record);
  }
}

export function updateHostInstance(
  instance: HostInstance,
  props: Readonly<Record<string, unknown>>,
): void {
  if (isLocked(instance)) return;
  if (isMountHostRecord(instance)) {
    updateMountInstance(instance, props);
    return;
  }
  // React commits all current JSX props at once, unlike author updates that patch only provided
  // fields. Replace the declaration so fields absent from this commit do not survive.
  replaceDeclarationProps(instance, withoutHostMetadata(props));
}

export function appendInitialHostChild(
  parent: HostInstance,
  child: AdapterChild,
): void {
  if (isLocked(parent)) return;
  if (isMountHostRecord(parent) && parent.slot === undefined) {
    throw mountError("Mount cannot accept children without a slot container");
  }
  const order = orderListOf(parent);
  recordOrderInsert(order, child, undefined);
  if (isMountHostRecord(child)) {
    child.hostParent = parent;
    return;
  }
  // Defer slot attachment until commitMount so aborted render attempts never mutate external nodes.
  if (isMountHostRecord(parent)) return;
  insertDeclarationChild(parent, child);
}

export function commitMountHostInstance(instance: HostInstance): void {
  if (!isMountHostRecord(instance) || instance.committed) return;
  if (isLocked(instance)) return;

  if (
    instance.hostParent !== undefined &&
    isMountHostRecord(instance.hostParent) &&
    !instance.hostParent.committed
  ) {
    commitMountHostInstance(instance.hostParent);
    if (instance.committed) return;
  }

  const parent = instance.hostParent;
  if (parent === undefined) return;

  const order = orderListOf(parent);
  const index = order.indexOf(instance);
  const anchor =
    index === -1
      ? undefined
      : resolveAnchorAfter(order, index + 1, realChildrenOf(parent));

  commitMountIntoParent(parent, instance, anchor);
}

export function insertHostChild(
  parent: HostParent,
  child: AdapterChild,
  before?: AdapterChild,
): void {
  if (isLocked(parent) || isLocked(child)) return;
  const order = orderListOf(parent);
  const liveChildren = realChildrenOf(parent);
  recordOrderInsert(order, child, before);

  const index = order.indexOf(child);
  const anchor =
    before === undefined || index === -1
      ? undefined
      : resolveAnchorAfter(order, index + 1, liveChildren);

  if (isMountHostRecord(child)) {
    commitMountIntoParent(parent, child, anchor);
    return;
  }

  if (isHostContainer(parent)) {
    spliceRootChildren(parent, [child], anchor);
  } else {
    insertDeclarationChild(realParentNode(parent), child, anchor);
  }
}

export function removeHostChild(parent: HostParent, child: AdapterChild): void {
  const locked = isLocked(parent) || isLocked(child);
  const order = orderListOf(parent);
  const index = order.indexOf(child);
  if (index !== -1) order.splice(index, 1);

  releaseMountSubtree(child, locked);

  if (locked) return;
  if (!isMountHostRecord(child)) {
    if (isHostContainer(parent)) {
      removeFromRootChildren(parent, [child]);
    } else {
      removeDeclarationChild(realParentNode(parent), child);
    }
  }
}

/**
 * Closes the React ownership window on the completed candidate before renderer teardown. The
 * candidate shares its nodes with the returned document; teardown must only detach the renderer
 * root, never clear this content.
 */
export function lockHandoffTree(container: HostContainer): void {
  container.locked = true;
}

/**
 * Releases active mount tracking for a finished or aborted root. When the root was not handed off
 * (for example, an error or abort interrupted commit), any adopted nodes attached during the failed
 * commit are detached cleanly so they remain unowned.
 */
export function releaseContainerMounts(container: HostContainer): void {
  for (const record of Array.from(container.mountedRecords)) {
    releaseMountSubtree(record, container.locked);
  }
}
