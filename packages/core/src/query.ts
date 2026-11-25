import type {
  BuildQuery,
  GeometryBox,
  LayoutObservation,
  Locator,
  LocatorResourceObservation,
  LocatorObservation,
  ObjectObservation,
  ResourceObservation,
  SlideObservation,
  SlideQuery,
  TextMetrics,
} from "#/types.ts";

export class BuildQueryError extends Error {
  readonly code:
    | "missing-slide"
    | "missing-match"
    | "ambiguous-match"
    | "invalid-query"
    | "invalid-reference"
    | "foreign-reference"
    | "unbound-observation";

  constructor(code: BuildQueryError["code"], message: string) {
    super(message);
    this.code = code;
    this.name = "BuildQueryError";
  }
}

export type QueryContentReference =
  | {
      readonly kind: "layout";
      readonly id: number;
      readonly objectIds: readonly number[];
    }
  | { readonly kind: "object"; readonly id: number }
  | {
      readonly kind: "part";
      readonly box?: GeometryBox;
      readonly text?: string;
      readonly textMetrics?: TextMetrics;
    };

export type QueryContentNode = {
  readonly authorId?: string;
  readonly reference: QueryContentReference;
  readonly children: readonly QueryContentNode[];
};

export type BuildQuerySlideInput = {
  readonly observation: SlideObservation;
  readonly content: readonly QueryContentNode[];
  readonly resources?: readonly (ResourceObservation & {
    readonly objectId: number;
  })[];
};

type QueryNode = {
  readonly authorId?: string;
  readonly reference?: QueryContentReference;
  readonly children: readonly QueryNode[];
};

type LocatorEntry = {
  readonly node: QueryNode;
};

function assertQueryName(value: string, label: string): void {
  if (typeof value !== "string" || value.trim() === "") {
    throw new BuildQueryError(
      "invalid-query",
      `${label} must be a non-empty string`,
    );
  }
}

function collectObjects(
  node: QueryNode,
  objectsById: ReadonlyMap<number, ObjectObservation>,
  seen: Set<number>,
  objects: ObjectObservation[],
): void {
  const addObject = (id: number) => {
    const object = objectsById.get(id);
    if (object === undefined || seen.has(id)) return;
    seen.add(id);
    objects.push(object);
    for (const member of object.members ?? []) addObject(member.id);
  };

  if (node.reference?.kind === "object") addObject(node.reference.id);
  if (node.reference?.kind === "layout") {
    for (const id of node.reference.objectIds) addObject(id);
  }
  for (const child of node.children) {
    collectObjects(child, objectsById, seen, objects);
  }
}

function collectTexts(
  node: QueryNode,
  objectsById: ReadonlyMap<number, ObjectObservation>,
  seen: Set<number>,
  texts: string[],
): void {
  const addMemberObjects = (object: ObjectObservation) => {
    if (!seen.has(object.id)) {
      seen.add(object.id);
      if (object.text !== undefined) texts.push(object.text);
    }
    for (const member of object.members ?? []) addMemberObjects(member);
  };

  if (node.reference?.kind === "object") {
    const object = objectsById.get(node.reference.id);
    if (object !== undefined && !seen.has(object.id)) {
      seen.add(object.id);
      if (object.text !== undefined) {
        texts.push(object.text);
      } else {
        for (const child of node.children) {
          if (
            child.reference?.kind === "part" &&
            child.reference.text !== undefined
          ) {
            texts.push(child.reference.text);
          }
        }
      }
    }
  }
  for (const child of node.children) {
    if (child.reference?.kind !== "part") {
      collectTexts(child, objectsById, seen, texts);
    }
  }
  if (node.reference?.kind === "object") {
    const object = objectsById.get(node.reference.id);
    if (object !== undefined) {
      for (const member of object.members ?? []) addMemberObjects(member);
    }
  }
  if (node.reference?.kind === "layout") {
    for (const id of node.reference.objectIds) {
      const object = objectsById.get(id);
      if (object !== undefined) addMemberObjects(object);
    }
  }
}

function collectDescendantEntries(node: QueryNode): LocatorEntry[] {
  return node.children.flatMap((child) => [
    { node: child },
    ...collectDescendantEntries(child),
  ]);
}

/** Entries visible from one author scope: anonymous nodes stay transparent, named nodes are boundaries. */
function collectScopedEntries(node: QueryNode): LocatorEntry[] {
  return node.children.flatMap((child) => [
    { node: child },
    ...(child.authorId === undefined ? collectScopedEntries(child) : []),
  ]);
}

function queryNode(node: QueryContentNode): QueryNode {
  const children = node.children.map((child) => queryNode(child));
  return Object.freeze({
    ...(node.authorId === undefined ? {} : { authorId: node.authorId }),
    reference: node.reference,
    children: Object.freeze(children),
  });
}

function makeLocator(
  entry: LocatorEntry,
  layouts: ReadonlyMap<number, LayoutObservation>,
  objectsById: ReadonlyMap<number, ObjectObservation>,
  resourcesByObjectId: ReadonlyMap<number, readonly ResourceObservation[]>,
): Locator {
  const childLocators = (): readonly Locator[] =>
    Object.freeze(
      entry.node.children.map((child) =>
        makeLocator({ node: child }, layouts, objectsById, resourcesByObjectId),
      ),
    );
  const descendants = (): readonly Locator[] =>
    Object.freeze(
      collectDescendantEntries(entry.node).map((descendant) =>
        makeLocator(descendant, layouts, objectsById, resourcesByObjectId),
      ),
    );
  const findAuthorIds = (id: string): readonly Locator[] => {
    assertQueryName(id, "id");
    return Object.freeze(
      collectScopedEntries(entry.node)
        .filter(({ node }) => node.authorId === id)
        .map(({ node }) =>
          makeLocator({ node }, layouts, objectsById, resourcesByObjectId),
        ),
    );
  };
  // oxlint-disable-next-line unicorn/consistent-function-scoping
  const oneMatch = (
    matches: readonly Locator[],
    label: string,
    value: string,
  ): Locator => {
    if (matches.length === 0) {
      throw new BuildQueryError(
        "missing-match",
        `No match found for ${label} ${JSON.stringify(value)}`,
      );
    }
    if (matches.length > 1) {
      throw new BuildQueryError(
        "ambiguous-match",
        `${label} ${JSON.stringify(value)} matched multiple items`,
      );
    }
    const [match] = matches;
    if (match === undefined) throw new Error("Query match item missing");
    return match;
  };
  const observe = (): LocatorObservation => {
    const objects: ObjectObservation[] = [];
    collectObjects(entry.node, objectsById, new Set(), objects);
    const layoutBox =
      entry.node.reference?.kind === "layout"
        ? layouts.get(entry.node.reference.id)?.box
        : undefined;
    const objectBox =
      entry.node.reference?.kind === "object"
        ? objectsById.get(entry.node.reference.id)?.box
        : undefined;
    const partRef =
      entry.node.reference?.kind === "part" ? entry.node.reference : undefined;
    const text: string[] = [];
    if (partRef === undefined) {
      collectTexts(entry.node, objectsById, new Set(), text);
    }
    const selectedText = text.length === 0 ? undefined : text.join("\n");
    const box = partRef !== undefined ? partRef.box : (layoutBox ?? objectBox);
    const observedText = partRef !== undefined ? partRef.text : selectedText;
    const textMetrics =
      partRef !== undefined
        ? partRef.textMetrics
        : entry.node.reference?.kind === "object"
          ? objectsById.get(entry.node.reference.id)?.textMetrics
          : undefined;
    const resources: LocatorResourceObservation[] = objects.flatMap((object) =>
      (resourcesByObjectId.get(object.id) ?? []).map((resource) =>
        Object.freeze({ objectId: object.id, ...resource }),
      ),
    );
    return Object.freeze({
      evidence: "generation",
      ...(entry.node.authorId === undefined
        ? {}
        : { authorId: entry.node.authorId }),
      ...(box === undefined ? {} : { box }),
      ...(observedText === undefined ? {} : { text: observedText }),
      ...(textMetrics === undefined ? {} : { textMetrics }),
      objects: Object.freeze(objects),
      ...(resources.length === 0
        ? {}
        : { resources: Object.freeze(resources) }),
    });
  };

  return Object.freeze({
    get(id: string) {
      return oneMatch(findAuthorIds(id), "id", id);
    },
    getAll(id: string) {
      return findAuthorIds(id);
    },
    children: childLocators,
    descendants,
    observe,
  });
}

function slideQuery(
  observation: SlideObservation,
  content: readonly QueryContentNode[],
  resources: readonly (ResourceObservation & { readonly objectId: number })[],
): SlideQuery {
  const objectsById = new Map<number, ObjectObservation>();
  const addObject = (object: ObjectObservation) => {
    if (objectsById.has(object.id)) return;
    objectsById.set(object.id, object);
    for (const member of object.members ?? []) addObject(member);
  };
  for (const object of observation.objects) addObject(object);
  const layouts = new Map(
    (observation.layouts ?? []).map((layout) => [layout.id, layout] as const),
  );
  const resourcesByObjectId = new Map<number, ResourceObservation[]>();
  for (const { objectId, ...resource } of resources) {
    const items = resourcesByObjectId.get(objectId) ?? [];
    items.push(Object.freeze({ ...resource }));
    resourcesByObjectId.set(objectId, items);
  }
  const root: QueryNode = Object.freeze({
    children: Object.freeze(content.map((node) => queryNode(node))),
  });
  const locator = makeLocator(
    { node: root },
    layouts,
    objectsById,
    resourcesByObjectId,
  );
  return Object.freeze({
    ...locator,
    id: observation.id,
    ...(observation.authorId === undefined
      ? {}
      : { authorId: observation.authorId }),
    index: observation.index,
    observation,
  });
}

/** Bind content references retained during the compilation phase to the same report without copying or rebuilding the deck tree. */
export function createBuildQuery(
  slides: readonly BuildQuerySlideInput[],
): BuildQuery {
  const queries = Object.freeze(
    slides.map(({ observation, content, resources = [] }) =>
      slideQuery(observation, content, resources),
    ),
  );
  const byId = new Map(queries.map((slide) => [slide.id, slide] as const));
  return Object.freeze({
    evidence: "generation" as const,
    slides: queries,
    getSlide(id: string) {
      assertQueryName(id, "id");
      const slide = byId.get(id);
      if (slide === undefined) {
        throw new BuildQueryError(
          "missing-slide",
          `No slide found for author id ${JSON.stringify(id)}`,
        );
      }
      return slide;
    },
  });
}
