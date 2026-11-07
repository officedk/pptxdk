import type {
  BuildQuery,
  LayoutObservation,
  Locator,
  LocatorResourceObservation,
  LocatorObservation,
  ObjectObservation,
  ResourceObservation,
  SlideObservation,
  SlideQuery,
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
      readonly kind: "table-part";
      readonly tableId: number;
      readonly partIndex: number;
    };

export type QueryContentNode = {
  readonly testId?: string;
  readonly part?: string;
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
  readonly testId?: string;
  readonly part?: string;
  readonly reference?: QueryContentReference;
  readonly children: readonly QueryNode[];
};

type LocatorEntry = {
  readonly node: QueryNode;
  readonly partOwner?: QueryNode;
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

function collectDescendantEntries(
  node: QueryNode,
  partOwner?: QueryNode,
): LocatorEntry[] {
  return node.children.flatMap((child) => {
    const owner = child.testId === undefined ? partOwner : child;
    return [
      { node: child, ...(owner === undefined ? {} : { partOwner: owner }) },
      ...collectDescendantEntries(child, owner),
    ];
  });
}

function queryNode(node: QueryContentNode): QueryNode {
  const children = node.children.map((child) => queryNode(child));
  return Object.freeze({
    ...(node.testId === undefined ? {} : { testId: node.testId }),
    ...(node.part === undefined ? {} : { part: node.part }),
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
        makeLocator(
          {
            node: child,
            ...(child.testId === undefined
              ? entry.partOwner === undefined
                ? {}
                : { partOwner: entry.partOwner }
              : { partOwner: child }),
          },
          layouts,
          objectsById,
          resourcesByObjectId,
        ),
      ),
    );
  const descendants = (): readonly Locator[] =>
    Object.freeze(
      collectDescendantEntries(entry.node, entry.partOwner).map((descendant) =>
        makeLocator(descendant, layouts, objectsById, resourcesByObjectId),
      ),
    );
  const findTestIds = (testId: string): readonly Locator[] => {
    assertQueryName(testId, "testId");
    return Object.freeze(
      collectDescendantEntries(entry.node, entry.partOwner)
        .filter(({ node }) => node.testId === testId)
        .map(({ node }) =>
          makeLocator(
            { node, partOwner: node },
            layouts,
            objectsById,
            resourcesByObjectId,
          ),
        ),
    );
  };
  const findParts = (name: string): readonly Locator[] => {
    assertQueryName(name, "part");
    const owner = entry.partOwner;
    const matches: LocatorEntry[] = [];
    const visit = (node: QueryNode, activeOwner?: QueryNode) => {
      const currentOwner = node.testId === undefined ? activeOwner : node;
      if (
        node.part === name &&
        (owner === undefined || currentOwner === owner)
      ) {
        matches.push({
          node,
          ...(currentOwner === undefined ? {} : { partOwner: currentOwner }),
        });
      }
      for (const child of node.children) visit(child, currentOwner);
    };
    visit(entry.node, owner);
    return Object.freeze(
      matches.map((match) =>
        makeLocator(match, layouts, objectsById, resourcesByObjectId),
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
    const tablePart =
      entry.node.reference?.kind === "table-part"
        ? objectsById.get(entry.node.reference.tableId)?.table?.parts?.[
            entry.node.reference.partIndex
          ]
        : undefined;
    const text = objects.flatMap((object) => {
      if (object.text !== undefined) return [object.text];
      if (object.table !== undefined) {
        return [object.table.cells.map((row) => row.join("\t")).join("\n")];
      }
      return [];
    });
    const selectedText = text.length === 0 ? undefined : text.join("\n");
    const box = tablePart?.box ?? layoutBox ?? objectBox;
    const observedText = tablePart?.text ?? selectedText;
    const resources: LocatorResourceObservation[] = objects.flatMap((object) =>
      (resourcesByObjectId.get(object.id) ?? []).map((resource) =>
        Object.freeze({ objectId: object.id, ...resource }),
      ),
    );
    return Object.freeze({
      evidence: "generation",
      ...(entry.node.testId === undefined ? {} : { testId: entry.node.testId }),
      ...(entry.node.part === undefined ? {} : { part: entry.node.part }),
      ...(box === undefined ? {} : { box }),
      ...(observedText === undefined ? {} : { text: observedText }),
      objects: Object.freeze(objects),
      ...(resources.length === 0
        ? {}
        : { resources: Object.freeze(resources) }),
    });
  };

  return Object.freeze({
    getByTestId(testId: string) {
      return oneMatch(findTestIds(testId), "testId", testId);
    },
    getAllByTestId(testId: string) {
      return findTestIds(testId);
    },
    part(name: string) {
      return oneMatch(findParts(name), "part", name);
    },
    parts(name: string) {
      return findParts(name);
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
    slide(slideId: string) {
      assertQueryName(slideId, "slideId");
      const slide = byId.get(slideId);
      if (slide === undefined) {
        throw new BuildQueryError(
          "missing-slide",
          `No slide found for slideId ${JSON.stringify(slideId)}`,
        );
      }
      return slide;
    },
  });
}
