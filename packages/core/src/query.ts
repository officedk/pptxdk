import type {
  Locator,
  LocatorResource,
  ObjectTransform,
  Rect,
  ResolutionDimensionsSummary,
  ResolvedLayout,
  ResolvedObject,
  ResolvedResource,
  ResolvedSlide,
  ScreenPoint,
  SlideLocator,
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
      readonly partKind?: string;
      readonly path?: string;
      readonly box?: Rect;
      readonly bounds?: Rect;
      readonly text?: string;
      readonly textMetrics?: TextMetrics;
    };

export type QueryContentNode = {
  readonly authorId?: string;
  readonly path?: string;
  readonly reference: QueryContentReference;
  readonly children: readonly QueryContentNode[];
};

export type BuildQuerySlideInput = {
  readonly observation: ResolvedSlide;
  readonly dimensions?: ResolutionDimensionsSummary;
  readonly content: readonly QueryContentNode[];
  readonly resources?: readonly (ResolvedResource & {
    readonly objectId: number;
  })[];
};

type QueryNode = {
  readonly authorId?: string;
  readonly path?: string;
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
  objectsById: ReadonlyMap<number, ResolvedObject>,
  seen: Set<number>,
  objects: ResolvedObject[],
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
  objectsById: ReadonlyMap<number, ResolvedObject>,
  seen: Set<number>,
  texts: string[],
): void {
  const addMemberObjects = (object: ResolvedObject) => {
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
    ...(node.path === undefined ? {} : { path: node.path }),
    reference: node.reference,
    children: Object.freeze(children),
  });
}

type LocatorSnapshot = {
  readonly id?: string;
  readonly kind: string;
  readonly path: string;
  readonly box?: Rect;
  readonly bounds?: Rect;
  readonly corners?: readonly [
    ScreenPoint,
    ScreenPoint,
    ScreenPoint,
    ScreenPoint,
  ];
  readonly transform?: ObjectTransform;
  readonly text?: string;
  readonly textMetrics?: TextMetrics;
  readonly gridArea?: Rect;
  readonly objects: readonly ResolvedObject[];
  readonly resources?: readonly LocatorResource[];
};

function computeLocatorSnapshot(
  node: QueryNode,
  layouts: ReadonlyMap<number, ResolvedLayout>,
  objectsById: ReadonlyMap<number, ResolvedObject>,
  resourcesByObjectId: ReadonlyMap<number, readonly ResolvedResource[]>,
  defaultKind: string,
  defaultPath: string,
): LocatorSnapshot {
  const objects: ResolvedObject[] = [];
  collectObjects(node, objectsById, new Set(), objects);

  const layout =
    node.reference?.kind === "layout"
      ? layouts.get(node.reference.id)
      : undefined;
  const object =
    node.reference?.kind === "object"
      ? objectsById.get(node.reference.id)
      : undefined;
  const partRef = node.reference?.kind === "part" ? node.reference : undefined;

  const textParts: string[] = [];
  if (partRef === undefined) {
    collectTexts(node, objectsById, new Set(), textParts);
  }
  const selectedText =
    textParts.length === 0 ? undefined : textParts.join("\n");

  const box =
    partRef !== undefined ? partRef.box : (layout?.box ?? object?.box);
  const bounds =
    partRef !== undefined
      ? (partRef.bounds ?? partRef.box)
      : (layout?.box ?? object?.bounds ?? object?.box);
  const corners = object?.corners;
  const transform = object?.transform;
  const gridArea = layout?.gridArea ?? object?.gridArea;
  const text = partRef !== undefined ? partRef.text : selectedText;
  const textMetrics =
    partRef !== undefined ? partRef.textMetrics : object?.textMetrics;
  const image = object?.image;
  const kind =
    partRef !== undefined
      ? (partRef.partKind ?? "part")
      : layout !== undefined
        ? "layout"
        : (object?.kind ?? defaultKind);
  const path =
    partRef?.path ?? layout?.path ?? object?.path ?? node.path ?? defaultPath;
  const resources: LocatorResource[] = objects.flatMap((item) =>
    (resourcesByObjectId.get(item.id) ?? []).map((resource) =>
      Object.freeze({ objectId: item.id, ...resource }),
    ),
  );

  return {
    ...(node.authorId === undefined ? {} : { id: node.authorId }),
    kind,
    path,
    ...(box === undefined ? {} : { box }),
    ...(bounds === undefined ? {} : { bounds }),
    ...(corners === undefined ? {} : { corners }),
    ...(transform === undefined ? {} : { transform }),
    ...(text === undefined ? {} : { text }),
    ...(textMetrics === undefined ? {} : { textMetrics }),
    ...(image === undefined ? {} : { image }),
    ...(gridArea === undefined ? {} : { gridArea }),
    objects: Object.freeze(objects),
    ...(resources.length === 0 ? {} : { resources: Object.freeze(resources) }),
  };
}

function makeLocator(
  entry: LocatorEntry,
  layouts: ReadonlyMap<number, ResolvedLayout>,
  objectsById: ReadonlyMap<number, ResolvedObject>,
  resourcesByObjectId: ReadonlyMap<number, readonly ResolvedResource[]>,
  defaultKind = "locator",
  defaultPath = "",
): Locator {
  const snapshot = computeLocatorSnapshot(
    entry.node,
    layouts,
    objectsById,
    resourcesByObjectId,
    defaultKind,
    defaultPath,
  );
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
  return Object.freeze({
    ...snapshot,
    get(id: string) {
      return oneMatch(findAuthorIds(id), "id", id);
    },
    getAll(id: string) {
      return findAuthorIds(id);
    },
    children: childLocators,
    descendants,
  });
}

const resolvedSlideByLocator = new WeakMap<SlideLocator, ResolvedSlide>();

/** Resolves a `SlideLocator` back to its backing `ResolvedSlide` without exposing `.observation`. */
export function resolveSlideFromLocator(
  slide: SlideLocator | ResolvedSlide,
): ResolvedSlide | undefined {
  if (typeof slide !== "object" || slide === null) return undefined;
  const registered = resolvedSlideByLocator.get(slide as SlideLocator);
  if (registered !== undefined) return registered;
  if ("get" in slide && typeof slide.get === "function") return undefined;
  return slide;
}

function slideQuery(
  observation: ResolvedSlide,
  dimensions: ResolutionDimensionsSummary | undefined,
  content: readonly QueryContentNode[],
  resources: readonly (ResolvedResource & { readonly objectId: number })[],
): SlideLocator {
  const objectsById = new Map<number, ResolvedObject>();
  const addObject = (object: ResolvedObject) => {
    if (objectsById.has(object.id)) return;
    objectsById.set(object.id, object);
    for (const member of object.members ?? []) addObject(member);
  };
  for (const object of observation.objects) addObject(object);
  const layouts = new Map(
    (observation.layouts ?? []).map((layout) => [layout.id, layout] as const),
  );
  const resourcesByObjectId = new Map<number, ResolvedResource[]>();
  for (const { objectId, ...resource } of resources) {
    const items = resourcesByObjectId.get(objectId) ?? [];
    items.push(Object.freeze({ ...resource }));
    resourcesByObjectId.set(objectId, items);
  }
  const root: QueryNode = Object.freeze({
    path: observation.path,
    children: Object.freeze(content.map((node) => queryNode(node))),
  });
  const locator = makeLocator(
    { node: root },
    layouts,
    objectsById,
    resourcesByObjectId,
    "slide",
    observation.path,
  );
  const slideLocator: SlideLocator = Object.freeze({
    ...locator,
    id: observation.id,
    kind: "slide",
    path: observation.path,
    index: observation.index,
    dimensions: dimensions ?? Object.freeze({ width: 0, height: 0 }),
    ...(observation.footer === undefined ? {} : { footer: observation.footer }),
  });
  resolvedSlideByLocator.set(slideLocator, observation);
  return slideLocator;
}

export type SlideLocatorIndex = {
  readonly slides: readonly SlideLocator[];
  slide(id: string): SlideLocator;
};

/** Bind content references retained during the compilation phase to the same report without copying or rebuilding the deck tree. */
export function createBuildQuery(
  slides: readonly BuildQuerySlideInput[],
): SlideLocatorIndex {
  const queries = Object.freeze(
    slides.map(({ observation, dimensions, content, resources = [] }) =>
      slideQuery(observation, dimensions, content, resources),
    ),
  );
  const byId = new Map(queries.map((slide) => [slide.id, slide] as const));
  return Object.freeze({
    slides: queries,
    slide(id: string) {
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
