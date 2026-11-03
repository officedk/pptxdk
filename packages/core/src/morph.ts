const morphKeyBrand: unique symbol = Symbol("MorphKey");
const morphKeyLabel: unique symbol = Symbol("MorphKeyLabel");

/** 由 MorphScope 创建、用于显式声明跨页对象连续性的句柄。 */
export type MorphKey = {
  readonly [morphKeyBrand]: true;
  readonly [morphKeyLabel]: string;
};

/**
 * 同一组件逻辑实例的连续性句柄工厂。
 *
 * 跨页复用时传递同一个 MorphScope；同页重复组件各自创建作用域。若在 React 组件中创建，
 * 请用 useMemo 或 useState 保持作用域稳定；普通函数组件边界不会自动建立作用域。
 */
export type MorphScope = {
  /** 在此作用域内按名称取得稳定句柄。 */
  handle(name: string): MorphKey;
};

/**
 * 创建独立的连续性作用域；句柄只在显式复用时才会跨页对应。
 *
 * 在 React 组件内创建时，应以 useMemo 或 useState 保持返回值稳定；跨页对应必须复用同一作用域。
 */
export function createMorphScope(): MorphScope {
  const handles = new Map<string, MorphKey>();
  return Object.freeze({
    handle(name: string): MorphKey {
      if (typeof name !== "string" || name.trim() === "") {
        throw new TypeError("MorphScope.handle(name) 需要非空名称");
      }
      let handle = handles.get(name);
      if (handle === undefined) {
        handle = Object.freeze({
          [morphKeyBrand]: true as const,
          [morphKeyLabel]: name,
        });
        handles.set(name, handle);
      }
      return handle;
    },
  });
}

export function isMorphKey(value: unknown): value is MorphKey {
  return (
    typeof value === "object" &&
    value !== null &&
    morphKeyBrand in value &&
    morphKeyLabel in value
  );
}

export function getMorphKeyLabel(value: MorphKey): string {
  return value[morphKeyLabel];
}
