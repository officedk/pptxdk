import type {
  EmitContext,
  EmitOutput,
  PaintContext,
  PaintOutput,
  SlideEmitContext,
  SlideEmitOutput,
} from "#/types.ts";

/**
 * Directional contract for a native PresentationML canvas mechanism.
 * Every canvas mechanism must declare production emit, paint, and import directions.
 */
export type CanvasMechanismDefinition<
  Key extends string = string,
  Model = unknown,
  ImportSpec = unknown,
  EmitCtx = EmitContext,
  EmitOut = EmitOutput,
  PaintCtx = PaintContext,
  PaintOut = PaintOutput,
> = {
  readonly key: Key;
  readonly role: "canvas";
  emit(model: Model, context: EmitCtx): EmitOut;
  paint(model: Model, context: PaintCtx): PaintOut;
  readonly import: ImportSpec;
};

/**
 * Directional contract for a native PresentationML slide/show mechanism without static paint output.
 * Every show mechanism must declare production emit and import directions and forbid paint.
 */
export type ShowMechanismDefinition<
  Key extends string = string,
  Model = unknown,
  ImportSpec = unknown,
  EmitCtx = SlideEmitContext,
  EmitOut = SlideEmitOutput,
> = {
  readonly key: Key;
  readonly role: "show";
  emit(model: Model, context: EmitCtx): EmitOut;
  readonly import: ImportSpec;
  readonly paint?: never;
};

/** Union of all production native mechanism definitions declared by capability modules. */
export type NativeMechanismDefinition<
  Key extends string = string,
  Model = any,
  ImportSpec = unknown,
> =
  | CanvasMechanismDefinition<Key, Model, ImportSpec, any, any, any, any>
  | ShowMechanismDefinition<Key, Model, ImportSpec, any, any>;
