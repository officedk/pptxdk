import {
  defineSlideCodegen,
  expr,
  type CodegenContext,
  type CodegenExpr,
} from "@pptxdk/codegen";

import {
  advanceAfterSchema,
  transitionSchema,
  type Transition,
} from "./transition-api.ts";

function serializeTransitionDeclaration(
  transition: Transition,
  ctx: CodegenContext,
): CodegenExpr {
  const transitionsVar = ctx.useImport("transitions");
  const optionEntries: [string, unknown][] = [];

  if (transition.duration !== undefined) {
    optionEntries.push(["duration", transition.duration]);
  }

  switch (transition.kind) {
    case "fade":
      break;
    case "strips":
      optionEntries.push(["direction", transition.direction]);
      break;
    case "push":
      optionEntries.push(["from", transition.from]);
      break;
    case "flip":
      optionEntries.push(["direction", transition.direction]);
      optionEntries.push([
        "fallback",
        serializeTransitionDeclaration(transition.fallback, ctx),
      ]);
      break;
    case "morph":
      optionEntries.push(["by", transition.by]);
      optionEntries.push([
        "fallback",
        serializeTransitionDeclaration(transition.fallback, ctx),
      ]);
      break;
    default: {
      const exhaustive: never = transition;
      return exhaustive;
    }
  }

  const args =
    optionEntries.length === 0 ? [] : [ctx.formatProps(optionEntries)];
  return expr(
    ctx.formatVanillaCall(`${transitionsVar}.${transition.kind}`, args),
  );
}

/**
 * Source codegen contribution for `transition` and `advanceAfter` slide/layout/master properties.
 */
export const transitionCodegen = defineSlideCodegen({
  slideProps: {
    transition: (value, ctx) => {
      const parsed = transitionSchema.safeParse(value);
      if (!parsed.success) {
        return ctx.fail(
          "unexpressible-value",
          parsed.error.issues[0]?.message ??
            "Transition declaration is invalid",
        );
      }
      return serializeTransitionDeclaration(parsed.data, ctx);
    },
    advanceAfter: (value, ctx) => {
      const parsed = advanceAfterSchema.safeParse(value);
      if (!parsed.success) {
        return ctx.fail(
          "unexpressible-value",
          parsed.error.issues[0]?.message ??
            "Slide.advanceAfter must be a finite millisecond value between 0 and 4294967295",
        );
      }
      return expr(ctx.emitValueExpr(parsed.data));
    },
  },
});
