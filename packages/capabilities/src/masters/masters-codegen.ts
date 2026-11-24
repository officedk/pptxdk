import {
  defineSlideCodegen,
  expr,
  type CodegenContext,
  type CodegenExpr,
} from "@pptxdk/codegen";
import {
  footerPlaceholdersSchema,
  headerFooterSchema,
  slideFooterValueSchema,
  slideNumberValueSchema,
} from "@pptxdk/core/capability";

function serializeValidatedValue(
  label: string,
  validate: (value: unknown) => {
    readonly success: boolean;
    readonly error?: {
      readonly issues: readonly { readonly message: string }[];
    };
  },
  value: unknown,
  ctx: CodegenContext,
): CodegenExpr {
  const parsed = validate(value);
  if (!parsed.success) {
    return ctx.fail(
      "unexpressible-value",
      `Invalid ${label} declaration: ${parsed.error?.issues[0]?.message ?? "invalid value"}`,
    );
  }
  return expr(ctx.emitValueExpr(value));
}

/** Source codegen contribution for `SlideMaster` and `SlideLayout` `headerFooter` metadata. */
export const mastersCodegen = defineSlideCodegen({
  slideProps: {
    headerFooter: (value, ctx) =>
      serializeValidatedValue(
        "headerFooter",
        (candidate) => headerFooterSchema.safeParse(candidate),
        value,
        ctx,
      ),
  },
});

/** Source codegen contribution for `footerPlaceholders` and slide `footer` / `dateTime` / `slideNumber`. */
export const footerPlaceholderCodegen = defineSlideCodegen({
  slideProps: {
    footerPlaceholders: (value, ctx) =>
      serializeValidatedValue(
        "footerPlaceholders",
        (candidate) => footerPlaceholdersSchema.safeParse(candidate),
        value,
        ctx,
      ),
    footer: (value, ctx) =>
      serializeValidatedValue(
        "footer",
        (candidate) => slideFooterValueSchema.safeParse(candidate),
        value,
        ctx,
      ),
    dateTime: (value, ctx) =>
      serializeValidatedValue(
        "dateTime",
        (candidate) => slideFooterValueSchema.safeParse(candidate),
        value,
        ctx,
      ),
    slideNumber: (value, ctx) =>
      serializeValidatedValue(
        "slideNumber",
        (candidate) => slideNumberValueSchema.safeParse(candidate),
        value,
        ctx,
      ),
  },
});
