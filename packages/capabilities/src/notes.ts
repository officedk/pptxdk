import {
  CapabilityError,
  themeXml,
  xmlEscape,
  type CapabilityInput,
  type ResolvedSlide,
  type SlideCapabilityModule,
  type SlideEmitContext,
  type SlideEmitOutput,
  type TextStyle,
  type ThemeContext,
} from "@react-pptx/core/capability";

import {
  notesMasterContentType,
  notesMasterPathTemplate,
  notesMasterRelationshipType,
  notesMasterXml,
  notesSlideContentType,
  notesSlidePathTemplate,
  notesSlideRelationshipType,
  notesSlideXml,
  slideRelationshipType,
  themeContentType,
  themePathTemplate,
  themeRelationshipType,
} from "./notes-xml.ts";
import {
  emitTextBody,
  parseTextContent,
  resolveTextBody,
  type Paragraph,
  type ResolvedTextStyle,
} from "./text-service.ts";

/**
 * Notes text defaults are independent of slide projection styles.
 *
 * Colors and font tokens still resolve through the owning page theme; the inherited projection
 * text defaults are replaced instead of being overridden property by property.
 */
const notesDefaultTextStyle: TextStyle = Object.freeze({
  fontFamily: "minor",
  fontSize: "12pt",
  fontWeight: "normal",
  fontStyle: "normal",
  color: "dk1",
  verticalAlign: "top",
  align: "left",
  spacingAfter: 0,
  letterSpacing: 0,
  textTransform: "none",
  textDecoration: "none",
  textShadow: "none",
  textWrap: "normal",
});

type BoundNotesEntry = {
  readonly paragraphs: readonly Paragraph[];
  readonly style: ResolvedTextStyle;
  readonly theme: ThemeContext;
};

// Module-private binding table: keeps parsed notes rich text encapsulated within notesCapability
// without storing raw declaration nodes or inspectable parsed state on core contributions.
const boundNotesEntries = new WeakMap<object, BoundNotesEntry>();

function parseNotes(
  input: CapabilityInput,
  pageTheme: ThemeContext,
  slideNumber: number,
): BoundNotesEntry {
  const unknown = Object.keys(input.props).find(
    (key) => key !== "id" && key !== "style",
  );
  if (unknown !== undefined) {
    throw new CapabilityError(
      "unknown-parameter",
      `Notes contains unsupported parameter ${unknown}`,
    );
  }
  const theme = pageTheme.withTextStyle(notesDefaultTextStyle);
  const parsed = parseTextContent(
    input.children,
    input.props["style"],
    theme,
    { slideNumber },
    "Notes.style",
  );
  return { ...parsed, theme };
}

function notesObservation(
  text: string,
  paragraphs: readonly Paragraph[],
): unknown {
  return {
    text,
    paragraphCount: paragraphs.length,
  };
}

/**
 * Page-affiliated rich-text notes capability.
 *
 * The visible slide tree never contains the notes body: core treats the declaration subtree as
 * owned by this module, while this module parses the shared P/Span content once into an opaque
 * per-build binding, resolves text metrics and resource claims during resolve, and emits the notes
 * part, body placeholder, notes master, relationships, and font resources during emit. Notes has
 * no canvas geometry or static slide preview layer.
 */
export const notesCapability: SlideCapabilityModule = Object.freeze({
  declaration: "notes",
  slideProp: "notes",
  slideProps: [],
  slot: "timing",
  always: true,
  transparentDeclarations: [
    {
      declaration: "rp:notes",
      children: "owned" as const,
      compileDeclaration(input, context) {
        if (context.theme === undefined || context.slideNumber === undefined) {
          throw new CapabilityError(
            "invalid-notes",
            "Notes declaration requires slide theme and number context",
          );
        }
        const entry = parseNotes(input, context.theme, context.slideNumber);
        const token = Object.freeze({});
        boundNotesEntries.set(token, entry);
        return [{ capability: "notes", value: token }];
      },
    },
  ],
  bindSlide(_input: CapabilityInput, parseContext) {
    const contributions = parseContext.contributions.filter(
      (contribution) => contribution.capability === "notes",
    );
    if (contributions.length === 0) {
      return {
        resolveSlide(): ResolvedSlide {
          return {
            extensions: [],
            emit(): SlideEmitOutput {
              return { xml: "" };
            },
          };
        },
      };
    }
    if (contributions.length > 1) {
      throw new CapabilityError(
        "duplicate-notes",
        "A slide can declare at most one Notes",
        contributions[1]!.path,
      );
    }
    const contribution = contributions[0]!;
    const entry =
      typeof contribution.value === "object" && contribution.value !== null
        ? boundNotesEntries.get(contribution.value)
        : undefined;
    if (entry === undefined) {
      throw new CapabilityError(
        "invalid-notes",
        "Notes declaration is missing its bound content",
        contribution.path,
      );
    }

    return {
      resolveSlide(): ResolvedSlide {
        const resolvedBody = resolveTextBody(
          entry.paragraphs,
          entry.style,
          entry.theme,
          "top",
          undefined,
          "text-box",
        );
        return {
          observation: notesObservation(resolvedBody.text, entry.paragraphs),
          extensions: [],
          resourceClaims: resolvedBody.resourceClaims,
          ...(resolvedBody.warnings.length === 0
            ? {}
            : { warnings: resolvedBody.warnings }),
          emit(emitContext: SlideEmitContext): SlideEmitOutput {
            const embedFonts = emitContext.options.embedFonts;
            const body = emitTextBody(resolvedBody, entry.theme, {
              createFieldId: emitContext.createTextFieldId,
              ...(embedFonts === undefined ? {} : { embedFonts }),
            });
            // The notes master needs its own theme part: PowerPoint rejects a notes master that
            // shares the slide master's theme. Register the theme first so the notes master
            // relationship can resolve to this distinct OPC part, then serialize the same author
            // theme values into it.
            emitContext.registerSharedPart("notes-theme", {
              part: {
                pathTemplate: themePathTemplate,
                contentType: themeContentType,
              },
              content: () => themeXml(emitContext.theme.theme),
            });
            emitContext.registerSharedPart("notes-master", {
              part: {
                pathTemplate: notesMasterPathTemplate,
                contentType: notesMasterContentType,
              },
              presentationRelationshipType: notesMasterRelationshipType,
              relationships: [
                {
                  relationshipType: themeRelationshipType,
                  target: { kind: "part", key: "notes-theme" },
                },
              ],
              presentationElement: {
                position: "before-slides",
                content: (relationshipId) =>
                  `<p:notesMasterIdLst><p:notesMasterId r:id="${xmlEscape(relationshipId)}"/></p:notesMasterIdLst>`,
              },
              content: () => notesMasterXml(),
            });
            emitContext.registerPackagePart({
              part: {
                pathTemplate: notesSlidePathTemplate,
                contentType: notesSlideContentType,
              },
              slideRelationshipType: notesSlideRelationshipType,
              relationships: [
                {
                  relationshipType: slideRelationshipType,
                  target: { kind: "slide" },
                },
                {
                  relationshipType: notesMasterRelationshipType,
                  target: { kind: "part", key: "notes-master" },
                },
              ],
              content: () => notesSlideXml(body.xml),
            });
            return {
              xml: "",
              ...(body.presentationContributions.length === 0
                ? {}
                : {
                    presentationContributions: body.presentationContributions,
                  }),
            };
          },
        };
      },
    };
  },
});
