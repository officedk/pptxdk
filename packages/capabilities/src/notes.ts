import {
  CapabilityError,
  themeXml,
  xmlEscape,
  type CapabilityInput,
  type SlideCapabilityModule,
  type TextStyle,
} from "@react-pptx/core/capability";
import type { ThemeContext } from "@react-pptx/core/capability";

import { fontPresentationContributions } from "./font-embedding.ts";
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
  parseTextContent,
  writeTextBodyFromParagraphs,
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

function parseNotes(
  input: CapabilityInput,
  pageTheme: ThemeContext,
  slideNumber: number,
  path: string,
): {
  readonly paragraphs: readonly Paragraph[];
  readonly style: ResolvedTextStyle;
  readonly theme: ThemeContext;
} {
  const unknown = Object.keys(input.props).find(
    (key) => key !== "id" && key !== "style",
  );
  if (unknown !== undefined) {
    throw new CapabilityError(
      "unknown-parameter",
      `Notes contains unsupported parameter ${unknown}`,
      path,
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
 * owned by this module, while this module parses the shared P/Span content once, writes the notes
 * part and body placeholder, and contributes the notes master, relationships, and font resources.
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
      compileDeclaration(input) {
        return [{ capability: "notes", value: input.props }];
      },
    },
  ],
  compileSlide(_input: CapabilityInput, context) {
    const contributions = context.contributions.filter(
      (contribution) => contribution.capability === "notes",
    );
    if (contributions.length === 0) {
      return { xml: "", extensions: [] };
    }
    if (contributions.length > 1) {
      throw new CapabilityError(
        "duplicate-notes",
        "A slide can declare at most one Notes",
        contributions[1]!.path,
      );
    }
    const contribution = contributions[0]!;
    const declaration = contribution.declaration;
    if (declaration === undefined) {
      throw new CapabilityError(
        "invalid-notes",
        "Notes declaration is missing its author content",
        contribution.path,
      );
    }
    let parsed: ReturnType<typeof parseNotes>;
    try {
      parsed = parseNotes(
        declaration,
        context.theme,
        context.slideNumber,
        contribution.path,
      );
    } catch (error) {
      if (error instanceof CapabilityError) {
        throw new CapabilityError(
          error.code,
          error.message,
          error.path ?? contribution.path,
        );
      }
      throw error;
    }
    const body = writeTextBodyFromParagraphs(
      parsed.paragraphs,
      parsed.style,
      parsed.theme,
      "top",
      undefined,
      "text-box",
      context.options.embedFonts,
      context.createTextFieldId,
    );
    // The notes master needs its own theme part: PowerPoint rejects a notes master that shares the
    // slide master's theme. Register the theme first so the notes master relationship can resolve
    // to this distinct OPC part, then serialize the same author theme values into it.
    context.registerSharedPart("notes-theme", {
      part: {
        pathTemplate: themePathTemplate,
        contentType: themeContentType,
      },
      content: () => themeXml(context.theme.theme),
    });
    context.registerSharedPart("notes-master", {
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
    context.registerPackagePart({
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
      extensions: [],
      observation: notesObservation(body.text, parsed.paragraphs),
      resourceClaims: body.resourceClaims,
      presentationContributions: fontPresentationContributions(
        body.fontUses,
        context.options.embedFonts,
      ),
      warnings: body.warnings,
    };
  },
});
