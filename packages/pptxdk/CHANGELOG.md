# @officedk/pptxdk

## 0.2.0

### Highlights

- **1D `Line` & connector capability**: Added `<Line>` / `line()` supporting flow dividers, arrow heads, straight/elbow/curved routing, and `from`/`to` object anchor bindings with full import, codegen, and preview parity.
- **Typography & font collections**: Added `fromFontCollection()` for `.ttc`/`.otc` collections, orthogonal `eastAsianFontFamily` and theme East Asian slots (`majorEastAsian` / `minorEastAsian`), proportional `lineHeight`, inline `<Sup>` / `<Sub>` and `baselineShift`, list indent parity, `1..1000` numeric font weights, static CFF shaping, and zero-copy sealed font bytes.
- **Motion ergonomics**: Added `appear`, exit effects (`disappear`, `fadeOut`, `flyOut`, `wipeOut`, `zoomOut`, `floatOut`, `spinOut`, `collapse`), `onStep` cross-step lifecycles, direct leaf `startAt` cues, and single-textbox `<P>` paragraph stepping.
- **Tables, charts, React interop & CLI**: Added cascading table/row/cell `align` and four-side borders, secondary value axis title rotation parity, `Mount` / `fromVanilla` / `expandFragment` React interop, `resolveFromSource` in `@officedk/pptxdk/node`, and structured CLI diagnostics.
- **Hardened PPTX import & static preview**: Enforced strong `importPptx` postconditions (out-of-spec placeholder unbinding, DrawingML bounds sanitization, `dsp:drawing` SmartArt recovery to `Group`, `<a:fmtScheme>` baseline completion), guarded preview SVG rasterization against degenerate geometry, and deduplicated repeated per-page raster images via `PaintContext.registerRasterResource`.

### Breaking Changes

- Unified layout and visual props onto single-track `style`, removed legacy author prop shims, closed the native OOXML `defineTheme()` contract (removing `ThemeScope`), and moved line arrow heads from `Shape.style.border` to `<Line>` / `line()`.
- Replaced legacy query/warning/observation surfaces with direct `Locator` APIs, `Resolved*` geometry snapshots (removing `*Observation` suffixes), and typed `BuildDiagnostic`, `PreviewError`, `PreviewPlaceholder`, and `CodegenDiagnostic` unions.

## 0.1.0

The first public release includes the SDK and the `pptxdk` command in one package. Author presentations with vanilla JavaScript or React, build editable PPTX files, preview slides, and generate source from existing presentations.
