# First-Party Reference Fixtures (Masters & Brand Inheritance)

| File                                     | Origin                                                                                                                          | Bytes | SHA-256                                                            |
| ---------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------- | ----- | ------------------------------------------------------------------ |
| `brand-master-inheritance-override.pptx` | Created with Microsoft PowerPoint x64 `16.0.17932.21140` on Windows 11 `10.0.26200` (using universal `Arial` font declarations) | 21313 | `de3376969353a2322ccce48182393f63eea9a44f5a231565985bcb5d3d336536` |

## `brand-master-inheritance-override.pptx`

### Authoring Steps

1. Created a 16:9 presentation (`720 x 405 pt` / `960 x 540 px` at 96 dpi) in Microsoft PowerPoint `16.0.17932.21140` with `Arial` major/minor theme fonts and `#1D4ED8` accent1.
2. Configured `Brand Master` with solid `#F8FAFC` background and a non-placeholder shared textbox `BrandBadge` (`"ACME BRAND"`).
3. Retained two preserved custom layouts on `Brand Master`:
   - `Brand Hero` (`type="title"`, initially configured with `28pt` title and `16pt` subtitle placeholders).
   - `Brand Idle Content` (`type="obj"`, `26pt` title and `16pt` content placeholders, intentionally unused by initial slides).
4. Added Slide 1 (`Inherited Brand Slide`, layout `Brand Hero`) with title `"Inherited Brand Title"` without setting a local title font size, plus footer `"Confidential Brand Deck"` and slide number.
5. Added Slide 2 (`Explicit Override Slide`, layout `Brand Hero`) with title `"Locked Brand Title"` and explicitly set title font size to `32pt`, plus footer `"Confidential Brand Deck"` and slide number.
6. Raised `Brand Hero` title placeholder font size from `28pt` to `32pt` and saved, so both slides start at effective `32pt` while Slide 1 omits local `sz` and Slide 2 retains explicit `sz="3200"`.

### Independent Expectations & Verified Native Behavior

- **Initial state**: 1 master (`Brand Master`), 2 layouts (`Brand Hero` + unused `Brand Idle Content`), master background `#F8FAFC`, master shared shape `BrandBadge` (absent from ordinary slide shapes), footer `"Confidential Brand Deck"`, and effective `32pt` title size on both Slide 1 and Slide 2.
- **Layout edit + idle layout reuse**: Changing `Brand Hero` title size `32pt -> 22pt` updates Slide 1 effective title size to `22pt` while Slide 2 retains `32pt`; creating Slide 3 from `Brand Idle Content` binds cleanly to the previously unused layout and retains populated placeholder text after SaveAs and reopen.
- **Native app parity**: Verified in Microsoft PowerPoint `16.0.17932.21140` (x64) and WPS Presentation `12.2.0.23196` (x86); both produced identical inheritance, override preservation, and idle-layout results, so only this single minimal PowerPoint fixture is retained.
