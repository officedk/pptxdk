# First-Party Test Fixtures

This directory contains a byte-identical copy of `brand-master-inheritance-override.pptx` from the masters capability test fixtures so `pptxdk` package tests remain self-contained within package boundaries.

| Filename                                 | Bytes | SHA-256                                                            | Origin                                                                                                     |
| ---------------------------------------- | ----- | ------------------------------------------------------------------ | ---------------------------------------------------------------------------------------------------------- |
| `brand-master-inheritance-override.pptx` | 21313 | `de3376969353a2322ccce48182393f63eea9a44f5a231565985bcb5d3d336536` | Microsoft PowerPoint x64 `16.0.17932.21140` (Windows 11 `10.0.26200`, universal `Arial` font declarations) |

See the fixture expectations documented in `@pptxdk/capabilities` (`brand-master-inheritance-override.pptx`): 1 brand master (`#F8FAFC` background, `BrandBadge` shared text object, footer `"Confidential Brand Deck"`), 2 layouts (`Brand Hero` used at `32pt` and unused `Brand Idle Content`), Slide 1 inheriting `32pt` title size, and Slide 2 holding an equal explicit `32pt` override.
