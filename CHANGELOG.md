# Changelog

## Unreleased

- Added `Bookmark`, `Hyperlink` and `PageReference` for links and page references within a document [#50](https://github.com/PumasAI/WriteDocx.jl/pull/50).
- Added `width`, `layout` and `grid` to `Table` and `width` to `TableCell`, along with the `Percent` unit [#50](https://github.com/PumasAI/WriteDocx.jl/pull/50).
- Added tab stops via `ParagraphProperties(tabs = [TabStop(...)])` [#50](https://github.com/PumasAI/WriteDocx.jl/pull/50).
- Added the `line` keyword to `Spacing` for setting line height, including `AtLeast` [#50](https://github.com/PumasAI/WriteDocx.jl/pull/50).
- Added `underline` and `strike` to `RunProperties` and `shading` to `TableCellProperties` [#51](https://github.com/PumasAI/WriteDocx.jl/pull/51).

## v1.2.0 - 2025-09-04

- Added `columns` and `margins` to `SectionProperties` for multi-column layouts and page margins [#31](https://github.com/PumasAI/WriteDocx.jl/pull/31).

## v1.1.0 - 2025-02-28

- Added `Image` for passing objects showable as svg or png (for example Makie figures) directly to `InlineDrawing`, rendered when the document is saved [#29](https://github.com/PumasAI/WriteDocx.jl/pull/29).

## v1.0.0 - 2024-03-18

- Initial public release.
