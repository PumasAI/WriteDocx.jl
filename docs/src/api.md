# API

!!! note
    
    Positional arguments are only considered part of the API where they are explicitly mentioned in the docstrings.
    Otherwise, the API is built on keyword arguments so that missing options and properties can be added in the future without breaking existing code.

## Types

```@autodocs
Modules = [WriteDocx]
Order   = [:type]
Filter = t -> !(t <: WriteDocx.Length)
```

## `Length`s

```@autodocs
Modules = [WriteDocx]
Order   = [:type]
Filter = t -> t <: WriteDocx.Length
```

## Enums

```@docs
WriteDocx.BorderStyle
WriteDocx.Justification
WriteDocx.UnderlinePattern
WriteDocx.ShadingPattern
WriteDocx.VerticalAlign
WriteDocx.VerticalAlignment
```
