module WriteDocx

import EzXML as E
using EnumX: @enumx
using MacroTools: @capture
using ZipFile: ZipFile
using OrderedCollections: OrderedDict

"""
    Length

The supertype for all length metrics that WriteDocx can handle.
Each `Length` can be converted to any other `Length` and can therefore
be passed to any struct that stores a specific length type.
"""
abstract type Length end

"""
    Point(value::Float64) <: Length

A length of one typographic point, or 1/72 of an inch.
For convenience, the constant `pt` is provided for `Point(1)`.
"""
struct Point <: Length
    value::Float64
end
const pt = Point(1)

"""
    HalfPoint(value::Float64) <: Length

A length of a half typographic point, or 1/144 of an inch.
For convenience, the constant `halfpt` is provided for `HalfPoint(1)`.
"""
struct HalfPoint <: Length
    value::Float64
end
const halfpt = HalfPoint(1)

"""
    EighthPoint(value::Float64) <: Length

A length of one eight of a typographic point, or 1/576 of an inch.
For convenience, the constant `eighthpt` is provided for `EighthPoint(1)`.
"""
struct EighthPoint <: Length
    value::Float64
end
const eighthpt = EighthPoint(1)

"""
    Twip(value::Float64) <: Length

A length of one twip, or twentieth of a point, or 1/1440 of an inch.
For convenience, the constant `twip` is provided for `Twip(1)`.
"""
struct Twip <: Length
    value::Float64
end
const twip = Twip(1)

"""
    EMU(value::Float64) <: Length

A length of one English metric unit, or 1/914400 of an inch.
For convenience, the constant `emu` is provided for `EMU(1)`.
"""
struct EMU <: Length
    value::Float64
end
const emu = EMU(1)

"""
    Inch(value::Float64) <: Length

A length of one inch.
For convenience, the constant `inch` is provided for `Inch(1)`.
"""
struct Inch <: Length
    value::Float64
end
const inch = Inch(1)

"""
    Centimeter(value::Float64) <: Length

A length of one centimeter, or 1/2.54 of an inch.
For convenience, the constants `cm` and `mm` are provided for
`Centimeter(1)` and `Centimeter(0.1)`, respectively.
"""
struct Centimeter <: Length
    value::Float64
end
const mm = Centimeter(0.1)
const cm = Centimeter(1)

inchfactor(::Type{Inch}) = 1.0
inchfactor(::Type{Point}) = 72.0
inchfactor(::Type{HalfPoint}) = 144.0
inchfactor(::Type{EighthPoint}) = 576.0
inchfactor(::Type{Twip}) = 20.0 * 72.0
inchfactor(::Type{EMU}) = 914400.0
inchfactor(::Type{Centimeter}) = 2.54

# explicit definitions due to ambiguities with (T::Type{<:Length})(...)
Point(l::Length) = convert(Point, l)
HalfPoint(l::Length) = convert(HalfPoint, l)
EighthPoint(l::Length) = convert(EighthPoint, l)
Twip(l::Length) = convert(Twip, l)
EMU(l::Length) = convert(EMU, l)
Inch(l::Length) = convert(Inch, l)
Centimeter(l::Length) = convert(Centimeter, l)

Base.convert(::Type{L}, l::L) where L <: Length = l
Base.convert(::Type{L1}, l::L2) where {L1 <: Length, L2 <: Length} = L1(inchfactor(L1) * l.value / inchfactor(L2))
Base.isless(l1::L1, l2::L2) where {L1 <: Length, L2 <: Length} = convert(L2, l1).value < l2.value

Base.:(*)(x::Real, l::L) where L <: Length = L(l.value * x)
Base.:(*)(l::L, x::Real) where L <: Length = x * l
Base.:(/)(l::L, l2::Length) where L <: Length = l.value / convert(L, l2).value
Base.:(/)(l::L, r::Real) where L <: Length = L(l.value / r)
# it doesn't matter if return types here are order dependent, this is just for convenience anyway
Base.:(+)(l::L, l2::Length) where L <: Length = L(l.value + convert(L, l2).value)
Base.:(-)(l::L, l2::Length) where L <: Length = L(l.value - convert(L, l2).value)

const Maybe = Union{Nothing, <:Any}

macro partialkw(expr::Expr)
    if expr.head === :struct
        name::Symbol = let
            maybe_parametric = if @capture expr.args[2] x_ <: y_
                x
            else
                expr.args[2]
            end
            if @capture maybe_parametric x_{y_}
                x
            else
                maybe_parametric
            end
        end
        fields = expr.args[3].args
        kwargs_boundary = findfirst(e -> (@capture e @kwargs), fields)
        kwargs_boundary = something(kwargs_boundary, length(fields) + 1)
        positional = fields[1:kwargs_boundary-1]
        kw = fields[kwargs_boundary+1:end]
        stripped_kw = map(kw) do field
            if @capture field x_ = y_
                x
            else
                field
            end
        end

        new_struct = Expr(
            :struct,
            expr.args[1],
            expr.args[2],
            Expr(:block, positional..., stripped_kw...)
        )

        kw_transformed = map(filter(x -> !(x isa LineNumberNode), kw)) do kw
            if @capture kw x_ = y_
                if !@capture x sym_::typ_
                    sym = x
                end
                Expr(:kw, sym::Symbol, y)
            else
                if @capture kw sym_::typ_
                    sym::Symbol
                else
                    kw::Symbol
                end
            end
        end

        positional_syms = map(filter(x -> !(x isa LineNumberNode), positional)) do arg
            if !@capture arg sym_::typ_
                sym = arg
            end
            sym::Symbol
        end

        kw_syms = map(kw_transformed) do kw
            kw isa Symbol ? kw : kw.args[1]::Symbol
        end

        func = Expr(
            :function,
            Expr(
                :call,
                name,
                Expr(
                    :parameters,
                    kw_transformed...
                ),
                positional_syms...
            ),
            quote
                $(name)($(positional_syms...), $(kw_syms...))
            end
        )

        result = quote
            $new_struct
            $func
        end
        esc(result)
    else
        error("Expected struct.")
    end
end

struct Text
    text::String
end

is_inline_element(_) = false
is_inline_element(::Type{Text}) = true

@enumx BreakType column page text_wrapping

struct Break
    type::BreakType.T
end

is_inline_element(::Type{Break}) = true

Break() = Break(BreakType.text_wrapping)

struct Size
    size::HalfPoint
end

"""
    VerticalAlignment

An enum that can be either `baseline`, `subscript` or `superscript`.
"""
@enumx VerticalAlignment baseline subscript superscript

"""
    UnderlinePattern

An enum that can be either `dash`, `dash_dot_dot_heavy`, `dash_dot_heavy`, `dashed_heavy`, `dash_long`, `dash_long_heavy`, `dot_dash`, `dot_dot_dash`, `dotted`, `dotted_heavy`, `double`, `none`, `single`, `thick`, `wave`, `wavy_double`, `wavy_heavy` or `words`.
"""
@enumx UnderlinePattern dash dash_dot_dot_heavy dash_dot_heavy dashed_heavy dash_long dash_long_heavy dot_dash dot_dot_dash dotted dotted_heavy double none single thick wave wavy_double wavy_heavy words
"""
    BorderStyle

An enum that can be either `single`, `dash_dot_stroked`, `dashed`, `dash_small_gap`, `dot_dash`, `dot_dot_dash`, `dotted`, `double`, `double_wave`, `inset`, `nil`, `none`, `outset`, `thick`, `thick_thin_large_gap`, `thick_thin_medium_gap`, `thick_thin_small_gap`, `thin_thick_large_gap`, `thin_thick_medium_gap`, `thin_thick_small_gap`, `thin_thick_thin_large_gap`, `thin_thick_thin_medium_gap`, `thin_thick_thin_small_gap`, `three_d_emboss`, `three_d_engrave`, `triple` or `wave`
"""
@enumx BorderStyle single dash_dot_stroked dashed dash_small_gap dot_dash dot_dot_dash dotted double double_wave inset nil none outset thick thick_thin_large_gap thick_thin_medium_gap thick_thin_small_gap thin_thick_large_gap thin_thick_medium_gap thin_thick_small_gap thin_thick_thin_large_gap thin_thick_thin_medium_gap thin_thick_thin_small_gap three_d_emboss three_d_engrave triple wave

snake_to_camel(string::String) = replace(string, r"_[a-z]" => uppercase ∘ last)

"""
    Fonts(; [ascii::String, high_ansi::String, complex::String, east_asia::String])
    Fonts(font; kwargs...)

Specifies fonts to use for four different Unicode character ranges.
The convenience constructor with one positional argument changes the font for
`ascii` and `high_ansi`, which should usually be the same.
"""
Base.@kwdef struct Fonts
    ascii::Maybe{String} = nothing
    high_ansi::Maybe{String} = nothing
    complex::Maybe{String} = nothing
    east_asia::Maybe{String} = nothing
end

Fonts(font; kwargs...) = Fonts(; ascii = font, high_ansi = font, kwargs...)

"""
    HexColor(s::String)

A color in hexadecimal RGB format, for example "FF0000" for red or
"333333" for a dark gray.
"""
struct HexColor
    hex::String
    function HexColor(s::String)
        if match(r"^[a-fA-F0-9]{6}$", s) === nothing
            error("Invalid color string $(repr(s)).")
        end
        new(s)
    end
end

struct Automatic end

const automatic = Automatic()

"""
    AutomaticDefault{T}

Signals that either a value of type `T` is accepted or
`automatic`, for which the viewer application chooses appropriate
behavior.
"""
struct AutomaticDefault{T}
    value::Union{Automatic, T}
end

struct Bold end
struct Italic end

Base.convert(::Type{AutomaticDefault{X}}, x::X) where X = AutomaticDefault{X}(x)
Base.convert(::Type{AutomaticDefault{X}}, ::Automatic) where X = AutomaticDefault{X}(automatic)

struct Underline
    color::AutomaticDefault{HexColor}
    pattern::UnderlinePattern.T
end

struct Color
    color::AutomaticDefault{HexColor}
end

struct ParagraphStyle
    name::String
end

struct RunStyle
    name::String
end

struct GridSpan
    n::Int
end

struct VerticalMerge
    restart::Bool
end

"""
    TableCellBorder(; kwargs...)

Holds properties for one border of a table cell and is used by [`TableCellBorders`](@ref).
All properties are optional.

## Keyword arguments

| Keyword | Description |
| :-- | :-- |
| `color::`[`AutomaticDefault`](@ref)`{`[`HexColor`]@ref`}` | The color of the border. |
| `shadow::Bool` | Applies a shadow effect if `true`. |
| `space::`[`Point`](@ref) | The spacing between border and content. |
| `size::`[`EighthPoint`](@ref) | The thickness of the border line. |
| `style::`[`BorderStyle`](@ref)`.T` | The line style of the border. |
"""
Base.@kwdef struct TableCellBorder
    color::Maybe{AutomaticDefault{HexColor}} = nothing
    shadow::Maybe{Bool} = nothing
    space::Maybe{Point} = nothing
    size::Maybe{EighthPoint} = nothing
    style::Maybe{BorderStyle.T}
end

"""
    ParagraphBorder(; kwargs...)

Holds properties for one border of a table cell and is used by [`ParagraphBorders`](@ref).
All properties are optional.

## Keyword arguments

| Keyword | Description |
| :-- | :-- |
| `color::`[`AutomaticDefault`](@ref)`{`[`HexColor`]@ref`}` | The color of the border. |
| `shadow::Bool` | Applies a shadow effect if `true`. |
| `space::`[`Point`](@ref) | The spacing between border and content. |
| `size::`[`EighthPoint`](@ref) | The thickness of the border line. |
| `style::`[`BorderStyle`](@ref)`.T` | The line style of the border. |
"""
Base.@kwdef struct ParagraphBorder
    color::Maybe{AutomaticDefault{HexColor}} = nothing
    shadow::Maybe{Bool} = nothing
    space::Maybe{Point} = nothing
    size::Maybe{EighthPoint} = nothing
    style::Maybe{BorderStyle.T}
end

"""
    TableCellBorders(; kwargs...)

Holds properties for the borders of a [`TableCell`](@ref) and is used by [`TableCellProperties`](@ref).
All properties are optional.

## Keyword arguments

| Keyword | Description |
| :-- | :-- |
| `top::`[`TableCellBorder`](@ref) | The properties of the top border. |
| `bottom::`[`TableCellBorder`](@ref) | The properties of the bottom border. |
| `start::`[`TableCellBorder`](@ref) | The properties of the left border in left-to-right text. |
| `stop::`[`TableCellBorder`](@ref) | The properties of the right border in left-to-right text. |
| `inside_h::`[`TableCellBorder`](@ref) | The properties of the horizontal border that lies between adjacent cells. |
| `inside_v::`[`TableCellBorder`](@ref) | The properties of the vertical border that lies between adjacent cells. |
| `tl2br::`[`TableCellBorder`](@ref) | The properties of the diagonal border going from the top left to the bottom right corner. |
| `tr2bl::`[`TableCellBorder`](@ref) | The properties of the diagonal border going from the top right to the bottom left corner. |
"""
Base.@kwdef struct TableCellBorders
    top::Maybe{TableCellBorder} = nothing
    start::Maybe{TableCellBorder} = nothing
    bottom::Maybe{TableCellBorder} = nothing
    stop::Maybe{TableCellBorder} = nothing
    inside_h::Maybe{TableCellBorder} = nothing
    inside_v::Maybe{TableCellBorder} = nothing
    tl2br::Maybe{TableCellBorder} = nothing
    tr2bl::Maybe{TableCellBorder} = nothing
end

"""
    ParagraphBorders(; kwargs...)

Holds properties for the borders of a [`Paragraph`](@ref) and is used by [`ParagraphProperties`](@ref).
All properties are optional.

## Keyword arguments

| Keyword | Description |
| :-- | :-- |
| `top::`[`ParagraphBorder`](@ref) | The properties of the top border. |
| `bottom::`[`ParagraphBorder`](@ref) | The properties of the bottom border. |
| `left::`[`ParagraphBorder`](@ref) | The properties of the left border. |
| `right::`[`ParagraphBorder`](@ref) | The properties of the right border. |
| `between::`[`ParagraphBorder`](@ref) | The properties of horizontal border that lies between adjacent paragraphs. |
"""
Base.@kwdef struct ParagraphBorders
    top::Maybe{ParagraphBorder} = nothing
    left::Maybe{ParagraphBorder} = nothing
    bottom::Maybe{ParagraphBorder} = nothing
    right::Maybe{ParagraphBorder} = nothing
    between::Maybe{ParagraphBorder} = nothing
end

"""
    TableCellMargins(; kwargs...)

Holds properties for the margins of a [`TableCell`](@ref) and is used by [`TableCellProperties`](@ref).
All properties are optional.

## Keyword arguments

| Keyword | Description |
| :-- | :-- |
| `top::`[`Twip`](@ref) | The top margin. |
| `bottom::`[`Twip`](@ref) | The bottom margin. |
| `start::`[`Twip`](@ref) | The left margin in left-to-right text. |
| `stop::`[`Twip`](@ref) | The right margin in left-to-right text. |
"""
Base.@kwdef struct TableCellMargins
    top::Maybe{Twip} = nothing
    start::Maybe{Twip} = nothing
    bottom::Maybe{Twip} = nothing
    stop::Maybe{Twip} = nothing
end

"""
    TableCellMargins(; kwargs...)

Holds properties for the default margins of all [`TableCell`](@ref)s
in a [`Table`](@ref) and is used by [`TableProperties`](@ref).
All properties are optional.

## Keyword arguments

| Keyword | Description |
| :-- | :-- |
| `top::`[`Twip`](@ref) | The top margin. |
| `bottom::`[`Twip`](@ref) | The bottom margin. |
| `start::`[`Twip`](@ref) | The left margin in left-to-right text. |
| `stop::`[`Twip`](@ref) | The right margin in left-to-right text. |
"""
Base.@kwdef struct TableLevelCellMargins
    top::Maybe{Twip} = nothing
    start::Maybe{Twip} = nothing
    bottom::Maybe{Twip} = nothing
    stop::Maybe{Twip} = nothing
end

struct Path
    path::String
end

"""
    Image(m::MIME, object)

Represents an image of MIME type `m` that can be written to an appropriate file
when writing out a docx document. The `object` needs to have a `show` method
for `m` defined for the default behavior to work.
"""
struct Image{MIME,T}
    m::MIME
    object::T
end

"""
    Image(path::String)

Create an `Image` pointing to the file at `path`. The MIME type is determined by file extension.
"""
Image(path::String) = Image(Path(path))
Image(i::Image) = i

function Image(path::Path)
    if endswith(path.path, r"\.svg"i)
        Image(MIME"image/svg+xml"(), path)
    elseif endswith(path.path, r"\.png"i)
        Image(MIME"image/png"(), path)
    else
        throw(ArgumentError("Unknown image file extension, only .svg and .png are allowed. Path was \"$(path.path)\""))
    end
end

# avoid writing file again which is already given as a path
function get_image_file(i::Image{M,Path}, tempdir) where M
    return i.object.path
end

file_extension(::Type{MIME"image/svg+xml"}) = "svg"
file_extension(::Type{MIME"image/png"}) = "png"

function get_image_file(i::Image{M}, tempdir) where M
    ext = file_extension(M)
    filepath = joinpath(tempdir, "temp.$(ext)")
    open(filepath, "w") do io
        Base.show(io::IO, M(), i.object)
    end
    return filepath
end

"""
    InlineDrawing{T}(; image::T, width::EMU, height::EMU)

Create an `InlineDrawing` object which, as the name implies, can be placed inline with
text inside [`Run`](@ref)s.

WriteDocx supports different types `T` for the `image` argument.
If `T` is a `String`, `image` is treated as the file path to an existing .png or .svg image.
You can pass an `Image` object which can hold a reference to an object that can be written
to a file with the desired MIME type at render time.
You can also use [`SVGWithPNGFallback`](@ref) to place .svg images with better fallback behavior.

Width and height of the placed image are set via `width` and `height`, note that you have to
determine these values yourself for any image you place, a correct aspect ratio will not
be determined automatically.
"""
struct InlineDrawing{T}
    image::T
    width::EMU
    height::EMU
end

function InlineDrawing(; image, width, height)
    InlineDrawing(image, width, height)
end

InlineDrawing(image::T, width, height) where T = InlineDrawing{T}(image, convert(EMU, width), convert(EMU, height))

# backwards-compatibility, interpret String as path to an image
function InlineDrawing(path::AbstractString, width, height)
    InlineDrawing(Image(path), width, height)
end

# fix ambiguity
function InlineDrawing(path::AbstractString, width::EMU, height::EMU)
    InlineDrawing(Image(path), width, height)
end

is_inline_element(::Type{<:InlineDrawing}) = true

"""
    SVGWithPNGFallback(; svg, png)

Create a `SVGWithPNGFallback` for the svg `svg` and the fallback `png`.
If `svg` or `png` are `AbstractString`s, they will be treated as paths to image files.
Otherwise, they should be `Image`s with the appropriate MIME types. Use `SVGImage` and
`PNGImage` as shortcuts to create these.

Word Online and other services like Slack preview don't work when a simple svg file is added via
[`InlineDrawing`](@ref)`{String}`.
`SVGWithPNGFallback` supplies a fallback png file which will be used for display in those situations.
Note that it is your responsibility to check whether the png file is an accurate
replacement for the svg.
"""
struct SVGWithPNGFallback{T1,T2}
    svg::Image{MIME"image/svg+xml",T1} # path to SVG
    png::Image{MIME"image/png",T2} # path to PNG
end

function SVGWithPNGFallback(; svg, png)
    SVGWithPNGFallback(Image(svg), Image(png))
end

@partialkw struct Style{T}
    id::String
    properties::T
    @kwargs
    name::String
    default::Bool = false
    based_on::Maybe{String} = nothing
    ui_priority::Maybe{Int} = nothing
    next::Maybe{String} = nothing
    link::Maybe{String} = nothing
    custom_style::Maybe{Bool} = nothing
    primary_style::Bool = true # important style (without this it doesn't show up in the Word user interface)
end

"""
    RunProperties(; kwargs...)

Holds properties for a [`Run`](@ref).
All properties are optional.

## Keyword arguments

| Keyword | Description |
| :-- | :-- |
| `style::String` | The name of the style applied to this `Run`. |
| `color::`[`AutomaticDefault`](@ref)`{HexColor}` | The color of the text. |
| `size::`[`HalfPoint`](@ref) | The font size. |
| `valign::`[`VerticalAlignment`](@ref)`.T` | Whether text is shown with baseline, superscript or subscript style. |
| `fonts::`[`Fonts`](@ref) | The font settings for this text. |
| `bold::Bool` | Whether text should be bold. Note that this works like a toggle when nested, turning boldness off again the second time it's `true`. |
| `italic::Bool` | Whether text should be italic. Note that this works like a toggle when nested, turning italic style off again the second time it's `true`. |
"""
Base.@kwdef struct RunProperties
    style::Maybe{String} = nothing
    color::Maybe{AutomaticDefault{HexColor}} = nothing
    size::Maybe{HalfPoint} = nothing
    valign::Maybe{VerticalAlignment.T} = nothing
    fonts::Maybe{Fonts} = nothing
    bold::Maybe{Bool} = nothing
    italic::Maybe{Bool} = nothing
end

"""
    Justification

An enum that can be either `start`, `stop`, `center`, `both` or `distribute`.
"""
@enumx Justification start stop center both distribute

Base.@kwdef struct Spacing
    before::Maybe{Twip} = nothing
    after::Maybe{Twip} = nothing
end

"""
ShadingPattern

An enum that can be either `clear`, `diag_cross`, `diag_stripe`, `horz_cross`, `horz_stripe`, `nil`, `thin_diag_cross`, or `solid`.
"""
@enumx ShadingPattern clear diag_cross diag_stripe horz_cross horz_stripe nil thin_diag_cross solid

Base.@kwdef struct Shading
    pattern::ShadingPattern.T = ShadingPattern.clear
    fill::AutomaticDefault{HexColor} = automatic
    color::AutomaticDefault{HexColor} = automatic
end

"""
    ParagraphProperties(; kwargs...)

Holds properties for a [`Paragraph`](@ref).
All properties are optional.

## Keyword arguments

| Keyword | Description |
| :-- | :-- |
| `style::String` | The name of the style applied to this `Paragraph`. |
| `justification::`[`Justification`](@ref)`.T` | The justification of the paragraph. |
"""
Base.@kwdef struct ParagraphProperties
    style::Maybe{String} = nothing
    justification::Maybe{Justification.T} = nothing
    keep_next::Maybe{Bool} = nothing
    keep_lines::Maybe{Bool} = nothing
    page_break_before::Maybe{Bool} = nothing
    run_properties::Maybe{RunProperties} = nothing
    spacing::Maybe{Spacing} = nothing
    borders::Maybe{ParagraphBorders} = nothing
    shading::Maybe{Shading} = nothing
end

"""
    DocDefaults(; kwargs...)

Holds the default run and paragraph properties for a [`Document`](@ref).
All properties are optional.

## Keyword arguments

| Keyword | Description |
| :-- | :-- |
| `run::`[`RunProperties`](@ref) | The default properties for every [`Run`](@ref). |
| `paragraph::`[`ParagraphProperties`](@ref) | The default properties for every [`Paragraph`](@ref). |
"""
Base.@kwdef struct DocDefaults
    run::Maybe{RunProperties} = nothing
    paragraph::Maybe{ParagraphProperties} = nothing
end

"""
    Styles(styles::Vector{Style}, doc_defaults::DocDefaults)
    Styles(styles; kwargs...)

Holds style information for a [`Document`](@ref).
The second convenience constructor forwards all keyword arguments to [`DocDefaults`](@ref).
"""
struct Styles
    styles::Vector{Style}
    doc_defaults::DocDefaults

    function Styles(styles, doc_defaults)
        style_ids = Set([style.id for style in styles])

        styles = Vector{Style}(styles)

        for style in default_styles()
            if style.id ∉ style_ids
                push!(styles, style)
            end
        end

        new(styles, doc_defaults)
    end
end

Styles(styles; kwargs...) = Styles(styles, DocDefaults(; kwargs...))

"""
    Run(children::AbstractVector, properties::RunProperties)
    Run(children::AbstractVector; kwargs...)

Create a `Run` with `children` who all have to satisfy `is_inline_element`.

The second convenience constructor forwards all keyword arguments to the
[`RunProperties`](@ref) constructor.
"""
struct Run
    children::Vector{Any}
    properties::RunProperties
    function Run(children::AbstractVector, properties::RunProperties)
        children = validate_elements(is_inline_element, children, :Run)
        new(children, properties)
    end
end

Run(children::AbstractVector; kwargs...) = Run(children, RunProperties(; kwargs...))

is_run_element(::Type{Run}) = true

"""
    TableProperties(; kwargs...)

Holds properties for a [`Table`](@ref).
All properties are optional.

## Keyword arguments

| Keyword | Description |
| :-- | :-- |
| `margins::TableLevelCellMargins` | Margins for all cells in the table. |
| `spacing::Twip` | The space between adjacent cells and the edges of the table. |
| `justification::`[`Justification`](@ref)`.T` | The justification of the table. |
"""
Base.@kwdef struct TableProperties
    margins::Maybe{TableLevelCellMargins} = nothing
    spacing::Maybe{Twip} = nothing
    justification::Maybe{Justification.T} = nothing
end

@enumx HeightRule at_least exact auto 

struct TableRowHeight
    height::Twip
    hrule::Maybe{HeightRule.T}
end

TableRowHeight(height; hrule = nothing) = TableRowHeight(height, hrule)

"""
    TableRowProperties(; kwargs...)

Holds properties for a [`TableRow`](@ref).
All properties are optional.

## Keyword arguments

| Keyword | Description |
| :-- | :-- |
| `header::Bool` | Whether this row should be part of the header section which is repeated after every page break. |
| `height::TableRowHeight` | The height of the table row. |
"""
Base.@kwdef struct TableRowProperties
    header::Maybe{Bool} = nothing
    height::Maybe{TableRowHeight} = nothing
end

"""
    VerticalAlign

An enum that can be `bottom`, `center` or `top`.
"""
@enumx VerticalAlign bottom center top

"""
    TableCellProperties(; kwargs...)

Holds properties for a [`TableCell`](@ref).
All properties are optional.

## Keyword arguments

| Keyword | Description |
| :-- | :-- |
| `borders::TableCellBorders` | The border style of the cell. |
| `vertical_merge::Bool` | Should be set to `true` if this cell should be merged with the one above it. |
| `gridspan::Int` | The number of cells this cell should span in horizontal direction. |
| `margins::TableCellMargins` | The margins of the cell. |
| `valign::`[`VerticalAlign`](@ref)`.T` | The vertical alignment of the content in the cell. |
| `hide_mark::Bool` | If `true`, hides the editor mark so that the table cell can fully collapse if it's empty. |
"""
Base.@kwdef struct TableCellProperties
    borders::Maybe{TableCellBorders} = nothing
    vertical_merge::Maybe{Bool} = nothing
    gridspan::Maybe{Int} = nothing
    margins::Maybe{TableCellMargins} = nothing
    valign::Maybe{VerticalAlign.T} = nothing
    hide_mark::Maybe{Bool} = nothing # cells can only collapse vertically if this is true, for example to use a cell only for its border
end

struct SimpleField
    instruction::String
    cached_value::Maybe{Run}
    dirty::Maybe{Bool}
end

SimpleField(instruction; dirty = true) = SimpleField(instruction, nothing, dirty)

is_run_element(::Type{SimpleField}) = true

function page_number_field()
    SimpleField("PAGE"; dirty = nothing) # it seems that page numbers are always recalculated, but if dirty is set they cause a warning
end

"""
    ComplexField(instruction::String; dirty = true)

Creates a complex field with a specific `instruction` that has an effect in the
viewer application. If `dirty === true`, the field will be reevaluated when opening the
docx file.

The `ComplexField` element must be paired with a following [`ComplexFieldEnd`](@ref). For some
purposes, other elements may appear between the two.
"""
struct ComplexField
    instruction::String
    dirty::Maybe{Bool}
end

ComplexField(instruction; dirty = true) = ComplexField(instruction, dirty)

is_run_element(::Type{ComplexField}) = true

"""
    ComplexFieldEnd()

Every [`ComplexField`](@ref) element must be paired with this element.
"""
struct ComplexFieldEnd end

is_run_element(::Type{ComplexFieldEnd}) = true

is_block_element(x) = false

is_run_element(x) = false

"""
    Paragraph(children::Vector{Any}, properties::ParagraphProperties)
    Paragraph(children::AbstractVector; kwargs...)

A paragraph can contain children that satisfy `is_run_element`.
The second convenience constructor forwards all keyword arguments to [`ParagraphProperties`](@ref).
"""
struct Paragraph
    children::Vector{Any}
    properties::ParagraphProperties
    function Paragraph(children::AbstractVector, properties::ParagraphProperties)
        children = validate_elements(is_run_element, children, :Paragraph)
        new(children, properties)
    end
end

is_block_element(::Type{Paragraph}) = true

Paragraph(children; kwargs...) = Paragraph(children, ParagraphProperties(; kwargs...))

"""
    TableCell(children::Vector{Any}, properties::TableCellProperties)
    TableCell(children::AbstractVector; kwargs...)

One cell of a [`Table`](@ref) which can hold elements that satisfy `is_block_element`.
The second convenience constructor forwards all keyword arguments to [`TableCellProperties`](@ref).
"""
struct TableCell
    children::Vector{Any}
    properties::TableCellProperties
    function TableCell(children::AbstractVector, properties::TableCellProperties)
        children = validate_elements(is_block_element, children, :TableCell)
        new(children, properties)
    end
end

TableCell(children::AbstractVector; kwargs...) = TableCell(children, TableCellProperties(; kwargs...))

"""
    TableRow(cells::Vector{TableCell}, properties::TableRowProperties)
    TableRow(cells; kwargs...)

One row of a [`Table`](@ref) which can hold a vector of [`TableCell`](@ref)s.
The second convenience constructor forwards all keyword arguments to [`TableRowProperties`](@ref).
"""
struct TableRow
    cells::Vector{TableCell}
    properties::TableRowProperties
end

TableRow(cells; kwargs...) = TableRow(cells, TableRowProperties(; kwargs...))

"""
    Table(rows::Vector{TableRow}, properties::TableProperties)
    Table(rows; kwargs...)

A table which can hold a vector of [`TableRow`](@ref)s.
The second convenience constructor forwards all keyword arguments to [`TableProperties`](@ref).
"""
struct Table
    rows::Vector{TableRow}
    properties::TableProperties
end

is_block_element(::Type{Table}) = true

Table(rows; kwargs...) = Table(rows, TableProperties(; kwargs...))

@enumx PageOrientation landscape portrait

struct PageSize
    width::Twip
    height::Twip
    orientation::PageOrientation.T
end

"""
    PageSize(width, height)

The size of a page. If `width > height`, the page is set to `PageOrientation.landscape`.
"""
function PageSize(width, height)
    orientation = height >= width ? PageOrientation.portrait : PageOrientation.landscape
    PageSize(width, height, orientation)
end

@enumx PageVerticalAlign both bottom center top

function validate_elements(predicate, elements::AbstractVector, target::Symbol)
    elements = convert(Vector{Any}, elements)
    for element in elements
        if !predicate(typeof(element))
            error("Element of type $(typeof(element)) does not satisfy `$predicate` and cannot be placed in a `$target`.")
        end
    end
    return elements
end

"""
    Header(children::AbstractVector)

Contains elements for use in a [`Section`](@ref)'s header section.
Each element should satisfy `is_block_element`.
"""
struct Header
    children::Vector{Any}

    function Header(children::AbstractVector)
        children = validate_elements(is_block_element, children, :Header)
        new(children)
    end
end

"""
    Headers(; default::Header, [first::Header, even::Header])

Holds information about the [`Header`](@ref)s of a [`Section`](@ref).
A `default` `Header` must always be specified. If `first` is set, the
first page of the section gets this separate header. If `even` is
set, every even-numbered page of the section gets this separate header, making
`default` effectively mean `odd`.
"""
Base.@kwdef struct Headers
    default::Header
    first::Maybe{Header} = nothing
    even::Maybe{Header} = nothing
end

"""
    Footer(children::AbstractVector)

Contains elements for use in a [`Section`](@ref)'s footer section.
Each element should satisfy `is_block_element`.
"""
struct Footer
    children::Vector{Any}
    function Footer(children::AbstractVector)
        children = validate_elements(is_block_element, children, :Footer)
        new(children)
    end
end

"""
    Footers(; default::Footer, [first::Footer, even::Footer])

Holds information about the [`Footer`](@ref)s of a [`Section`](@ref).
A `default` `Footer` must always be specified. If `first` is set, the
first page of the section gets this separate footer. If `even` is
set, every even-numbered page of the section gets this separate footer, making
`default` effectively mean `odd`.
"""
Base.@kwdef struct Footers
    default::Footer
    first::Maybe{Footer} = nothing
    even::Maybe{Footer} = nothing
end

"""
    Column(; [width::Twip, space::Twip])

Describes a single column. `width` sets the column width, and `space` sets the whitespace after the column (before the next column).

See also: [`Columns`](@ref)
"""
Base.@kwdef struct Column
    width::Maybe{Twip} = nothing
    space::Maybe{Twip} = nothing
end

"""
    Columns(; kwargs...)

Sets the columns for a [`Section`](@ref).

## Keyword arguments

| Keyword | Description |
| :-- | :-- |
| `equal::Bool = true` | Whether all columns are equal-width with the same space between every column |
| `num::Int` | The number of equal-width columns to layout. |
| `space::`[`Twip`](@ref) | The space between equal-width columns. |
| `sep::Bool = false` | Sets whether a vertical separator line is drawn between columns |
| `cols::Vector{`[`Column`](@ref)`}` | A vector of custom columns. `equal`, `num`, and `space` are ignored if `cols` is provided. |
"""
Base.@kwdef struct Columns
    num::Int = 1
    space::Maybe{Twip} = nothing
    sep::Bool = false
    cols::Vector{Column} = Column[]
    equal::Bool = isempty(cols)
end

"""
    PageMargins(; top, right, bottom, left, kwargs...)

Describes page margins in a [`Section`](@ref).

## Keyword arguments

| Keyword | Description |
| :-- | :-- |
| `top::`[`Twip`](@ref) | The top margin. |
| `right::`[`Twip`](@ref) | The right margin. |
| `bottom::`[`Twip`](@ref) | The bottom margin. |
| `left::`[`Twip`](@ref) | The left margin. |
| `header::`[`Twip`](@ref)`=Twip(0)` | The header margin. |
| `footer::`[`Twip`](@ref)`=Twip(0)` | The footer margin. |
| `gutter::`[`Twip`](@ref)`=Twip(0)` | The gutter margin. |
"""
Base.@kwdef struct PageMargins
    top::Twip
    right::Twip
    bottom::Twip
    left::Twip
    header::Twip = Twip(0)
    footer::Twip = Twip(0)
    gutter::Twip = Twip(0)
end

"""
    SectionProperties(; kwargs...)

Holds properties for a [`Section`](@ref).

## Keyword arguments

| Keyword | Description |
| :-- | :-- |
| `pagesize::PageSize` | The size of each page in the section. |
| `margins::PageMargins` | The margins for each page in the section |
| `valign::PageVerticalAlign.T` | The vertical alignment of content on each page of the section. |
| `headers::`[`Headers`](@ref) | Defines the header content shown at the top of each page of the section. |
| `footers::`[`Footers`](@ref) | Defines the footer content shown at the bottom of each page of the section. |
| `columns::`[`Columns`](@ref) | Configures the columns in the section |
"""
Base.@kwdef struct SectionProperties
    pagesize::Maybe{PageSize} = nothing
    margins::Maybe{PageMargins} = nothing
    valign::Maybe{PageVerticalAlign.T} = nothing
    headers::Maybe{Headers} = nothing
    footers::Maybe{Footers} = nothing
    columns::Maybe{Columns} = nothing
end

"""
    Section(children::AbstractVector, properties::SectionProperties)
    Section(children::AbstractVector; kwargs...)

A section of a document contains a vector of `children` which should satisfy `is_block_element`.
The docx format does not have a concept of individual pages, although `Section` might be
thought of as a group of related "page"s.

The content within a document's `Section`s is laid out into actual pages dynamically in the viewer application.
A `Section` has [`SectionProperties`](@ref) which then control how those pages are rendered.

The second convenience constructor forwards all keyword
arguments to the [`SectionProperties`](@ref) constructor.
"""
struct Section
    children::Vector{Any}
    properties::SectionProperties

    function Section(children::AbstractVector, properties::SectionProperties)
        children = validate_elements(is_block_element, children, :Section)
        new(children, properties)
    end
end

Section(children::AbstractVector; kwargs...) = Section(children, SectionProperties(; kwargs...))

"""
    Body(sections::Vector{Section})

The document body which contains the sections of the document.
"""
struct Body
    sections::Vector{Section}
end

struct Document
    body::Body
    styles::Styles
end

"""
    Document(body::Body; styles::Styles = Styles([]))

The root object containing all other elements that make up the document.
"""
Document(body::Body; styles::Styles = Styles([])) = Document(body, styles)

function save(path, document::Document)
    if !endswith(path, ".docx")
        throw(ArgumentError("Path does not end in .docx"))
    end

    check_styles(document)

    mktempdir() do dir
        rels = gather_rels(document, dir)
        resolved_rels = resolve_rels!(rels, dir, "")

        zipwriter = ZipFile.Writer(path)
    
        f = ZipFile.addfile(zipwriter, "[Content_Types].xml"; method = ZipFile.Deflate)
        print(f, content_types(rels))

        f = ZipFile.addfile(zipwriter, "_rels/.rels"; method = ZipFile.Deflate)
        print(f, global_rels_xml())
        
        f = ZipFile.addfile(zipwriter, "word/_rels/document.xml.rels"; method = ZipFile.Deflate)
        E.prettyprint(f, rels_xml(resolved_rels))

        for (rel, i) in rels
            if rel isa Union{Header,Footer}
                render_header_or_footer(rel, resolved_rels[i], zipwriter, dir)
            end
        end

        f = ZipFile.addfile(zipwriter, "word/styles.xml"; method = ZipFile.Deflate)
        E.prettyprint(f, styles_xml(document.styles))

        f = ZipFile.addfile(zipwriter, "word/document.xml"; method = ZipFile.Deflate)
        E.prettyprint(f, to_xml(document, rels))

        for (root, _, files) in walkdir(dir)
            for file in files
                relp = relpath(joinpath(root, file), dir)
                open(joinpath(root, file)) do io
                    f = ZipFile.addfile(zipwriter, relp; method = ZipFile.Deflate)
                    write(f, io)
                end
            end
        end

        close(zipwriter)
    end
end

function render_header_or_footer(x::Union{Header,Footer}, resolvedrel, zipwriter, zipdir)
    rels = gather_rels(x, zipdir)
    name = splitext(basename(resolvedrel.target))[1]
    prefix = name * "_"
    resolved_rels = resolve_rels!(rels, zipdir, prefix)

    if !isempty(resolved_rels)
        f = ZipFile.addfile(zipwriter, "word/_rels/$(resolvedrel.target).rels"; method = ZipFile.Deflate)
        E.prettyprint(f, rels_xml(resolved_rels))
    end

    xmlpath = joinpath(zipdir, "word", resolvedrel.target)

    doc = E.XMLDocument()

    node = to_xml(x, rels)
    E.setroot!(doc, node)
    mkpath(dirname(xmlpath))
    open(xmlpath, "w") do io
        E.prettyprint(io, doc)
    end
end

struct DuplicateStyleError <: Exception
    id::String
    type::Symbol
end

struct MissingStyleError <: Exception
    id::String
    type::Symbol
end

function check_styles(doc::Document)
    parastyles = Set{String}()
    charstyles = Set{String}()
    for style in doc.styles.styles
        if style isa Style{RunProperties}
            if style.id in charstyles
                throw(DuplicateStyleError(style.id, :character))
            else
                push!(charstyles, style.id)
            end
        elseif style isa Style{ParagraphProperties}
            if style.id in parastyles
                throw(DuplicateStyleError(style.id, :paragraph))
            else
                push!(parastyles, style.id)
            end
        else
            error("Unknown style type $(typeof(style))")
        end
    end

    function chk(x)
         _chk(x)
         children(x) === nothing || foreach(chk, children(x))
    end
    _chk(x) = nothing
    function _chk(r::Run)
        s = r.properties.style
        if s isa String && s ∉ charstyles
            throw(MissingStyleError(s, :character))
        end
    end
    function _chk(p::Paragraph)
        s = p.properties.style
        if s isa String && s ∉ parastyles
            throw(MissingStyleError(s, :paragraph))
        end
    end
    chk(doc.body)
    return
end

function linkedstyles(abbreviation, name, para_properties, run_properties)
    [
        Style(
            abbreviation,
            para_properties,
            name = name,
            link = abbreviation * "Char",
            based_on = "Normal",
        ),
        Style(
            abbreviation * "Char",
            run_properties,
            name = name,
            link = abbreviation,
            based_on = "Normal",
        )
    ]
end

function default_styles()
    vcat(
        Style(
            "Normal",
            ParagraphProperties(
                run_properties = RunProperties(
                    size = 10 * pt
                )
            ),
            name = "Normal",
            link = "NormalChar",
            default = true
        ),
        Style(
            "NormalChar",
            RunProperties(
                size = 10 * pt
            ),
            name = "Normal",
            link = "Normal",
            default = true
        ),
        linkedstyles(
            "Heading1", "Heading 1",
            ParagraphProperties(
                spacing = Spacing(
                    before = 12 * pt,
                    after = 10 * pt,
                ),
                run_properties = RunProperties(
                    size = 20 * pt
                )
            ),
            RunProperties()
        ),
        linkedstyles(
            "Heading2", "Heading 2",
            ParagraphProperties(
                spacing = Spacing(
                    before = 10 * pt,
                    after = 8 * pt,
                ),
                run_properties = RunProperties(
                    size = 16 * pt
                )
            ),
            RunProperties()
        ),
        linkedstyles(
            "Heading3", "Heading 3",
            ParagraphProperties(
                spacing = Spacing(
                    before = 8 * pt,
                    after = 6 * pt,
                ),
                run_properties = RunProperties(
                    size = 14 * pt
                )
            ),
            RunProperties()
        ),
        linkedstyles(
            "Heading4", "Heading 4",
            ParagraphProperties(
                spacing = Spacing(
                    before = 6 * pt,
                    after = 4 * pt,
                ),
                run_properties = RunProperties(
                    size = 12 * pt
                )
            ),
            RunProperties()
        ),
        linkedstyles(
            "Heading5", "Heading 5",
            ParagraphProperties(
                spacing = Spacing(
                    before = 6 * pt,
                    after = 4 * pt,
                ),
                run_properties = RunProperties(
                    size = 11 * pt,
                    italic = true,
                )
            ),
            RunProperties()
        ),
        linkedstyles(
            "Heading6", "Heading 6",
            ParagraphProperties(
                spacing = Spacing(
                    before = 6 * pt,
                    after = 4 * pt,
                ),
                run_properties = RunProperties(
                    size = 10 * pt,
                    italic = true,
                )
            ),
            RunProperties()
        ),
    )
end

function resolve_rels!(rels, dir, prefix)
    resolved = Vector{Relationship}(undef, length(rels))
    for (x, i) in rels
        resolved[i] = resolve_rel!(x, i, dir, prefix)
    end
    return resolved
end

struct Relationship
    type::String
    target::String
end

struct StyleRel end

function resolve_rel!(x::StyleRel, i, dir, prefix)
    return Relationship(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
        prefix * "styles.xml",
    )
end

function resolve_rel!(x::Image, i, dir, prefix)
    mktempdir() do tempdir
        imgpath = get_image_file(x, tempdir)
        mediapath = joinpath(dir, "word", "media")
        mkpath(mediapath)
        _, extension = splitext(imgpath)
        targetpath = joinpath(mediapath, "$(prefix)image_$i" * extension)
        cp(imgpath, targetpath)
        type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
        target = relpath(targetpath, joinpath(dir, "word"))
        return Relationship(type, target)
    end
end

function resolve_rel!(h::Union{Header, Footer}, i, dir, prefix)
    kind = h isa Header ? "header" : "footer"
    xmlpath = joinpath(dir, "word", "$kind$i.xml")
    type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/$kind"

    target = relpath(xmlpath, joinpath(dir, "word"))
    return Relationship(type, target)
end

function add_rel!(rels, x)
    if !haskey(rels, x)
        rels[x] = length(rels) + 1
    end
    return
end

function gather_rels(document::Document, zipdir)
    rels = OrderedDict{Any, Int}()
    add_rel!(rels, StyleRel())
    gather_rels!(rels, zipdir, document.body)
    return rels
end

function gather_rels(x::Union{Header,Footer}, zipdir)
    rels = OrderedDict{Any, Int}()
    for child in x.children
        gather_rels!(rels, zipdir, child)
    end
    return rels
end

function gather_rels!(rels, zipdir, x)
    _children = children(x)
    for child in _children
        gather_rels!(rels, zipdir, child)
    end
    return
end

gather_rels!(rels, zipdir, ::Nothing) = nothing

function gather_rels!(rels, zipdir, i::InlineDrawing{<:Image})
    add_rel!(rels, i.image)
    return
end

function gather_rels!(rels, zipdir, i::InlineDrawing{<:SVGWithPNGFallback})
    add_rel!(rels, i.image.png)
    add_rel!(rels, i.image.svg)
    return
end

function gather_rels!(rels, zipdir, s::Section)
    _children = children(s)
    for child in _children
        gather_rels!(rels, zipdir, child)
    end
    gather_rels!(rels, zipdir, s.properties.headers)
    gather_rels!(rels, zipdir, s.properties.footers)
    return
end

function gather_rels!(rels, zipdir, h::Union{Headers, Footers})
    add_rel!(rels, h.default)
    h.first === nothing || add_rel!(rels, h.first)
    h.even === nothing || add_rel!(rels, h.even)
    return
end

function rel_target(i::InlineDrawing{String})
    i.image
end

function content_types(rels)
    
    io = IOBuffer()

    print(io, """
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Types
        xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Default Extension="png" ContentType="image/png" />
        <Default Extension="svg" ContentType="image/svg+xml" />
        <Default Extension="xml" ContentType="application/xml" />
        <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" />
        <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml" />
        <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml" />
    """)

    for (rel, i) in rels
        if rel isa Header
            println(io, "    <Override PartName=\"/word/header$i.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml\" />")
        elseif rel isa Footer
            println(io, "    <Override PartName=\"/word/footer$i.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml\" />")
        end
    end

    println(io, "</Types>")

    return String(take!(io))
end

function rels_xml(resolved_rels)
    
    doc = E.XMLDocument()
    node = E.ElementNode("Relationships")
    node["xmlns"] = "http://schemas.openxmlformats.org/package/2006/relationships"
    E.setroot!(doc, node)
    for (i, x) in enumerate(resolved_rels)
        relnode = relationship_node(x, i)
        E.link!(node, relnode)
    end

    return doc
end

function global_rels_xml()
    """
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Relationships
        xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
    </Relationships>
    """
end

function styles_xml(styles::Styles)
    doc = E.XMLDocument()
    node = xml("w:styles")
    node["xmlns:w"] = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    E.setroot!(doc, node)

    E.link!(node, to_xml(styles.doc_defaults))

    for style in styles.styles
        E.link!(node, to_xml(style))
    end

    E.link!(node, xml("w:latentStyles"))

    return doc
end

function to_xml(d::DocDefaults)
    node = xml("w:docDefaults")
    rprdefault = xml("w:rPrDefault")
    E.link!(node, rprdefault)
    if d.run !== nothing
        E.link!(rprdefault, to_xml(d.run, nothing))
    end
    pprdefault = xml("w:pPrDefault")
    E.link!(node, pprdefault)
    if d.paragraph !== nothing
        E.link!(pprdefault, to_xml(d.paragraph, nothing))
    end
    return node
end

forwardslashpath(path) = replace(path, r"\\\\" => "/")

function relationship_node(rel::Relationship, i)
    node = E.ElementNode("Relationship")
    node["Id"] = "rId$i"
    node["Type"] = rel.type
    node["Target"] = forwardslashpath(rel.target)
    return node
end

# function document(table_xml)
#     """
#     <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
#     <w:document
#         xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
#         <w:body>
#             $table_xml
#         </w:body>
#     </w:document>
#     """
# end

to_xml(x) = to_xml(x, nothing)

function to_xml(document::Document, rels)
    doc = E.XMLDocument()
    node = xmlnode(document)
    node["xmlns:w"] = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    node["xmlns:r"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    node["xmlns:wp"] = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    E.setroot!(doc, node)
    E.link!(node, to_xml(document.body, rels))
    return doc
end

function to_xml(body::Body, rels)
    bodynode = xmlnode(body)
    n_sections = length(body.sections)
    for (i, section) in enumerate(body.sections)
        props = section.properties
        section_params_node = xml("w:sectPr")
        if props.pagesize !== nothing
            E.link!(section_params_node, to_xml(props.pagesize, rels))
        end
        if props.margins !== nothing
            E.link!(section_params_node, to_xml(props.margins, rels))
        end
        if props.valign !== nothing
            E.link!(section_params_node, to_xml(props.valign, rels))
        end
        if props.columns !== nothing
            E.link!(section_params_node, to_xml(props.columns, rels))
        end
        if props.headers !== nothing
            for type in (:default, :first, :even)
                x = getproperty(props.headers, type)
                x === nothing && continue
                rel_id = rels[x]
                E.link!(section_params_node, xml("w:headerReference", "r:id" => "rId$rel_id", "w:type" => string(type)))
            end
        end
        if props.footers !== nothing
            for type in (:default, :first, :even)
                x = getproperty(props.footers, type)
                x === nothing && continue
                rel_id = rels[x]
                E.link!(section_params_node, xml("w:footerReference", "r:id" => "rId$rel_id", "w:type" => string(type)))
            end
        end
        for child in section.children
            E.link!(bodynode, to_xml(child, rels))
        end
        if i < n_sections
            # For all but the last section, the section parameters go into an empty
            # paragraph trailing the remaining section content
            E.link!(bodynode, xml("w:p", [xml("w:pPr", [section_params_node])]))
        else
            # The last section parameter node needs to go into the body directly
            E.link!(bodynode, section_params_node)
        end
    end
    return bodynode
end

function to_xml(x, rels)
    node = xmlnode(x)
    props = properties(x)
    if props !== nothing
        linkall!(node, to_xml(props, rels))
    end
    _children = children(x)
    for child in _children
        linkall!(node, to_xml(child, rels))
    end
    for (attribute_key, attribute) in attributes(x)
        node[attribute_key] = xmlstring(attribute)
    end
    return node
end

properties(x) = hasfield(typeof(x), :properties) ? x.properties : nothing

to_xml(s::String, rels) = E.TextNode(s)

function to_xml(t::TableCellBorders, rels)
    node = xmlnode(t)

    t.bottom === nothing || E.link!(node, to_xml((t.bottom, :bottom), rels))
    t.top === nothing || E.link!(node, to_xml((t.top, :top), rels))
    t.stop === nothing || E.link!(node, to_xml((t.stop, :end), rels))
    t.start === nothing || E.link!(node, to_xml((t.start, :start), rels))
    t.tl2br === nothing || E.link!(node, to_xml((t.tl2br, :tl2br), rels))
    t.tr2bl === nothing || E.link!(node, to_xml((t.tr2bl, :tr2bl), rels))
    t.inside_h === nothing || E.link!(node, to_xml((t.inside_h, :insideH), rels))
    t.inside_v === nothing || E.link!(node, to_xml((t.inside_v, :insideV), rels))

    return node
end

function to_xml(t::ParagraphBorders, rels)
    node = xmlnode(t)

    t.left === nothing || E.link!(node, to_xml((t.left, :left), rels))
    t.top === nothing || E.link!(node, to_xml((t.top, :top), rels))
    t.right === nothing || E.link!(node, to_xml((t.right, :right), rels))
    t.bottom === nothing || E.link!(node, to_xml((t.bottom, :bottom), rels))
    t.between === nothing || E.link!(node, to_xml((t.between, :between), rels))

    return node
end

function xml(name, children::Vector, pairs::Pair{String,<:Any}...)
    node = E.ElementNode(name)
    for (key, value) in pairs
        node[key] = xmlstring(value)
    end
    for child in children
        linkall!(node, child)
    end
    return node
end

function linkall!(node, child::E.Node)
    E.link!(node, child)
    return
end
function linkall!(node, children::AbstractVector)
    for child in children
        E.link!(node, child)
    end
    return
end

xml(name, pairs::Pair{String,<:Any}...) = xml(name, [], pairs...)

to_xml(xml::E.Node, _) = xml

function get_ablip(i::InlineDrawing{<:Image{M}}, rels) where M
    index = rels[i.image]
    ablip = if M <: MIME"image/svg+xml"
        xml("a:blip", [
            xml("a:extLst", [
                # this is the empty scaffolding of a png thumbnail that doesn't actually need to be there in modern Word as it seems
                xml("a:ext", "uri" => "{28A0092B-C50C-407E-A947-70E740481C1C}"),
                # this is the svg specific part
                xml("a:ext", [
                    xml("asvg:svgBlip", "xmlns:asvg" => "http://schemas.microsoft.com/office/drawing/2016/SVG/main", "r:embed" => "rId$index")
                ], "uri" => "{96DAC541-7B7A-43D3-8B79-37D633B846F1}")
            ])
        ])
    elseif M <: MIME"image/png"
        xml("a:blip", "r:embed" => "rId$index")
    else
        error("Cannot deal with image of MIME type $M")
    end
    return (; ablip, index)
end

function get_ablip(i::InlineDrawing{<:SVGWithPNGFallback}, rels)
    index_png = rels[i.image.png]
    index_svg = rels[i.image.svg]
    ablip = 
        xml("a:blip", [
            xml("a:extLst", [
                xml("a:ext", "uri" => "{28A0092B-C50C-407E-A947-70E740481C1C}"),
                # this is the svg specific part
                xml("a:ext", [
                    xml("asvg:svgBlip", "xmlns:asvg" => "http://schemas.microsoft.com/office/drawing/2016/SVG/main", "r:embed" => "rId$index_svg")
                ], "uri" => "{96DAC541-7B7A-43D3-8B79-37D633B846F1}")
            ])
        ],  "r:embed" => "rId$index_png")
    return (; ablip, index = index_svg)
end

function to_xml(i::InlineDrawing, rels)

    a = get_ablip(i, rels)
    ablip = a.ablip
    index = a.index

    name = "Picture $index" # it seems not to matter to reuse `index` here

    drawing = xml("w:drawing", [
        xml("wp:inline", [
            xml("wp:extent", "cx" => i.width, "cy" => i.height),
            xml("wp:docPr", "id" => index, "name" => name),
            xml("a:graphic", [
                xml("a:graphicData", [
                    xml("pic:pic", [
                        xml("pic:nvPicPr", [
                            xml("pic:cNvPr", "id" => index, "name" => name),
                            xml("pic:cNvPicPr"),
                        ]),
                        xml("pic:blipFill", [
                            ablip,
                            xml("a:stretch", [
                                xml("a:fillRect")
                            ])
                        ]),
                        xml("pic:spPr", [
                            xml("a:xfrm", [
                                xml("a:off", "x" => 0, "y" => 0),
                                xml("a:ext", "cx" => i.width, "cy" => i.height)
                            ]),
                            xml("a:prstGeom", [
                                xml("a:avLst")
                            ], "prst" => "rect")
                        ])
                    ],
                    "xmlns:pic" => "http://schemas.openxmlformats.org/drawingml/2006/picture")
                ], "uri" => "http://schemas.openxmlformats.org/drawingml/2006/picture"
            )
            ], "xmlns:a" => "http://schemas.openxmlformats.org/drawingml/2006/main")
        ])
    ])
    return drawing
end

function to_xml(c::ComplexField, rels)
    [
        xml("w:r", [xml("w:fldChar", "w:fldCharType" => "begin", (c.dirty === nothing ? () : ("w:dirty" => c.dirty,))...)]),
        xml("w:r", [xml("w:instrText", [E.TextNode(" $(c.instruction) ")], "xml:space" => "preserve")]),
        xml("w:r", [xml("w:fldChar", "w:fldCharType" => "separate")]),
    ]
end
function to_xml(c::ComplexFieldEnd, rels)
    xml("w:r", [xml("w:fldChar", "w:fldCharType" => "end")])
end

children(body::Body) = body.sections
children(section::Section) = section.children
children(paragraph::Paragraph) = paragraph.children
children(run::Run) = run.children
children(text::Text) = (text.text,)
children(table::Table) = table.rows
children(tablerow::TableRow) = tablerow.cells
children(tablecell::TableCell) = tablecell.children
children(h::Header) = h.children
children(f::Footer) = f.children
children(_) = ()

function children(r::RunProperties)
    c = []
    r.style === nothing || push!(c, RunStyle(r.style))
    r.color === nothing || push!(c, Color(r.color))
    r.size === nothing || push!(c, Size(r.size))
    r.valign === nothing || push!(c, r.valign)
    r.fonts === nothing || push!(c, r.fonts)
    r.bold === nothing || push!(c, Bold())
    r.italic === nothing || push!(c, Italic())
    return c
end

function children(p::ParagraphProperties)
    c = []
    p.style === nothing || push!(c, ParagraphStyle(p.style))
    p.run_properties === nothing || push!(c, p.run_properties)
    p.justification === nothing || push!(c, p.justification)
    p.spacing === nothing || push!(c, p.spacing)
    p.borders === nothing || push!(c, p.borders)
    p.shading === nothing || push!(c, p.shading)
    something(p.keep_next, false) && push!(c, xml("w:keepNext"))
    something(p.page_break_before, false) && push!(c, xml("w:pageBreakBefore"))
    something(p.keep_lines, false) && push!(c, xml("w:keepLines"))
    return c
end

function children(p::Columns)
    c = []
    if !p.equal
        append!(c, p.cols)
    end
    return c
end

function children(p::TableCellProperties)
    c = []
    p.borders === nothing || push!(c, p.borders)
    p.vertical_merge === nothing || push!(c, VerticalMerge(p.vertical_merge))
    p.gridspan === nothing || push!(c, GridSpan(p.gridspan))
    p.margins === nothing || push!(c, p.margins)
    p.valign === nothing || push!(c, p.valign)
    something(p.hide_mark, false) && push!(c, xml("w:hideMark"))
    return c
end

function children(p::TableProperties)
    c = []
    p.margins === nothing || push!(c, p.margins)
    p.spacing === nothing || push!(c, xml("w:tblCellSpacing", "w:w" => p.spacing))
    p.justification === nothing || push!(c, p.justification)
    return c
end

function children(p::TableRowProperties)
    c = []
    p.header === true && push!(c, xml("w:tblHeader"))
    p.height === nothing || push!(c, p.height)
    return c
end

function children(s::Style)
    c = []
    push!(c, xml("w:name", "w:val" => s.name))
    s.link === nothing || push!(c, xml("w:link", "w:val" => s.link))
    s.primary_style && push!(c, xml("w:qFormat"))
    s.next === nothing || push!(c, xml("w:next", "w:val" => s.next))
    s.based_on === nothing || push!(c, xml("w:basedOn", "w:val" => s.based_on))
    s.ui_priority === nothing || push!(c, xml("w:uiPriority", "w:val" => s.ui_priority))
    return c
end

function children(t::TableCellMargins)
    c = []
    t.top === nothing || push!(c, xml("w:top", "w:w" => t.top))
    t.bottom === nothing || push!(c, xml("w:bottom", "w:w" => t.bottom))
    t.start === nothing || push!(c, xml("w:start", "w:w" => t.start))
    t.stop === nothing || push!(c, xml("w:end", "w:w" => t.stop))
    return c
end

function children(t::TableLevelCellMargins)
    c = []
    t.top === nothing || push!(c, xml("w:top", "w:w" => t.top))
    t.bottom === nothing || push!(c, xml("w:bottom", "w:w" => t.bottom))
    t.start === nothing || push!(c, xml("w:start", "w:w" => t.start))
    t.stop === nothing || push!(c, xml("w:end", "w:w" => t.stop))
    return c
end

attributes(_) = ()
attributes(s::Size) = (("w:val", s.size),)
attributes(u::Underline) = (("w:val", u.pattern), ("w:color", u.color))
attributes(v::VerticalMerge) = (("w:val", v.restart ? "restart" : "continue"),)
attributes(v::VerticalAlign.T) = (("w:val", v),)
attributes(c::Color) = (("w:val", c.color),)
attributes(g::GridSpan) = (("w:val", g.n),)
attributes(t::Tuple{<:Union{TableCellBorder,ParagraphBorder}, Symbol}) = attributes(t[1])
function attributes(t::Union{TableCellBorder, ParagraphBorder})
    attrs = Tuple{String, Any}[("w:val", t.style)]
    t.color === nothing || push!(attrs, ("w:color", t.color))
    t.shadow === nothing || push!(attrs, ("w:shadow", t.shadow))
    t.space === nothing || push!(attrs, ("w:space", t.space))
    t.size === nothing || push!(attrs, ("w:sz", t.size))
    return attrs
end
attributes(p::PageSize) = (("w:h", p.height), ("w:w", p.width), ("w:orient", p.orientation))
attributes(p::PageMargins) = (("w:top", p.top), ("w:right", p.right), ("w:bottom", p.bottom), ("w:left", p.left), ("w:header", p.header), ("w:footer", p.footer), ("w:gutter", p.gutter))
attributes(p::PageVerticalAlign.T) = (("w:val", p),)
attributes(p::Justification.T) = (("w:val", p),)
function attributes(s::Style)
    attrstyle(::Style{RunProperties}) = "character"
    attrstyle(::Style{ParagraphProperties}) = "paragraph"
    attrstyle(::Style{TableProperties}) = "table"
    attrs = Tuple{String, Any}[
        ("w:type", attrstyle(s)),
        ("w:styleId", s.id),
    ]
    s.custom_style === nothing || push!(attrs, ("w:customStyle", s.custom_style))
    return attrs
end
attributes(p::ParagraphStyle) = (("w:val", p.name),)
attributes(p::RunStyle) = (("w:val", p.name),)
attributes(p::VerticalAlignment.T) = (("w:val", p),)
function attributes(p::Column)
    attrs = Tuple{String, Any}[]
    p.width === nothing || push!(attrs, ("w:w", p.width))
    p.space === nothing || push!(attrs, ("w:space", p.space))
    return attrs
end
function attributes(p::Columns)
    attrs = Tuple{String, Any}[]
    if p.num > 1 && p.equal
        push!(attrs, ("w:equalWidth", p.equal))
        push!(attrs, ("w:num", p.num))
        p.sep === false || push!(attrs, ("w:sep", p.sep))
        p.space === nothing || push!(attrs, ("w:space", p.space))
    elseif !p.equal
        push!(attrs, ("w:equalWidth", p.equal))
        p.sep === false || push!(attrs, ("w:sep", p.sep))
    end
    return attrs
end
function attributes(f::Fonts)
    attrs = Tuple{String, Any}[]
    f.ascii === nothing || push!(attrs, ("w:ascii", f.ascii))
    f.high_ansi === nothing || push!(attrs, ("w:hAnsi", f.high_ansi))
    f.east_asia === nothing || push!(attrs, ("w:eastAsia", f.east_asia))
    f.complex === nothing || push!(attrs, ("w:cs", f.complex))
    return attrs
end
function attributes(t::Text)
    attrs = Tuple{String, String}[]
    # if there's leading or trailing whitespace, this needs to be preserved explicitly
    if match(r"^\s|\s$", t.text) !== nothing
        push!(attrs, ("xml:space", "preserve"))
    end
    return attrs
end
attributes(b::Break) = (("w:type", b.type),)
function attributes(x::Union{Header, Footer})
    [
        ("xmlns:w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"),
        ("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"),
        ("xmlns:wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"),
    ]
end
function attributes(s::SimpleField)
    attrs = Tuple{String, Any}[("w:instr", s.instruction)]
    s.dirty === nothing || push!(attrs, ("w:dirty", s.dirty))
    return attrs
end
function attributes(s::Spacing)
    attrs = Tuple{String, Any}[]
    s.before === nothing || push!(attrs, ("w:before", s.before))
    s.after === nothing || push!(attrs, ("w:after", s.after))
    return attrs
end
function attributes(s::Shading)
    Tuple{String, Any}[
        ("w:val", s.pattern),
        ("w:color", s.color),
        ("w:fill", s.fill),
    ]
end
function attributes(t::TableRowHeight)
    attrs = Tuple{String, Any}[("w:val", t.height)]
    t.hrule === nothing || push!(attrs, ("w:hRule", t.hrule))
    return attrs
end

xmlnode(x) = E.ElementNode(xmltag(x))

xmltag(::Document) = "w:document"
xmltag(::Body) = "w:body"
xmltag(::Paragraph) = "w:p"
xmltag(::Run) = "w:r"
xmltag(::Text) = "w:t"
xmltag(::Size) = "w:sz"
xmltag(::VerticalAlignment.T) = "w:vertAlign"
xmltag(::Underline) = "w:u"
xmltag(::Color) = "w:color"
xmltag(::Bold) = "w:b"
xmltag(::Italic) = "w:i"
xmltag(::ParagraphStyle) = "w:pStyle"
xmltag(::RunStyle) = "w:rStyle"
xmltag(::Table) = "w:tbl"
xmltag(::TableRow) = "w:tr"
xmltag(::TableCell) = "w:tc"
xmltag(::ParagraphProperties) = "w:pPr"
xmltag(::RunProperties) = "w:rPr"
xmltag(::TableCellProperties) = "w:tcPr"
xmltag(::TableProperties) = "w:tblPr"
xmltag(::GridSpan) = "w:gridSpan"
xmltag(::VerticalMerge) = "w:vMerge"
xmltag(::TableCellBorders) = "w:tcBorders"
xmltag(::ParagraphBorders) = "w:pBdr"
function xmltag(t::Tuple{TableCellBorder, Symbol})
    sym = t[2]
    sym === :top ? "w:top" :
    sym === :start ? "w:start" :
    sym === :end ? "w:end" :
    sym === :bottom ? "w:bottom" :
    sym === :tl2br ? "w:tl2br" :
    sym === :tr2bl ? "w:tr2bl" :
    sym === :insideH ? "w:insideH" :
    sym === :insideV ? "w:insideV" :
    throw(ArgumentError("Invalid symbol $(repr(sym)), options are :top, :start, :end, :bottom, :tl2br, :tr2bl, :insideH, :insideV."))
end
function xmltag(t::Tuple{ParagraphBorder, Symbol})
    sym = t[2]
    sym === :top ? "w:top" :
    sym === :left ? "w:left" :
    sym === :right ? "w:right" :
    sym === :bottom ? "w:bottom" :
    sym === :between ? "w:between" :
    throw(ArgumentError("Invalid symbol $(repr(sym)), options are :top, :left, :right, :bottom, :between."))
end
xmltag(::PageSize) = "w:pgSz"
xmltag(::PageMargins) = "w:pgMar"
xmltag(::PageVerticalAlign.T) = "w:vAlign"
xmltag(::Columns) = "w:cols"
xmltag(::Column) = "w:col"
xmltag(::Style) = "w:style"
xmltag(::Justification.T) = "w:jc"
xmltag(::Fonts) = "w:rFonts"
xmltag(::TableCellMargins) = "w:tcMar"
xmltag(::TableLevelCellMargins) = "w:tblCellMar"
xmltag(::VerticalAlign.T) = "w:vAlign"
xmltag(::Break) = "w:br"
xmltag(::Header) = "w:hdr"
xmltag(::Footer) = "w:ftr"
xmltag(::SimpleField) = "w:fldSimple"
xmltag(::Spacing) = "w:spacing"
xmltag(::TableRowProperties) = "w:trPr"
xmltag(::TableRowHeight) = "w:trHeight"
xmltag(::Shading) = "w:shd"

xmlstring(h::HexColor) = h.hex
xmlstring(s::String) = s
xmlstring(b::Bool) = string(b)
xmlstring(number::Union{Int, Float64}) = string(number)
xmlstring(a::AutomaticDefault) = xmlstring(a.value)
xmlstring(::Automatic) = "auto"
xmlstring(v::VerticalAlignment.T) = string(v)
xmlstring(v::BorderStyle.T) = snake_to_camel(string(v))
xmlstring(u::UnderlinePattern.T) = snake_to_camel(string(u))
xmlstring(p::PageOrientation.T) = string(p)
xmlstring(p::PageVerticalAlign.T) = string(p)
xmlstring(p::Justification.T) = p === Justification.stop ? "end" : string(p)
xmlstring(l::Length) = string(round(Int, l.value)) # all sizes should be converted at this point and word only wants integers
xmlstring(v::VerticalAlign.T) = string(v)
xmlstring(b::BreakType.T) = snake_to_camel(string(b))
xmlstring(b::HeightRule.T) = snake_to_camel(string(b))
xmlstring(b::ShadingPattern.T) = snake_to_camel(string(b))
end
