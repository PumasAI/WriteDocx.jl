# Paragraphs

```@meta
CurrentModule = WriteDocx
```

## Tab stops

A [`Tab`](@ref) in a [`Run`](@ref) advances to the next [`TabStop`](@ref) of its paragraph.
Each stop has a position, a [`TabAlignment`](@ref) that decides how the text sits relative to it, and a [`TabLeader`](@ref) that fills the space the tab jumps:

```@example
import WriteDocx as W

stops = [
    W.TabStop(6 * W.cm, alignment = W.TabAlignment.stop, leader = W.TabLeader.dot),
    W.TabStop(12 * W.cm, alignment = W.TabAlignment.decimal),
]

function row(label, amount)
    W.Paragraph(
        [W.Run([W.Text(label), W.Tab(), W.Text(""), W.Tab(), W.Text(amount)])],
        W.ParagraphProperties(tabs = stops),
    )
end

doc = W.Document(W.Body([W.Section([
    row("Coffee", "3.50"),
    row("Sandwich", "12.75"),
])]))

W.save("tab_stops.docx", doc)
```

Download `tab_stops.docx`:

```@raw html
<a href="../tab_stops.docx"><img src="./../../assets/icon_docx.png" width="60">
```

## Line spacing

The `line` of a [`Spacing`](@ref) sets how tall the lines of a paragraph are.
A [`Percent`](@ref) is a multiple of single spacing, a [`Length`](@ref) fixes the height exactly, and an [`AtLeast`](@ref) lets Word grow the line to fit content taller than the given height:

```@example
import WriteDocx as W

function paragraph(line)
    W.Paragraph(
        [W.Run([W.Text("The quick brown fox jumps over the lazy dog. "^3)])],
        W.ParagraphProperties(spacing = W.Spacing(line = line)),
    )
end

doc = W.Document(W.Body([W.Section([
    paragraph(150 * W.percent),
    paragraph(14 * W.pt),
    paragraph(W.AtLeast(20 * W.pt)),
])]))

W.save("line_spacing.docx", doc)
```

Download `line_spacing.docx`:

```@raw html
<a href="../line_spacing.docx"><img src="./../../assets/icon_docx.png" width="60">
```
