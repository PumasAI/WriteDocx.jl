# Tables

```@meta
CurrentModule = WriteDocx
```

A [`Table`](@ref) consists of [`TableRow`](@ref)s with [`TableCell`](@ref)s inside.
Each cell can be styled separately, for example with borders and margins:

```@example
import WriteDocx as W

border() = W.TableCellBorder(
    color = W.HexColor("000000"),
    style = W.BorderStyle.single,
)

function cell(string)
    paragraph = W.Paragraph([W.Run([W.Text(string)])])
    return W.TableCell(
        [paragraph],
        borders = W.TableCellBorders(
            top = border(),
            bottom = border(),
            start = border(),
            stop = border(),
        ),
        margins = W.TableCellMargins(
            top = 10 * W.pt,
            bottom = 10 * W.pt,
            start = 10 * W.pt,
            stop = 10 * W.pt,
        ),
    )
end

cells = [cell("$col$row") for row = 1:8, col = 'A':'H']

rows = [W.TableRow(row) for row in eachrow(cells)]

doc = W.Document(W.Body([W.Section([W.Table(rows)])]))

W.save("table.docx", doc)
```

Download `table.docx`:

```@raw html
<a href="../table.docx"><img src="./../../assets/icon_docx.png" width="60">
```

## Width and column layout

By default, Word sizes a table and its columns to fit their content.
A [`TableWidth`](@ref) sets how wide the table itself is, either as a [`Length`](@ref) or as a [`Percent`](@ref) of the surrounding text column.
The columns follow the widths in the table's `grid` if its [`TableLayout`](@ref) is `fixed`:

```@example
import WriteDocx as W

function cell(string)
    W.TableCell([W.Paragraph([W.Run([W.Text(string)])])])
end

filling = W.Table(
    [W.TableRow([cell("A"), cell("B")])],
    width = 100 * W.percent,
)

fixed = W.Table(
    [W.TableRow([cell("narrow"), cell("wide")])],
    grid = [3 * W.cm, 9 * W.cm],
    width = 12 * W.cm,
    layout = W.TableLayout.fixed,
)

doc = W.Document(W.Body([W.Section([filling, fixed])]))

W.save("table_width.docx", doc)
```

Download `table_width.docx`:

```@raw html
<a href="../table_width.docx"><img src="./../../assets/icon_docx.png" width="60">
```
