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
<a href="table.docx"><img src="./../../assets/icon_docx.png" width="60">
```
