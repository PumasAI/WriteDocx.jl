# WriteDocx

WriteDocx is a utility package that lets you create .docx files compliant with [ECMA-376](https://ecma-international.org/publications-and-standards/standards/ecma-376/), for use with Microsoft Office Word and other compatible software.
Under the hood, these files are zip files containing a standardized folder structure with XML files and other assets.

WriteDocx contains many Julia types that mirror the types of XML nodes commonly found in docx files, without the user having to write any XML manually.

!!! note
    
    The types of WriteDocx.jl do not exhaustively cover the vast [ECMA-376](https://ecma-international.org/publications-and-standards/standards/ecma-376/) spec.
    Instead, we have implemented the parts most useful to us and will consider extending this set more and more when there's a specific need.

## Example

Here's a simple document with two paragraphs, one of which has pink-colored text:

```@example
import WriteDocx as W

doc = W.Document(
    W.Body([
        W.Section([
            W.Paragraph([
                W.Run([W.Text("Hello world, from WriteDocx.jl")]),
            ]),
            W.Paragraph([
                W.Run(
                    [W.Text("Goodbye!")],
                    color = W.HexColor("FF00FF"),
                ),
            ]),
        ]),
    ]),
)

W.save("example.docx", doc)
nothing # hide
```

Download `example.docx`:

```@raw html
<a href="example.docx"><img src="./assets/icon_docx.png" width="60">
```
