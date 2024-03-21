# Headers & Footers

```@meta
CurrentModule = WriteDocx
```

Headers and footers are specified using the [`Headers`](@ref) and [`Footers`](@ref) objects.
You can specify different styles for the first page and for every even page, but in this example we only create a default header and footer:

```@example
import WriteDocx as W

headers = W.Headers(
    default = W.Header([
        W.Paragraph([W.Run([W.Text("The header")])]),
    ]),
)

footers = W.Footers(
    default = W.Footer([
        W.Paragraph([W.Run([W.Text("The footer")])]),
    ]),
)

doc = W.Document(W.Body([W.Section([]; headers, footers)]))

W.save("headers_and_footers.docx", doc)
```

Download `headers_and_footers.docx`:

```@raw html
<a href="../headers_and_footers.docx"><img src="./../../assets/icon_docx.png" width="60">
```
