# Bookmarks

```@meta
CurrentModule = WriteDocx
```

A [`Bookmark`](@ref) gives a name to a part of a document so that other parts can point at it.
A [`Hyperlink`](@ref) jumps to a bookmark and a [`PageReference`](@ref) shows the page it is on, which together with a dotted [`TabStop`](@ref) is all a table of contents needs:

```@example
import WriteDocx as W

function entry(title, anchor)
    W.Paragraph(
        [
            W.Hyperlink([W.Run([W.Text(title)])], anchor = anchor),
            W.Run([W.Tab()]),
            W.PageReference(anchor),
        ],
        W.ParagraphProperties(tabs = [
            W.TabStop(15 * W.cm, alignment = W.TabAlignment.stop, leader = W.TabLeader.dot),
        ]),
    )
end

function chapter(title, anchor)
    W.Bookmark(anchor, [
        W.Paragraph([W.Run([W.Text(title)])], style = "Heading1"),
        W.Paragraph([W.Run([W.Text("The quick brown fox jumps over the lazy dog. "^20)])]),
        W.Paragraph([W.Run([W.Break(W.BreakType.page)])]),
    ])
end

doc = W.Document(W.Body([W.Section([
    entry("Introduction", "introduction"),
    entry("Methods", "methods"),
    W.Paragraph([W.Run([W.Break(W.BreakType.page)])]),
    chapter("Introduction", "introduction"),
    chapter("Methods", "methods"),
])]))

W.save("bookmarks.docx", doc)
```

Download `bookmarks.docx`:

```@raw html
<a href="../bookmarks.docx"><img src="./../../assets/icon_docx.png" width="60">
```

!!! note
    A [`PageReference`](@ref) is a field whose result Word computes, so the page numbers
    appear once it recalculates the document's fields, which it does when printing or
    exporting. Until then they may show as empty. Pressing `Ctrl+A` and then `F9` updates
    them by hand.

Bookmark names identify their bookmark, so they have to be unique within a document, and links pointing at a name that no bookmark defines are an error when the document is saved.
Word truncates names longer than 40 characters and replaces whitespace in them with underscores, which would silently break the links pointing at them, so both are rejected instead.
