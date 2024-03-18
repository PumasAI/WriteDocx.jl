using WriteDocx
using Documenter

DocMeta.setdocmeta!(WriteDocx, :DocTestSetup, :(using WriteDocx); recursive = true)

makedocs(;
    modules = [WriteDocx],
    authors = "Julius Krumbiegel <julius.krumbiegel@gmail.com> and contributors",
    repo = "https://github.com/PumasAI/WriteDocx.jl/blob/{commit}{path}#{line}",
    sitename = "WriteDocx.jl",
    format = Documenter.HTML(;
        prettyurls = get(ENV, "CI", "false") == "true",
        canonical = "https://pumasai.github.io/WriteDocx.jl",
        edit_link = "main",
        assets = String[],
    ),
    pages = [
        "Home" => "index.md",
        "Examples" => [
            "examples/tables.md",
            "examples/drawings.md",
            "examples/headers_and_footers.md",
        ],
        "api.md",
    ],
)

deploydocs(; repo = "github.com/PumasAI/WriteDocx.jl", devbranch = "main", push_preview = true)
