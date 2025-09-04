import WriteDocx as W
using Test
using ImageIO
using FileIO
using ReferenceTests: @test_reference
using ZipFile: ZipFile

# Set ENV["WRITEDOCX_OPEN_TEST_FILES"] = true to open each test docx file for a visual check

function unzip(zipfile, dir)
    zipreader = ZipFile.Reader(zipfile)
    for file in zipreader.files
        mkpath(joinpath(dir, splitdir(file.name)[1]))
        open(joinpath(dir, file.name), "w") do io
            write(io, file)
        end
    end
    close(zipreader)
end

function reftest_docx(doc::W.Document, reference_name)
    reference_dir = joinpath(@__DIR__, "references", reference_name)
    mktempdir() do dir
        docxpath = joinpath(dir, "temp.docx")
        zipdir = joinpath(dir, "unzipped")
        W.save(docxpath, doc)

        if get(ENV, "WRITEDOCX_OPEN_TEST_FILES", "false") == "true"
            run(`open $docxpath`)
            println("Opening $(repr(reference_name)). Press enter to continue testing")
            readline()
            println("Continuing.")
        end

        unzip(docxpath, zipdir)

        # test that all reference files exist in the test docx
        for (root, _, files) in walkdir(reference_dir)
            relroot = relpath(root, reference_dir)
            for file in files
                if !isfile(joinpath(zipdir, relroot, file))
                    error("Reference file $(joinpath(relroot, file)) is missing in docx.")
                end
            end
        end
        # now test that all references match
        for (root, _, files) in walkdir(zipdir)
            relroot = relpath(root, zipdir)
            for file in files
                extension = splitext(file)[2]
                reference_path = joinpath(reference_dir, relroot, file)
                test_path = joinpath(root, file)
                if extension in [".png", ".jpg", ".svg"]
                    if !isfile(reference_path)
                        if get(ENV, "JULIA_REFERENCETESTS_UPDATE", "false") == "true"
                            mkpath(splitdir(reference_path)[1])
                            cp(test_path, reference_path)
                        else
                            @warn "Image at $reference_path does not exist, set ENV[\"JULIA_REFERENCETESTS_UPDATE\"] = true to add it."
                        end
                    end
                    # we just care about exact equality of images, because they should
                    # be copied (but direct comparison of files seemed to fail, maybe
                    # because of slightly different metadata)
                    if extension == ".svg"
                        @test read(reference_path, String) == read(test_path, String)
                    else
                        @test FileIO.load(reference_path) == FileIO.load(test_path)
                    end
                else
                    content = read(test_path, String)
                    @test_reference reference_path content
                end
            end
        end
    end
    return
end

struct SVG
    svg::String
end

Base.show(io::IO, ::MIME"image/svg+xml", s::SVG) = print(io, s.svg)

struct PNG
    bytes::Vector{UInt8}
end

Base.show(io::IO, ::MIME"image/png", p::PNG) = write(io, p.bytes)

@testset "WriteDocx.jl" begin
    @testset "Basic document" begin
        doc = W.Document(
            W.Body(
                [
                    W.Section(
                        [
                            W.Paragraph(
                                [
                                    W.Run(
                                        [
                                            W.Text("Test")
                                        ]
                                    )
                                ],
                            )
                        ]
                    )
                ]
            )
        )
        reftest_docx(doc, "basic_document")
    end
    @testset "Run properties" begin
        doc = W.Document(
            W.Body(
                [
                    W.Section(
                        [
                            W.Paragraph(
                                [
                                    W.Run(
                                        [
                                            W.Text("Pink")
                                        ],
                                        W.RunProperties(
                                            color = W.HexColor("FF00AA"),
                                            size = 30W.pt
                                        )
                                    )
                                ],
                            ),
                            W.Paragraph(
                                [  
                                    W.Run(
                                        [
                                            W.Text("Normal text")
                                        ],
                                    )
                                    W.Run(
                                        [
                                            W.Text("superscript")
                                        ],
                                        W.RunProperties(
                                            valign = W.VerticalAlignment.superscript,
                                        )
                                    )
                                ],
                            ),
                            W.Paragraph(
                                [  
                                    W.Run(
                                        [
                                            W.Text("Arial ")
                                        ],
                                        W.RunProperties(
                                            fonts = W.Fonts("Arial"),
                                        )
                                    )
                                    W.Run(
                                        [
                                            W.Text("Times New Roman")
                                        ],
                                        W.RunProperties(
                                            fonts = W.Fonts("Times New Roman"),
                                        )
                                    )
                                ],
                            ),
                            W.Paragraph(
                                [  
                                    W.Run(
                                        [
                                            W.Text("Bold ")
                                        ],
                                        W.RunProperties(
                                            bold = true,
                                        )
                                    )
                                    W.Run(
                                        [
                                            W.Text("Italic")
                                        ],
                                        W.RunProperties(
                                            italic = true,
                                        )
                                    )
                                ],
                            )
                        ]
                    )
                ]
            )
        )
        reftest_docx(doc, "run_properties")
    end

    @testset "Paragraph properties" begin
        doc = W.Document(
            W.Body(
                [
                    W.Section(
                        [
                            W.Paragraph(
                                [
                                    W.Run(
                                        [
                                            W.Text("Some test text")
                                        ],
                                    )
                                ],
                                W.ParagraphProperties(
                                    justification = W.Justification.distribute,
                                    borders = W.ParagraphBorders(
                                        top = W.ParagraphBorder(size = 1 * W.pt, style = W.BorderStyle.single),
                                        bottom = W.ParagraphBorder(size = 1.5 * W.pt, style = W.BorderStyle.wave),
                                        left = W.ParagraphBorder(size = 2 * W.pt, style = W.BorderStyle.double),
                                        right = W.ParagraphBorder(size = 2.5 * W.pt, style = W.BorderStyle.triple),
                                    )
                                )
                            ),
                            W.Paragraph(
                                [
                                    W.Run(
                                        [
                                            W.Text("Some test text")
                                        ],
                                    )
                                ],
                                W.ParagraphProperties(
                                    shading = W.Shading(
                                        color = W.HexColor("FFDDDD"),
                                        pattern = W.ShadingPattern.diag_cross,
                                        fill = W.HexColor("002255"),
                                    )
                                )
                            ),
                            W.Paragraph(
                                [
                                    W.Run(
                                        [
                                            W.Text("Some test text")
                                        ],
                                    )
                                ],
                                W.ParagraphProperties(
                                    shading = W.Shading(
                                        fill = W.HexColor("EEEEEE"),
                                    )
                                )
                            ),
                        ]
                    )
                ]
            )
        )
        reftest_docx(doc, "paragraph_properties")
    end

    @testset "Basic table" begin

        black_cell = W.TableCellBorders(
            top = W.TableCellBorder(
                color = W.HexColor("000000"),
                style = W.BorderStyle.single,
            ),
            bottom = W.TableCellBorder(
                color = W.HexColor("000000"),
                style = W.BorderStyle.single,
            ),
            start = W.TableCellBorder(
                color = W.HexColor("000000"),
                style = W.BorderStyle.single,
            ),
            stop = W.TableCellBorder(
                color = W.HexColor("000000"),
                style = W.BorderStyle.single,
            ),
        )

        doc = W.Document(
            W.Body(
                [
                    W.Section(
                        [
                            W.Table(
                                [
                                    W.TableRow([
                                        W.TableCell([
                                            W.Paragraph([
                                                W.Run([
                                                    W.Text("Vertically Merged")
                                                ])
                                            ])
                                        ],
                                        W.TableCellProperties(
                                            vertical_merge = true,
                                            borders = black_cell,
                                            valign = W.VerticalAlign.center,
                                        )
                                        ),
                                        W.TableCell(
                                            [
                                                W.Paragraph([
                                                    W.Run([
                                                        W.Text("Horizontally Merged")
                                                    ])
                                                ])
                                            ],
                                            W.TableCellProperties(
                                                gridspan = 2,
                                                borders = W.TableCellBorders(
                                                    top = W.TableCellBorder(
                                                        color = W.HexColor("FF0000"),
                                                        style = W.BorderStyle.single,
                                                        size = W.eighthpt * 2
                                                    ),
                                                    bottom = W.TableCellBorder(
                                                        color = W.HexColor("00FF00"),
                                                        style = W.BorderStyle.single,
                                                        size = W.eighthpt * 6,
                                                    ),
                                                    start = W.TableCellBorder(
                                                        color = W.HexColor("0000FF"),
                                                        style = W.BorderStyle.single,
                                                        size = W.eighthpt * 12,
                                                    ),
                                                    stop = W.TableCellBorder(
                                                        color = W.HexColor("FF00FF"),
                                                        style = W.BorderStyle.single,
                                                        size = W.eighthpt * 16,
                                                        shadow = true,
                                                    ),
                                                )
                                            )
                                        ),
                                    ]),
                                    W.TableRow([
                                        W.TableCell(
                                            [
                                                W.Paragraph([])
                                            ],
                                            W.TableCellProperties(
                                                vertical_merge = false,
                                                borders = black_cell,
                                            )
                                        ),
                                        W.TableCell([
                                            W.Paragraph([
                                                W.Run([
                                                    W.Text("Normal Cell 1")
                                                ])
                                            ])
                                        ], W.TableCellProperties(
                                            margins = W.TableCellMargins(
                                                top = 5W.pt,
                                                bottom = 10W.pt,
                                                start = 15W.pt,
                                                stop = 20W.pt,
                                            )
                                        )),
                                        W.TableCell([
                                            W.Paragraph([
                                                W.Run([
                                                    W.Text("Normal Cell 2")
                                                ])
                                            ])
                                        ]),
                                    ]),
                                ],
                                W.TableProperties(
                                    margins = W.TableLevelCellMargins(
                                        top = 5W.pt,
                                        bottom = 6W.pt,
                                        start = 7W.pt,
                                        stop = 8W.pt,
                                    )
                                )
                            )
                        ]
                    )
                ],
            ),
        )

        reftest_docx(doc, "basic_table")
    end

    @testset "Table justification" begin
        tbl(just) = W.Table([
            W.TableRow([
                W.TableCell([
                    W.Paragraph([
                        W.Run([
                            W.Text(string(just))
                        ])
                    ])
                ])
            ])
        ], W.TableProperties(justification = just))

        doc = W.Document(W.Body([
            W.Section(
                mapreduce(vcat, [W.Justification.start, W.Justification.center, W.Justification.stop]) do just
                    [
                        tbl(just),
                        W.Paragraph([])
                    ]
                end
            )
        ]))

        reftest_docx(doc, "table_justification")
    end

    @testset "Table hide mark" begin
        doc = W.Document(
            W.Body(
                [
                    W.Section(
                        [
                            W.Table(
                                [
                                    W.TableRow([
                                        W.TableCell([
                                            W.Paragraph([
                                                W.Run([
                                                    W.Text("Row 1 (row 2 should stay small if selected)")
                                                ])
                                            ])
                                        ])
                                    ]),
                                    W.TableRow([
                                        W.TableCell([
                                            W.Paragraph([
                                            ])
                                        ], W.TableCellProperties(
                                            hide_mark = true,
                                            borders = W.TableCellBorders(
                                                bottom = W.TableCellBorder(
                                                    color = W.automatic,
                                                    size = W.eighthpt * 8,
                                                    style = W.BorderStyle.single,
                                                )
                                            ),
                                            valign = W.VerticalAlign.center,
                                        ))
                                    ], W.TableRowProperties(
                                        height = W.TableRowHeight(W.Twip(1), hrule = W.HeightRule.exact)
                                    )),
                                    W.TableRow([
                                        W.TableCell([
                                            W.Paragraph([
                                                W.Run([
                                                    W.Text("Row 3")
                                                ])
                                            ])
                                        ])
                                    ])
                                ]
                            )
                        ]
                    )
                ]
            )
        )
        reftest_docx(doc, "table_hide_mark")
    end

    @testset "Table header" begin
        doc = W.Document(W.Body([
            W.Section([
                W.Table([
                    W.TableRow([
                        W.TableCell([
                            W.Paragraph([
                                W.Run([
                                    W.Text("HEADER ROW 1 (should repeat each page)")
                                ])
                            ])
                        ])
                    ], W.TableRowProperties(header = true));
                    W.TableRow([
                        W.TableCell([
                            W.Paragraph([
                                W.Run([
                                    W.Text("HEADER ROW 2 (should repeat each page)")
                                ])
                            ])
                        ], W.TableCellProperties(
                            borders = W.TableCellBorders(
                                bottom = W.TableCellBorder(
                                    color = W.automatic,
                                    size = W.eighthpt * 8,
                                    style = W.BorderStyle.single,
                                )
                            ),
                        ))
                    ], W.TableRowProperties(header = true));
                    map(1:100) do i
                        W.TableRow([
                            W.TableCell([
                                W.Paragraph([
                                    W.Run([
                                        W.Text("Content cell $i")
                                    ])
                                ])
                            ])
                        ])
                    end
                ])
            ])
        ]))
        reftest_docx(doc, "table_header")
    end

    @testset "Inline drawing png" begin

        doc = d = W.Document(
            W.Body(
                [
                    W.Section(
                        [
                            W.Paragraph(
                                [
                                    W.Run(
                                        [
                                            W.InlineDrawing(joinpath(@__DIR__, "pumas.png"), 700_000W.emu, 300_000W.emu),
                                        ],
                                    ),
                                    W.Run(
                                        [
                                            W.InlineDrawing(joinpath(@__DIR__, "pumas.png"), 5W.cm, 1.2W.inch),
                                        ],
                                    ),
                                ],
                            )
                        ]
                    )
                ]
            )
        )

        reftest_docx(doc, "inline_drawing_png")
    end

    @testset "Drawing in header and footer" begin
        doc = d = W.Document(
            W.Body(
                [
                    W.Section(
                        [W.Paragraph([])],
                        W.SectionProperties(
                            headers = W.Headers(
                                default = W.Header([
                                    W.Paragraph([
                                        W.Run(
                                            [
                                                W.InlineDrawing(joinpath(@__DIR__, "pumas.png"), 40 * W.pt, 15 * W.pt),
                                            ],
                                        ), 
                                    ])
                                ])
                            ),
                            footers = W.Footers(
                                default = W.Footer([
                                    W.Paragraph([
                                        W.Run(
                                            [
                                                W.InlineDrawing(joinpath(@__DIR__, "pumas.png"), 700_000W.emu, 300_000W.emu),
                                            ],
                                        ), 
                                    ])
                                ])
                            ),
                        )
                    )
                ]
            )
        )

        reftest_docx(doc, "drawing_in_header_and_footer")
    end

    @testset "Inline drawing svg" begin

        doc = d = W.Document(
            W.Body(
                [
                    W.Section(
                        [
                            W.Paragraph(
                                [
                                    W.Run(
                                        [
                                            W.InlineDrawing(joinpath(@__DIR__, "pumas.svg"), 700_000W.emu, 300_000W.emu),
                                        ],
                                    ),
                                    W.Run(
                                        [
                                            W.InlineDrawing(joinpath(@__DIR__, "pumas.svg"), 5W.cm, 1.2W.inch),
                                        ],
                                    ),
                                ],
                            )
                        ]
                    )
                ]
            )
        )

        reftest_docx(doc, "inline_drawing_svg")
    end

    @testset "Inline drawing svg with png fallback" begin

        doc = d = W.Document(
            W.Body(
                [
                    W.Section(
                        [
                            W.Paragraph(
                                [
                                    W.Run(
                                        [
                                            W.InlineDrawing(
                                                W.SVGWithPNGFallback(
                                                    svg = joinpath(@__DIR__, "pumas.svg"),
                                                    png = joinpath(@__DIR__, "pumas.png"),
                                                ),
                                                700_000W.emu,
                                                300_000W.emu
                                            ),
                                        ],
                                    ),
                                ],
                            )
                        ]
                    )
                ]
            )
        )

        reftest_docx(doc, "inline_drawing_svg_with_png_fallback")
    end

    @testset "Inline drawings with showable objects" begin

        svg = W.Image(MIME"image/svg+xml"(), SVG(read(joinpath(@__DIR__, "pumas.svg"), String)))
        png = W.Image(MIME"image/png"(), PNG(read(joinpath(@__DIR__, "pumas.png"))))

        doc = d = W.Document(
            W.Body(
                [
                    W.Section(
                        [
                            W.Paragraph(
                                [
                                    W.Run(
                                        [
                                            W.InlineDrawing(svg, 700_000W.emu, 300_000W.emu),
                                        ],
                                    ),
                                    W.Run(
                                        [
                                            W.InlineDrawing(png, 5W.cm, 1.2W.inch),
                                        ],
                                    ),
                                    W.Run(
                                        [
                                            W.InlineDrawing(
                                                W.SVGWithPNGFallback(
                                                    svg = svg,
                                                    png = png,
                                                ),
                                                700_000W.emu,
                                                300_000W.emu
                                            ),
                                        ],
                                    ),
                                ],
                            )
                        ]
                    )
                ]
            )
        )

        reftest_docx(doc, "inline_drawings_with_showable_objects")
    end

    @testset "Multiple sections" begin
        doc = d = W.Document(
            W.Body(
                map(1:3) do i
                    W.Section([
                        W.Paragraph([
                            W.Run([
                                W.Text("Section $i")
                            ])
                        ])
                    ])
                end
            )
        )

        reftest_docx(doc, "multiple_sections")
    end

    @testset "Columns" begin
        doc = W.Document(
            W.Body([
                W.Section([
                    W.Paragraph([
                        W.Run(
                            [W.Text("No columns " ^ 10)],
                            W.RunProperties(size = 40W.pt)
                        )
                    ])
                ])
                W.Section([
                    W.Paragraph([
                        W.Run(
                            [W.Text("2 columns " ^ 20)],
                            W.RunProperties(size = 40W.pt)
                        )
                    ])
                ]; columns=W.Columns(;num=2))
                W.Section([
                    W.Paragraph([
                        W.Run(
                            [W.Text("different width columns " ^ 12)],
                            W.RunProperties(size = 40W.pt)
                        )
                    ])
                    ]; columns=W.Columns(;cols=[W.Column(;width=4W.inch, space=0.5W.inch), W.Column(;width=2W.inch)]))
            ])
        )

        reftest_docx(doc, "columns")
    end

    @testset "Page margins" begin
        doc = W.Document(
            W.Body([
                W.Section([
                    W.Paragraph([
                        W.Run([
                            W.Text("no margins")
                        ])
                    ])
                ])
                W.Section([
                    W.Paragraph([
                        W.Run([
                            W.Text("thiccc margins")
                        ])
                    ])
                ]; margins=W.PageMargins(;top=2.5W.inch, right=2.5W.inch, bottom=2.5W.inch, left=2.5W.inch))
            ])
        )

        reftest_docx(doc, "page_margins")
    end

    @testset "Page sizes" begin
        doc = W.Document(
            W.Body(
                [
                    W.Section([
                        W.Paragraph([
                            W.Run([
                                W.Text("Portrait")
                            ])
                        ])
                    ], W.SectionProperties(pagesize = W.PageSize(5000W.twip, 10000W.twip))),
                    W.Section([
                        W.Paragraph([
                            W.Run([
                                W.Text("Landscape")
                            ])
                        ])
                    ], W.SectionProperties(pagesize = W.PageSize(29.7W.cm, 21W.cm))),

                ]
            )
        )

        reftest_docx(doc, "page_sizes")
    end

    @testset "Paragraph style" begin
        doc = W.Document(
            W.Body(
                [
                    W.Section([
                        W.Paragraph([
                            W.Run([
                                W.Text("Test")
                            ]),
                        ],
                        W.ParagraphProperties(
                            style = "TestStyle"
                        )
                        )
                    ])
                ]
            ),
            styles = W.Styles([
                W.Style(
                    "TestStyle",
                    W.ParagraphProperties(
                        justification = W.Justification.center,
                        run_properties = W.RunProperties(
                            color = W.HexColor("0033FF"),
                        )
                    ),
                    name = "Test Style",
                )
            ])
        )

        reftest_docx(doc, "paragraph_style")
    end

    @testset "Run style" begin
        doc = W.Document(
            W.Body(
                [
                    W.Section([
                        W.Paragraph([
                            W.Run(
                                [
                                    W.Text("Test")
                                ],
                                W.RunProperties(
                                    style = "TestStyle"
                                )
                            )
                        ])
                    ])
                ]
            ),
            styles = W.Styles([
                W.Style(
                    "TestStyle",
                    W.RunProperties(
                        color = W.HexColor("FF0000"),
                    ),
                    name = "Test Style",
                    custom_style = true,
                )
            ])
        )

        reftest_docx(doc, "run_style")
    end

    @testset "Missing styles" begin
        doc = W.Document(
            W.Body(
                [
                    W.Section([
                        W.Paragraph([
                            W.Run(
                                [
                                    W.Text("Test")
                                ],
                                W.RunProperties(
                                    style = "FirstTestStyle"
                                )
                            ),
                        ])
                    ])
                ]
            ),
        )
        @test_throws W.MissingStyleError("FirstTestStyle", :character) W.save("missing.docx", doc)

        doc = W.Document(
            W.Body(
                [
                    W.Section([
                        W.Paragraph([
                            W.Run(
                                [
                                    W.Text("Test")
                                ]
                            ),
                        ],
                        W.ParagraphProperties(
                                style = "SecondTestStyle"
                            )
                        )
                    ])
                ]
            ),
        )
        @test_throws W.MissingStyleError("SecondTestStyle", :paragraph) W.save("missing.docx", doc)
    end

    @testset "Duplicate styles" begin

        for (props, type) in [W.RunProperties => :character, W.ParagraphProperties => :paragraph]
            doc = W.Document(
                W.Body(
                    [
                        W.Section([
                            W.Paragraph([])
                        ])
                    ]
                ),
                styles = W.Styles(
                    [
                        W.Style(
                            "TestStyle",
                            props(),
                            name = "Test Style",
                        ),
                        W.Style(
                            "TestStyle",
                            props(),
                            name = "Test Style",
                        )
                    ]
                )
            )

            @test_throws W.DuplicateStyleError("TestStyle", type) W.save("duplicate.docx", doc)
        end
    end

    @testset "Breaks" begin
        doc = W.Document(
            W.Body(
                [
                    W.Section([
                        W.Paragraph([
                            W.Run(
                                [
                                    W.Text("Before line break"),
                                    W.Break(),
                                    W.Text("after line break, before page break"),
                                    W.Break(W.BreakType.page),
                                    W.Text("after page break"),
                                ]
                            )
                        ])
                    ])
                    W.Section([
                        W.Paragraph([
                            W.Run(
                                [
                                    W.Text("column 1"),
                                    W.Break(W.BreakType.column),
                                    W.Text("column 2"),
                                ]
                            )
                        ])
                    ]; columns=W.Columns(;num=2))
                ]
            )
        )
        reftest_docx(doc, "breaks")
    end

    @testset "Default styles" begin
        doc = W.Document(
            W.Body([
                W.Section([
                    mapreduce(vcat, 1:6) do level
                        [
                            W.Paragraph([
                                W.Run([
                                    W.Text("Heading $level")
                                ])
                            ], W.ParagraphProperties(
                                style = "Heading$level"
                            )),
                            W.Paragraph([
                                W.Run([
                                    W.Text("This is some text for a paragraph body. It is just meaningless text. Some more nonsensical characters follow - look at this one, it's also not interesting.")
                                ])
                            ], W.ParagraphProperties(
                                style = "Normal"
                            )),
                            W.Paragraph([
                                W.Run([
                                    W.Text("And this is another paragraph. It's not much better than the first in terms of writing quality. It just looks like text that is important enough to read it.")
                                ])
                            ], W.ParagraphProperties(
                                style = "Normal"
                            )),
                        ]
                    end...,
                ])
            ])
        )

        reftest_docx(doc, "default_styles")
    end

    @testset "Images with captions" begin
        doc = W.Document(
            W.Body([
                W.Section([
                    mapreduce(vcat, 1:3) do i
                        [
                            W.Paragraph([
                                W.Run([
                                    W.InlineDrawing(joinpath(@__DIR__, "pumas.png"), 7W.cm, 3W.cm)
                                ])
                            ], W.ParagraphProperties(
                                keep_next = true,
                                keep_lines = true,
                                page_break_before = true,
                            )),
                            W.Paragraph([
                                W.Run([
                                    W.Text("Caption number $i")
                                ])
                            ])
                        ]
                    end...,
                ])
            ])
        )

        reftest_docx(doc, "images_with_captions")
    end

    @testset "Header and footer" begin
        doc = W.Document(
            W.Body([
                W.Section([
                    W.Paragraph([
                        W.Run([
                            W.Text("Page content")
                        ])
                    ])
                ], W.SectionProperties(
                    headers = W.Headers(
                        default = W.Header([
                            W.Paragraph([
                                W.Run([
                                    W.Text("Header text")
                                ])
                            ])
                        ])
                    ),
                    footers = W.Footers(
                        default = W.Footer([
                            W.Paragraph([
                                W.Run([
                                    W.Text("Footer text")
                                ])
                            ])
                        ])
                    ),
                ))
            ])
        )

        reftest_docx(doc, "header_and_footer")
    end

    @testset "Page numbers" begin
        doc = W.Document(
            W.Body([
                W.Section(map(1:3) do i
                    W.Paragraph([
                        W.Run([
                            W.Text("Page $i"),
                            W.Break(W.BreakType.page),
                        ])
                    ])
                end, W.SectionProperties(
                    footers = W.Footers(
                        default = W.Footer([
                            W.Paragraph([
                                W.page_number_field(),
                            ])
                        ])
                    )
                ))
            ])
        )

        reftest_docx(doc, "page_numbers")
    end

    @testset "Table of contents" begin
        doc = W.Document(
            W.Body([
                W.Section([
                    W.Paragraph([
                        W.ComplexField("TOC \\o \"1-3\" \\h \\z \\u"),
                        W.ComplexFieldEnd()
                    ]),
                    W.Paragraph([
                            W.Run([
                                W.Text("Heading 1")
                            ])
                        ], W.ParagraphProperties(
                        style = "Heading1",
                    )),
                    W.Paragraph([
                            W.Run([
                                W.Text("Subheading 1")
                            ])
                        ], W.ParagraphProperties(
                        style = "Heading2",
                    )),
                    W.Paragraph([
                            W.Run([
                                W.Text("Heading 2")
                            ])
                        ], W.ParagraphProperties(
                        style = "Heading1",
                    )),
                    W.Paragraph([
                            W.Run([
                                W.Text("Subheading 2")
                            ])
                        ], W.ParagraphProperties(
                        style = "Heading2",
                    )),
                    W.Paragraph([
                            W.Run([
                                W.Text("Subsubheading 1")
                            ])
                        ], W.ParagraphProperties(
                        style = "Heading3",
                    )),
                ])
            ])
        )

        reftest_docx(doc, "table_of_contents")
    end

    @testset "Table of figures" begin
        doc = W.Document(
            W.Body([
                W.Section([
                    W.Paragraph([
                        # The \t flag here makes a TOC using only the paragraph styles behind that flag
                        # The problem with auto-numbered figures using a SEQ simple field (that's what Word
                        # does by default) is that the TOC is before the SEQ fields in the document, so
                        # upon opening the document, the TOC is updated before the SEQ fields are, which means
                        # the numbers show up in the captions but not in the TOC. Only after "update field" is
                        # manually selected is the TOC updated correctly. The alternative is to hardcode numbers
                        # and instead of via SEQ field (for example \s "Figure") gather the captions by some style
                        # name as done below with "Figure Caption Title".
                        W.ComplexField("TOC \\h \\z \\t \"Figure Caption Title,1\""),
                        W.ComplexFieldEnd()
                    ]),
                    reduce(vcat, map(1:3) do i
                        [
                            W.Paragraph([
                                W.Run([
                                    W.InlineDrawing(joinpath(@__DIR__, "pumas.png"), 10W.cm, 4W.cm),
                                ])
                            ]),
                            W.Paragraph([
                                W.Run([
                                    W.Text("Figure $i"),
                                ]),
                            ], W.ParagraphProperties(style = "FigureCaptionTitle"))
                        ]
                    end)...
                ])
            ]),
            W.Styles([
                W.Style("FigureCaptionTitle", W.ParagraphProperties(), name = "Figure Caption Title")
            ])
        )

        reftest_docx(doc, "table_of_figures")
    end

    @testset "Length calculations" begin
        @test 4 * W.cm == W.Centimeter(4)
        @test W.cm * 4 == 4 * W.cm
        @test (6 * W.cm) / (2 * W.cm) == 3.0
        @test (6 * W.cm) / 2 == (3 * W.cm)
        @test W.Point(6 * W.cm) / W.Inch(2 * W.cm) ≈ 3.0
        @test 2 * W.inch - 144 * W.pt == 0 * W.inch
        @test -2 * W.inch + 144 * W.pt == 0 * W.inch
    end
end
