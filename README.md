# docx

Write Word `.docx` files from Go. Pure Go, standard library only, MIT licensed.

```sh
go get github.com/hostfog/docx
```

## Usage

```go
doc := docx.New()
doc.Properties.Title = "Quarterly report"

doc.AddHeading("Quarterly report", 0)
p := doc.AddParagraph("Revenue grew by ")
p.AddText("12%").Bold = true

list := doc.AddList(docx.Bullets)
list.AddItem("New customers", 0)
list.AddItem("UK", 1)

t := doc.AddTable(3, 2)
t.Rows[0].Header = true
t.Cell(0, 0).AddParagraph("Region")
t.Merge(1, 0, 2, 0)

doc.AddImage(pngBytes, 3*docx.Inch, 0) // height keeps the aspect ratio

footer := doc.Section().Footer().AddParagraph("Page ")
footer.AddPageNumber()

err := doc.Save("report.docx")
```

## Features

- **Text.** Bold, italic, underline, strike, font, size, color and highlight; alignment, spacing, line spacing and indents.
- **Styles.** Title, Heading 1-6, Normal, List Paragraph and Quote use Word's built-in names, so headings appear in the navigation pane and tables of contents.
- **Lists.** Bulleted and numbered, nested up to 9 levels. Each numbered list restarts at 1.
- **Tables.** Column widths, borders, shading, vertical alignment, repeating header rows, and merged cells across rows and columns.
- **Images.** Inline PNG and JPEG, sized in any unit or from pixels at 96 DPI. Identical images are stored once.
- **Pages.** Sections with their own page size, orientation, margins, headers and footers; page breaks; page number and page count fields.
- **Links and metadata.** External hyperlinks; title, author and company in the document properties.
- **Safe output.** Text is escaped and characters that XML does not allow are dropped. Invalid colors or sizes are returned as errors from `Write`, never written.

## License

MIT
