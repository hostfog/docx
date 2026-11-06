package docx_test

import (
	"bytes"
	"fmt"
	"log"

	"github.com/hostfog/docx"
)

func Example() {
	doc := docx.New()
	doc.Properties.Title = "Quarterly report"
	doc.Properties.Author = "HostFog Ltd"

	doc.AddHeading("Quarterly report", 0)
	p := doc.AddParagraph("Revenue grew by ")
	r := p.AddText("12%")
	r.Bold = true
	p.AddText(" on the previous quarter.")

	var buf bytes.Buffer
	if err := doc.Write(&buf); err != nil {
		log.Fatal(err)
	}
	fmt.Println(buf.Len() > 0)
	// Output: true
}

func ExampleBody_AddList() {
	doc := docx.New()
	steps := doc.AddList(docx.Numbers)
	steps.AddItem("Download the installer", 0)
	steps.AddItem("Run it", 0)
	steps.AddItem("Accept the licence", 1)
}

func ExampleTable_Merge() {
	doc := docx.New()
	t := doc.AddTable(3, 3)
	t.Rows[0].Header = true
	t.Cell(0, 0).AddParagraph("Region")
	t.Cell(0, 1).Shading = "D9E2F3"

	if err := t.Merge(1, 0, 2, 0); err != nil {
		log.Fatal(err)
	}
}

func ExampleSection_Footer() {
	doc := docx.New()
	p := doc.Section().Footer().AddParagraph("Page ")
	p.Align = docx.AlignCenter
	p.AddPageNumber()
	p.AddText(" of ")
	p.AddPageCount()
}
