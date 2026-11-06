package docx

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"errors"
	"image"
	"image/color"
	"image/png"
	"io"
	"path/filepath"
	"strings"
	"testing"
)

// unzip writes d and returns its parts, checking that each XML part is well
// formed.
func unzip(t *testing.T, d *Document) map[string]string {
	t.Helper()
	var buf bytes.Buffer
	if err := d.Write(&buf); err != nil {
		t.Fatal(err)
	}
	zr, err := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		t.Fatal(err)
	}
	parts := make(map[string]string)
	for _, f := range zr.File {
		rc, err := f.Open()
		if err != nil {
			t.Fatal(err)
		}
		b, err := io.ReadAll(rc)
		rc.Close()
		if err != nil {
			t.Fatal(err)
		}
		parts[f.Name] = string(b)
		if strings.HasSuffix(f.Name, ".xml") || strings.HasSuffix(f.Name, ".rels") {
			if err := wellFormed(b); err != nil {
				t.Fatalf("%s: %v", f.Name, err)
			}
		}
	}
	return parts
}

func wellFormed(b []byte) error {
	dec := xml.NewDecoder(bytes.NewReader(b))
	for {
		if _, err := dec.Token(); err == io.EOF {
			return nil
		} else if err != nil {
			return err
		}
	}
}

func contains(t *testing.T, part, s string, want ...string) {
	t.Helper()
	for _, w := range want {
		if !strings.Contains(s, w) {
			t.Errorf("%s does not contain %s", part, w)
		}
	}
}

func testPNG(t *testing.T, w, h int) []byte {
	t.Helper()
	img := image.NewRGBA(image.Rect(0, 0, w, h))
	img.Set(0, 0, color.RGBA{255, 0, 0, 255})
	var buf bytes.Buffer
	if err := png.Encode(&buf, img); err != nil {
		t.Fatal(err)
	}
	return buf.Bytes()
}

func TestPackage(t *testing.T) {
	parts := unzip(t, New())
	for _, name := range []string{
		"[Content_Types].xml", "_rels/.rels", "docProps/core.xml", "docProps/app.xml",
		"word/document.xml", "word/styles.xml", "word/numbering.xml", "word/settings.xml",
		"word/_rels/document.xml.rels",
	} {
		if _, ok := parts[name]; !ok {
			t.Errorf("missing %s", name)
		}
	}
	contains(t, "document", parts["word/document.xml"],
		`<w:pgSz w:w="11906" w:h="16838"/>`, `w:top="1440"`)
	contains(t, "styles", parts["word/styles.xml"],
		`w:styleId="Heading1"><w:name w:val="heading 1"/>`, `<w:outlineLvl w:val="5"/>`,
		`w:styleId="Title"`, `w:styleId="ListParagraph"`, `w:styleId="Quote"`)
}

func TestRuns(t *testing.T) {
	d := New()
	p := d.AddParagraph("")
	r := p.AddText("a < b & \"c\"\x00\x1f\nnext\tcol")
	r.Bold, r.Italic, r.Underline, r.Strike = true, true, true, true
	r.Font, r.Size, r.Color, r.Highlight = "Arial", 10.5, "FF0000", HighlightYellow
	p.Align = AlignCenter
	p.Spacing = &Spacing{Before: 6 * Pt, After: 0, Line: 1.5}
	p.Indent = &Indent{Left: Inch, FirstLine: -Inch / 4}

	doc := unzip(t, d)["word/document.xml"]
	contains(t, "document", doc,
		`<w:t xml:space="preserve">a &lt; b &amp; &quot;c&quot;</w:t><w:br/>`,
		`<w:t xml:space="preserve">next</w:t><w:tab/>`,
		`<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:b/><w:i/><w:strike/>`,
		`<w:color w:val="FF0000"/><w:sz w:val="21"/><w:szCs w:val="21"/><w:highlight w:val="yellow"/><w:u w:val="single"/>`,
		`<w:spacing w:before="120" w:after="0" w:line="360" w:lineRule="auto"/>`,
		`<w:ind w:left="1440" w:right="0" w:hanging="360"/><w:jc w:val="center"/>`)
	if strings.ContainsAny(doc, "\x00\x1f") {
		t.Error("invalid XML characters were not stripped")
	}
}

func TestHeadings(t *testing.T) {
	d := New()
	d.AddHeading("T", 0)
	d.AddHeading("H2", 2)
	d.AddHeading("H9", 9)
	contains(t, "document", unzip(t, d)["word/document.xml"],
		`<w:pStyle w:val="Title"/>`, `<w:pStyle w:val="Heading2"/>`, `<w:pStyle w:val="Heading6"/>`)
}

func TestLists(t *testing.T) {
	d := New()
	b := d.AddList(Bullets)
	b.AddItem("one", 0)
	b.AddItem("nested", 1)
	n := d.AddList(Numbers)
	n.AddItem("first", 0)
	n.AddItem("deep", 20)

	parts := unzip(t, d)
	contains(t, "document", parts["word/document.xml"],
		`<w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr>`,
		`<w:numPr><w:ilvl w:val="8"/><w:numId w:val="2"/></w:numPr>`)
	contains(t, "numbering", parts["word/numbering.xml"],
		`<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>`,
		`<w:num w:numId="2"><w:abstractNumId w:val="1"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/>`,
		`<w:numFmt w:val="bullet"/>`, `<w:lvlText w:val="%1."/>`)
}

func TestTable(t *testing.T) {
	d := New()
	tb := d.AddTable(3, 3)
	tb.Rows[0].Header = true
	tb.Cell(0, 0).AddParagraph("head")
	tb.Cell(1, 1).Shading = "D9E2F3"
	tb.Cell(1, 1).VAlign = VAlignCenter
	if err := tb.Merge(1, 0, 2, 1); err != nil {
		t.Fatal(err)
	}
	if err := tb.Merge(0, 0, 1, 0); err == nil {
		t.Error("overlapping merge succeeded")
	}
	if err := tb.Merge(0, 0, 5, 0); err == nil {
		t.Error("out of range merge succeeded")
	}
	if tb.Cell(3, 0) != nil {
		t.Error("Cell out of range is not nil")
	}

	doc := unzip(t, d)["word/document.xml"]
	contains(t, "document", doc,
		`<w:gridCol w:w="3009"/>`,
		`<w:trPr><w:tblHeader/></w:trPr>`,
		`<w:tcW w:w="6018" w:type="dxa"/><w:gridSpan w:val="2"/><w:vMerge w:val="restart"/>`,
		`<w:gridSpan w:val="2"/><w:vMerge/></w:tcPr>`,
		`<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>`,
		`</w:tbl><w:p/>`)
	if got := strings.Count(doc, "<w:tc>"); got != 7 {
		t.Errorf("got %d cells, want 7", got)
	}
}

func TestImage(t *testing.T) {
	d := New()
	data := testPNG(t, 200, 100)
	if _, err := d.AddImage(data, 0, 0); err != nil {
		t.Fatal(err)
	}
	if _, err := d.AddImage(data, 2*Inch, 0); err != nil {
		t.Fatal(err)
	}
	if _, err := d.AddImage([]byte("GIF89a"), 0, 0); !errors.Is(err, ErrUnsupportedImage) {
		t.Errorf("got %v, want ErrUnsupportedImage", err)
	}

	parts := unzip(t, d)
	if _, ok := parts["word/media/image1.png"]; !ok {
		t.Error("missing word/media/image1.png")
	}
	if _, ok := parts["word/media/image2.png"]; ok {
		t.Error("identical image stored twice")
	}
	contains(t, "document", parts["word/document.xml"],
		`<wp:extent cx="1905000" cy="952500"/>`,
		`<wp:extent cx="1828800" cy="914400"/>`,
		`<wp:docPr id="2" name="Picture 2"/>`,
		`<a:blip r:embed="rId4"/>`)
	contains(t, "rels", parts["word/_rels/document.xml.rels"],
		`Id="rId4" Type="`+relImage+`" Target="media/image1.png"`)
	contains(t, "content types", parts["[Content_Types].xml"],
		`<Default Extension="png" ContentType="image/png"/>`)
}

func TestLinkAndBreak(t *testing.T) {
	d := New()
	d.AddParagraph("see ").AddLink("site", "https://example.com/?a=1&b=2")
	d.AddPageBreak()

	parts := unzip(t, d)
	contains(t, "document", parts["word/document.xml"],
		`<w:hyperlink r:id="rId4" w:history="1"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/>`,
		`<w:br w:type="page"/>`)
	contains(t, "rels", parts["word/_rels/document.xml.rels"],
		`Target="https://example.com/?a=1&amp;b=2" TargetMode="External"`)
}

func TestSections(t *testing.T) {
	d := New()
	d.Section().Header().AddParagraph("Report")
	f := d.Section().Footer().AddParagraph("Page ")
	f.AddPageNumber()
	f.AddText(" of ")
	f.AddPageCount()
	d.Section().Footer().AddImage(testPNG(t, 4, 4), 0, 0)
	d.AddParagraph("portrait")

	s := d.AddSection()
	s.PageSize = Letter
	s.Landscape = true
	s.Margins.Left = 2 * Cm
	d.AddParagraph("landscape")

	parts := unzip(t, d)
	doc := parts["word/document.xml"]
	contains(t, "document", doc,
		`<w:p><w:pPr><w:sectPr><w:headerReference w:type="default" r:id="rId4"/><w:footerReference w:type="default" r:id="rId5"/><w:pgSz w:w="11906" w:h="16838"/>`,
		`<w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>`,
		`w:left="1134"`)
	if strings.Count(doc, "<w:sectPr>") != 2 {
		t.Error("want 2 sections")
	}
	contains(t, "header", parts["word/header1.xml"], `<w:hdr `, "Report")
	contains(t, "footer", parts["word/footer1.xml"],
		`<w:fldSimple w:instr=" PAGE ">`, `<w:fldSimple w:instr=" NUMPAGES ">`, `r:embed="rId1"`)
	contains(t, "footer rels", parts["word/_rels/footer1.xml.rels"], `Target="media/image1.png"`)
	contains(t, "content types", parts["[Content_Types].xml"],
		`<Override PartName="/word/header1.xml" ContentType="`+ctBase+`header+xml"/>`,
		`<Override PartName="/word/footer1.xml" ContentType="`+ctBase+`footer+xml"/>`)
}

func TestProperties(t *testing.T) {
	d := New()
	d.Properties = Properties{Title: "Q3 <Report>", Author: "Param", Company: "HostFog Ltd"}
	parts := unzip(t, d)
	contains(t, "core", parts["docProps/core.xml"],
		`<dc:title>Q3 &lt;Report&gt;</dc:title>`, `<dc:creator>Param</dc:creator>`, `<dcterms:created`)
	contains(t, "app", parts["docProps/app.xml"], `<Company>HostFog Ltd</Company>`)
}

func TestInvalidValues(t *testing.T) {
	tests := []func(*Document){
		func(d *Document) { d.AddParagraph("").AddText("x").Color = "red" },
		func(d *Document) { d.AddParagraph("").AddText("x").Size = 5000 },
		func(d *Document) { d.AddTable(1, 1).Cell(0, 0).Shading = "#FFF" },
		func(d *Document) { d.AddTable(1, 2).Widths = []Length{Inch} },
		func(d *Document) { d.Section().PageSize = PageSize{} },
	}
	for i, fn := range tests {
		d := New()
		fn(d)
		if err := d.Write(io.Discard); err == nil {
			t.Errorf("case %d: Write succeeded", i)
		}
	}
}

func TestSave(t *testing.T) {
	name := filepath.Join(t.TempDir(), "out.docx")
	if err := New().Save(name); err != nil {
		t.Fatal(err)
	}
	if err := New().Save(filepath.Join(t.TempDir(), "missing", "out.docx")); err == nil {
		t.Error("Save to missing directory succeeded")
	}
}
