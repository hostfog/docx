package docx

import "fmt"

// Margins are the page margins of a section. Header and Footer are the
// distances from the page edge to the header and footer.
type Margins struct {
	Top, Right, Bottom, Left Length
	Header, Footer           Length
}

// Section is a run of pages that share a page setup, headers and footers.
type Section struct {
	PageSize  PageSize
	Landscape bool
	Margins   Margins

	doc    *Document
	header *HeaderFooter
	footer *HeaderFooter
}

func newSection(d *Document, size PageSize) *Section {
	return &Section{
		PageSize: size,
		Margins: Margins{
			Top: Inch, Right: Inch, Bottom: Inch, Left: Inch,
			Header: Inch / 2, Footer: Inch / 2,
		},
		doc: d,
	}
}

// Header returns the header shown on every page of the section, creating it
// on first use.
func (s *Section) Header() *HeaderFooter {
	if s.header == nil {
		s.header = s.doc.newHeaderFooter(false)
	}
	return s.header
}

// Footer returns the footer shown on every page of the section, creating it
// on first use.
func (s *Section) Footer() *HeaderFooter {
	if s.footer == nil {
		s.footer = s.doc.newHeaderFooter(true)
	}
	return s.footer
}

func (s *Section) pageWidth() Length {
	if s.Landscape {
		return max(s.PageSize.Width, s.PageSize.Height)
	}
	return min(s.PageSize.Width, s.PageSize.Height)
}

func (s *Section) pageHeight() Length {
	if s.Landscape {
		return min(s.PageSize.Width, s.PageSize.Height)
	}
	return max(s.PageSize.Width, s.PageSize.Height)
}

func (s *Section) writeXML(x *xmlWriter) {
	x.raw("<w:sectPr>")
	if s.header != nil {
		x.raw(`<w:headerReference w:type="default" r:id="`, s.header.rel, `"/>`)
	}
	if s.footer != nil {
		x.raw(`<w:footerReference w:type="default" r:id="`, s.footer.rel, `"/>`)
	}

	w, h := s.pageWidth(), s.pageHeight()
	if w <= 0 || h <= 0 {
		x.fail(fmt.Errorf("docx: invalid page size %dx%d", w, h))
	}
	x.raw(`<w:pgSz w:w="`, itoa(w.twips()), `" w:h="`, itoa(h.twips()), `"`)
	if s.Landscape {
		x.raw(` w:orient="landscape"`)
	}
	m := s.Margins
	x.raw(`/><w:pgMar w:top="`, itoa(m.Top.twips()),
		`" w:right="`, itoa(m.Right.twips()),
		`" w:bottom="`, itoa(m.Bottom.twips()),
		`" w:left="`, itoa(m.Left.twips()),
		`" w:header="`, itoa(m.Header.twips()),
		`" w:footer="`, itoa(m.Footer.twips()),
		`" w:gutter="0"/><w:cols w:space="720"/><w:docGrid w:linePitch="360"/></w:sectPr>`)
}

// sectionBreak ends a section in the body.
type sectionBreak struct{ s *Section }

func (b sectionBreak) writeXML(x *xmlWriter) {
	x.raw("<w:p><w:pPr>")
	b.s.writeXML(x)
	x.raw("</w:pPr></w:p>")
}

// HeaderFooter is the content of a page header or footer.
type HeaderFooter struct {
	Body

	tag  string // "hdr" or "ftr"
	name string // part name, such as "header1.xml"
	rel  string // relationship ID in the main document
	own  part
}

func (d *Document) newHeaderFooter(footer bool) *HeaderFooter {
	kind, tag, typ := "header", "hdr", relHeader
	if footer {
		kind, tag, typ = "footer", "ftr", relFooter
	}
	n := 1
	for _, hf := range d.hdrftr {
		if hf.tag == tag {
			n++
		}
	}

	hf := &HeaderFooter{tag: tag, name: fmt.Sprintf("%s%d.xml", kind, n)}
	hf.own.doc = d
	hf.part = &hf.own
	hf.rel = d.main.addRel(typ, hf.name, false)
	d.hdrftr = append(d.hdrftr, hf)
	return hf
}

func (hf *HeaderFooter) writeXML(x *xmlWriter) {
	x.raw(xmlHeader, `<w:`, hf.tag, ` xmlns:w="`, nsW, `" xmlns:r="`, nsR, `" xmlns:wp="`, nsWP, `">`)
	hf.writeContent(x)
	x.raw(`</w:`, hf.tag, `>`)
}
