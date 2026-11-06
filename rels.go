package docx

import "fmt"

const (
	relBase      = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
	relDocument  = relBase + "officeDocument"
	relStyles    = relBase + "styles"
	relNumbering = relBase + "numbering"
	relSettings  = relBase + "settings"
	relHeader    = relBase + "header"
	relFooter    = relBase + "footer"
	relImage     = relBase + "image"
	relHyperlink = relBase + "hyperlink"
	relCore      = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
	relApp       = relBase + "extended-properties"
)

type rel struct {
	id, typ, target string
	external        bool
}

// part is a package part that can own relationships: the main document or a
// header or footer.
type part struct {
	doc  *Document
	rels []rel
}

func (p *part) addRel(typ, target string, external bool) string {
	for _, r := range p.rels {
		if r.typ == typ && r.target == target {
			return r.id
		}
	}
	id := fmt.Sprintf("rId%d", len(p.rels)+1)
	p.rels = append(p.rels, rel{id: id, typ: typ, target: target, external: external})
	return id
}

func (p *part) writeRels(x *xmlWriter) {
	writeRels(x, p.rels)
}

func writePackageRels(x *xmlWriter) {
	writeRels(x, []rel{
		{id: "rId1", typ: relDocument, target: "word/document.xml"},
		{id: "rId2", typ: relCore, target: "docProps/core.xml"},
		{id: "rId3", typ: relApp, target: "docProps/app.xml"},
	})
}

func writeRels(x *xmlWriter, rels []rel) {
	x.raw(xmlHeader, `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`)
	for _, r := range rels {
		x.raw(`<Relationship Id="`, r.id, `" Type="`, r.typ, `" Target="`)
		x.esc(r.target)
		if r.external {
			x.raw(`" TargetMode="External`)
		}
		x.raw(`"/>`)
	}
	x.raw(`</Relationships>`)
}
