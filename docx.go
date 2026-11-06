// Package docx writes Office Open XML (.docx) documents that open in
// Microsoft Word, LibreOffice and Google Docs.
//
// A document is built from paragraphs, lists, tables and images added to its
// body, then written with Save or Write:
//
//	doc := docx.New()
//	doc.AddHeading("Report", 1)
//	doc.AddParagraph("Hello, world.")
//	err := doc.Save("report.docx")
package docx

import (
	"archive/zip"
	"bufio"
	"crypto/sha256"
	"fmt"
	"io"
	"os"
	"time"
)

// Properties are the document metadata shown in Word's File > Info panel.
type Properties struct {
	Title       string
	Subject     string
	Author      string
	Keywords    string
	Description string
	Company     string
	Created     time.Time // defaults to the time of writing
}

// Document is a Word document. Create one with New.
type Document struct {
	Body
	Properties Properties

	main     *part
	sections []*Section
	hdrftr   []*HeaderFooter
	lists    []*List
	media    []media
	mediaIDs map[[32]byte]string
	drawings int
}

type media struct {
	name string
	data []byte
}

// New returns an empty A4 portrait document with 1 inch margins.
func New() *Document {
	d := &Document{mediaIDs: make(map[[32]byte]string)}
	d.main = &part{doc: d}
	d.main.addRel(relStyles, "styles.xml", false)
	d.main.addRel(relNumbering, "numbering.xml", false)
	d.main.addRel(relSettings, "settings.xml", false)
	d.Body = Body{part: d.main}
	d.sections = []*Section{newSection(d, A4)}
	return d
}

// Section returns the current (last) section of the document.
func (d *Document) Section() *Section {
	return d.sections[len(d.sections)-1]
}

// AddSection ends the current section and starts a new one on the next page.
// The new section copies the page setup of the previous one and inherits its
// headers and footers unless it sets its own.
func (d *Document) AddSection() *Section {
	prev := d.Section()
	d.blocks = append(d.blocks, sectionBreak{prev})

	s := newSection(d, prev.PageSize)
	s.Landscape = prev.Landscape
	s.Margins = prev.Margins
	d.sections = append(d.sections, s)
	return s
}

// Save writes the document to the named file.
func (d *Document) Save(name string) error {
	f, err := os.Create(name)
	if err != nil {
		return err
	}
	if err := d.Write(f); err != nil {
		f.Close()
		return err
	}
	return f.Close()
}

// Write writes the document as a .docx package to w.
func (d *Document) Write(w io.Writer) error {
	bw := bufio.NewWriter(w)
	zw := zip.NewWriter(bw)

	parts := []zipPart{
		{"[Content_Types].xml", d.writeContentTypes},
		{"_rels/.rels", writePackageRels},
		{"docProps/core.xml", d.writeCore},
		{"docProps/app.xml", d.writeApp},
		{"word/document.xml", d.writeDocument},
		{"word/styles.xml", writeStyles},
		{"word/numbering.xml", d.writeNumbering},
		{"word/settings.xml", writeSettings},
		{"word/_rels/document.xml.rels", d.main.writeRels},
	}
	for _, hf := range d.hdrftr {
		parts = append(parts,
			zipPart{"word/" + hf.name, hf.writeXML},
			zipPart{"word/_rels/" + hf.name + ".rels", hf.own.writeRels},
		)
	}

	for _, p := range parts {
		x := &xmlWriter{}
		p.fn(x)
		if x.err != nil {
			return x.err
		}
		if err := writeZipFile(zw, p.name, x.buf.Bytes()); err != nil {
			return err
		}
	}
	for _, m := range d.media {
		if err := writeZipFile(zw, "word/"+m.name, m.data); err != nil {
			return err
		}
	}

	if err := zw.Close(); err != nil {
		return err
	}
	return bw.Flush()
}

type zipPart struct {
	name string
	fn   func(*xmlWriter)
}

func writeZipFile(zw *zip.Writer, name string, data []byte) error {
	f, err := zw.Create(name)
	if err != nil {
		return fmt.Errorf("docx: %s: %w", name, err)
	}
	_, err = f.Write(data)
	return err
}

func (d *Document) addMedia(data []byte, ext string) string {
	sum := sha256.Sum256(data)
	if name, ok := d.mediaIDs[sum]; ok {
		return name
	}
	name := fmt.Sprintf("media/image%d.%s", len(d.media)+1, ext)
	d.media = append(d.media, media{name: name, data: data})
	d.mediaIDs[sum] = name
	return name
}
