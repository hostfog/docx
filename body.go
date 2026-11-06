package docx

import (
	"fmt"
	"strings"
)

// Built-in paragraph styles. Headings use Word's own style names, so they
// appear in the navigation pane and in generated tables of contents.
const (
	StyleNormal        = "Normal"
	StyleTitle         = "Title"
	StyleHeading1      = "Heading1"
	StyleHeading2      = "Heading2"
	StyleHeading3      = "Heading3"
	StyleHeading4      = "Heading4"
	StyleHeading5      = "Heading5"
	StyleHeading6      = "Heading6"
	StyleListParagraph = "ListParagraph"
	StyleQuote         = "Quote"
)

type block interface {
	writeXML(x *xmlWriter)
}

// Body holds block content: paragraphs, lists, tables and images. It is
// embedded in Document, Cell and HeaderFooter.
type Body struct {
	part   *part
	blocks []block
}

// AddParagraph appends a paragraph. An empty text gives an empty paragraph.
func (b *Body) AddParagraph(text string) *Paragraph {
	p := &Paragraph{part: b.part}
	if text != "" {
		p.AddText(text)
	}
	b.blocks = append(b.blocks, p)
	return p
}

// AddHeading appends a heading. Level 0 is the document title and levels 1
// to 6 are Heading 1 to Heading 6; other values are clamped to that range.
func (b *Body) AddHeading(text string, level int) *Paragraph {
	p := b.AddParagraph(text)
	switch {
	case level <= 0:
		p.Style = StyleTitle
	case level > 6:
		p.Style = StyleHeading6
	default:
		p.Style = fmt.Sprintf("Heading%d", level)
	}
	return p
}

// AddPageBreak appends a paragraph containing a page break.
func (b *Body) AddPageBreak() {
	p := b.AddParagraph("")
	p.runs = append(p.runs, &Run{kind: runPageBreak})
}

// AddImage appends a paragraph containing a PNG or JPEG image. See
// Paragraph.AddImage for how width and height are used.
func (b *Body) AddImage(data []byte, width, height Length) (*Paragraph, error) {
	p := &Paragraph{part: b.part}
	if _, err := p.AddImage(data, width, height); err != nil {
		return nil, err
	}
	b.blocks = append(b.blocks, p)
	return p, nil
}

// AddList starts a bulleted or numbered list. Items are appended to b, in
// order, as List.AddItem is called. Unknown kinds give a bulleted list.
func (b *Body) AddList(kind ListKind) *List {
	if kind != Numbers {
		kind = Bullets
	}
	d := b.part.doc
	l := &List{body: b, kind: kind, id: len(d.lists) + 1}
	d.lists = append(d.lists, l)
	return l
}

// writeContent writes the blocks, ending with a paragraph as Word requires
// in cells, headers and footers.
func (b *Body) writeContent(x *xmlWriter) {
	for _, bl := range b.blocks {
		bl.writeXML(x)
	}
	if len(b.blocks) == 0 {
		x.raw("<w:p/>")
		return
	}
	if _, ok := b.blocks[len(b.blocks)-1].(*Table); ok {
		x.raw("<w:p/>")
	}
}

// Alignment is the horizontal alignment of a paragraph.
type Alignment string

const (
	AlignLeft    Alignment = "left"
	AlignCenter  Alignment = "center"
	AlignRight   Alignment = "right"
	AlignJustify Alignment = "both"
)

// Spacing is the space around and between the lines of a paragraph.
type Spacing struct {
	Before, After Length
	Line          float64 // multiple of single line spacing, such as 1.5; 0 keeps the style's
}

// Indent is the indentation of a paragraph. A negative FirstLine gives a
// hanging indent.
type Indent struct {
	Left, Right, FirstLine Length
}

// Paragraph is a block of runs. Zero-valued fields inherit from the style.
type Paragraph struct {
	Style   string
	Align   Alignment
	Spacing *Spacing
	Indent  *Indent

	part  *part
	runs  []*Run
	list  *List
	level int
}

// AddText appends a run of text. Newlines become line breaks and tabs
// become tab stops.
func (p *Paragraph) AddText(text string) *Run {
	r := &Run{Text: text}
	p.runs = append(p.runs, r)
	return r
}

// AddLink appends a run that links to an external URL.
func (p *Paragraph) AddLink(text, url string) *Run {
	r := &Run{Text: text, link: p.part.addRel(relHyperlink, url, true)}
	p.runs = append(p.runs, r)
	return r
}

// AddPageNumber appends a field showing the current page number.
func (p *Paragraph) AddPageNumber() *Run {
	return p.addField("PAGE")
}

// AddPageCount appends a field showing the total number of pages.
func (p *Paragraph) AddPageCount() *Run {
	return p.addField("NUMPAGES")
}

func (p *Paragraph) addField(instr string) *Run {
	r := &Run{kind: runField, Text: instr}
	p.runs = append(p.runs, r)
	return r
}

func (p *Paragraph) writeXML(x *xmlWriter) {
	x.raw("<w:p>")
	p.writeProps(x)
	for _, r := range p.runs {
		r.writeXML(x)
	}
	x.raw("</w:p>")
}

func (p *Paragraph) writeProps(x *xmlWriter) {
	style := p.Style
	if style == "" && p.list != nil {
		style = StyleListParagraph
	}
	if style == "" && p.Spacing == nil && p.Indent == nil && p.Align == "" {
		return
	}

	x.raw("<w:pPr>")
	if style != "" {
		x.val("w:pStyle", style)
	}
	if p.list != nil {
		x.raw("<w:numPr>")
		x.valInt("w:ilvl", int64(p.level))
		x.valInt("w:numId", int64(p.list.id))
		x.raw("</w:numPr>")
	}
	if s := p.Spacing; s != nil {
		x.raw(`<w:spacing w:before="`, itoa(s.Before.twips()), `" w:after="`, itoa(s.After.twips()), `"`)
		if s.Line > 0 {
			x.raw(` w:line="`, itoa(int64(s.Line*240+0.5)), `" w:lineRule="auto"`)
		}
		x.raw("/>")
	}
	if in := p.Indent; in != nil {
		x.raw(`<w:ind w:left="`, itoa(in.Left.twips()), `" w:right="`, itoa(in.Right.twips()), `"`)
		if in.FirstLine < 0 {
			x.raw(` w:hanging="`, itoa((-in.FirstLine).twips()), `"`)
		} else if in.FirstLine > 0 {
			x.raw(` w:firstLine="`, itoa(in.FirstLine.twips()), `"`)
		}
		x.raw("/>")
	}
	if p.Align != "" {
		x.val("w:jc", string(p.Align))
	}
	x.raw("</w:pPr>")
}

// Highlight is a text highlight color from Word's fixed palette.
type Highlight string

const (
	HighlightYellow    Highlight = "yellow"
	HighlightGreen     Highlight = "green"
	HighlightCyan      Highlight = "cyan"
	HighlightMagenta   Highlight = "magenta"
	HighlightBlue      Highlight = "blue"
	HighlightRed       Highlight = "red"
	HighlightLightGray Highlight = "lightGray"
)

type runKind int

const (
	runText runKind = iota
	runPageBreak
	runField
	runImage
)

// Run is a span of text with the same formatting. Zero-valued fields
// inherit from the paragraph style.
type Run struct {
	Text      string
	Bold      bool
	Italic    bool
	Underline bool
	Strike    bool
	Font      string
	Size      float64 // in points
	Color     string  // hex RGB, such as "1F4E79"
	Highlight Highlight

	kind runKind
	link string
	pic  *picture
}

func (r *Run) writeXML(x *xmlWriter) {
	if r.link != "" {
		x.raw(`<w:hyperlink r:id="`, r.link, `" w:history="1">`)
		defer x.raw("</w:hyperlink>")
	}
	if r.kind == runField {
		x.raw(`<w:fldSimple w:instr=" `)
		x.esc(r.Text)
		x.raw(` ">`)
		defer x.raw("</w:fldSimple>")
	}

	x.raw("<w:r>")
	r.writeProps(x)
	switch r.kind {
	case runPageBreak:
		x.raw(`<w:br w:type="page"/>`)
	case runField:
		x.raw("<w:t>1</w:t>")
	case runImage:
		r.pic.writeXML(x)
	default:
		writeText(x, r.Text)
	}
	x.raw("</w:r>")
}

func writeText(x *xmlWriter, s string) {
	s = strings.ReplaceAll(s, "\r\n", "\n")
	for s != "" {
		i := strings.IndexAny(s, "\n\r\t")
		if i < 0 {
			i = len(s)
		}
		if i > 0 {
			x.raw(`<w:t xml:space="preserve">`)
			x.esc(s[:i])
			x.raw("</w:t>")
		}
		if i == len(s) {
			break
		}
		if s[i] == '\t' {
			x.raw("<w:tab/>")
		} else {
			x.raw("<w:br/>")
		}
		s = s[i+1:]
	}
}

func (r *Run) writeProps(x *xmlWriter) {
	if r.link == "" && !r.Bold && !r.Italic && !r.Underline && !r.Strike &&
		r.Font == "" && r.Size == 0 && r.Color == "" && r.Highlight == "" {
		return
	}

	x.raw("<w:rPr>")
	if r.link != "" {
		x.val("w:rStyle", "Hyperlink")
	}
	if r.Font != "" {
		x.raw(`<w:rFonts w:ascii="`)
		x.esc(r.Font)
		x.raw(`" w:hAnsi="`)
		x.esc(r.Font)
		x.raw(`" w:cs="`)
		x.esc(r.Font)
		x.raw(`"/>`)
	}
	if r.Bold {
		x.raw("<w:b/>")
	}
	if r.Italic {
		x.raw("<w:i/>")
	}
	if r.Strike {
		x.raw("<w:strike/>")
	}
	if r.Color != "" {
		if !validColor(r.Color) {
			x.fail(fmt.Errorf("docx: invalid run color %q", r.Color))
		}
		x.val("w:color", r.Color)
	}
	if r.Size != 0 {
		hp := int64(r.Size*2 + 0.5)
		if hp < 2 || hp > 3276 {
			x.fail(fmt.Errorf("docx: font size %gpt out of range", r.Size))
		}
		x.valInt("w:sz", hp)
		x.valInt("w:szCs", hp)
	}
	if r.Highlight != "" {
		x.val("w:highlight", string(r.Highlight))
	}
	if r.Underline {
		x.val("w:u", "single")
	}
	x.raw("</w:rPr>")
}

func validColor(s string) bool {
	if len(s) != 6 {
		return false
	}
	for _, c := range s {
		if !strings.ContainsRune("0123456789abcdefABCDEF", c) {
			return false
		}
	}
	return true
}

// ListKind selects bullets or numbers for a list.
type ListKind int

const (
	Bullets ListKind = iota
	Numbers
)

// List is a bulleted or numbered list. Each numbered list starts again at 1.
type List struct {
	body *Body
	kind ListKind
	id   int
}

// AddItem appends a list item at the given nesting level, from 0 to 8.
// Levels outside that range are clamped.
func (l *List) AddItem(text string, level int) *Paragraph {
	p := l.body.AddParagraph(text)
	p.list = l
	p.level = min(max(level, 0), 8)
	return p
}
