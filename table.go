package docx

import (
	"errors"
	"fmt"
	"strconv"
)

// BorderStyle is the line style of a table border.
type BorderStyle string

const (
	BorderNone   BorderStyle = "none"
	BorderSingle BorderStyle = "single"
	BorderDouble BorderStyle = "double"
	BorderDashed BorderStyle = "dashed"
	BorderDotted BorderStyle = "dotted"
)

// Border is the outer and inner border of a table.
type Border struct {
	Style BorderStyle
	Width float64 // in points, from 0.25 to 12
	Color string  // hex RGB; empty is automatic
}

// VAlign is the vertical alignment of content in a table cell.
type VAlign string

const (
	VAlignTop    VAlign = "top"
	VAlignCenter VAlign = "center"
	VAlignBottom VAlign = "bottom"
)

// Table is a grid of cells. Create one with Body.AddTable.
type Table struct {
	Rows   []*Row
	Widths []Length // one per column
	Border Border
	Align  Alignment

	part *part
}

// Row is a table row.
type Row struct {
	Cells  []*Cell
	Header bool   // repeat this row at the top of each page
	Height Length // minimum height; 0 fits the content
}

// Cell is a table cell. Its Body holds the cell content; an empty cell gets
// an empty paragraph when written.
type Cell struct {
	Body
	Shading string // background as hex RGB, such as "D9E2F3"
	VAlign  VAlign

	span    int  // columns covered, from gridSpan
	vmerge  int  // 0 none, 1 restart, 2 continue
	covered bool // hidden by a horizontal merge
}

// AddTable appends a table with single-line borders and equal column widths
// spanning the text area of the current section. rows and cols below 1 are
// treated as 1.
func (b *Body) AddTable(rows, cols int) *Table {
	rows, cols = max(rows, 1), max(cols, 1)
	t := &Table{
		Border: Border{Style: BorderSingle, Width: 0.5},
		part:   b.part,
	}

	s := b.part.doc.Section()
	text := s.pageWidth() - s.Margins.Left - s.Margins.Right
	for range cols {
		t.Widths = append(t.Widths, text/Length(cols))
	}
	for range rows {
		t.AddRow()
	}
	b.blocks = append(b.blocks, t)
	return t
}

// AddRow appends a row with one cell per column.
func (t *Table) AddRow() *Row {
	r := &Row{}
	for range t.Widths {
		r.Cells = append(r.Cells, &Cell{Body: Body{part: t.part}, span: 1})
	}
	t.Rows = append(t.Rows, r)
	return r
}

// Cell returns the cell at row r and column c, or nil if out of range.
func (t *Table) Cell(r, c int) *Cell {
	if r < 0 || r >= len(t.Rows) || c < 0 || c >= len(t.Rows[r].Cells) {
		return nil
	}
	return t.Rows[r].Cells[c]
}

// Merge merges the rectangle of cells from (r1, c1) to (r2, c2) inclusive
// into its top-left cell. Content of the other cells is discarded.
func (t *Table) Merge(r1, c1, r2, c2 int) error {
	if r1 > r2 || c1 > c2 {
		return fmt.Errorf("docx: invalid merge range (%d,%d)-(%d,%d)", r1, c1, r2, c2)
	}
	for r := r1; r <= r2; r++ {
		for c := c1; c <= c2; c++ {
			cell := t.Cell(r, c)
			if cell == nil {
				return fmt.Errorf("docx: merge cell (%d,%d) out of range", r, c)
			}
			if cell.span != 1 || cell.vmerge != 0 || cell.covered {
				return errors.New("docx: merge overlaps an existing merge")
			}
		}
	}

	for r := r1; r <= r2; r++ {
		for c := c1; c <= c2; c++ {
			t.Rows[r].Cells[c].covered = c > c1
		}
		first := t.Rows[r].Cells[c1]
		first.span = c2 - c1 + 1
		switch {
		case r1 == r2:
		case r == r1:
			first.vmerge = 1
		default:
			first.vmerge = 2
		}
	}
	return nil
}

func (t *Table) writeXML(x *xmlWriter) {
	x.raw(`<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>`)
	if t.Align != "" {
		x.val("w:jc", string(t.Align))
	}
	x.raw("<w:tblBorders>")
	for _, side := range []string{"top", "left", "bottom", "right", "insideH", "insideV"} {
		writeBorder(x, side, t.Border)
	}
	x.raw(`</w:tblBorders><w:tblLayout w:type="fixed"/>`,
		`<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>`,
		`</w:tblPr><w:tblGrid>`)
	for _, w := range t.Widths {
		x.raw(`<w:gridCol w:w="`, itoa(w.twips()), `"/>`)
	}
	x.raw("</w:tblGrid>")

	for i, r := range t.Rows {
		if len(r.Cells) != len(t.Widths) {
			x.fail(fmt.Errorf("docx: table row %d has %d cells, want %d", i, len(r.Cells), len(t.Widths)))
			continue
		}
		r.writeXML(x, t.Widths)
	}
	x.raw("</w:tbl>")
}

func writeBorder(x *xmlWriter, side string, b Border) {
	style := b.Style
	if style == "" {
		style = BorderNone
	}
	color := b.Color
	if color == "" {
		color = "auto"
	} else if !validColor(color) {
		x.fail(fmt.Errorf("docx: invalid border color %q", color))
	}
	size := min(max(int64(b.Width*8+0.5), 2), 96)
	x.raw(`<w:`, side, ` w:val="`, string(style), `" w:sz="`, itoa(size), `" w:space="0" w:color="`)
	x.esc(color)
	x.raw(`"/>`)
}

func (r *Row) writeXML(x *xmlWriter, widths []Length) {
	x.raw("<w:tr>")
	if r.Header || r.Height > 0 {
		x.raw("<w:trPr>")
		if r.Height > 0 {
			x.raw(`<w:trHeight w:val="`, itoa(r.Height.twips()), `" w:hRule="atLeast"/>`)
		}
		if r.Header {
			x.raw("<w:tblHeader/>")
		}
		x.raw("</w:trPr>")
	}
	for i, c := range r.Cells {
		if c.covered {
			continue
		}
		var w int64
		for _, cw := range widths[i : i+c.span] {
			w += cw.twips()
		}
		c.writeXML(x, w)
	}
	x.raw("</w:tr>")
}

func (c *Cell) writeXML(x *xmlWriter, twips int64) {
	x.raw(`<w:tc><w:tcPr><w:tcW w:w="`, itoa(twips), `" w:type="dxa"/>`)
	if c.span > 1 {
		x.raw(`<w:gridSpan w:val="`, strconv.Itoa(c.span), `"/>`)
	}
	switch c.vmerge {
	case 1:
		x.val("w:vMerge", "restart")
	case 2:
		x.raw("<w:vMerge/>")
	}
	if c.Shading != "" {
		if !validColor(c.Shading) {
			x.fail(fmt.Errorf("docx: invalid cell shading %q", c.Shading))
		}
		x.raw(`<w:shd w:val="clear" w:color="auto" w:fill="`)
		x.esc(c.Shading)
		x.raw(`"/>`)
	}
	if c.VAlign != "" {
		x.val("w:vAlign", string(c.VAlign))
	}
	x.raw("</w:tcPr>")

	c.writeContent(x)
	x.raw("</w:tc>")
}
