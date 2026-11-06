package docx

import (
	"bytes"
	"strconv"
)

const xmlHeader = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n"

const (
	nsW   = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
	nsR   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
	nsWP  = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
	nsA   = "http://schemas.openxmlformats.org/drawingml/2006/main"
	nsPic = "http://schemas.openxmlformats.org/drawingml/2006/picture"
)

// xmlWriter builds one XML part. The first error is kept and later writes
// still happen, so callers check err once at the end.
type xmlWriter struct {
	buf bytes.Buffer
	err error
}

func (x *xmlWriter) fail(err error) {
	if x.err == nil {
		x.err = err
	}
}

func (x *xmlWriter) raw(ss ...string) {
	for _, s := range ss {
		x.buf.WriteString(s)
	}
}

// esc writes s escaped for use in text or a double-quoted attribute.
// Characters that XML 1.0 does not allow are dropped.
func (x *xmlWriter) esc(s string) {
	for _, r := range s {
		switch {
		case r == '&':
			x.buf.WriteString("&amp;")
		case r == '<':
			x.buf.WriteString("&lt;")
		case r == '>':
			x.buf.WriteString("&gt;")
		case r == '"':
			x.buf.WriteString("&quot;")
		case r == '\t' || r == '\n' || r == '\r':
			x.buf.WriteRune(r)
		case !validXMLChar(r):
		default:
			x.buf.WriteRune(r)
		}
	}
}

// val writes <name w:val="v"/>.
func (x *xmlWriter) val(name, v string) {
	x.raw("<", name, ` w:val="`)
	x.esc(v)
	x.raw(`"/>`)
}

func (x *xmlWriter) valInt(name string, v int64) {
	x.raw("<", name, ` w:val="`, strconv.FormatInt(v, 10), `"/>`)
}

func validXMLChar(r rune) bool {
	switch {
	case r >= 0x20 && r <= 0xD7FF:
		return true
	case r >= 0xE000 && r <= 0xFFFD:
		return true
	case r >= 0x10000 && r <= 0x10FFFF:
		return true
	}
	return r == '\t' || r == '\n' || r == '\r'
}

func itoa(v int64) string { return strconv.FormatInt(v, 10) }
