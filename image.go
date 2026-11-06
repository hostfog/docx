package docx

import (
	"bytes"
	"errors"
	"fmt"
	"image"
	"image/jpeg"
	"image/png"
)

// ErrUnsupportedImage is returned for image data that is not PNG or JPEG.
var ErrUnsupportedImage = errors.New("docx: unsupported image format")

// emuPerPixel assumes images are 96 DPI, as Word does.
const emuPerPixel = Inch / 96

type picture struct {
	rel           string
	id            int
	name          string
	width, height Length
}

// AddImage appends an inline PNG or JPEG image. If width and height are both
// zero the image is sized from its pixel dimensions at 96 DPI; if only one is
// zero it is computed to keep the aspect ratio.
func (p *Paragraph) AddImage(data []byte, width, height Length) (*Run, error) {
	if width < 0 || height < 0 {
		return nil, fmt.Errorf("docx: negative image size %dx%d", width, height)
	}

	var (
		cfg image.Config
		ext string
		err error
	)
	switch {
	case bytes.HasPrefix(data, []byte("\x89PNG\r\n\x1a\n")):
		cfg, err = png.DecodeConfig(bytes.NewReader(data))
		ext = "png"
	case bytes.HasPrefix(data, []byte("\xff\xd8\xff")):
		cfg, err = jpeg.DecodeConfig(bytes.NewReader(data))
		ext = "jpeg"
	default:
		return nil, ErrUnsupportedImage
	}
	if err != nil {
		return nil, fmt.Errorf("docx: decoding %s: %w", ext, err)
	}
	if cfg.Width <= 0 || cfg.Height <= 0 {
		return nil, fmt.Errorf("docx: %s has no pixels", ext)
	}

	px, py := Length(cfg.Width), Length(cfg.Height)
	switch {
	case width == 0 && height == 0:
		width, height = px*emuPerPixel, py*emuPerPixel
	case height == 0:
		height = width * py / px
	case width == 0:
		width = height * px / py
	}

	d := p.part.doc
	target := d.addMedia(bytes.Clone(data), ext)
	d.drawings++
	r := &Run{kind: runImage, pic: &picture{
		rel:    p.part.addRel(relImage, target, false),
		id:     d.drawings,
		name:   fmt.Sprintf("Picture %d", d.drawings),
		width:  width,
		height: height,
	}}
	p.runs = append(p.runs, r)
	return r, nil
}

func (pic *picture) writeXML(x *xmlWriter) {
	cx, cy := itoa(int64(pic.width)), itoa(int64(pic.height))
	id := itoa(int64(pic.id))
	x.raw(`<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">`,
		`<wp:extent cx="`, cx, `" cy="`, cy, `"/>`,
		`<wp:effectExtent l="0" t="0" r="0" b="0"/>`,
		`<wp:docPr id="`, id, `" name="`, pic.name, `"/>`,
		`<wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="`, nsA, `" noChangeAspect="1"/></wp:cNvGraphicFramePr>`,
		`<a:graphic xmlns:a="`, nsA, `"><a:graphicData uri="`, nsPic, `">`,
		`<pic:pic xmlns:pic="`, nsPic, `">`,
		`<pic:nvPicPr><pic:cNvPr id="`, id, `" name="`, pic.name, `"/><pic:cNvPicPr/></pic:nvPicPr>`,
		`<pic:blipFill><a:blip r:embed="`, pic.rel, `"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>`,
		`<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="`, cx, `" cy="`, cy, `"/></a:xfrm>`,
		`<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>`,
		`</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>`)
}
