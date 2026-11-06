package docx

import (
	"strings"
	"time"
)

const ctBase = "application/vnd.openxmlformats-officedocument.wordprocessingml."

func (d *Document) writeContentTypes(x *xmlWriter) {
	x.raw(xmlHeader,
		`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">`,
		`<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>`,
		`<Default Extension="xml" ContentType="application/xml"/>`,
		`<Default Extension="png" ContentType="image/png"/>`,
		`<Default Extension="jpeg" ContentType="image/jpeg"/>`,
		`<Override PartName="/word/document.xml" ContentType="`, ctBase, `document.main+xml"/>`,
		`<Override PartName="/word/styles.xml" ContentType="`, ctBase, `styles+xml"/>`,
		`<Override PartName="/word/numbering.xml" ContentType="`, ctBase, `numbering+xml"/>`,
		`<Override PartName="/word/settings.xml" ContentType="`, ctBase, `settings+xml"/>`)
	for _, hf := range d.hdrftr {
		kind := "header"
		if hf.tag == "ftr" {
			kind = "footer"
		}
		x.raw(`<Override PartName="/word/`, hf.name, `" ContentType="`, ctBase, kind, `+xml"/>`)
	}
	x.raw(`<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>`,
		`<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>`,
		`</Types>`)
}

func (d *Document) writeCore(x *xmlWriter) {
	p := d.Properties
	created := p.Created
	if created.IsZero() {
		created = time.Now()
	}
	ts := created.UTC().Format("2006-01-02T15:04:05Z")

	x.raw(xmlHeader, `<cp:coreProperties`,
		` xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"`,
		` xmlns:dc="http://purl.org/dc/elements/1.1/"`,
		` xmlns:dcterms="http://purl.org/dc/terms/"`,
		` xmlns:dcmitype="http://purl.org/dc/dcmitype/"`,
		` xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">`)
	for _, e := range []struct{ tag, v string }{
		{"dc:title", p.Title},
		{"dc:subject", p.Subject},
		{"dc:creator", p.Author},
		{"cp:keywords", p.Keywords},
		{"dc:description", p.Description},
		{"cp:lastModifiedBy", p.Author},
	} {
		if e.v != "" {
			x.raw("<", e.tag, ">")
			x.esc(e.v)
			x.raw("</", e.tag, ">")
		}
	}
	x.raw(`<dcterms:created xsi:type="dcterms:W3CDTF">`, ts, `</dcterms:created>`,
		`<dcterms:modified xsi:type="dcterms:W3CDTF">`, ts, `</dcterms:modified>`,
		`</cp:coreProperties>`)
}

func (d *Document) writeApp(x *xmlWriter) {
	x.raw(xmlHeader,
		`<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"`,
		` xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">`,
		`<Application>hostfog/docx</Application>`)
	if d.Properties.Company != "" {
		x.raw("<Company>")
		x.esc(d.Properties.Company)
		x.raw("</Company>")
	}
	x.raw(`</Properties>`)
}

func (d *Document) writeDocument(x *xmlWriter) {
	x.raw(xmlHeader, `<w:document xmlns:w="`, nsW, `" xmlns:r="`, nsR, `" xmlns:wp="`, nsWP, `"><w:body>`)
	d.writeContent(x)
	d.Section().writeXML(x)
	x.raw(`</w:body></w:document>`)
}

func writeSettings(x *xmlWriter) {
	x.raw(xmlHeader, `<w:settings xmlns:w="`, nsW, `">`,
		`<w:defaultTabStop w:val="720"/>`,
		`<w:characterSpacingControl w:val="doNotCompress"/>`,
		`<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>`,
		`</w:settings>`)
}

func (d *Document) writeNumbering(x *xmlWriter) {
	x.raw(xmlHeader, `<w:numbering xmlns:w="`, nsW, `">`)
	writeAbstractNum(x, Bullets)
	writeAbstractNum(x, Numbers)
	for _, l := range d.lists {
		x.raw(`<w:num w:numId="`, itoa(int64(l.id)), `">`)
		x.valInt("w:abstractNumId", int64(l.kind))
		if l.kind == Numbers {
			for lvl := range 9 {
				x.raw(`<w:lvlOverride w:ilvl="`, itoa(int64(lvl)), `"><w:startOverride w:val="1"/></w:lvlOverride>`)
			}
		}
		x.raw(`</w:num>`)
	}
	x.raw(`</w:numbering>`)
}

func writeAbstractNum(x *xmlWriter, kind ListKind) {
	bullets := []string{"•", "◦", "▪"}
	formats := []string{"decimal", "lowerLetter", "lowerRoman"}

	x.raw(`<w:abstractNum w:abstractNumId="`, itoa(int64(kind)), `">`)
	x.val("w:multiLevelType", "hybridMultilevel")
	for lvl := range 9 {
		x.raw(`<w:lvl w:ilvl="`, itoa(int64(lvl)), `">`)
		x.val("w:start", "1")
		if kind == Bullets {
			x.val("w:numFmt", "bullet")
			x.val("w:lvlText", bullets[lvl%3])
		} else {
			x.val("w:numFmt", formats[lvl%3])
			x.val("w:lvlText", "%"+itoa(int64(lvl+1))+".")
		}
		x.val("w:lvlJc", "left")
		x.raw(`<w:pPr><w:ind w:left="`, itoa(int64(720*(lvl+1))), `" w:hanging="360"/></w:pPr>`)
		x.raw(`</w:lvl>`)
	}
	x.raw(`</w:abstractNum>`)
}

func writeStyles(x *xmlWriter) {
	x.raw(xmlHeader, `<w:styles xmlns:w="`, nsW, `">`, strings.Join(styles, ""), `</w:styles>`)
}

var styles = []string{
	`<w:docDefaults><w:rPrDefault><w:rPr>`,
	`<w:rFonts w:ascii="Calibri" w:eastAsia="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>`,
	`<w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="ar-SA"/>`,
	`</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr>`,
	`<w:spacing w:after="160" w:line="259" w:lineRule="auto"/>`,
	`</w:pPr></w:pPrDefault></w:docDefaults>`,

	`<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>`,

	`<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>`,
	`<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:contextualSpacing/></w:pPr>`,
	`<w:rPr><w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light" w:cs="Calibri Light"/><w:spacing w:val="-10"/><w:kern w:val="28"/><w:sz w:val="56"/><w:szCs w:val="56"/></w:rPr></w:style>`,

	heading(1, 240, 32, "2F5496", false),
	heading(2, 40, 26, "2F5496", false),
	heading(3, 40, 24, "1F3763", false),
	heading(4, 40, 22, "2F5496", true),
	heading(5, 40, 22, "2F5496", false),
	heading(6, 40, 22, "1F3763", false),

	`<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="34"/><w:qFormat/>`,
	`<w:pPr><w:ind w:left="720"/><w:contextualSpacing/></w:pPr></w:style>`,

	`<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="29"/><w:qFormat/>`,
	`<w:pPr><w:spacing w:before="200" w:after="160"/><w:ind w:left="864" w:right="864"/><w:jc w:val="center"/></w:pPr>`,
	`<w:rPr><w:i/><w:iCs/><w:color w:val="404040"/></w:rPr></w:style>`,

	`<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style>`,

	`<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/>`,
	`<w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>`,

	`<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/>`,
	`<w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/>`,
	`<w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>`,

	`<w:style w:type="numbering" w:default="1" w:styleId="NoList"><w:name w:val="No List"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/></w:style>`,
}

func heading(level int, before, size int64, color string, italic bool) string {
	n := itoa(int64(level))
	var b strings.Builder
	b.WriteString(`<w:style w:type="paragraph" w:styleId="Heading` + n + `"><w:name w:val="heading ` + n + `"/>`)
	b.WriteString(`<w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/>`)
	b.WriteString(`<w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="` + itoa(before) + `" w:after="0"/>`)
	b.WriteString(`<w:outlineLvl w:val="` + itoa(int64(level-1)) + `"/></w:pPr><w:rPr>`)
	b.WriteString(`<w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light" w:cs="Calibri Light"/>`)
	if italic {
		b.WriteString(`<w:i/><w:iCs/>`)
	}
	b.WriteString(`<w:color w:val="` + color + `"/><w:sz w:val="` + itoa(size) + `"/><w:szCs w:val="` + itoa(size) + `"/></w:rPr></w:style>`)
	return b.String()
}
