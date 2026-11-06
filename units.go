package docx

import "math"

// Length is a distance in English Metric Units, the native unit of DrawingML.
// Build one from the unit constants, such as 2*docx.Cm or 12*docx.Pt.
type Length int64

const (
	EMU  Length = 1
	Twip Length = 635
	Pt   Length = 12700
	Mm   Length = 36000
	Cm   Length = 360000
	Inch Length = 914400
)

// twips returns l in twentieths of a point, the unit used by most of
// WordprocessingML.
func (l Length) twips() int64 {
	return int64(math.Round(float64(l) / float64(Twip)))
}

// PageSize is the width and height of a page in portrait orientation.
type PageSize struct {
	Width, Height Length
}

var (
	A4     = PageSize{210 * Mm, 297 * Mm}
	A5     = PageSize{148 * Mm, 210 * Mm}
	Letter = PageSize{Inch * 17 / 2, 11 * Inch}
	Legal  = PageSize{Inch * 17 / 2, 14 * Inch}
)
