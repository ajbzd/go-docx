package docx

import "encoding/xml"

const (
	HYPERLINK_STYLE = "a1"
)

type RunProperties struct {
	XMLName  xml.Name  `xml:"w:rPr"`
	Color    *Color    `xml:"w:color,omitempty"`
	Size     *Size     `xml:"w:sz,omitempty"`
	Bold     *Bold     `xml:"w:b,omitempty"`
	Font     *Font     `xml:"w:rFonts,omitempty"`
	RunStyle *RunStyle `xml:"w:rStyle,omitempty"`
}

type RunStyle struct {
	XMLName xml.Name `xml:"w:rStyle"`
	Val     string   `xml:"w:val,attr"`
}

type Color struct {
	XMLName xml.Name `xml:"w:color"`
	Val     string   `xml:"w:val,attr"`
}

type Size struct {
	XMLName xml.Name `xml:"w:sz"`
	Val     int      `xml:"w:val,attr"`
}

type Bold struct {
	XMLName xml.Name `xml:"w:b"`
	Val     bool     `xml:"w:val,attr"`
}

type Font struct {
	XMLName xml.Name `xml:"w:rFonts"`
	Ascii   string   `xml:"w:ascii,attr"`
	HAnsi   string   `xml:"w:hAnsi,attr"`
	Cs      string   `xml:"w:cs,attr"`
}
