package docx

import "encoding/xml"

const (
	XMLNS_W = `http://schemas.openxmlformats.org/wordprocessingml/2006/main`
	XMLNS_R = `http://schemas.openxmlformats.org/officeDocument/2006/relationships`
)

type Document struct {
	XMLName xml.Name `xml:"w:document"`
	XMLW    string   `xml:"xmlns:w,attr"`
	XMLR    string   `xml:"xmlns:r,attr"`
	Body    *Body
}

type Body struct {
	XMLName xml.Name `xml:"w:body"`
	Content []interface{}
}

type SectionProperties struct {
	XMLName    xml.Name `xml:"w:sectPr"`
	PageMargin *PageMargin
}

type PageMargin struct {
	XMLName xml.Name `xml:"w:pgMar"`
	Header  int      `xml:"w:header,attr"`
	Bottom  int      `xml:"w:bottom,attr"`
	Top     int      `xml:"w:top,attr"`
	Right   int      `xml:"w:right,attr"`
	Left    int      `xml:"w:left,attr"`
	Footer  int      `xml:"w:footer,attr"`
	Gutter  int      `xml:"w:gutter,attr"`
}
