package docx

import "encoding/xml"

const (
	BLACK                                   = "000000"
	TABLE_BORDER_VAL_SINGLE                 = "single"                 // a single line
	TABLE_BORDER_VAL_DASHDOTSTROKED         = "dashDotStroked"         // a line with a series of alternating thin and thick strokes
	TABLE_BORDER_VAL_DASHED                 = "dashed"                 // a dashed line
	TABLE_BORDER_VAL_DASHSMALLGAP           = "dashSmallGap"           // a dashed line with small gaps
	TABLE_BORDER_VAL_DOTDASH                = "dotDash"                // a line with alternating dots and dashes
	TABLE_BORDER_VAL_DOTDOTDASH             = "dotDotDash"             // a line with a repeating dot - dot - dash sequence
	TABLE_BORDER_VAL_DOTTED                 = "dotted"                 // a dotted line
	TABLE_BORDER_VAL_DOUBLE                 = "double"                 // a double line
	TABLE_BORDER_VAL_DOUBLEWAVE             = "doubleWave"             // a double wavy line
	TABLE_BORDER_VAL_INSET                  = "inset"                  // an inset set of lines
	TABLE_BORDER_VAL_NIL                    = "nil"                    // no border
	TABLE_BORDER_VAL_NONE                   = "none"                   // no border
	TABLE_BORDER_VAL_OUTSET                 = "outset"                 // an outset set of lines
	TABLE_BORDER_VAL_THICK                  = "thick"                  // a single line
	TABLE_BORDER_VAL_THICKTHINLARGEGAP      = "thickThinLargeGap"      // a thick line contained within a thin line with a large-sized intermediate gap
	TABLE_BORDER_VAL_THICKTHINMEDIUMGAP     = "thickThinMediumGap"     // a thick line contained within a thin line with a medium-sized intermediate gap
	TABLE_BORDER_VAL_THICKTHINSMALLGAP      = "thickThinSmallGap"      // a thick line contained within a thin line with a small-sized intermediate gap
	TABLE_BORDER_VAL_THINTHICKLARGEGAP      = "thinThickLargeGap"      // a thin line contained within a thick line with a large-sized intermediate gap
	TABLE_BORDER_VAL_THINTHICKMEDIUMGAP     = "thinThickMediumGap"     // a thin line contained within a thick line with a medium-sized intermediate gap
	TABLE_BORDER_VAL_THINTHICKSMALLGAP      = "thinThickSmallGap"      // a thin line contained within a thick line with a small-sized intermediate gap
	TABLE_BORDER_VAL_THINTHICKTHINLARGEGAP  = "thinThickThinLargeGap"  // a thin-thick-thin line with a large gap
	TABLE_BORDER_VAL_THINTHICKTHINMEDIUMGAP = "thinThickThinMediumGap" // a thin-thick-thin line with a medium gap
	TABLE_BORDER_VAL_THINTHICKTHINSMALLGAP  = "thinThickThinSmallGap"  // a thin-thick-thin line with a small gap
	TABLE_BORDER_VAL_THREEDEEMBOSS          = "threeDEmboss"           // a three-staged gradient line, getting darker towards the paragraph
	TABLE_BORDER_VAL_THREEDEENGRAVE         = "threeDEngrave"          // a three-staged gradient like, getting darker away from the paragraph
	TABLE_BORDER_VAL_TRIPLE                 = "triple"                 // a triple line
	TABLE_BORDER_VAL_WAVE                   = "wave"                   // a wavy line
	TABLE_LAYOUT_TYPE_FIXED                 = "fixed"                  // fixed cell width
	TABLE_LAYOUT_TYPE_AUTO                  = "auto"                   // auto cell width

	CELL_VERTICAL_ALIGN_TOP    = "top"    // align content to the top of cell vertically
	CELL_VERTICAL_ALIGN_BOTTOM = "bottom" // align content to the bottom of cell vertically
	CELL_VERTICAL_ALIGN_CENTER = "center" // align content to the center of cell vertically
)

type Table struct {
	XMLName    xml.Name `xml:"w:tbl"`
	Data       []interface{}
	grid       *TableGrid
	Properties *TableProperties
}

type TableProperties struct {
	XMLName    xml.Name `xml:"w:tblPr"`
	Width      *TableWidth
	Borders    *TableBorders
	Layout     *TableLayout
	CellMargin *TableCellMargin
}

type TableBorders struct {
	XMLName xml.Name `xml:"w:tblBorders"`
	Top     *TableBordersTop
	Bottom  *TableBordersBottom
	Start   *TableBordersStart
	End     *TableBordersEnd
	InsideV *TableBordersInsideV
	InsideH *TableBordersInsideH
}

type TableLayout struct {
	XMLName xml.Name `xml:"w:tblLayout"`
	Type    string   `xml:"w:type,attr"`
}

type TableCellMargin struct {
	XMLName xml.Name `xml:"w:tblCellMar"`
	Top     *TableCellMarginTop
	Bottom  *TableCellMarginBottom
	Start   *TableCellMarginStart
	End     *TableCellMarginEnd
}

// TableBordersTop specifies the border displayed at the top of a table.
type TableBordersTop struct {
	XMLName xml.Name `xml:"w:top"`
	TableBordersAttributes
}

// TableBordersBottom specifies the border displayed at the bottom of a table.
type TableBordersBottom struct {
	XMLName xml.Name `xml:"w:bottom"`
	TableBordersAttributes
}

// TableBordersStart specifies the border displayed to the left for left-to-right tables and right for right-to-left tables.
type TableBordersStart struct {
	XMLName xml.Name `xml:"w:start"`
	TableBordersAttributes
}

// TableBordersEnd specifies the border displayed on the right for left-to-right tables and left for right-to-left tables
type TableBordersEnd struct {
	XMLName xml.Name `xml:"w:end"`
	TableBordersAttributes
}

// TableBordersInsideH specifies the border displayed on all inside horizontal table cell borders.
type TableBordersInsideH struct {
	XMLName xml.Name `xml:"w:insideH"`
	TableBordersAttributes
}

// TableBordersInsideV specifies the border displayed on all inside vertical table cell borders.
type TableBordersInsideV struct {
	XMLName xml.Name `xml:"w:insideV"`
	TableBordersAttributes
}

type TableCellMarginTop struct {
	XMLName xml.Name `xml:"w:top"`
	TableCellMarginAttributes
}

type TableCellMarginBottom struct {
	XMLName xml.Name `xml:"w:bottom"`
	TableCellMarginAttributes
}

type TableCellMarginStart struct {
	XMLName xml.Name `xml:"w:start"`
	TableCellMarginAttributes
}

type TableCellMarginEnd struct {
	XMLName xml.Name `xml:"w:end"`
	TableCellMarginAttributes
}

type TableBordersAttributes struct {
	Color  string `xml:"w:color,attr"`
	Shadow string `xml:"w:shadow,attr"`
	Space  int    `xml:"w:space,attr"`
	Size   int    `xml:"w:sz,attr"`
	Val    string `xml:"w:val,attr"`
}

type TableCellMarginAttributes struct {
	Width int    `xml:"w:w,attr"`
	Type  string `xml:"w:type,attr"`
}

type TableWidth struct {
	XMLName xml.Name `xml:"w:tblW"`
	Width   int      `xml:"w:w,attr"`
	Type    string   `xml:"w:type,attr"`
}

type TableGrid struct {
	XMLName xml.Name `xml:"w:tblGrid"`
	Data    []interface{}
}

type tableGridCol struct {
	XMLName xml.Name `xml:"w:gridCol"`
	Width   int      `xml:"w:w,attr"`
	Type    string   `xml:"w:type,attr"`
}

type Row struct {
	XMLName xml.Name `xml:"w:tr"`
	Data    []interface{}
	table   *Table
}

type RowProperties struct {
	XMLName xml.Name `xml:"w:trPr"`
	Data    []interface{}
	width   *TableWidth
}

type Cell struct {
	XMLName    xml.Name `xml:"w:tc"`
	Data       []interface{}
	Properties *CellProperties
	Paragraph  *Paragraph
	row        *Row
}

type CellProperties struct {
	XMLName xml.Name `xml:"w:tcPr"`
	Data    []interface{}
}

type CellVerticalAlignment struct {
	XMLName xml.Name `xml:"w:vAlign"`
	Val     string   `xml:"w:val,attr"`
}

type cellWidth struct {
	XMLName xml.Name `xml:"w:tcW"`
	Width   int      `xml:"w:w,attr"`
	Type    string   `xml:"w:type,attr"`
}

// AddTable is adding table to document
func (f *File) AddTable() *Table {
	grid := &TableGrid{
		Data: make([]interface{}, 0),
	}
	tblW := &TableWidth{}
	// default attributes of table: single black borders
	attrs := TableBordersAttributes{Color: BLACK, Val: TABLE_BORDER_VAL_SINGLE}
	tblBorders := &TableBorders{
		Top:     &TableBordersTop{TableBordersAttributes: attrs},
		Bottom:  &TableBordersBottom{TableBordersAttributes: attrs},
		End:     &TableBordersEnd{TableBordersAttributes: attrs},
		Start:   &TableBordersStart{TableBordersAttributes: attrs},
		InsideV: &TableBordersInsideV{TableBordersAttributes: attrs},
		InsideH: &TableBordersInsideH{TableBordersAttributes: attrs},
	}
	tblLayout := &TableLayout{}
	mAttrs := TableCellMarginAttributes{Type: "dxa"}
	tblCellMargin := &TableCellMargin{
		Top:    &TableCellMarginTop{TableCellMarginAttributes: mAttrs},
		Bottom: &TableCellMarginBottom{TableCellMarginAttributes: mAttrs},
		End:    &TableCellMarginEnd{TableCellMarginAttributes: mAttrs},
		Start:  &TableCellMarginStart{TableCellMarginAttributes: mAttrs},
	}
	props := &TableProperties{
		Width:      tblW,
		Borders:    tblBorders,
		Layout:     tblLayout,
		CellMargin: tblCellMargin,
	}
	t := &Table{
		Data:       make([]interface{}, 0),
		grid:       grid,
		Properties: props,
	}
	t.Data = append(t.Data, props)
	t.Data = append(t.Data, grid)

	f.Document.Body.Content = append(f.Document.Body.Content, t)
	return t
}

// AddRow is adding row to table
func (t *Table) AddRow() *Row {
	r := &Row{
		table: t,
	}
	t.Data = append(t.Data, r)

	return r
}

// AddCell is adding cell to row, takes width in twips
func (r *Row) AddCell(width int) *Cell {
	tcProps := &CellProperties{}
	props := &ParagraphProperties{}
	p := &Paragraph{
		Data:       make([]interface{}, 0),
		Properties: props,
	}
	c := &Cell{
		row:        r,
		Properties: tcProps,
		Paragraph:  p,
	}
	p.Data = append(p.Data, props)
	r.Data = append(r.Data, c)
	gridCol := &tableGridCol{
		Width: width,
		Type:  "dxa",
	}
	tcW := &cellWidth{
		Width: width,
		Type:  "dxa",
	}
	tcProps.Data = append(tcProps.Data, tcW)
	// Do not add more width than we have columns
	if len(r.Data) > len(r.table.grid.Data) {
		r.table.grid.Data = append(r.table.grid.Data, gridCol)
		r.table.Properties.Width.Width += width
	}

	return c
}

// AddText is adding text to cell
func (c *Cell) AddText(text string) *Run {
	p := &Paragraph{
		Data: make([]interface{}, 0),
	}
	txt := &Text{
		Text: text,
	}
	run := &Run{
		Text:          txt,
		RunProperties: &RunProperties{},
	}

	p.Data = append(p.Data, run)
	c.Data = append(c.Data, p)

	return run
}

// CELL_VERTICAL_ALIGN_TOP align content to the top of cell vertically
// CELL_VERTICAL_ALIGN_CENTER align content to the bottom of cell vertically
// CELL_VERTICAL_ALIGN_BOTTM align content to the center of cell vertically
func (cp *CellProperties) VerticalAlignment(align string) {
	cp.Data = append(cp.Data, &CellVerticalAlignment{Val: align})
}
