package main

import (
	"fmt"

	"github.com/ChotiwatMajor/go-docx"
)

func main() {
	f := docx.NewFile()
	sp := docx.SectionProperties{
		PageMargin: &docx.PageMargin{
			Top:    docx.INCH,
			Bottom: docx.INCH,
			Left:   docx.INCH * 0.5,
			Right:  docx.INCH * 0.5,
		},
	}
	f.Document.Body.Content = append(f.Document.Body.Content, &sp)
	// add new table
	t := f.AddTable()
	t.Properties.Layout.Type = docx.TABLE_LAYOUT_TYPE_FIXED
	t.Properties.Borders.InsideH.Val = docx.TABLE_BORDER_VAL_DASHED
	t.Properties.Borders.InsideV.Color = "ff0000"
	t.Properties.Borders.InsideV.Size = 15
	t.Properties.CellMargin.Start.TableCellMarginAttributes.Width = 100
	t.Properties.CellMargin.End.TableCellMarginAttributes.Width = 100
	t.Properties.CellMargin.Top.TableCellMarginAttributes.Width = 100
	t.Properties.CellMargin.Bottom.TableCellMarginAttributes.Width = 100

	for i := 0; i < 3; i++ {
		row := t.AddRow()
		for i := 0; i < 3; i++ {
			c := row.AddCell(2 * docx.CM)
			c.Properties.VerticalAlignment(docx.CELL_VERTICAL_ALIGN_CENTER)
			if i == 0 {
				c.Paragraph.Properties.Justification(docx.JUSTIFY_CENTER)
				c.Paragraph.AddText(fmt.Sprint(i)).Size(14)
			} else {
				c.Paragraph.AddText("Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer").Size(14)
			}

		}
	}
	row := t.AddRow()
	c := row.AddCell(2 * docx.CM)
	c.AddText("One column")
	row = t.AddRow()
	c = row.AddCell(2 * docx.CM)
	c.AddText("Two columns")
	c = row.AddCell(2 * docx.CM)
	c.AddText("Two columns")
	// Adding new paragraph
	para := f.AddParagraph()
	// Adding new page
	para.AddNewPage()
	// Setting Justification/Alignment of paragraph
	para.Properties.Justification(docx.JUSTIFY_END)
	para.AddText("Paragraph starting from right")
	para = f.AddParagraph()
	para.Properties.Justification(docx.JUSTIFY_START)
	para.AddText("Paragraph starting from left")
	para.AddNewLine()
	para = f.AddParagraph()
	para.Properties.Justification(docx.JUSTIFY_BOTH)
	para.AddText("Paragraph both justified very long line with").Size(18)
	para.AddNewLine()
	para = f.AddParagraph()
	para.Properties.Justification(docx.JUSTIFY_CENTER)
	para.AddText("Paragraph with centered text").Size(18)
	para.AddNewLine()
	para = f.AddParagraph()
	para.Properties.Justification(docx.JUSTIFY_DISTRIBUTE)
	para.Properties.Indentation(2*docx.CM, 1*docx.INCH)
	para.AddText("Paragraph with indentation and distrubution").Size(12)
	para.AddNewLine()
	para = f.AddParagraph()
	para.AddText("test font size bold").Size(22).Bold(true)
	para.AddNewLine()
	para.AddText("test color").Color("808080")
	para.AddNewLine()
	para.AddText("test font size and color").Size(22).Color("ff0f0f")
	nextPara := f.AddParagraph()
	nextPara.AddNewPage()
	nextPara.AddLink("google", `http://google.com`)

	f.Save("./test.docx")
}
