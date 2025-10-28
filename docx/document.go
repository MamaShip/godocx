package docx

import (
	"encoding/xml"

	"github.com/MamaShip/godocx/internal"
	"github.com/MamaShip/godocx/wml/stypes"
)

var docAttrs = []xml.Attr{
	{Name: xml.Name{Local: "xmlns:w"}, Value: "http://schemas.openxmlformats.org/wordprocessingml/2006/main"},
	{Name: xml.Name{Local: "xmlns:o"}, Value: "urn:schemas-microsoft-com:office:office"},
	{Name: xml.Name{Local: "xmlns:r"}, Value: "http://schemas.openxmlformats.org/officeDocument/2006/relationships"},
	{Name: xml.Name{Local: "xmlns:v"}, Value: "urn:schemas-microsoft-com:vml"},
	{Name: xml.Name{Local: "xmlns:w10"}, Value: "urn:schemas-microsoft-com:office:word"},
	{Name: xml.Name{Local: "xmlns:wp"}, Value: "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"},
	{Name: xml.Name{Local: "xmlns:wps"}, Value: "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"},
	{Name: xml.Name{Local: "xmlns:wpg"}, Value: "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"},
	{Name: xml.Name{Local: "xmlns:mc"}, Value: "http://schemas.openxmlformats.org/markup-compatibility/2006"},
	{Name: xml.Name{Local: "xmlns:wp14"}, Value: "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"},
	{Name: xml.Name{Local: "xmlns:w14"}, Value: "http://schemas.microsoft.com/office/word/2010/wordml"},
	{Name: xml.Name{Local: "xmlns:w15"}, Value: "http://schemas.microsoft.com/office/word/2012/wordml"},
	{Name: xml.Name{Local: "mc:Ignorable"}, Value: "w14 wp14 w15"},
}

// This element specifies the contents of a main document part in a WordprocessingML document.
type Document struct {
	// Reference to the RootDoc
	Root *RootDoc

	// Elements
	Background *Background
	Body       *Body

	// Non elements - helper fields
	DocRels      Relationships // DocRels represents relationships specific to the document.
	RID          int
	relativePath string
}

// IncRelationID increments the relation ID of the document and returns the new ID.
// This method is used to generate unique IDs for relationships within the document.
func (doc *Document) IncRelationID() int {
	doc.RID += 1
	return doc.RID
}

// MarshalXML implements the xml.Marshaler interface for the Document type.
func (doc Document) MarshalXML(e *xml.Encoder, start xml.StartElement) (err error) {
	start.Name.Local = "w:document"

	start.Attr = append(start.Attr, docAttrs...)

	err = e.EncodeToken(start)
	if err != nil {
		return err
	}

	if doc.Background != nil {
		if err = doc.Background.MarshalXML(e, xml.StartElement{}); err != nil {
			return err
		}
	}

	if doc.Body != nil {
		bodyElement := xml.StartElement{Name: xml.Name{Local: "w:body"}}
		if err = e.EncodeElement(doc.Body, bodyElement); err != nil {
			return err
		}
	}

	return e.EncodeToken(xml.EndElement{Name: start.Name})
}

func (d *Document) UnmarshalXML(decoder *xml.Decoder, start xml.StartElement) (err error) {

	for {
		currentToken, err := decoder.Token()
		if err != nil {
			return err
		}

		switch elem := currentToken.(type) {
		case xml.StartElement:
			switch elem.Name.Local {
			case "body":
				body := NewBody(d.Root)
				if err := decoder.DecodeElement(body, &elem); err != nil {
					return err
				}
				d.Body = body
			case "background":
				bg := NewBackground()
				if err := decoder.DecodeElement(bg, &elem); err != nil {
					return err
				}
				d.Background = bg
			default:
				if err = decoder.Skip(); err != nil {
					return err
				}
			}
		case xml.EndElement:
			return nil
		}
	}

}

// AddPageBreak adds a page break to the document by inserting a paragraph containing only a page break.
//
// Returns:
//   - *Paragraph: A pointer to the newly created Paragraph object containing the page break.
//
// Example:
//
//	document := godocx.NewDocument()
//	para := document.AddPageBreak()
func (rd *RootDoc) AddPageBreak() *Paragraph {
	p := rd.AddEmptyParagraph()
	p.AddRun().AddBreak(internal.ToPtr(stypes.BreakTypePage))

	return p
}

// AddHorizontalLine adds a simple horizontal line (divider) to the document.
//
// This creates an empty paragraph with a bottom border styled as a single line.
// The default line is a single solid line with automatic color and standard width (0.75pt).
//
// Returns:
//   - *Paragraph: A pointer to the newly created Paragraph object with a horizontal line.
//
// Example:
//
//	document := godocx.NewDocument()
//	document.AddHorizontalLine()
func (rd *RootDoc) AddHorizontalLine() *Paragraph {
	p := rd.AddEmptyParagraph()
	p.BottomBorder(stypes.BorderStyleSingle, 6, "auto")
	return p
}

// AddDoubleHorizontalLine adds a double horizontal line (divider) to the document.
//
// This creates an empty paragraph with a bottom border styled as a double line.
//
// Returns:
//   - *Paragraph: A pointer to the newly created Paragraph object with a double horizontal line.
//
// Example:
//
//	document := godocx.NewDocument()
//	document.AddDoubleHorizontalLine()
func (rd *RootDoc) AddDoubleHorizontalLine() *Paragraph {
	p := rd.AddEmptyParagraph()
	p.BottomBorder(stypes.BorderStyleDouble, 6, "auto")
	return p
}

// AddThickHorizontalLine adds a thick horizontal line (divider) to the document.
//
// This creates an empty paragraph with a bottom border styled as a thick line.
//
// Returns:
//   - *Paragraph: A pointer to the newly created Paragraph object with a thick horizontal line.
//
// Example:
//
//	document := godocx.NewDocument()
//	document.AddThickHorizontalLine()
func (rd *RootDoc) AddThickHorizontalLine() *Paragraph {
	p := rd.AddEmptyParagraph()
	p.BottomBorder(stypes.BorderStyleThick, 12, "auto")
	return p
}

// AddDashedHorizontalLine adds a dashed horizontal line (divider) to the document.
//
// This creates an empty paragraph with a bottom border styled as a dashed line.
//
// Returns:
//   - *Paragraph: A pointer to the newly created Paragraph object with a dashed horizontal line.
//
// Example:
//
//	document := godocx.NewDocument()
//	document.AddDashedHorizontalLine()
func (rd *RootDoc) AddDashedHorizontalLine() *Paragraph {
	p := rd.AddEmptyParagraph()
	p.BottomBorder(stypes.BorderStyleDashed, 6, "auto")
	return p
}

// AddCustomHorizontalLine adds a custom horizontal line (divider) to the document with specified properties.
//
// This allows full customization of the horizontal line's appearance.
//
// Parameters:
//   - style: The border style from stypes.BorderStyle (e.g., BorderStyleSingle, BorderStyleDouble, BorderStyleWave).
//   - size: The border width in eighths of a point (e.g., 6 = 0.75pt, 12 = 1.5pt, 24 = 3pt).
//   - color: The border color in hex format (e.g., "FF0000" for red, "0000FF" for blue) or "auto" for automatic color.
//
// Returns:
//   - *Paragraph: A pointer to the newly created Paragraph object with a custom horizontal line.
//
// Example:
//
//	document := godocx.NewDocument()
//	// Add a red wavy line at 1.5pt thickness
//	document.AddCustomHorizontalLine(stypes.BorderStyleWave, 12, "FF0000")
func (rd *RootDoc) AddCustomHorizontalLine(style stypes.BorderStyle, size int, color string) *Paragraph {
	p := rd.AddEmptyParagraph()
	p.BottomBorder(style, size, color)
	return p
}
