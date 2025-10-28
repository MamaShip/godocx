package docx

import (
	"testing"

	"github.com/MamaShip/godocx/internal"
	"github.com/MamaShip/godocx/wml/ctypes"
	"github.com/MamaShip/godocx/wml/stypes"
	"github.com/stretchr/testify/assert"
)

// Note: setupRootDoc is defined in docx_test.go

// assertTightSpacing verifies that paragraph has tight spacing settings to avoid empty line effect
func assertTightSpacing(t *testing.T, p *Paragraph) {
	t.Helper()
	assert.NotNil(t, p.ct.Property.Spacing, "Paragraph should have spacing")
	assert.NotNil(t, p.ct.Property.Spacing.Before, "Spacing Before should be set")
	assert.Equal(t, uint64(0), *p.ct.Property.Spacing.Before, "Before spacing should be 0")
	assert.NotNil(t, p.ct.Property.Spacing.After, "Spacing After should be set")
	assert.Equal(t, uint64(0), *p.ct.Property.Spacing.After, "After spacing should be 0")
	assert.NotNil(t, p.ct.Property.Spacing.Line, "Line spacing should be set")
	assert.Equal(t, 20, *p.ct.Property.Spacing.Line, "Line spacing should be 20 (1pt)")
	assert.NotNil(t, p.ct.Property.Spacing.LineRule, "LineRule should be set")
	assert.Equal(t, stypes.LineSpacingRuleExact, *p.ct.Property.Spacing.LineRule, "LineRule should be exact")
}

// TestAddHorizontalLine tests the AddHorizontalLine method
func TestAddHorizontalLine(t *testing.T) {
	doc := setupRootDoc(t)
	p := doc.AddHorizontalLine()

	assert.NotNil(t, p, "AddHorizontalLine should return a non-nil Paragraph")
	assert.NotNil(t, p.ct.Property, "Paragraph should have properties")
	assert.NotNil(t, p.ct.Property.Border, "Paragraph should have border")
	assert.NotNil(t, p.ct.Property.Border.Bottom, "Paragraph should have bottom border")
	assert.Equal(t, stypes.BorderStyleSingle, p.ct.Property.Border.Bottom.Val, "Border style should be single")
	assert.Equal(t, 6, *p.ct.Property.Border.Bottom.Size, "Border size should be 6")
	assert.Equal(t, "auto", *p.ct.Property.Border.Bottom.Color, "Border color should be auto")

	// Verify tight spacing to avoid empty line effect
	assertTightSpacing(t, p)
}

// TestAddDoubleHorizontalLine tests the AddDoubleHorizontalLine method
func TestAddDoubleHorizontalLine(t *testing.T) {
	doc := setupRootDoc(t)
	p := doc.AddDoubleHorizontalLine()

	assert.NotNil(t, p, "AddDoubleHorizontalLine should return a non-nil Paragraph")
	assert.NotNil(t, p.ct.Property.Border.Bottom, "Paragraph should have bottom border")
	assert.Equal(t, stypes.BorderStyleDouble, p.ct.Property.Border.Bottom.Val, "Border style should be double")
	assert.Equal(t, 6, *p.ct.Property.Border.Bottom.Size, "Border size should be 6")

	// Verify tight spacing to avoid empty line effect
	assertTightSpacing(t, p)
}

// TestAddThickHorizontalLine tests the AddThickHorizontalLine method
func TestAddThickHorizontalLine(t *testing.T) {
	doc := setupRootDoc(t)
	p := doc.AddThickHorizontalLine()

	assert.NotNil(t, p, "AddThickHorizontalLine should return a non-nil Paragraph")
	assert.NotNil(t, p.ct.Property.Border.Bottom, "Paragraph should have bottom border")
	assert.Equal(t, stypes.BorderStyleThick, p.ct.Property.Border.Bottom.Val, "Border style should be thick")
	assert.Equal(t, 12, *p.ct.Property.Border.Bottom.Size, "Border size should be 12")

	// Verify tight spacing to avoid empty line effect
	assertTightSpacing(t, p)
}

// TestAddDashedHorizontalLine tests the AddDashedHorizontalLine method
func TestAddDashedHorizontalLine(t *testing.T) {
	doc := setupRootDoc(t)
	p := doc.AddDashedHorizontalLine()

	assert.NotNil(t, p, "AddDashedHorizontalLine should return a non-nil Paragraph")
	assert.NotNil(t, p.ct.Property.Border.Bottom, "Paragraph should have bottom border")
	assert.Equal(t, stypes.BorderStyleDashed, p.ct.Property.Border.Bottom.Val, "Border style should be dashed")
	assert.Equal(t, 6, *p.ct.Property.Border.Bottom.Size, "Border size should be 6")

	// Verify tight spacing to avoid empty line effect
	assertTightSpacing(t, p)
}

// TestAddCustomHorizontalLine tests the AddCustomHorizontalLine method
func TestAddCustomHorizontalLine(t *testing.T) {
	tests := []struct {
		name          string
		style         stypes.BorderStyle
		size          int
		color         string
		expectedStyle stypes.BorderStyle
		expectedSize  int
		expectedColor string
	}{
		{
			name:          "Red wavy line",
			style:         stypes.BorderStyleWave,
			size:          12,
			color:         "FF0000",
			expectedStyle: stypes.BorderStyleWave,
			expectedSize:  12,
			expectedColor: "FF0000",
		},
		{
			name:          "Blue thick single line",
			style:         stypes.BorderStyleSingle,
			size:          24,
			color:         "0000FF",
			expectedStyle: stypes.BorderStyleSingle,
			expectedSize:  24,
			expectedColor: "0000FF",
		},
		{
			name:          "Green dotted line",
			style:         stypes.BorderStyleDotted,
			size:          8,
			color:         "00FF00",
			expectedStyle: stypes.BorderStyleDotted,
			expectedSize:  8,
			expectedColor: "00FF00",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			doc := setupRootDoc(t)
			p := doc.AddCustomHorizontalLine(tt.style, tt.size, tt.color)

			assert.NotNil(t, p, "AddCustomHorizontalLine should return a non-nil Paragraph")
			assert.NotNil(t, p.ct.Property.Border.Bottom, "Paragraph should have bottom border")
			assert.Equal(t, tt.expectedStyle, p.ct.Property.Border.Bottom.Val, "Border style should match")
			assert.Equal(t, tt.expectedSize, *p.ct.Property.Border.Bottom.Size, "Border size should match")
			assert.Equal(t, tt.expectedColor, *p.ct.Property.Border.Bottom.Color, "Border color should match")

			// Verify tight spacing to avoid empty line effect
			assertTightSpacing(t, p)
		})
	}
}

// TestParagraphLineSpacing tests the LineSpacing method
func TestParagraphLineSpacing(t *testing.T) {
	doc := setupRootDoc(t)
	p := doc.AddParagraph("Test paragraph")

	result := p.LineSpacing(240, stypes.LineSpacingRuleExact)

	// Verify method returns paragraph for chaining
	assert.Equal(t, p, result, "LineSpacing should return the paragraph for chaining")

	// Verify spacing properties
	assert.NotNil(t, p.ct.Property, "Paragraph should have properties")
	assert.NotNil(t, p.ct.Property.Spacing, "Paragraph should have spacing")
	assert.NotNil(t, p.ct.Property.Spacing.Line, "Line spacing should be set")
	assert.Equal(t, 240, *p.ct.Property.Spacing.Line, "Line spacing should be 240")
	assert.NotNil(t, p.ct.Property.Spacing.LineRule, "LineRule should be set")
	assert.Equal(t, stypes.LineSpacingRuleExact, *p.ct.Property.Spacing.LineRule, "LineRule should be exact")
}

// TestParagraphBottomBorder tests the BottomBorder method
func TestParagraphBottomBorder(t *testing.T) {
	doc := setupRootDoc(t)
	p := doc.AddParagraph("Test paragraph")
	p.BottomBorder(stypes.BorderStyleSingle, 6, "auto")

	assert.NotNil(t, p.ct.Property, "Paragraph should have properties")
	assert.NotNil(t, p.ct.Property.Border, "Paragraph should have border")
	assert.NotNil(t, p.ct.Property.Border.Bottom, "Paragraph should have bottom border")
	assert.Equal(t, stypes.BorderStyleSingle, p.ct.Property.Border.Bottom.Val, "Border style should be single")
	assert.Equal(t, 6, *p.ct.Property.Border.Bottom.Size, "Border size should be 6")
	assert.Equal(t, "auto", *p.ct.Property.Border.Bottom.Color, "Border color should be auto")
	assert.Equal(t, "1", *p.ct.Property.Border.Bottom.Space, "Border space should be 1")
}

// TestParagraphBottomBorder_Chaining tests that BottomBorder returns the paragraph for chaining
func TestParagraphBottomBorder_Chaining(t *testing.T) {
	doc := setupRootDoc(t)
	p := doc.AddParagraph("Test paragraph").
		BottomBorder(stypes.BorderStyleDouble, 8, "000000")

	// Justification doesn't return the paragraph, so call it separately
	p.Justification(stypes.JustificationCenter)

	assert.NotNil(t, p.ct.Property.Border.Bottom, "Paragraph should have bottom border")
	assert.Equal(t, stypes.BorderStyleDouble, p.ct.Property.Border.Bottom.Val, "Border style should be double")
	assert.NotNil(t, p.ct.Property.Justification, "Paragraph should have justification")
	assert.Equal(t, stypes.JustificationCenter, p.ct.Property.Justification.Val, "Justification should be center")
}

// TestParagraphBorder tests the Border method with a complete ParaBorder
func TestParagraphBorder(t *testing.T) {
	doc := setupRootDoc(t)
	p := doc.AddParagraph("Test paragraph")

	border := &ctypes.ParaBorder{
		Top: &ctypes.Border{
			Val:   stypes.BorderStyleSingle,
			Size:  internal.ToPtr(6),
			Color: internal.ToPtr("FF0000"),
		},
		Bottom: &ctypes.Border{
			Val:   stypes.BorderStyleDouble,
			Size:  internal.ToPtr(12),
			Color: internal.ToPtr("0000FF"),
		},
		Left: &ctypes.Border{
			Val:   stypes.BorderStyleDashed,
			Size:  internal.ToPtr(8),
			Color: internal.ToPtr("00FF00"),
		},
		Right: &ctypes.Border{
			Val:   stypes.BorderStyleDotted,
			Size:  internal.ToPtr(8),
			Color: internal.ToPtr("FFFF00"),
		},
	}

	p.Border(border)

	assert.NotNil(t, p.ct.Property.Border, "Paragraph should have border")
	assert.Equal(t, stypes.BorderStyleSingle, p.ct.Property.Border.Top.Val, "Top border style should be single")
	assert.Equal(t, stypes.BorderStyleDouble, p.ct.Property.Border.Bottom.Val, "Bottom border style should be double")
	assert.Equal(t, stypes.BorderStyleDashed, p.ct.Property.Border.Left.Val, "Left border style should be dashed")
	assert.Equal(t, stypes.BorderStyleDotted, p.ct.Property.Border.Right.Val, "Right border style should be dotted")
}

// TestHorizontalLine_Integration tests creating a document with multiple horizontal lines
func TestHorizontalLine_Integration(t *testing.T) {
	doc := setupRootDoc(t)

	doc.AddParagraph("Section 1")
	doc.AddHorizontalLine()
	doc.AddParagraph("Section 2")
	doc.AddDoubleHorizontalLine()
	doc.AddParagraph("Section 3")
	doc.AddThickHorizontalLine()
	doc.AddParagraph("Section 4")
	doc.AddCustomHorizontalLine(stypes.BorderStyleWave, 12, "FF0000")

	// Verify document has correct number of children
	assert.Equal(t, 8, len(doc.Document.Body.Children), "Document should have 8 children (4 paragraphs + 4 lines)")

	// Verify that each line is a paragraph with a bottom border
	assert.NotNil(t, doc.Document.Body.Children[1].Para.ct.Property.Border.Bottom, "Second child should have bottom border")
	assert.NotNil(t, doc.Document.Body.Children[3].Para.ct.Property.Border.Bottom, "Fourth child should have bottom border")
	assert.NotNil(t, doc.Document.Body.Children[5].Para.ct.Property.Border.Bottom, "Sixth child should have bottom border")
	assert.NotNil(t, doc.Document.Body.Children[7].Para.ct.Property.Border.Bottom, "Eighth child should have bottom border")
}
