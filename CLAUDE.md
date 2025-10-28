# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Godocx is a pure Go library for creating, reading, and manipulating Microsoft Word DOCX files. It implements the Office Open XML (OOXML) WordprocessingML standard without dependencies on Microsoft Office or external libraries (beyond testify for tests).

## Development Commands

### Testing
```bash
# Run all tests
go test ./...

# Run tests with verbose output
go test -v ./...

# Run tests for a specific package
go test ./docx
go test ./wml/ctypes

# Run a specific test
go test -run TestFunctionName ./package

# Run tests with coverage
go test -cover ./...

# Generate test output as JSON (CI format)
go test -json ./...
```

### Building
```bash
# Install dependencies
go get .

# Build (verify compilation)
go build ./...

# Format code
go fmt ./...

# Run linter (if installed)
golint ./...
```

## Architecture

### Package Structure

The codebase follows a layered architecture reflecting the OOXML standard structure:

```
godocx/                    # Top-level API (NewDocument, OpenDocument)
├── docx/                  # High-level document manipulation API
├── wml/                   # WordprocessingML types
│   ├── ctypes/           # Complex types (paragraphs, tables, styles, etc.)
│   ├── stypes/           # Simple types (enums, basic types)
│   └── color/            # Color definitions
├── dml/                   # DrawingML for graphics/images
│   ├── dmlpic/           # Picture elements
│   ├── dmlct/            # DrawingML complex types
│   ├── dmlst/            # DrawingML simple types
│   ├── dmlprops/         # Drawing properties
│   ├── shapes/           # Shape elements
│   └── geom/             # Geometry definitions
├── packager/             # OPC (Open Packaging Conventions) handling
├── common/               # Shared utilities
│   ├── constants/        # OOXML constants
│   └── units/            # Measurement unit conversions
├── internal/             # Internal utilities
└── templates/            # Embedded default.docx template
```

### Key Architectural Concepts

**1. Document Lifecycle**

Documents flow through three stages:
- **Unpacking** (`packager.Unpack`): ZIP → in-memory structs with FileMap
- **Manipulation** (`docx.RootDoc` API): High-level operations
- **Packing** (`RootDoc.WriteTo`): In-memory structs → ZIP output

**2. Root Document (docx.RootDoc)**

The central structure managing the entire document:
- `Path`: Document file path
- `FileMap`: sync.Map holding all parts of the OOXML package
- `Document`: Main document.xml structure
- `DocStyles`: styles.xml structure
- `Numbering`: Numbering/list management
- `RootRels`: Root-level relationships (.rels)
- `ContentType`: [Content_Types].xml

All document operations go through RootDoc methods.

**3. Type Separation Pattern**

- **High-level API** (`docx/`): User-facing methods (AddParagraph, AddTable, etc.)
- **Complex Types** (`wml/ctypes/`): OOXML WordprocessingML complex types with XML marshaling
- **Simple Types** (`wml/stypes/`): OOXML simple types (enums, basic values)

Each `docx` type wraps a corresponding `ctypes` structure:
```go
type Paragraph struct {
    root *RootDoc         // Back-reference to document
    ct   ctypes.Paragraph // Underlying OOXML structure
}
```

**4. Relationship Management**

OOXML uses relationships to link document parts:
- Managed via `Relationships` type and `.rels` files
- Each relationship has unique rId (relationship ID)
- Images, styles, numbering all connected via relationships
- Use `IncRelationID()` to generate new IDs

**5. XML Marshaling**

All OOXML types implement custom `MarshalXML`/`UnmarshalXML`:
- Preserves exact XML structure per ECMA-376 spec
- Handles namespace prefixes (w:, r:, wp:, etc.)
- Maintains compatibility with Microsoft Word

## OOXML Standards Compliance

All features must conform to **ECMA-376 Office Open XML File Formats**:

- **Official Spec**: https://www.ecma-international.org/publications-and-standards/standards/ecma-376/
- **Schema Reference**: https://www.datypic.com/sc/ooxml/ (quick lookups)
- **Validation Tool**: OOXML Validator for VS Code (marketplace extension)

When implementing features:
1. Consult ECMA-376 for correct element structure
2. Check Datypic for attribute/child element requirements
3. Test with actual Word documents to verify compatibility
4. Ensure deterministic output (sort maps, stable ordering)

## Common Patterns

### Adding Document Elements

High-level elements are added through RootDoc:
```go
doc.AddParagraph("text")
doc.AddHeading("title", level)
doc.AddTable()
doc.AddPageBreak()
```

Child elements added through parent:
```go
para.AddText("text").Bold(true)
table.AddRow().AddCell().AddParagraph("cell text")
```

### Working with Images

Images require relationship management:
```go
// Images are packaged with unique IDs
// Drawing elements reference images via relationship IDs
// See docx/pic.go and dml/dmlpic/ for implementation
```

### List/Numbering System

Lists use abstract numbering definitions with instances:
```go
numId := doc.NewListInstance(abstractNumId)
para.Numbering(numId, level)
```

### Deterministic Output

The codebase enforces deterministic document generation (see docx/determinism_test.go):
- Sort keys when iterating maps
- Use stable ordering for relationships
- Ensures identical input produces byte-for-byte identical output
- Critical for version control and testing

## Testing Practices

- Every public function should have `_test.go` coverage
- Use `testdata/` for sample DOCX files
- Test both XML marshaling and high-level API
- Include round-trip tests (create → save → open → verify)
- Use table-driven tests for similar cases
- Verify output with actual Microsoft Word when possible

## Code Conventions

- Package doc.go files explain package purpose
- XML tags match OOXML spec exactly (case-sensitive)
- Use pointer receivers for methods that modify state
- Embed default template (`templates/default.docx`)
- Thread-safe operations where needed (FileMap uses sync.Map)
