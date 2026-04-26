# PbDoc Library v1.3

PbDoc

\- PureBasic Word (.docx) Document Processing Library

- Author  : lcode.cn
- Version  : 1.3
- License  : Apache 2.0
- Compiler  : PureBasic 6.40 (Windows - x86)

***

## Introduction

PbDoc is a Word document processing library implemented entirely in PureBasic. It can create .docx files conforming to the Office Open XML standard without requiring Microsoft Office or any third-party dependencies.

This library is written with reference to the Python python-docx project, using PureBasic's built-in Packer (ZIP compression) library.

## Key Features

- Create Word Documents  : Create .docx files conforming to the Office Open XML standard from scratch
- Document Properties  : Set title, author, subject, company and other document metadata
- Paragraph Operations  : Add paragraphs, set alignment, set paragraph styles
- Heading Operations  : Add heading levels 1-9, automatically apply built-in heading styles
- Font Formatting  : Bold, italic, underline (7 types), strikethrough, superscript/subscript, font name, color, highlight, font size
- Paragraph Formatting  : First line indent, left/right indent, hanging indent, space before/after, line spacing, page break control
- Tab Stops  : Left/right/center/decimal aligned tabs, 5 leader types
- Table Operations  : Create tables, set style/alignment/column width, fill cells
- Hyperlinks  : Add external hyperlinks
- Page Breaks  : Insert page breaks, line breaks, column breaks
- Page Setup  : Paper size, margins, page orientation (portrait/landscape)
- Unit Conversion  : EMU/twips/points/cm/inches conversion

## System Requirements

- This project has been compiled successfully with PureBasic 6.40 (Windows x86). Other environments need to be tested independently.

## Quick Start

For details, please refer to the documentation: docs\PbDoc_Help_en.html

### Creating a Word Document

```purebasic
XIncludeFile "PbDoc.pb"

; Initialize the library
PbDoc_Init()

; Create a new document
*doc.PbDocument = PbDocument_Create()

; Set document properties
PbDocument_SetTitle(*doc, "My Document")
PbDocument_SetAuthor(*doc, "PbDoc User")

; Add heading and paragraph
PbDoc_Document_AddHeading(*doc, "Chapter 1 Introduction", 1)
PbDoc_Document_AddParagraph(*doc, "This is the first paragraph of the document.")

; Save the file
PbDocument_Save(*doc, "output.docx")

; Release resources
PbDocument_Free(*doc)
PbDoc_Cleanup()
```

### Adding Formatted Text

```purebasic
; Add a multi-Run paragraph (different formatting)
*para.PbDocParagraph = PbDoc_Document_AddParagraph(*doc, "")
*run.PbDocRun = PbDoc_Paragraph_AddRun(*para, "Bold text ")
PbDoc_Font_SetBold(*run, #True)
*run = PbDoc_Paragraph_AddRun(*para, "Red text ")
PbDoc_Font_SetColorHex(*run, "FF0000")
*run = PbDoc_Paragraph_AddRun(*para, "18pt text")
PbDoc_Font_SetSize(*run, 18)
```

### Creating a Table

```purebasic
; Create a 3-row, 4-column table
*table.PbDocTable = PbDoc_Document_AddTable(*doc, 3, 4)
PbDoc_Table_SetStyle(*table, "TableGrid")

; Fill the header row
PbDoc_Cell_SetText(PbDoc_Table_GetCell(*table, 0, 0), "Name")
PbDoc_Cell_SetText(PbDoc_Table_GetCell(*table, 0, 1), "Age")
PbDoc_Cell_SetText(PbDoc_Table_GetCell(*table, 0, 2), "City")
PbDoc_Cell_SetText(PbDoc_Table_GetCell(*table, 0, 3), "Occupation")

; Fill data rows
PbDoc_Cell_SetText(PbDoc_Table_GetCell(*table, 1, 0), "Zhang San")
PbDoc_Cell_SetText(PbDoc_Table_GetCell(*table, 1, 1), "28")
PbDoc_Cell_SetText(PbDoc_Table_GetCell(*table, 1, 2), "Beijing")
PbDoc_Cell_SetText(PbDoc_Table_GetCell(*table, 1, 3), "Engineer")
```

## API Documentation

### Initialization and Cleanup

| Function | Description |
| --- | --- |
| `PbDoc_Init()` | Initialize the PbDoc library, must be called before use |
| `PbDoc_Cleanup()` | Clean up library resources, call before program exit |

### Document Operations

| Function | Description |
| --- | --- |
| `PbDocument_Create()` | Create a new blank document, returns document pointer |
| `PbDocument_Save(*doc, filePath.s)` | Save document to a .docx file |
| `PbDocument_Free(*doc)` | Release document memory |
| `PbDocument_SetTitle(*doc, title.s)` | Set document title |
| `PbDocument_SetAuthor(*doc, author.s)` | Set document author |
| `PbDocument_SetSubject(*doc, subject.s)` | Set document subject |
| `PbDocument_SetCompany(*doc, company.s)` | Set company name |

### Paragraph Operations

| Function | Description |
| --- | --- |
| `PbDoc_Document_AddParagraph(*doc, text.s)` | Add a paragraph, returns paragraph pointer |
| `PbDoc_Paragraph_GetText(*para)` | Get paragraph text |
| `PbDoc_Paragraph_SetAlignment(*para, alignment)` | Set paragraph alignment |
| `PbDoc_Paragraph_SetStyle(*para, styleName.s)` | Set paragraph style |
| `PbDoc_Document_AddHeading(*doc, text.s, level.l)` | Add heading (levels 1-9) |

### Run Operations

| Function | Description |
| --- | --- |
| `PbDoc_Paragraph_AddRun(*para, text.s)` | Add a text run, returns Run pointer |
| `PbDoc_Run_AddTab(*run)` | Add a Tab character |
| `PbDoc_Run_AddBreak(*run, breakType.l)` | Add a break (line/page/column) |

### Font Formatting

| Function | Description |
| --- | --- |
| `PbDoc_Font_SetBold(*run, bold.l)` | Set bold |
| `PbDoc_Font_SetItalic(*run, italic.l)` | Set italic |
| `PbDoc_Font_SetUnderline(*run, underline.l)` | Set underline type |
| `PbDoc_Font_SetStrike(*run, strike.l)` | Set strikethrough |
| `PbDoc_Font_SetDoubleStrike(*run, doubleStrike.l)` | Set double strikethrough |
| `PbDoc_Font_SetSuperscript(*run, superscript.l)` | Set superscript |
| `PbDoc_Font_SetSubscript(*run, subscript.l)` | Set subscript |
| `PbDoc_Font_SetName(*run, fontName.s)` | Set font name |
| `PbDoc_Font_SetColorHex(*run, colorHex.s)` | Set font color (hexadecimal) |
| `PbDoc_Font_SetColorRGB(*run, r.a, g.a, b.a)` | Set font color (RGB) |
| `PbDoc_Font_SetHighlight(*run, highlight.s)` | Set highlight color |
| `PbDoc_Font_SetSize(*run, size.l)` | Set font size (points) |

### Paragraph Formatting

| Function | Description |
| --- | --- |
| `PbDoc_ParaFormat_SetFirstLineIndent(*para, emu.q)` | Set first line indent |
| `PbDoc_ParaFormat_SetLeftIndent(*para, emu.q)` | Set left indent |
| `PbDoc_ParaFormat_SetRightIndent(*para, emu.q)` | Set right indent |
| `PbDoc_ParaFormat_SetHangingIndent(*para, emu.q)` | Set hanging indent |
| `PbDoc_ParaFormat_SetSpaceBefore(*para, emu.q)` | Set space before paragraph |
| `PbDoc_ParaFormat_SetSpaceAfter(*para, emu.q)` | Set space after paragraph |
| `PbDoc_ParaFormat_SetLineSpacing(*para, spacing.q, rule.l)` | Set line spacing |
| `PbDoc_ParaFormat_SetKeepNext(*para, keepNext.l)` | Keep with next paragraph |
| `PbDoc_ParaFormat_SetKeepLines(*para, keepLines.l)` | Keep lines together |
| `PbDoc_ParaFormat_SetWidowControl(*para, widowControl.l)` | Widow/orphan control |
| `PbDoc_ParaFormat_AddTabStop(*para, position.q, alignment.l, leader.l)` | Add tab stop |

### Table Operations

| Function | Description |
| --- | --- |
| `PbDoc_Document_AddTable(*doc, rows.l, cols.l)` | Create a table |
| `PbDoc_Table_SetStyle(*tbl, style.s)` | Set table style |
| `PbDoc_Table_SetAlignment(*tbl, alignment.l)` | Set table alignment |
| `PbDoc_Table_SetColumnWidth(*tbl, colIndex.l, widthEmu.q)` | Set column width |
| `PbDoc_Table_GetCell(*tbl, row.l, col.l)` | Get cell reference |
| `PbDoc_Cell_SetText(*cell, text.s)` | Set cell text |

### Hyperlinks and Page Breaks

| Function | Description |
| --- | --- |
| `PbDoc_Document_AddHyperlink(*doc, url.s, text.s)` | Add a hyperlink |
| `PbDoc_Document_AddPageBreak(*doc)` | Add a page break |

### Page Setup

| Function | Description |
| --- | --- |
| `PbDoc_Section_SetPageSize(*doc, widthTwips.q, heightTwips.q)` | Set page size |
| `PbDoc_Section_SetMargins(*doc, top.q, bottom.q, left.q, right.q)` | Set page margins |
| `PbDoc_Section_SetOrientation(*doc, orientation.l)` | Set page orientation |
| `PbDoc_Section_SetHeaderDistance(*doc, dist.q)` | Set header distance |
| `PbDoc_Section_SetFooterDistance(*doc, dist.q)` | Set footer distance |
| `PbDoc_Section_SetGutter(*doc, gutter.q)` | Set gutter |

### Unit Conversion

| Function | Description |
| --- | --- |
| `PbDoc_InchesToEmu(inches.d)` | Inches to EMU |
| `PbDoc_CmToEmu(cm.d)` | Centimeters to EMU |
| `PbDoc_MmToEmu(mm.d)` | Millimeters to EMU |
| `PbDoc_PtToEmu(pt.d)` | Points to EMU |
| `PbDoc_PtToTwips(pt.d)` | Points to twips |
| `PbDoc_CmToTwips(cm.d)` | Centimeters to twips |
| `PbDoc_InchesToTwips(inches.d)` | Inches to twips |

### Enumeration Constants

| Constant | Value | Description |
| --- | --- | --- |
| `#PbDoc_ALIGN_LEFT` | 1 | Left align |
| `#PbDoc_ALIGN_CENTER` | 2 | Center align |
| `#PbDoc_ALIGN_RIGHT` | 3 | Right align |
| `#PbDoc_ALIGN_JUSTIFY` | 4 | Justify |
| `#PbDoc_ALIGN_DISTRIBUTE` | 5 | Distribute |
| `#PbDoc_TAB_LEFT` | 1 | Left tab |
| `#PbDoc_TAB_CENTER` | 2 | Center tab |
| `#PbDoc_TAB_RIGHT` | 3 | Right tab |
| `#PbDoc_TAB_DECIMAL` | 4 | Decimal tab |
| `#PbDoc_TAB_LEADER_NONE` | 1 | No leader |
| `#PbDoc_TAB_LEADER_DOTS` | 2 | Dot leader |
| `#PbDoc_TAB_LEADER_DASHES` | 4 | Dash leader |
| `#PbDoc_UNDERLINE_SINGLE` | 1 | Single underline |
| `#PbDoc_UNDERLINE_DOUBLE` | 2 | Double underline |
| `#PbDoc_BREAK_LINE` | 1 | Line break |
| `#PbDoc_BREAK_PAGE` | 2 | Page break |
| `#PbDoc_BREAK_COLUMN` | 3 | Column break |
| `#PbDoc_LINE_SPACING_AUTO` | 1 | Auto line spacing |
| `#PbDoc_LINE_SPACING_EXACT` | 2 | Exact line spacing |
| `#PbDoc_ORIENT_PORTRAIT` | 1 | Portrait |
| `#PbDoc_ORIENT_LANDSCAPE` | 2 | Landscape |

## File Structure

The PbDoc.pb file is organized into the following 35 sections:

| Section | Content |
| --- | --- |
| 1 | Compiler directives and initialization |
| 2 | Constant definitions (EMU/twips conversion, version number) |
| 3 | Enumeration definitions (alignment, tabs, underline, breaks, line spacing, page orientation) |
| 4 | Data structure definitions (Run, paragraph, table, cell, section, document, etc.) |
| 5 | Global variables |
| 6 | Utility functions - Unit conversion |
| 7 | Utility functions - Error handling |
| 8 | Utility functions - XML helpers |
| 9 | Initialization and cleanup |
| 10 | Relationship management |
| 11 | Document creation and release |
| 12 | Document property settings |
| 13 | Paragraph operations |
| 14 | Run operations |
| 15 | Font formatting operations |
| 16 | Paragraph formatting operations |
| 17 | Heading operations |
| 18 | Hyperlink operations |
| 19 | Page break and line break operations |
| 20 | Table operations |
| 21 | Section and page setup |
| 22 | Document statistics |
| 23-33 | XML generation (Run/paragraph/table/page break/document/relationships/content types/properties/styles/static files) |
| 34 | ZIP package helper functions |
| 35 | Document saving |

## Example Files

| File | Description |
| --- | --- |
| `01_基础文档操作.pb` | Create document, properties, paragraphs/headings, alignment |
| `02_段落格式设置.pb` | Indentation, spacing, line spacing, page break control |
| `03_字体格式设置.pb` | Bold/italic/underline/strikethrough/superscript-subscript/color/highlight/font size |
| `04_表格操作.pb` | Create tables, styles, alignment, column width, cells |
| `05_超链接与分页符.pb` | Hyperlinks, page breaks, line breaks, column breaks |
| `06_TabStop与页面设置.pb` | Tab stops, leaders, page size/margins/orientation |
| `07_控制台全面测试.pb` | Console output test |
| `08_全面测试.pb` | Full feature test |

## Version History

### v1.3 (2026-04-23)

- \[Added] Console comprehensive test example (07\_控制台全面测试.pb)
- \[Added] Comprehensive test example (08\_全面测试.pb)
- \[Added] README.md project documentation
- \[Added] HTML help documentation (docs\PbDoc\_Help.html)
- \[Improved] Enhanced Chinese code comments
- \[Improved] File section partitioning, 35 sections clearly defined

### v1.2 (2026-04-23)

- \[Added] Underline types (7 types: single/double/dotted/dash/wave/thick)
- \[Added] Strikethrough/double strikethrough (PbDoc\_Font\_SetStrike/SetDoubleStrike)
- \[Added] Superscript/subscript (PbDoc\_Font\_SetSuperscript/SetSubscript)
- \[Added] Font name setting (PbDoc\_Font\_SetName)
- \[Added] RGB color setting (PbDoc\_Font\_SetColorRGB)
- \[Added] Highlight color (PbDoc\_Font\_SetHighlight)
- \[Added] Left indent/right indent/hanging indent
- \[Added] Line spacing setting (auto/exact/at least three rules)
- \[Added] Page break control (keep with next/keep lines together/widow control)
- \[Added] Paragraph style setting (PbDoc\_Paragraph\_SetStyle)
- \[Added] Break types (line/page/column)
- \[Added] Table alignment (PbDoc\_Table\_SetAlignment)
- \[Added] Table column width setting (PbDoc\_Table\_SetColumnWidth)
- \[Added] Page setup (paper size/margins/orientation/header-footer distance/gutter)
- \[Added] Title style definition
- \[Added] 6 detailed example files (01-06)
- \[Improved] Description text update: "ported from python-docx" → "written with reference to python-docx project"
- \[Improved] Chinese localization of prompt messages
- \[Improved] Code section partitioning, 35 section directory index

### v1.1 (2026-04-22)

- \[Fixed] Fixed invalid memory access error (address 1101004800)
- \[Fixed] Fixed #PB\_Packer\_Zip constant not existing (changed to #PB\_PackerPlugin\_Zip)
- \[Fixed] Fixed PbDoc\_Cell\_SetText parameter mismatch
- \[Fixed] Fixed demon-test.pb EnableExplicit variable declaration issues
- \[Improved] Abandoned MSXML COM, switched to string-based XML construction to avoid memory management issues
- \[Improved] Used PureBasic built-in ZipPacker instead of manual ZIP operations
- \[Improved] Data structure-driven design, all content stored in structures

### v1.0 (2026-04-21)

- \[Added] Initial project creation, written with reference to the python-docx project
- \[Added] Document creation and saving (PbDocument\_Create/Save/Free)
- \[Added] Document property settings (title/author/subject/company)
- \[Added] Paragraph operations (add/get text/set alignment)
- \[Added] Run operations (add Run/Tab)
- \[Added] Font formatting (bold/italic/color/font size)
- \[Added] Paragraph formatting (first line indent/space before-after/TabStop)
- \[Added] Heading operations (levels 1-9)
- \[Added] Hyperlink operations
- \[Added] Page break operations
- \[Added] Table operations (create/style/cells)
- \[Added] Unit conversion functions (EMU/twips/points/cm/inches)
- \[Added] Complete XML generation (document/styles/settings/fontTable/numbering/theme/core/app)
- \[Added] ZIP packaging and saving
- \[Added] Enumeration definitions (alignment/Tab/Tab leader/Body element types)
- \[Added] Data structure definitions (Run/paragraph/table/cell/relationship/document)

## License

This library is licensed under the Apache 2.0 License.

```
Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
```

The python-docx project referenced by this library is licensed under the MIT License.

```
The MIT License (MIT)
Copyright (c) 2013 Steve Canny, https://github.com/scanny

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.

```

## Support & Donate

If PbDoc is helpful to you, please consider supporting the developer to continue maintaining and improving this project.

- **PayPal**: [https://www.paypal.me/lcodecn](https://www.paypal.me/lcodecn)
- **WeChat**: #付款:lcodecn(经营_lcodecn)/openlib/003

Thank you for every bit of support!

## Acknowledgments

- Thanks to the python-docx project for providing an excellent reference implementation
- Thanks to the PureBasic QQ group for their support
