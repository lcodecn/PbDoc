; ============================================================================
; 示例 03: 字体格式设置
; 演示: 粗体/斜体/下划线/删除线/上下标/字体名/颜色/高亮/字号
; ============================================================================
XIncludeFile "../PbDoc.pb"

Define *doc.PbDocument
Define *para.PbDocParagraph
Define *run.PbDocRun
Define outputPath.s

PbDoc_Init()
*doc = PbDocument_Create()
PbDocument_SetTitle(*doc, "字体格式设置示例")

; --- 粗体与斜体 ---
PbDoc_Document_AddHeading(*doc, "粗体与斜体", 1)
*para = PbDoc_Document_AddParagraph(*doc, "")
*run = PbDoc_Paragraph_AddRun(*para, "这是粗体文字。")
PbDoc_Font_SetBold(*run, #True)
*run = PbDoc_Paragraph_AddRun(*para, " 这是斜体文字。")
PbDoc_Font_SetItalic(*run, #True)
*run = PbDoc_Paragraph_AddRun(*para, " 这是粗斜体文字。")
PbDoc_Font_SetBold(*run, #True)
PbDoc_Font_SetItalic(*run, #True)

; --- 下划线类型 ---
PbDoc_Document_AddHeading(*doc, "下划线类型", 1)
*para = PbDoc_Document_AddParagraph(*doc, "")
*run = PbDoc_Paragraph_AddRun(*para, "单下划线 ")
PbDoc_Font_SetUnderline(*run, #PbDoc_UNDERLINE_SINGLE)
*run = PbDoc_Paragraph_AddRun(*para, "  双下划线 ")
PbDoc_Font_SetUnderline(*run, #PbDoc_UNDERLINE_DOUBLE)
*run = PbDoc_Paragraph_AddRun(*para, "  点线下划线 ")
PbDoc_Font_SetUnderline(*run, #PbDoc_UNDERLINE_DOTTED)
*run = PbDoc_Paragraph_AddRun(*para, "  短划线下划线 ")
PbDoc_Font_SetUnderline(*run, #PbDoc_UNDERLINE_DASH)
*run = PbDoc_Paragraph_AddRun(*para, "  波浪线下划线 ")
PbDoc_Font_SetUnderline(*run, #PbDoc_UNDERLINE_WAVE)
*run = PbDoc_Paragraph_AddRun(*para, "  粗下划线")
PbDoc_Font_SetUnderline(*run, #PbDoc_UNDERLINE_THICK)

; --- 删除线 ---
PbDoc_Document_AddHeading(*doc, "删除线", 1)
*para = PbDoc_Document_AddParagraph(*doc, "")
*run = PbDoc_Paragraph_AddRun(*para, "单删除线 ")
PbDoc_Font_SetStrike(*run, #True)
*run = PbDoc_Paragraph_AddRun(*para, "  双删除线")
PbDoc_Font_SetDoubleStrike(*run, #True)

; --- 上标与下标 ---
PbDoc_Document_AddHeading(*doc, "上标与下标", 1)
*para = PbDoc_Document_AddParagraph(*doc, "")
*run = PbDoc_Paragraph_AddRun(*para, "水的化学式: H")
PbDoc_Font_SetSubscript(*run, #True)
*run = PbDoc_Paragraph_AddRun(*para, "2")
PbDoc_Font_SetSubscript(*run, #True)
*run = PbDoc_Paragraph_AddRun(*para, "O")
*run = PbDoc_Paragraph_AddRun(*para, "  勾股定理: a")
PbDoc_Font_SetSuperscript(*run, #True)
*run = PbDoc_Paragraph_AddRun(*para, "2")
PbDoc_Font_SetSuperscript(*run, #True)
*run = PbDoc_Paragraph_AddRun(*para, " + b")
PbDoc_Font_SetSuperscript(*run, #True)
*run = PbDoc_Paragraph_AddRun(*para, "2")
PbDoc_Font_SetSuperscript(*run, #True)
*run = PbDoc_Paragraph_AddRun(*para, " = c")
PbDoc_Font_SetSuperscript(*run, #True)
*run = PbDoc_Paragraph_AddRun(*para, "2")
PbDoc_Font_SetSuperscript(*run, #True)

; --- 字体名称 ---
PbDoc_Document_AddHeading(*doc, "字体名称", 1)
*para = PbDoc_Document_AddParagraph(*doc, "")
*run = PbDoc_Paragraph_AddRun(*para, "宋体字体 ")
PbDoc_Font_SetName(*run, "宋体")
*run = PbDoc_Paragraph_AddRun(*para, "黑体字体 ")
PbDoc_Font_SetName(*run, "黑体")
*run = PbDoc_Paragraph_AddRun(*para, "Arial字体 ")
PbDoc_Font_SetName(*run, "Arial")
*run = PbDoc_Paragraph_AddRun(*para, "Courier New字体")
PbDoc_Font_SetName(*run, "Courier New")

; --- 字体颜色 ---
PbDoc_Document_AddHeading(*doc, "字体颜色", 1)
*para = PbDoc_Document_AddParagraph(*doc, "")
*run = PbDoc_Paragraph_AddRun(*para, "红色 ")
PbDoc_Font_SetColorHex(*run, "FF0000")
*run = PbDoc_Paragraph_AddRun(*para, "蓝色 ")
PbDoc_Font_SetColorHex(*run, "0000FF")
*run = PbDoc_Paragraph_AddRun(*para, "绿色 ")
PbDoc_Font_SetColorHex(*run, "008000")
*run = PbDoc_Paragraph_AddRun(*para, "橙色")
PbDoc_Font_SetColorRGB(*run, 255, 165, 0)

; --- 高亮颜色 ---
PbDoc_Document_AddHeading(*doc, "高亮颜色", 1)
*para = PbDoc_Document_AddParagraph(*doc, "")
*run = PbDoc_Paragraph_AddRun(*para, "黄色高亮 ")
PbDoc_Font_SetHighlight(*run, "yellow")
*run = PbDoc_Paragraph_AddRun(*para, "绿色高亮 ")
PbDoc_Font_SetHighlight(*run, "green")
*run = PbDoc_Paragraph_AddRun(*para, "青色高亮 ")
PbDoc_Font_SetHighlight(*run, "cyan")
*run = PbDoc_Paragraph_AddRun(*para, "红色高亮")
PbDoc_Font_SetHighlight(*run, "red")

; --- 字号 ---
PbDoc_Document_AddHeading(*doc, "字号", 1)
*para = PbDoc_Document_AddParagraph(*doc, "")
*run = PbDoc_Paragraph_AddRun(*para, "8pt ")
PbDoc_Font_SetSize(*run, 8)
*run = PbDoc_Paragraph_AddRun(*para, "10pt ")
PbDoc_Font_SetSize(*run, 10)
*run = PbDoc_Paragraph_AddRun(*para, "12pt ")
PbDoc_Font_SetSize(*run, 12)
*run = PbDoc_Paragraph_AddRun(*para, "16pt ")
PbDoc_Font_SetSize(*run, 16)
*run = PbDoc_Paragraph_AddRun(*para, "24pt ")
PbDoc_Font_SetSize(*run, 24)
*run = PbDoc_Paragraph_AddRun(*para, "36pt")
PbDoc_Font_SetSize(*run, 36)

; --- 综合效果 ---
PbDoc_Document_AddHeading(*doc, "综合效果", 1)
*para = PbDoc_Document_AddParagraph(*doc, "")
*run = PbDoc_Paragraph_AddRun(*para, "粗体+红色+18pt+下划线")
PbDoc_Font_SetBold(*run, #True)
PbDoc_Font_SetColorHex(*run, "CC0000")
PbDoc_Font_SetSize(*run, 18)
PbDoc_Font_SetUnderline(*run, #PbDoc_UNDERLINE_SINGLE)

; --- 保存 ---
outputPath = GetCurrentDirectory() + "示例03_字体格式.docx"
If PbDocument_Save(*doc, outputPath)
  Debug "保存成功: " + outputPath
Else
  Debug "保存失败: " + PbDoc_GetLastError()
EndIf

PbDocument_Free(*doc)
PbDoc_Cleanup()
Debug "示例03完成"
