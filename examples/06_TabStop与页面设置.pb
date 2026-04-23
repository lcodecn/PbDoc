; ============================================================================
; 示例 06: TabStop 与页面设置
; 演示: Tab制表符设置、页面大小/边距/方向
; ============================================================================
XIncludeFile "../PbDoc.pb"

Define *doc.PbDocument
Define *para.PbDocParagraph
Define *run.PbDocRun
Define outputPath.s

PbDoc_Init()
*doc = PbDocument_Create()
PbDocument_SetTitle(*doc, "TabStop与页面设置示例")

; ============================================================================
; TabStop 制表符设置
; ============================================================================

; --- 左对齐Tab ---
PbDoc_Document_AddHeading(*doc, "左对齐 Tab", 1)
*para = PbDoc_Document_AddParagraph(*doc, "")
PbDoc_ParaFormat_AddTabStop(*para, PbDoc_CmToEmu(5), #PbDoc_TAB_LEFT, #PbDoc_TAB_LEADER_NONE)
*run = PbDoc_Paragraph_AddRun(*para, "项目名称")
PbDoc_Run_AddTab(*run)
*run = PbDoc_Paragraph_AddRun(*para, "说明")
PbDoc_Paragraph_SetAlignment(*para, #PbDoc_ALIGN_LEFT)

; --- 带前导符的Tab ---
PbDoc_Document_AddHeading(*doc, "带前导符的 Tab", 1)

; 点线前导符 (目录常用)
*para = PbDoc_Document_AddParagraph(*doc, "")
PbDoc_ParaFormat_AddTabStop(*para, PbDoc_CmToEmu(14), #PbDoc_TAB_RIGHT, #PbDoc_TAB_LEADER_DOTS)
*run = PbDoc_Paragraph_AddRun(*para, "第一章 引言")
PbDoc_Run_AddTab(*run)
*run = PbDoc_Paragraph_AddRun(*para, "1")

*para = PbDoc_Document_AddParagraph(*doc, "")
PbDoc_ParaFormat_AddTabStop(*para, PbDoc_CmToEmu(14), #PbDoc_TAB_RIGHT, #PbDoc_TAB_LEADER_DOTS)
*run = PbDoc_Paragraph_AddRun(*para, "第二章 设计")
PbDoc_Run_AddTab(*run)
*run = PbDoc_Paragraph_AddRun(*para, "15")

*para = PbDoc_Document_AddParagraph(*doc, "")
PbDoc_ParaFormat_AddTabStop(*para, PbDoc_CmToEmu(14), #PbDoc_TAB_RIGHT, #PbDoc_TAB_LEADER_DOTS)
*run = PbDoc_Paragraph_AddRun(*para, "第三章 实现")
PbDoc_Run_AddTab(*run)
*run = PbDoc_Paragraph_AddRun(*para, "32")

; 短划线前导符
PbDoc_Document_AddParagraph(*doc, "")
*para = PbDoc_Document_AddParagraph(*doc, "")
PbDoc_ParaFormat_AddTabStop(*para, PbDoc_CmToEmu(14), #PbDoc_TAB_RIGHT, #PbDoc_TAB_LEADER_DASHES)
*run = PbDoc_Paragraph_AddRun(*para, "附录A")
PbDoc_Run_AddTab(*run)
*run = PbDoc_Paragraph_AddRun(*para, "50")

; --- 居中Tab ---
PbDoc_Document_AddHeading(*doc, "居中 Tab", 1)
*para = PbDoc_Document_AddParagraph(*doc, "")
PbDoc_ParaFormat_AddTabStop(*para, PbDoc_CmToEmu(8), #PbDoc_TAB_CENTER, #PbDoc_TAB_LEADER_NONE)
*run = PbDoc_Paragraph_AddRun(*para, "左")
PbDoc_Run_AddTab(*run)
*run = PbDoc_Paragraph_AddRun(*para, "中")
PbDoc_Run_AddTab(*run)
*run = PbDoc_Paragraph_AddRun(*para, "右")

; --- 小数点对齐Tab ---
PbDoc_Document_AddHeading(*doc, "小数点对齐 Tab", 1)
*para = PbDoc_Document_AddParagraph(*doc, "")
PbDoc_ParaFormat_AddTabStop(*para, PbDoc_CmToEmu(8), #PbDoc_TAB_DECIMAL, #PbDoc_TAB_LEADER_NONE)
*run = PbDoc_Paragraph_AddRun(*para, "数值1:")
PbDoc_Run_AddTab(*run)
*run = PbDoc_Paragraph_AddRun(*para, "123.45")
*para = PbDoc_Document_AddParagraph(*doc, "")
PbDoc_ParaFormat_AddTabStop(*para, PbDoc_CmToEmu(8), #PbDoc_TAB_DECIMAL, #PbDoc_TAB_LEADER_NONE)
*run = PbDoc_Paragraph_AddRun(*para, "数值2:")
PbDoc_Run_AddTab(*run)
*run = PbDoc_Paragraph_AddRun(*para, "9.8")

; ============================================================================
; 页面设置
; ============================================================================

PbDoc_Document_AddHeading(*doc, "页面设置", 1)
PbDoc_Document_AddParagraph(*doc, "以下演示不同的页面设置选项。")

; --- A4纸张设置 ---
; A4: 21cm x 29.7cm
PbDoc_Section_SetPageSize(*doc, PbDoc_CmToTwips(21), PbDoc_CmToTwips(29.7))

; --- 页面边距 ---
; 上下2.54cm, 左右3.17cm (Word默认值)
PbDoc_Section_SetMargins(*doc, PbDoc_CmToTwips(2.54), PbDoc_CmToTwips(2.54), PbDoc_CmToTwips(3.17), PbDoc_CmToTwips(3.17))

; --- 页眉页脚距离 ---
PbDoc_Section_SetHeaderDistance(*doc, PbDoc_CmToTwips(1.27))
PbDoc_Section_SetFooterDistance(*doc, PbDoc_CmToTwips(1.27))

; --- 装订线 ---
PbDoc_Section_SetGutter(*doc, PbDoc_CmToTwips(0))

; 注意: 如需横向页面，可调用:
; PbDoc_Section_SetOrientation(*doc, #PbDoc_ORIENT_LANDSCAPE)

PbDoc_Document_AddParagraph(*doc, "本文档已设置为 A4 纵向纸张，边距为 Word 默认值。")

; --- 保存 ---
outputPath = GetCurrentDirectory() + "示例06_TabStop与页面设置.docx"
If PbDocument_Save(*doc, outputPath)
  Debug "保存成功: " + outputPath
Else
  Debug "保存失败: " + PbDoc_GetLastError()
EndIf

PbDocument_Free(*doc)
PbDoc_Cleanup()
Debug "示例06完成"
