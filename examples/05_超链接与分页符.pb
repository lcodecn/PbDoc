; ============================================================================
; 示例 05: 超链接与分页符
; 演示: 添加超链接、分页符、换行符、分栏符
; ============================================================================
XIncludeFile "../PbDoc.pb"

Define *doc.PbDocument
Define *para.PbDocParagraph
Define *run.PbDocRun
Define outputPath.s

PbDoc_Init()
*doc = PbDocument_Create()
PbDocument_SetTitle(*doc, "超链接与分页符示例")

; --- 超链接 ---
PbDoc_Document_AddHeading(*doc, "超链接", 1)
PbDoc_Document_AddParagraph(*doc, "下面是几个超链接示例:")

; 添加超链接: 参数(文档, URL, 显示文本)
PbDoc_Document_AddHyperlink(*doc, "https://www.baidu.com", "百度搜索")
PbDoc_Document_AddHyperlink(*doc, "https://github.com", "GitHub")
PbDoc_Document_AddHyperlink(*doc, "https://docs.microsoft.com", "微软文档")

; --- 分页符 ---
PbDoc_Document_AddHeading(*doc, "分页符", 1)
PbDoc_Document_AddParagraph(*doc, "下面的内容将在新页面显示。")

; 添加分页符
PbDoc_Document_AddPageBreak(*doc)

; 新页面的内容
PbDoc_Document_AddHeading(*doc, "这是第二页", 1)
PbDoc_Document_AddParagraph(*doc, "这段内容出现在分页符之后的新页面上。")

; --- 换行符 (软回车) ---
PbDoc_Document_AddHeading(*doc, "换行符", 1)
*para = PbDoc_Document_AddParagraph(*doc, "")
*run = PbDoc_Paragraph_AddRun(*para, "第一行文字")
PbDoc_Run_AddBreak(*run, #PbDoc_BREAK_LINE)
*run = PbDoc_Paragraph_AddRun(*para, "第二行文字（换行符后）")
PbDoc_Run_AddBreak(*run, #PbDoc_BREAK_LINE)
*run = PbDoc_Paragraph_AddRun(*para, "第三行文字（仍在同一段落中）")
PbDoc_Document_AddParagraph(*doc, "注意: 换行符与分段不同，以上三行属于同一段落，共享段落格式。")

; --- 分栏符 ---
PbDoc_Document_AddHeading(*doc, "分栏符", 1)
*para = PbDoc_Document_AddParagraph(*doc, "")
*run = PbDoc_Paragraph_AddRun(*para, "分栏符前的内容")
PbDoc_Run_AddBreak(*run, #PbDoc_BREAK_COLUMN)
*run = PbDoc_Paragraph_AddRun(*para, "分栏符后的内容（在多栏布局中会跳到下一栏）")

; --- 保存 ---
outputPath = GetCurrentDirectory() + "示例05_超链接与分页符.docx"
If PbDocument_Save(*doc, outputPath)
  Debug "保存成功: " + outputPath
Else
  Debug "保存失败: " + PbDoc_GetLastError()
EndIf

PbDocument_Free(*doc)
PbDoc_Cleanup()
Debug "示例05完成"
