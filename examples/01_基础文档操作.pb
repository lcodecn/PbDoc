; ============================================================================
; 示例 01: 基础文档操作
; 演示: 创建文档、设置属性、添加段落/标题、保存文件
; ============================================================================
XIncludeFile "../PbDoc.pb"

Define *doc.PbDocument
Define *para.PbDocParagraph
Define outputPath.s

; 初始化 PbDoc 库
PbDoc_Init()

; --- 创建空白文档 ---
*doc = PbDocument_Create()
Debug "文档创建: " + Str(Bool(*doc <> 0))

; --- 设置文档属性 ---
; 这些属性会写入 docProps/core.xml 和 docProps/app.xml
PbDocument_SetTitle(*doc, "基础文档操作示例")
PbDocument_SetAuthor(*doc, "PbDoc 用户")
PbDocument_SetSubject(*doc, "PbDoc 基础功能演示")
PbDocument_SetCompany(*doc, "lcode.cn")
Debug "文档属性设置完成"

; --- 添加标题 ---
; 参数: 文档指针, 标题文本, 标题级别(1-9)
PbDoc_Document_AddHeading(*doc, "第一章 PbDoc 基础操作", 1)
PbDoc_Document_AddHeading(*doc, "1.1 创建文档", 2)
PbDoc_Document_AddHeading(*doc, "1.1.1 详细说明", 3)

; --- 添加普通段落 ---
; 参数: 文档指针, 段落文本
PbDoc_Document_AddParagraph(*doc, "这是第一个段落。PbDoc 可以轻松创建 Word 文档。")
PbDoc_Document_AddParagraph(*doc, "这是第二个段落。每个段落都是独立的文本块。")

; --- 添加空段落 ---
PbDoc_Document_AddParagraph(*doc, "")

; --- 添加居中段落 ---
*para = PbDoc_Document_AddParagraph(*doc, "这段文字居中显示")
PbDoc_Paragraph_SetAlignment(*para, #PbDoc_ALIGN_CENTER)

; --- 添加右对齐段落 ---
*para = PbDoc_Document_AddParagraph(*doc, "这段文字右对齐")
PbDoc_Paragraph_SetAlignment(*para, #PbDoc_ALIGN_RIGHT)

; --- 添加两端对齐段落 ---
*para = PbDoc_Document_AddParagraph(*doc, "这段文字两端对齐。两端对齐会使文本在左右两边都对齐边界，这是中文排版中常用的对齐方式。")
PbDoc_Paragraph_SetAlignment(*para, #PbDoc_ALIGN_JUSTIFY)

; --- 获取段落数量 ---
Debug "段落数量: " + Str(PbDoc_Document_GetParagraphCount(*doc))

; --- 保存文档 ---
outputPath = GetCurrentDirectory() + "示例01_基础文档.docx"
If PbDocument_Save(*doc, outputPath)
  Debug "文档保存成功: " + outputPath
Else
  Debug "保存失败: " + PbDoc_GetLastError()
EndIf

; --- 释放文档 ---
PbDocument_Free(*doc)

; --- 清理库资源 ---
PbDoc_Cleanup()

Debug "示例01完成"
