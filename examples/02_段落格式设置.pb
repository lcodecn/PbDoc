; ============================================================================
; 示例 02: 段落格式设置
; 演示: 缩进(首行/左/右/悬挂)、间距(段前/段后/行距)、分页控制
; ============================================================================
XIncludeFile "../PbDoc.pb"

Define *doc.PbDocument
Define *para.PbDocParagraph
Define outputPath.s

PbDoc_Init()
*doc = PbDocument_Create()
PbDocument_SetTitle(*doc, "段落格式设置示例")

; --- 首行缩进 ---
; 中文排版常用: 首行缩进2字符 ≈ 0.74cm
PbDoc_Document_AddHeading(*doc, "首行缩进", 1)
*para = PbDoc_Document_AddParagraph(*doc, "本段落设置了首行缩进2字符（约0.74厘米）。这是中文文档排版的标准格式，每段开头空两格。")
PbDoc_ParaFormat_SetFirstLineIndent(*para, PbDoc_CmToEmu(0.74))

; --- 左缩进 ---
PbDoc_Document_AddHeading(*doc, "左缩进", 1)
*para = PbDoc_Document_AddParagraph(*doc, "本段落设置了左缩进2厘米，整个段落从左边向右偏移。")
PbDoc_ParaFormat_SetLeftIndent(*para, PbDoc_CmToEmu(2))

; --- 右缩进 ---
PbDoc_Document_AddHeading(*doc, "右缩进", 1)
*para = PbDoc_Document_AddParagraph(*doc, "本段落设置了右缩进2厘米，整个段落从右边向左偏移。")
PbDoc_ParaFormat_SetRightIndent(*para, PbDoc_CmToEmu(2))

; --- 悬挂缩进 ---
PbDoc_Document_AddHeading(*doc, "悬挂缩进", 1)
*para = PbDoc_Document_AddParagraph(*doc, "本段落设置了悬挂缩进1厘米。悬挂缩进常用于项目列表，第一行不缩进，后续行缩进。")
PbDoc_ParaFormat_SetHangingIndent(*para, PbDoc_CmToEmu(1))

; --- 段前间距 ---
PbDoc_Document_AddHeading(*doc, "段前间距", 1)
*para = PbDoc_Document_AddParagraph(*doc, "本段落设置了段前间距12磅。注意本段与上方标题之间的距离。")
PbDoc_ParaFormat_SetSpaceBefore(*para, PbDoc_PtToEmu(12))

; --- 段后间距 ---
PbDoc_Document_AddHeading(*doc, "段后间距", 1)
*para = PbDoc_Document_AddParagraph(*doc, "本段落设置了段后间距18磅。注意本段与下方内容之间的距离。")
PbDoc_ParaFormat_SetSpaceAfter(*para, PbDoc_PtToEmu(18))

; --- 行间距设置 ---
PbDoc_Document_AddHeading(*doc, "行间距", 1)

; 1.5倍行距
*para = PbDoc_Document_AddParagraph(*doc, "本段落设置了1.5倍行距。行距值360表示1.5倍行距(240=1倍, 360=1.5倍, 480=2倍)。")
PbDoc_ParaFormat_SetLineSpacing(*para, 360, #PbDoc_LINE_SPACING_AUTO)

; 固定行距20磅
*para = PbDoc_Document_AddParagraph(*doc, "本段落设置了固定行距20磅。无论字号大小，行距固定为20磅。")
PbDoc_ParaFormat_SetLineSpacing(*para, PbDoc_PtToTwips(20), #PbDoc_LINE_SPACING_EXACT)

; 最小行距15磅
*para = PbDoc_Document_AddParagraph(*doc, "本段落设置了最小行距15磅。行距不会小于15磅，但可以更大。")
PbDoc_ParaFormat_SetLineSpacing(*para, PbDoc_PtToTwips(15), #PbDoc_LINE_SPACING_AT_LEAST)

; --- 分页控制 ---
PbDoc_Document_AddHeading(*doc, "分页控制", 1)

; 与下段同页
*para = PbDoc_Document_AddParagraph(*doc, "本段落设置了「与下段同页」，确保不会与下一段落分页显示。")
PbDoc_ParaFormat_SetKeepNext(*para, #True)

; 段中不分页
*para = PbDoc_Document_AddParagraph(*doc, "本段落设置了「段中不分页」，确保段落内部不会出现分页符。")
PbDoc_ParaFormat_SetKeepLines(*para, #True)

; 孤行控制
*para = PbDoc_Document_AddParagraph(*doc, "本段落设置了「孤行控制」，防止段落首行或末行单独出现在页面顶部或底部。")
PbDoc_ParaFormat_SetWidowControl(*para, #True)

; --- 保存 ---
outputPath = GetCurrentDirectory() + "示例02_段落格式.docx"
If PbDocument_Save(*doc, outputPath)
  Debug "保存成功: " + outputPath
Else
  Debug "保存失败: " + PbDoc_GetLastError()
EndIf

PbDocument_Free(*doc)
PbDoc_Cleanup()
Debug "示例02完成"
