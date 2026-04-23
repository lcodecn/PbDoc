; ============================================================================
; PbDoc 完整功能测试 - 测试所有新增功能
; ============================================================================
XIncludeFile "../PbDoc.pb"

Define outputPath$ = GetCurrentDirectory() + "示例08_PbDoc_Full_Test.docx"

Debug "=== PbDoc 完整功能测试开始 ==="

; 初始化
If PbDoc_Init() = #False
  Debug "[FAIL] 初始化失败"
  End
EndIf

; 创建文档
Define *doc.PbDocument = PbDocument_Create()
If *doc = 0
  Debug "[FAIL] 创建文档失败"
  End
EndIf

Debug "[OK] 文档创建成功"

; ============================================================================
; 1. 设置文档属性
; ============================================================================
PbDocument_SetTitle(*doc, "PbDoc 完整功能测试")
PbDocument_SetAuthor(*doc, "lcode.cn")
PbDocument_SetSubject(*doc, "PbDoc 库测试")
PbDocument_SetCompany(*doc, "lcode.cn")
Debug "[OK] 文档属性设置完成"

; ============================================================================
; 2. 添加标题
; ============================================================================
PbDoc_Document_AddHeading(*doc, "PbDoc 完整功能测试报告", 1)
Debug "[OK] 添加标题完成"

; ============================================================================
; 3. 添加带Tab Stops的段落
; ============================================================================
PbDoc_Document_AddHeading(*doc, "3. Tab Stops 测试", 2)

; 创建带Tab的段落
Define *tabPara.PbDocParagraph = PbDoc_Document_AddParagraph(*doc, "")
PbDoc_ParaFormat_AddTabStop(*tabPara, PbDoc_InchesToEmu(1.5), #PbDoc_TAB_LEFT, #PbDoc_TAB_LEADER_DOTS)
PbDoc_ParaFormat_AddTabStop(*tabPara, PbDoc_InchesToEmu(3.0), #PbDoc_TAB_RIGHT, #PbDoc_TAB_LEADER_DASHES)

; 添加带Tab的run
Define *run1.PbDocRun = PbDoc_Paragraph_AddRun(*tabPara, "项目名称:")
PbDoc_Run_AddTab(*run1)
Define *run2.PbDocRun = PbDoc_Paragraph_AddRun(*tabPara, "PbDoc")
PbDoc_Run_AddTab(*run2)
PbDoc_Paragraph_AddRun(*tabPara, "v0.1")
Debug "[OK] Tab Stops测试完成"

; ============================================================================
; 4. 添加超链接
; ============================================================================
PbDoc_Document_AddHeading(*doc, "4. 超链接测试", 2)

; 添加超链接段落
Define *hlPara.PbDocParagraph = PbDoc_Document_AddHyperlink(*doc, "https://www.example.com", "访问示例网站")
Debug "[OK] 超链接添加完成"

; ============================================================================
; 5. 添加段落和设置格式
; ============================================================================
PbDoc_Document_AddHeading(*doc, "5. 段落格式测试", 2)

; 普通段落
Define *para1.PbDocParagraph = PbDoc_Document_AddParagraph(*doc, "这是普通段落，使用默认格式。")
Debug "[OK] 添加普通段落"

; 居中段落带间距
Define *para2.PbDocParagraph = PbDoc_Document_AddParagraph(*doc, "这是居中对齐的段落，带有段前段后间距。")
PbDoc_Paragraph_SetAlignment(*para2, #PbDoc_ALIGN_CENTER)
PbDoc_ParaFormat_SetSpaceBefore(*para2, PbDoc_PtToEmu(12))
PbDoc_ParaFormat_SetSpaceAfter(*para2, PbDoc_PtToEmu(12))
Debug "[OK] 居中段落设置完成"

; 带首行缩进的段落
Define *para3.PbDocParagraph = PbDoc_Document_AddParagraph(*doc, "这是首行缩进段落。这是一段较长的文本，用于测试段落格式设置功能。通过设置首行缩进，可以使段落看起来更加美观和专业。")
PbDoc_ParaFormat_SetFirstLineIndent(*para3, PbDoc_CmToEmu(1.5))
Debug "[OK] 首行缩进段落完成"

; 多Run段落（不同格式）
Define *multiPara.PbDocParagraph = PbDoc_Document_AddParagraph(*doc, "")
PbDoc_Paragraph_AddRun(*multiPara, "正常文本 ")
Define *boldRun.PbDocRun = PbDoc_Paragraph_AddRun(*multiPara, "粗体文本 ")
PbDoc_Font_SetBold(*boldRun, #True)
Define *italicRun.PbDocRun = PbDoc_Paragraph_AddRun(*multiPara, "斜体文本 ")
PbDoc_Font_SetItalic(*italicRun, #True)
Define *redRun.PbDocRun = PbDoc_Paragraph_AddRun(*multiPara, "红色文本")
PbDoc_Font_SetColorHex(*redRun, "FF0000")
Define *largeRun.PbDocRun = PbDoc_Paragraph_AddRun(*multiPara, " 大字号")
PbDoc_Font_SetSize(*largeRun, 20)
Debug "[OK] 多Run段落完成"

; ============================================================================
; 6. 添加表格
; ============================================================================
PbDoc_Document_AddHeading(*doc, "6. 表格测试", 2)

; 创建表格
Define *table.PbDocTable = PbDoc_Document_AddTable(*doc, 4, 3)
If *table
  PbDoc_Table_SetStyle(*table, "TableGrid")
  
  ; 表头
  PbDoc_Cell_SetText(PbDoc_Table_GetCell(*table, 0, 0), "功能模块")
  PbDoc_Cell_SetText(PbDoc_Table_GetCell(*table, 0, 1), "状态")
  PbDoc_Cell_SetText(PbDoc_Table_GetCell(*table, 0, 2), "备注")
  
  ; 数据行
  PbDoc_Cell_SetText(PbDoc_Table_GetCell(*table, 1, 0), "段落操作")
  PbDoc_Cell_SetText(PbDoc_Table_GetCell(*table, 1, 1), "已完成")
  PbDoc_Cell_SetText(PbDoc_Table_GetCell(*table, 1, 2), "支持对齐、缩进、间距")
  
  PbDoc_Cell_SetText(PbDoc_Table_GetCell(*table, 2, 0), "字体格式")
  PbDoc_Cell_SetText(PbDoc_Table_GetCell(*table, 2, 1), "已完成")
  PbDoc_Cell_SetText(PbDoc_Table_GetCell(*table, 2, 2), "支持粗体、斜体、颜色、字号")
  
  PbDoc_Cell_SetText(PbDoc_Table_GetCell(*table, 3, 0), "表格操作")
  PbDoc_Cell_SetText(PbDoc_Table_GetCell(*table, 3, 1), "已完成")
  PbDoc_Cell_SetText(PbDoc_Table_GetCell(*table, 3, 2), "支持样式、对齐")
  
  Debug "[OK] 创建表格完成"
EndIf

; ============================================================================
; 7. 添加分页符
; ============================================================================
PbDoc_Document_AddPageBreak(*doc)
Debug "[OK] 添加分页符"

; ============================================================================
; 8. 第二部分标题和内容
; ============================================================================
PbDoc_Document_AddHeading(*doc, "7. 高级功能", 2)

PbDoc_Document_AddParagraph(*doc, "PbDoc 支持以下高级功能：")
PbDoc_Document_AddParagraph(*doc, "1. 段落格式设置（对齐、缩进、间距）")
PbDoc_Document_AddParagraph(*doc, "2. 字体格式设置（粗体、斜体、颜色、字号）")
PbDoc_Document_AddParagraph(*doc, "3. 表格操作（创建、填充、样式）")
PbDoc_Document_AddParagraph(*doc, "4. 超链接支持")
PbDoc_Document_AddParagraph(*doc, "5. Tab Stops支持")
PbDoc_Document_AddParagraph(*doc, "6. 图片插入（待完善DrawingML）")

; ============================================================================
; 9. 统计信息
; ============================================================================
PbDoc_Document_AddHeading(*doc, "8. 统计信息", 2)
PbDoc_Document_AddParagraph(*doc, "段落总数: " + Str(PbDoc_Document_GetParagraphCount(*doc)))
PbDoc_Document_AddParagraph(*doc, "表格总数: " + Str(PbDoc_Document_GetTableCount(*doc)))

; ============================================================================
; 10. 保存文档
; ============================================================================
If PbDocument_Save(*doc, outputPath$)
  Debug "[OK] 文档保存成功: " + outputPath$
Else
  Debug "[FAIL] 保存失败: " + PbDoc_GetLastError()
EndIf

; 清理
PbDocument_Free(*doc)
PbDoc_Cleanup()

Debug "=== PbDoc 完整功能测试完成 ==="
; IDE Options = PureBasic 6.40 (Windows - x86)
; CursorPosition = 170
; FirstLine = 143
; EnableThread
; EnableXP
; DPIAware
; CompileSourceDirectory