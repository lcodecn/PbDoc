
; ============================================================================
; PbDoc 测试程序 - 控制台输出版本
; ============================================================================
XIncludeFile "../PbDoc.pb"

OpenConsole("PbDoc 测试")
PrintN("=== PbDoc 完整功能测试开始 ===")

; 初始化
If PbDoc_Init() = #False
  PrintN("[FAIL] 初始化失败")
  CloseConsole()
  End
EndIf

; 创建文档
Define *doc.PbDocument = PbDocument_Create()
If *doc = 0
  PrintN("[FAIL] 创建文档失败")
  CloseConsole()
  End
EndIf

PrintN("[OK] 文档创建成功")

; 设置文档属性
PbDocument_SetTitle(*doc, "PbDoc 完整功能测试")
PbDocument_SetAuthor(*doc, "lcode.cn")
PbDocument_SetSubject(*doc, "PbDoc 库测试")
PbDocument_SetCompany(*doc, "lcode.cn")
PrintN("[OK] 文档属性设置完成")

; 添加标题
PbDoc_Document_AddHeading(*doc, "PbDoc 完整功能测试报告", 1)
PrintN("[OK] 添加标题完成")

; 添加测试段落
PbDoc_Document_AddParagraph(*doc, "这是一个简单的测试段落。")
PrintN("[OK] 添加段落完成")

; 添加分页符
PbDoc_Document_AddPageBreak(*doc)
PrintN("[OK] 添加分页符")

; 保存文档
Define outputPath$ = GetCurrentDirectory() + "示例07_Console_Test.docx"
PrintN("正在保存文档到: " + outputPath$)
If PbDocument_Save(*doc, outputPath$)
  PrintN("[OK] 文档保存成功: " + outputPath$)
Else
  PrintN("[FAIL] 保存失败: " + PbDoc_GetLastError())
EndIf

; 清理
PbDocument_Free(*doc)
PbDoc_Cleanup()

PrintN("=== PbDoc 完整功能测试完成 ===")
PrintN("按任意键退出...")
Input()
CloseConsole()

; IDE Options = PureBasic 6.40 (Windows - x86)
; CursorPosition = 62
; FirstLine = 35
; EnableThread
; EnableXP
; DPIAware
; CompileSourceDirectory