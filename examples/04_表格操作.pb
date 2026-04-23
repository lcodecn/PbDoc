; ============================================================================
; 示例 04: 表格操作
; 演示: 创建表格、设置样式/对齐/列宽、填充单元格
; ============================================================================
XIncludeFile "../PbDoc.pb"

Define *doc.PbDocument
Define *table.PbDocTable
Define *cell.PbDocCell
Define outputPath.s
Define i.l, j.l

PbDoc_Init()
*doc = PbDocument_Create()
PbDocument_SetTitle(*doc, "表格操作示例")

; --- 基础表格 ---
PbDoc_Document_AddHeading(*doc, "基础表格", 1)
*table = PbDoc_Document_AddTable(*doc, 3, 4)
PbDoc_Table_SetStyle(*table, "TableGrid")
PbDoc_Cell_SetText(PbDoc_Table_GetCell(*table, 0, 0), "姓名")
PbDoc_Cell_SetText(PbDoc_Table_GetCell(*table, 0, 1), "年龄")
PbDoc_Cell_SetText(PbDoc_Table_GetCell(*table, 0, 2), "城市")
PbDoc_Cell_SetText(PbDoc_Table_GetCell(*table, 0, 3), "职业")
PbDoc_Cell_SetText(PbDoc_Table_GetCell(*table, 1, 0), "张三")
PbDoc_Cell_SetText(PbDoc_Table_GetCell(*table, 1, 1), "28")
PbDoc_Cell_SetText(PbDoc_Table_GetCell(*table, 1, 2), "北京")
PbDoc_Cell_SetText(PbDoc_Table_GetCell(*table, 1, 3), "工程师")
PbDoc_Cell_SetText(PbDoc_Table_GetCell(*table, 2, 0), "李四")
PbDoc_Cell_SetText(PbDoc_Table_GetCell(*table, 2, 1), "35")
PbDoc_Cell_SetText(PbDoc_Table_GetCell(*table, 2, 2), "上海")
PbDoc_Cell_SetText(PbDoc_Table_GetCell(*table, 2, 3), "设计师")

; --- 居中表格 ---
PbDoc_Document_AddHeading(*doc, "居中表格", 1)
*table = PbDoc_Document_AddTable(*doc, 2, 3)
PbDoc_Table_SetStyle(*table, "TableGrid")
PbDoc_Table_SetAlignment(*table, #PbDoc_ALIGN_CENTER)
PbDoc_Cell_SetText(PbDoc_Table_GetCell(*table, 0, 0), "产品")
PbDoc_Cell_SetText(PbDoc_Table_GetCell(*table, 0, 1), "价格")
PbDoc_Cell_SetText(PbDoc_Table_GetCell(*table, 0, 2), "库存")
PbDoc_Cell_SetText(PbDoc_Table_GetCell(*table, 1, 0), "笔记本")
PbDoc_Cell_SetText(PbDoc_Table_GetCell(*table, 1, 1), "5999")
PbDoc_Cell_SetText(PbDoc_Table_GetCell(*table, 1, 2), "100")

; --- 自定义列宽表格 ---
PbDoc_Document_AddHeading(*doc, "自定义列宽表格", 1)
*table = PbDoc_Document_AddTable(*doc, 2, 3)
PbDoc_Table_SetStyle(*table, "TableGrid")
; 设置列宽: 第0列5cm, 第1列3cm, 第2列4cm
PbDoc_Table_SetColumnWidth(*table, 0, PbDoc_CmToEmu(5))
PbDoc_Table_SetColumnWidth(*table, 1, PbDoc_CmToEmu(3))
PbDoc_Table_SetColumnWidth(*table, 2, PbDoc_CmToEmu(4))
PbDoc_Cell_SetText(PbDoc_Table_GetCell(*table, 0, 0), "描述（宽列）")
PbDoc_Cell_SetText(PbDoc_Table_GetCell(*table, 0, 1), "数量")
PbDoc_Cell_SetText(PbDoc_Table_GetCell(*table, 0, 2), "备注")
PbDoc_Cell_SetText(PbDoc_Table_GetCell(*table, 1, 0), "这是一段较长的描述文字")
PbDoc_Cell_SetText(PbDoc_Table_GetCell(*table, 1, 1), "50")
PbDoc_Cell_SetText(PbDoc_Table_GetCell(*table, 1, 2), "正常")

; --- 无边框表格 ---
PbDoc_Document_AddHeading(*doc, "无边框表格（默认样式）", 1)
*table = PbDoc_Document_AddTable(*doc, 2, 2)
; 不设置 TableGrid 样式即为无边框
PbDoc_Cell_SetText(PbDoc_Table_GetCell(*table, 0, 0), "项目A")
PbDoc_Cell_SetText(PbDoc_Table_GetCell(*table, 0, 1), "项目B")
PbDoc_Cell_SetText(PbDoc_Table_GetCell(*table, 1, 0), "数据1")
PbDoc_Cell_SetText(PbDoc_Table_GetCell(*table, 1, 1), "数据2")

; --- 大型表格（循环填充） ---
PbDoc_Document_AddHeading(*doc, "大型表格", 1)
*table = PbDoc_Document_AddTable(*doc, 5, 5)
PbDoc_Table_SetStyle(*table, "TableGrid")
For i = 0 To 4
  For j = 0 To 4
    PbDoc_Cell_SetText(PbDoc_Table_GetCell(*table, i, j), "R" + Str(i) + "C" + Str(j))
  Next
Next

; --- 保存 ---
outputPath = GetCurrentDirectory() + "示例04_表格操作.docx"
If PbDocument_Save(*doc, outputPath)
  Debug "保存成功: " + outputPath
Else
  Debug "保存失败: " + PbDoc_GetLastError()
EndIf

PbDocument_Free(*doc)
PbDoc_Cleanup()
Debug "示例04完成"
