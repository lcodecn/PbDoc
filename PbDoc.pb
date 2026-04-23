; ============================================================================
; PbDoc v1.3 - PureBasic Word (.docx) 文档处理库
; 参考 Python-docx 项目编写，适用于 PureBasic 平台
; 开发者: lcode.cn
; 协议: Apache 2.0
; 编译环境: 本项目在 PureBasic 6.4 (Windows x86) 环境测试通过，其他环境请自行测试
; ============================================================================
; 本文件实现创建和修改 Microsoft Word (.docx) 文件的功能
; .docx 文件本质是 ZIP 压缩包，包含 XML 格式的文档内容
; ============================================================================
; 版块目录:
;   一、编译器指令与初始化
;   二、常量定义
;   三、枚举定义
;   四、数据结构定义
;   五、全局变量
;   六、工具函数 - 单位转换
;   七、工具函数 - 错误处理
;   八、工具函数 - XML 辅助
;   九、初始化与清理
;   十、关系管理
;   十一、文档创建与释放
;   十二、文档属性设置
;   十三、段落操作
;   十四、Run 操作
;   十五、字体格式操作
;   十六、段落格式操作
;   十七、标题操作
;   十八、超链接操作
;   十九、分页符与断行操作
;   二十、表格操作
;   二十一、节与页面设置
;   二十二、文档统计
;   二十三、XML 生成 - Run 与段落
;   二十四、XML 生成 - 表格
;   二十五、XML 生成 - 分页符与断行
;   二十六、XML 生成 - 文档主体
;   二十七、XML 生成 - 完整文档
;   二十八、XML 生成 - 关系文件
;   二十九、XML 生成 - 内容类型
;   三十、XML 生成 - 核心属性
;   三十一、XML 生成 - 扩展属性
;   三十二、XML 生成 - 样式定义
;   三十三、XML 生成 - 其他静态文件
;   三十四、ZIP 包辅助函数
;   三十五、文档保存
; ============================================================================

EnableExplicit

; ============================================================================
; 一、编译器指令与初始化
; ============================================================================
UseZipPacker()

; ============================================================================
; 二、常量定义
; ============================================================================
; 版本号
#PbDoc_VERSION$ = "1.3"

; EMU (English Metric Units) 单位换算常量
; EMU 是 Open XML 中使用的通用长度单位，1 英寸 = 914400 EMU
#PbDoc_EMU_PER_INCH = 914400    ; 每英寸对应的 EMU 值
#PbDoc_EMU_PER_CM = 360000      ; 每厘米对应的 EMU 值
#PbDoc_EMU_PER_MM = 36000       ; 每毫米对应的 EMU 值
#PbDoc_EMU_PER_PT = 12700       ; 每磅对应的 EMU 值
#PbDoc_EMU_PER_TWIP = 635       ; 每缇对应的 EMU 值

; 缇 (Twips) 单位换算常量
; 缇是 Word 文档中常用的长度单位，1 英寸 = 1440 缇
#PbDoc_TWIPS_PER_INCH = 1440    ; 每英寸对应的缇值
#PbDoc_TWIPS_PER_PT = 20        ; 每磅对应的缇值
#PbDoc_TWIPS_PER_CM = 567       ; 每厘米对应的缇值

; ============================================================================
; 三、枚举定义
; ============================================================================

; --- 段落对齐方式 (对应 python-docx WD_ALIGN_PARAGRAPH) ---
Enumeration
  #PbDoc_ALIGN_LEFT = 1         ; 左对齐
  #PbDoc_ALIGN_CENTER           ; 居中对齐
  #PbDoc_ALIGN_RIGHT            ; 右对齐
  #PbDoc_ALIGN_JUSTIFY          ; 两端对齐
  #PbDoc_ALIGN_DISTRIBUTE       ; 分散对齐
EndEnumeration

; --- Tab 对齐方式 (对应 python-docx WD_TAB_ALIGNMENT) ---
Enumeration
  #PbDoc_TAB_LEFT = 1           ; 左对齐制表符
  #PbDoc_TAB_CENTER             ; 居中对齐制表符
  #PbDoc_TAB_RIGHT              ; 右对齐制表符
  #PbDoc_TAB_DECIMAL            ; 小数点对齐制表符
  #PbDoc_TAB_BAR                ; 竖线制表符
  #PbDoc_TAB_CLEAR              ; 清除制表符
  #PbDoc_TAB_NUM                ; 编号制表符
EndEnumeration

; --- Tab 前导符类型 (对应 python-docx WD_TAB_LEADER) ---
Enumeration
  #PbDoc_TAB_LEADER_NONE = 1    ; 无前导符
  #PbDoc_TAB_LEADER_DOTS        ; 点线前导符 (......)
  #PbDoc_TAB_LEADER_HEAVY       ; 粗点线前导符
  #PbDoc_TAB_LEADER_DASHES      ; 短划线前导符 (------)
  #PbDoc_TAB_LEADER_LINES       ; 下划线前导符 (______)
EndEnumeration

; --- Body 元素类型 (内部使用) ---
Enumeration
  #PbDoc_ELEM_PARAGRAPH = 1     ; 段落元素
  #PbDoc_ELEM_TABLE             ; 表格元素
  #PbDoc_ELEM_PAGE_BREAK        ; 分页符元素
EndEnumeration

; --- 断行类型 (对应 python-docx WD_BREAK) ---
Enumeration
  #PbDoc_BREAK_LINE = 1         ; 换行符 (软回车)
  #PbDoc_BREAK_PAGE             ; 分页符
  #PbDoc_BREAK_COLUMN           ; 分栏符
EndEnumeration

; --- 下划线类型 (对应 python-docx WD_UNDERLINE) ---
Enumeration
  #PbDoc_UNDERLINE_NONE = 0     ; 无下划线
  #PbDoc_UNDERLINE_SINGLE = 1   ; 单下划线
  #PbDoc_UNDERLINE_DOUBLE = 2   ; 双下划线
  #PbDoc_UNDERLINE_DOTTED = 3   ; 点线下划线
  #PbDoc_UNDERLINE_DASH = 4     ; 短划线下划线
  #PbDoc_UNDERLINE_WAVE = 5     ; 波浪线下划线
  #PbDoc_UNDERLINE_THICK = 6    ; 粗下划线
EndEnumeration

; --- 行间距规则 ---
Enumeration
  #PbDoc_LINE_SPACING_AUTO = 1  ; 自动行距 (多倍行距)
  #PbDoc_LINE_SPACING_EXACT = 2 ; 精确行距 (固定值)
  #PbDoc_LINE_SPACING_AT_LEAST = 3 ; 最小行距
EndEnumeration

; --- 页面方向 ---
Enumeration
  #PbDoc_ORIENT_PORTRAIT = 1    ; 纵向
  #PbDoc_ORIENT_LANDSCAPE       ; 横向
EndEnumeration

; ============================================================================
; 四、数据结构定义
; ============================================================================

; --- RGB 颜色结构 ---
Structure PbDocRGB
  r.a                ; 红色分量 (0-255)
  g.a                ; 绿色分量 (0-255)
  b.a                ; 蓝色分量 (0-255)
EndStructure

; --- Run 数据结构 - 存储文本运行信息 ---
Structure PbDocRun
  text.s             ; Run 文本内容
  bold.l             ; 粗体标志 (#True/#False)
  italic.l           ; 斜体标志 (#True/#False)
  underline.l        ; 下划线类型 (#PbDoc_UNDERLINE_*)
  strike.l           ; 删除线标志 (#True/#False)
  doubleStrike.l     ; 双删除线标志 (#True/#False)
  superscript.l      ; 上标标志 (#True/#False)
  subscript.l        ; 下标标志 (#True/#False)
  fontName.s         ; 字体名称 (如 "宋体", "Arial")
  colorHex.s         ; 字体颜色 (十六进制, 如 "FF0000" 表示红色)
  highlight.s        ; 高亮颜色 (如 "yellow", "green")
  fontSize.l         ; 字体大小 (半磅单位, 如 24 = 12pt)
  hasTab.l           ; 是否包含 Tab 字符
  breakType.l        ; 断行类型 (0=无, #PbDoc_BREAK_*)
EndStructure

; --- Tab Stop 数据结构 ---
Structure PbDocTabStop
  position.q         ; 位置 (缇, twips)
  alignment.l        ; 对齐方式 (#PbDoc_TAB_*)
  leader.l           ; 前导符类型 (#PbDoc_TAB_LEADER_*)
EndStructure

; --- 段落数据结构 ---
Structure PbDocParagraph
  text.s             ; 段落初始文本
  alignment.l        ; 对齐方式 (#PbDoc_ALIGN_*)
  style.s            ; 样式名称 (如 "Heading1", "Normal")
  firstLineIndent.q  ; 首行缩进 (缇)
  leftIndent.q       ; 左缩进 (缇)
  rightIndent.q      ; 右缩进 (缇)
  hangingIndent.q    ; 悬挂缩进 (缇)
  spaceBefore.q      ; 段前间距 (缇)
  spaceAfter.q       ; 段后间距 (缇)
  lineSpacing.q      ; 行间距值
  lineSpacingRule.l  ; 行间距规则 (#PbDoc_LINE_SPACING_*)
  keepNext.l         ; 与下段同页标志
  keepLines.l        ; 段中不分页标志
  widowControl.l     ; 孤行控制标志
  List runs.PbDocRun()        ; Run 列表
  List tabStops.PbDocTabStop() ; Tab Stop 列表
  isHyperlink.l      ; 是否为超链接段落
  hyperlinkUrl.s     ; 超链接 URL
  hyperlinkRid.s     ; 超链接关系 ID
EndStructure

; --- 表格数据结构 ---
Structure PbDocTable
  rows.l             ; 行数
  cols.l             ; 列数
  style.s            ; 表格样式名称 (如 "TableGrid")
  alignment.l        ; 表格对齐方式 (#PbDoc_ALIGN_*)
  List colWidths.q() ; 列宽列表 (缇)
  Map cellText.s()   ; 单元格文本, 键格式 "row,col"
EndStructure

; --- 单元格引用结构 - 用于 PbDoc_Table_GetCell 返回值 ---
Structure PbDocCell
  *table.PbDocTable  ; 所属表格指针
  row.l              ; 行索引
  col.l              ; 列索引
EndStructure

; --- Body 元素结构 - 用于按顺序存储文档内容 ---
Structure PbDocBodyElement
  type.l             ; 元素类型 (#PbDoc_ELEM_*)
  paraIndex.l        ; 段落在 paragraphs 列表中的索引
  tableIndex.l       ; 表格在 tables 列表中的索引
EndStructure

; --- 关系数据结构 - 用于管理 OPC 关系 ---
Structure PbDocRel
  rid.s              ; 关系 ID (如 "rId10")
  relType.s          ; 关系类型 URL
  target.s           ; 目标路径
EndStructure

; --- 节属性数据结构 - 页面设置 ---
Structure PbDocSection
  pageWidth.q        ; 页面宽度 (缇)
  pageHeight.q       ; 页面高度 (缇)
  marginTop.q        ; 上边距 (缇)
  marginBottom.q     ; 下边距 (缇)
  marginLeft.q       ; 左边距 (缇)
  marginRight.q      ; 右边距 (缇)
  headerDist.q       ; 页眉距离 (缇)
  footerDist.q       ; 页脚距离 (缇)
  gutter.q           ; 装订线 (缇)
  orientation.l      ; 页面方向 (#PbDoc_ORIENT_*)
EndStructure

; --- 文档数据结构 - 核心文档对象 ---
Structure PbDocument
  title.s            ; 文档标题
  author.s           ; 文档作者
  subject.s          ; 文档主题
  company.s          ; 公司名称
  List paragraphs.PbDocParagraph()  ; 段落列表
  List tables.PbDocTable()          ; 表格列表
  List bodyOrder.PbDocBodyElement() ; Body 元素顺序列表
  section.PbDocSection              ; 节属性 (页面设置)
  paragraphCount.l   ; 段落计数
  tableCount.l       ; 表格计数
  nextRid.l          ; 下一个关系 ID 编号
  List rels.PbDocRel()             ; 关系列表
EndStructure

; ============================================================================
; 五、全局变量
; ============================================================================
Global g_pbdocInitialized.l = #False   ; 库初始化标志
Global g_pbdocLastError.s = ""         ; 最近一次错误信息
Global g_pbdocTempDir.s = ""           ; 临时目录路径
Global g_q.s = Chr(34)                 ; 双引号字符, 用于 XML 属性拼接
Global g_pbdocCell.PbDocCell           ; 单元格引用, 用于 PbDoc_Table_GetCell

; ============================================================================
; 六、工具函数 - 单位转换
; ============================================================================
; 这些函数用于在不同长度单位之间进行转换
; EMU 是 Open XML 内部使用的通用单位，缇是 Word XML 中常用的单位

; 英寸转 EMU
Procedure.q PbDoc_InchesToEmu(inches.d)
  ProcedureReturn Round(inches * #PbDoc_EMU_PER_INCH, #PB_Round_Nearest)
EndProcedure

; 厘米转 EMU
Procedure.q PbDoc_CmToEmu(cm.d)
  ProcedureReturn Round(cm * #PbDoc_EMU_PER_CM, #PB_Round_Nearest)
EndProcedure

; 毫米转 EMU
Procedure.q PbDoc_MmToEmu(mm.d)
  ProcedureReturn Round(mm * #PbDoc_EMU_PER_MM, #PB_Round_Nearest)
EndProcedure

; 磅转 EMU
Procedure.q PbDoc_PtToEmu(pt.d)
  ProcedureReturn Round(pt * #PbDoc_EMU_PER_PT, #PB_Round_Nearest)
EndProcedure

; 缇转 EMU
Procedure.q PbDoc_TwipsToEmu(twips.q)
  ProcedureReturn twips * #PbDoc_EMU_PER_TWIP
EndProcedure

; EMU 转缇
Procedure.q PbDoc_EmuToTwips(emu.q)
  ProcedureReturn emu / #PbDoc_EMU_PER_TWIP
EndProcedure

; 磅转缇
Procedure.q PbDoc_PtToTwips(pt.d)
  ProcedureReturn Round(pt * #PbDoc_TWIPS_PER_PT, #PB_Round_Nearest)
EndProcedure

; 厘米转缇
Procedure.q PbDoc_CmToTwips(cm.d)
  ProcedureReturn Round(cm * #PbDoc_TWIPS_PER_CM, #PB_Round_Nearest)
EndProcedure

; 英寸转缇
Procedure.q PbDoc_InchesToTwips(inches.d)
  ProcedureReturn Round(inches * #PbDoc_TWIPS_PER_INCH, #PB_Round_Nearest)
EndProcedure

; ============================================================================
; 七、工具函数 - 错误处理
; ============================================================================

; 获取最近一次错误信息
Procedure.s PbDoc_GetLastError()
  ProcedureReturn g_pbdocLastError
EndProcedure

; 设置错误信息 (内部使用)
Procedure PbDoc_SetError(errorMsg.s)
  g_pbdocLastError = errorMsg
EndProcedure

; ============================================================================
; 八、工具函数 - XML 辅助
; ============================================================================

; XML 转义特殊字符
; 将 &, <, >, ", ' 等字符转换为 XML 实体
Procedure.s PbDoc_XmlEscape(text.s)
  Protected result.s
  result = ReplaceString(text, "&", "&amp;")
  result = ReplaceString(result, "<", "&lt;")
  result = ReplaceString(result, ">", "&gt;")
  result = ReplaceString(result, g_q, "&quot;")
  result = ReplaceString(result, "'", "&apos;")
  ProcedureReturn result
EndProcedure

; 生成 XML 属性字符串: name="value"
; 此函数简化了 XML 属性的拼接操作
Procedure.s PbDoc_A(name.s, value.s)
  ProcedureReturn " " + name + "=" + g_q + value + g_q
EndProcedure

; 将段落对齐方式枚举值转换为 XML 属性值
Procedure.s PbDoc_AlignmentToXml(alignment.l)
  Select alignment
    Case #PbDoc_ALIGN_LEFT      : ProcedureReturn "left"
    Case #PbDoc_ALIGN_CENTER    : ProcedureReturn "center"
    Case #PbDoc_ALIGN_RIGHT     : ProcedureReturn "right"
    Case #PbDoc_ALIGN_JUSTIFY   : ProcedureReturn "both"
    Case #PbDoc_ALIGN_DISTRIBUTE : ProcedureReturn "distribute"
    Default                      : ProcedureReturn "left"
  EndSelect
EndProcedure

; 将 Tab 对齐方式枚举值转换为 XML 属性值
Procedure.s PbDoc_TabAlignToXml(alignment.l)
  Select alignment
    Case #PbDoc_TAB_LEFT   : ProcedureReturn "left"
    Case #PbDoc_TAB_CENTER : ProcedureReturn "center"
    Case #PbDoc_TAB_RIGHT  : ProcedureReturn "right"
    Case #PbDoc_TAB_DECIMAL : ProcedureReturn "decimal"
    Case #PbDoc_TAB_BAR    : ProcedureReturn "bar"
    Case #PbDoc_TAB_NUM    : ProcedureReturn "num"
    Default                : ProcedureReturn "left"
  EndSelect
EndProcedure

; 将 Tab 前导符枚举值转换为 XML 属性值
Procedure.s PbDoc_TabLeaderToXml(leader.l)
  Select leader
    Case #PbDoc_TAB_LEADER_NONE   : ProcedureReturn "none"
    Case #PbDoc_TAB_LEADER_DOTS   : ProcedureReturn "dot"
    Case #PbDoc_TAB_LEADER_HEAVY  : ProcedureReturn "heavy"
    Case #PbDoc_TAB_LEADER_DASHES : ProcedureReturn "hyphen"
    Case #PbDoc_TAB_LEADER_LINES  : ProcedureReturn "underscore"
    Default                        : ProcedureReturn "none"
  EndSelect
EndProcedure

; 将下划线类型枚举值转换为 XML 属性值
Procedure.s PbDoc_UnderlineToXml(underline.l)
  Select underline
    Case #PbDoc_UNDERLINE_SINGLE : ProcedureReturn "single"
    Case #PbDoc_UNDERLINE_DOUBLE : ProcedureReturn "double"
    Case #PbDoc_UNDERLINE_DOTTED : ProcedureReturn "dotted"
    Case #PbDoc_UNDERLINE_DASH   : ProcedureReturn "dash"
    Case #PbDoc_UNDERLINE_WAVE   : ProcedureReturn "wave"
    Case #PbDoc_UNDERLINE_THICK  : ProcedureReturn "thick"
    Default                       : ProcedureReturn "single"
  EndSelect
EndProcedure

; 将行间距规则枚举值转换为 XML 属性值
Procedure.s PbDoc_LineSpacingRuleToXml(rule.l)
  Select rule
    Case #PbDoc_LINE_SPACING_AUTO     : ProcedureReturn "auto"
    Case #PbDoc_LINE_SPACING_EXACT    : ProcedureReturn "exact"
    Case #PbDoc_LINE_SPACING_AT_LEAST : ProcedureReturn "atLeast"
    Default                            : ProcedureReturn "auto"
  EndSelect
EndProcedure

; ============================================================================
; 九、初始化与清理
; ============================================================================

; 初始化 PbDoc 库
; 在使用任何 PbDoc 功能之前必须调用此函数
; 返回值: #True 成功, #False 失败
Procedure.l PbDoc_Init()
  If g_pbdocInitialized
    ProcedureReturn #True
  EndIf
  g_pbdocTempDir = GetTemporaryDirectory() + "PbDoc_" + Str(GetTickCount_()) + "\"
  CreateDirectory(g_pbdocTempDir)
  g_pbdocInitialized = #True
  g_pbdocLastError = ""
  ProcedureReturn #True
EndProcedure

; 清理 PbDoc 库资源
; 在程序结束前调用此函数释放临时资源
Procedure PbDoc_Cleanup()
  If g_pbdocTempDir <> ""
    DeleteDirectory(g_pbdocTempDir, "", #PB_FileSystem_Recursive)
    g_pbdocTempDir = ""
  EndIf
  g_pbdocInitialized = #False
EndProcedure

; ============================================================================
; 十、关系管理
; ============================================================================
; OPC (Open Packaging Convention) 关系是 .docx 文件中各部件之间的关联

; 添加关系并返回 rId
; 参数 relType: 关系类型 URL
; 参数 target: 目标路径
; 返回值: 新生成的关系 ID (如 "rId9")
Procedure.s PbDoc_AddRelationship(*doc.PbDocument, relType.s, target.s)
  Protected rid.s
  *doc\nextRid + 1
  rid = "rId" + Str(*doc\nextRid)
  AddElement(*doc\rels())
  *doc\rels()\rid = rid
  *doc\rels()\relType = relType
  *doc\rels()\target = target
  ProcedureReturn rid
EndProcedure

; ============================================================================
; 十一、文档创建与释放
; ============================================================================

; 创建新的空白文档
; 返回值: 文档对象指针, 失败返回 0
Procedure PbDocument_Create()
  Protected *doc.PbDocument
  *doc = AllocateMemory(SizeOf(PbDocument))
  If *doc = 0
    PbDoc_SetError("无法分配文档内存")
    ProcedureReturn 0
  EndIf
  InitializeStructure(*doc, PbDocument)
  ; 设置默认属性
  *doc\title = ""
  *doc\author = "PbDoc"
  *doc\subject = ""
  *doc\company = ""
  *doc\paragraphCount = 0
  *doc\tableCount = 0
  *doc\nextRid = 8
  ; 设置默认页面设置 (Letter 纵向)
  *doc\section\pageWidth = 12240   ; 8.5 英寸
  *doc\section\pageHeight = 15840  ; 11 英寸
  *doc\section\marginTop = 1440    ; 1 英寸
  *doc\section\marginBottom = 1440
  *doc\section\marginLeft = 1800   ; 1.25 英寸
  *doc\section\marginRight = 1800
  *doc\section\headerDist = 720    ; 0.5 英寸
  *doc\section\footerDist = 720
  *doc\section\gutter = 0
  *doc\section\orientation = #PbDoc_ORIENT_PORTRAIT
  ProcedureReturn *doc
EndProcedure

; 释放文档对象
; 释放文档占用的所有内存资源
Procedure PbDocument_Free(*doc.PbDocument)
  If *doc = 0 : ProcedureReturn : EndIf
  ClearStructure(*doc, PbDocument)
  FreeMemory(*doc)
EndProcedure

; ============================================================================
; 十二、文档属性设置
; ============================================================================
; 这些属性对应 docProps/core.xml 和 docProps/app.xml 中的元数据

; 设置文档标题
Procedure PbDocument_SetTitle(*doc.PbDocument, title.s)
  If *doc : *doc\title = title : EndIf
EndProcedure

; 设置文档作者
Procedure PbDocument_SetAuthor(*doc.PbDocument, author.s)
  If *doc : *doc\author = author : EndIf
EndProcedure

; 设置文档主题
Procedure PbDocument_SetSubject(*doc.PbDocument, subject.s)
  If *doc : *doc\subject = subject : EndIf
EndProcedure

; 设置公司名称
Procedure PbDocument_SetCompany(*doc.PbDocument, company.s)
  If *doc : *doc\company = company : EndIf
EndProcedure

; ============================================================================
; 十三、段落操作
; ============================================================================

; 添加段落
; 参数 text: 段落文本内容 (可为空字符串)
; 返回值: 段落对象指针
Procedure PbDoc_Document_AddParagraph(*doc.PbDocument, text.s)
  Protected *para.PbDocParagraph
  If *doc = 0 : ProcedureReturn 0 : EndIf
  ; 创建段落数据
  AddElement(*doc\paragraphs())
  *para = @*doc\paragraphs()
  *para\text = text
  *para\alignment = 0
  *para\style = ""
  *para\firstLineIndent = 0
  *para\leftIndent = 0
  *para\rightIndent = 0
  *para\hangingIndent = 0
  *para\spaceBefore = 0
  *para\spaceAfter = 0
  *para\lineSpacing = 0
  *para\lineSpacingRule = 0
  *para\keepNext = #False
  *para\keepLines = #False
  *para\widowControl = #False
  *para\isHyperlink = #False
  *para\hyperlinkUrl = ""
  *para\hyperlinkRid = ""
  ; 如果有初始文本，创建默认 Run
  If text <> ""
    AddElement(*para\runs())
    *para\runs()\text = text
    *para\runs()\bold = #False
    *para\runs()\italic = #False
    *para\runs()\underline = #PbDoc_UNDERLINE_NONE
    *para\runs()\strike = #False
    *para\runs()\doubleStrike = #False
    *para\runs()\superscript = #False
    *para\runs()\subscript = #False
    *para\runs()\fontName = ""
    *para\runs()\colorHex = ""
    *para\runs()\highlight = ""
    *para\runs()\fontSize = 0
    *para\runs()\hasTab = #False
    *para\runs()\breakType = 0
  EndIf
  ; 添加到 Body 顺序列表
  AddElement(*doc\bodyOrder())
  *doc\bodyOrder()\type = #PbDoc_ELEM_PARAGRAPH
  *doc\bodyOrder()\paraIndex = *doc\paragraphCount
  *doc\bodyOrder()\tableIndex = -1
  *doc\paragraphCount + 1
  ProcedureReturn *para
EndProcedure

; 获取段落文本
; 拼接段落中所有 Run 的文本内容
Procedure.s PbDoc_Paragraph_GetText(*para.PbDocParagraph)
  Protected result.s
  If *para = 0 : ProcedureReturn "" : EndIf
  If ListSize(*para\runs()) = 0
    ProcedureReturn *para\text
  EndIf
  ForEach *para\runs()
    result + *para\runs()\text
  Next
  ProcedureReturn result
EndProcedure

; 设置段落对齐方式
Procedure PbDoc_Paragraph_SetAlignment(*para.PbDocParagraph, alignment.l)
  If *para : *para\alignment = alignment : EndIf
EndProcedure

; 设置段落样式
; 参数 styleName: 样式 ID (如 "Heading1", "Normal", "Title")
Procedure PbDoc_Paragraph_SetStyle(*para.PbDocParagraph, styleName.s)
  If *para : *para\style = styleName : EndIf
EndProcedure

; ============================================================================
; 十四、Run 操作
; ============================================================================

; 添加 Run 到段落
; Run 是段落中具有相同格式的一段文本
; 返回值: Run 对象指针
Procedure PbDoc_Paragraph_AddRun(*para.PbDocParagraph, text.s)
  Protected *run.PbDocRun
  If *para = 0 : ProcedureReturn 0 : EndIf
  AddElement(*para\runs())
  *run = @*para\runs()
  *run\text = text
  *run\bold = #False
  *run\italic = #False
  *run\underline = #PbDoc_UNDERLINE_NONE
  *run\strike = #False
  *run\doubleStrike = #False
  *run\superscript = #False
  *run\subscript = #False
  *run\fontName = ""
  *run\colorHex = ""
  *run\highlight = ""
  *run\fontSize = 0
  *run\hasTab = #False
  *run\breakType = 0
  ProcedureReturn *run
EndProcedure

; Run 添加 Tab 字符
Procedure PbDoc_Run_AddTab(*run.PbDocRun)
  If *run : *run\hasTab = #True : EndIf
EndProcedure

; Run 添加断行符
; 参数 breakType: 断行类型 (#PbDoc_BREAK_LINE, #PbDoc_BREAK_PAGE, #PbDoc_BREAK_COLUMN)
Procedure PbDoc_Run_AddBreak(*run.PbDocRun, breakType.l)
  If *run : *run\breakType = breakType : EndIf
EndProcedure

; ============================================================================
; 十五、字体格式操作
; ============================================================================
; 这些函数用于设置 Run 的字体属性

; 设置粗体
Procedure PbDoc_Font_SetBold(*run.PbDocRun, bold.l)
  If *run : *run\bold = bold : EndIf
EndProcedure

; 设置斜体
Procedure PbDoc_Font_SetItalic(*run.PbDocRun, italic.l)
  If *run : *run\italic = italic : EndIf
EndProcedure

; 设置下划线
; 参数 underline: 下划线类型 (#PbDoc_UNDERLINE_*)
Procedure PbDoc_Font_SetUnderline(*run.PbDocRun, underline.l)
  If *run : *run\underline = underline : EndIf
EndProcedure

; 设置删除线
Procedure PbDoc_Font_SetStrike(*run.PbDocRun, strike.l)
  If *run : *run\strike = strike : EndIf
EndProcedure

; 设置双删除线
Procedure PbDoc_Font_SetDoubleStrike(*run.PbDocRun, doubleStrike.l)
  If *run : *run\doubleStrike = doubleStrike : EndIf
EndProcedure

; 设置上标
Procedure PbDoc_Font_SetSuperscript(*run.PbDocRun, superscript.l)
  If *run : *run\superscript = superscript : EndIf
EndProcedure

; 设置下标
Procedure PbDoc_Font_SetSubscript(*run.PbDocRun, subscript.l)
  If *run : *run\subscript = subscript : EndIf
EndProcedure

; 设置字体名称
; 参数 fontName: 字体名称 (如 "宋体", "Arial", "Times New Roman")
Procedure PbDoc_Font_SetName(*run.PbDocRun, fontName.s)
  If *run : *run\fontName = fontName : EndIf
EndProcedure

; 设置字体颜色 (十六进制)
; 参数 colorHex: 颜色十六进制值 (如 "FF0000" 红色, "0000FF" 蓝色)
Procedure PbDoc_Font_SetColorHex(*run.PbDocRun, colorHex.s)
  If *run : *run\colorHex = colorHex : EndIf
EndProcedure

; 设置字体颜色 (RGB)
Procedure PbDoc_Font_SetColorRGB(*run.PbDocRun, r.a, g.a, b.a)
  If *run = 0 : ProcedureReturn : EndIf
  *run\colorHex = RSet(Hex(r), 2, "0") + RSet(Hex(g), 2, "0") + RSet(Hex(b), 2, "0")
EndProcedure

; 设置高亮颜色
; 参数 highlight: 高亮颜色名称 (如 "yellow", "green", "cyan", "red")
Procedure PbDoc_Font_SetHighlight(*run.PbDocRun, highlight.s)
  If *run : *run\highlight = highlight : EndIf
EndProcedure

; 设置字体大小
; 参数 size: 字体大小 (磅值, 如 12 表示 12pt)
Procedure PbDoc_Font_SetSize(*run.PbDocRun, size.l)
  If *run : *run\fontSize = size * 2 : EndIf
EndProcedure

; ============================================================================
; 十六、段落格式操作
; ============================================================================
; 这些函数用于设置段落的缩进、间距、分页控制等格式属性

; 设置首行缩进 (EMU 单位)
Procedure PbDoc_ParaFormat_SetFirstLineIndent(*para.PbDocParagraph, emu.q)
  If *para : *para\firstLineIndent = PbDoc_EmuToTwips(emu) : EndIf
EndProcedure

; 设置左缩进 (EMU 单位)
Procedure PbDoc_ParaFormat_SetLeftIndent(*para.PbDocParagraph, emu.q)
  If *para : *para\leftIndent = PbDoc_EmuToTwips(emu) : EndIf
EndProcedure

; 设置右缩进 (EMU 单位)
Procedure PbDoc_ParaFormat_SetRightIndent(*para.PbDocParagraph, emu.q)
  If *para : *para\rightIndent = PbDoc_EmuToTwips(emu) : EndIf
EndProcedure

; 设置悬挂缩进 (EMU 单位)
Procedure PbDoc_ParaFormat_SetHangingIndent(*para.PbDocParagraph, emu.q)
  If *para : *para\hangingIndent = PbDoc_EmuToTwips(emu) : EndIf
EndProcedure

; 设置段前间距 (EMU 单位)
Procedure PbDoc_ParaFormat_SetSpaceBefore(*para.PbDocParagraph, emu.q)
  If *para : *para\spaceBefore = PbDoc_EmuToTwips(emu) : EndIf
EndProcedure

; 设置段后间距 (EMU 单位)
Procedure PbDoc_ParaFormat_SetSpaceAfter(*para.PbDocParagraph, emu.q)
  If *para : *para\spaceAfter = PbDoc_EmuToTwips(emu) : EndIf
EndProcedure

; 设置行间距
; 参数 spacing: 行间距值
;   - 自动行距: 240 = 1倍行距, 360 = 1.5倍, 480 = 2倍
;   - 精确行距/最小行距: 缇值 (如 PbDoc_PtToTwips(20) = 20pt固定行距)
; 参数 rule: 行间距规则 (#PbDoc_LINE_SPACING_*)
Procedure PbDoc_ParaFormat_SetLineSpacing(*para.PbDocParagraph, spacing.q, rule.l)
  If *para
    *para\lineSpacing = spacing
    *para\lineSpacingRule = rule
  EndIf
EndProcedure

; 设置与下段同页
; 防止当前段落与下一段落之间出现分页符
Procedure PbDoc_ParaFormat_SetKeepNext(*para.PbDocParagraph, keepNext.l)
  If *para : *para\keepNext = keepNext : EndIf
EndProcedure

; 设置段中不分页
; 防止段落内部出现分页符
Procedure PbDoc_ParaFormat_SetKeepLines(*para.PbDocParagraph, keepLines.l)
  If *para : *para\keepLines = keepLines : EndIf
EndProcedure

; 设置孤行控制
; 防止段落首行或末行单独出现在页面顶部或底部
Procedure PbDoc_ParaFormat_SetWidowControl(*para.PbDocParagraph, widowControl.l)
  If *para : *para\widowControl = widowControl : EndIf
EndProcedure

; 添加 Tab Stop
; 参数 position: 位置 (EMU 单位)
; 参数 alignment: 对齐方式 (#PbDoc_TAB_*)
; 参数 leader: 前导符类型 (#PbDoc_TAB_LEADER_*)
Procedure PbDoc_ParaFormat_AddTabStop(*para.PbDocParagraph, position.q, alignment.l, leader.l)
  If *para = 0 : ProcedureReturn : EndIf
  AddElement(*para\tabStops())
  *para\tabStops()\position = PbDoc_EmuToTwips(position)
  *para\tabStops()\alignment = alignment
  *para\tabStops()\leader = leader
EndProcedure

; ============================================================================
; 十七、标题操作
; ============================================================================

; 添加标题段落
; 参数 text: 标题文本
; 参数 level: 标题级别 (1-9, 1 为最高级)
; 返回值: 段落对象指针
Procedure PbDoc_Document_AddHeading(*doc.PbDocument, text.s, level.l)
  Protected *para.PbDocParagraph
  If *doc = 0 : ProcedureReturn 0 : EndIf
  If level < 1 : level = 1 : EndIf
  If level > 9 : level = 9 : EndIf
  *para = PbDoc_Document_AddParagraph(*doc, text)
  If *para
    *para\style = "Heading" + Str(level)
  EndIf
  ProcedureReturn *para
EndProcedure

; ============================================================================
; 十八、超链接操作
; ============================================================================

; 添加超链接
; 参数 url: 超链接目标 URL
; 参数 text: 超链接显示文本
; 返回值: 段落对象指针
Procedure PbDoc_Document_AddHyperlink(*doc.PbDocument, url.s, text.s)
  Protected *para.PbDocParagraph
  Protected rid.s
  If *doc = 0 : ProcedureReturn 0 : EndIf
  *para = PbDoc_Document_AddParagraph(*doc, "")
  If *para
    *para\isHyperlink = #True
    *para\hyperlinkUrl = url
    rid = PbDoc_AddRelationship(*doc, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", url)
    *para\hyperlinkRid = rid
    If text <> ""
      AddElement(*para\runs())
      *para\runs()\text = text
      *para\runs()\bold = #False
      *para\runs()\italic = #False
      *para\runs()\underline = #PbDoc_UNDERLINE_NONE
      *para\runs()\strike = #False
      *para\runs()\doubleStrike = #False
      *para\runs()\superscript = #False
      *para\runs()\subscript = #False
      *para\runs()\fontName = ""
      *para\runs()\colorHex = ""
      *para\runs()\highlight = ""
      *para\runs()\fontSize = 0
      *para\runs()\hasTab = #False
      *para\runs()\breakType = 0
    EndIf
  EndIf
  ProcedureReturn *para
EndProcedure

; ============================================================================
; 十九、分页符与断行操作
; ============================================================================

; 添加分页符
; 在文档中插入一个分页符，后续内容从新页开始
Procedure PbDoc_Document_AddPageBreak(*doc.PbDocument)
  If *doc = 0 : ProcedureReturn : EndIf
  AddElement(*doc\bodyOrder())
  *doc\bodyOrder()\type = #PbDoc_ELEM_PAGE_BREAK
  *doc\bodyOrder()\paraIndex = -1
  *doc\bodyOrder()\tableIndex = -1
EndProcedure

; ============================================================================
; 二十、表格操作
; ============================================================================

; 添加表格
; 参数 rows: 行数
; 参数 cols: 列数
; 返回值: 表格对象指针
Procedure PbDoc_Document_AddTable(*doc.PbDocument, rows.l, cols.l)
  Protected *tbl.PbDocTable
  Protected i.l
  If *doc = 0 : ProcedureReturn 0 : EndIf
  If rows < 1 Or cols < 1 : ProcedureReturn 0 : EndIf
  AddElement(*doc\tables())
  *tbl = @*doc\tables()
  *tbl\rows = rows
  *tbl\cols = cols
  *tbl\style = ""
  *tbl\alignment = 0
  ; 设置默认列宽 (每列 2000 缇 ≈ 3.5cm)
  For i = 1 To cols
    AddElement(*tbl\colWidths())
    *tbl\colWidths() = 2000
  Next
  ; 添加到 Body 顺序列表
  AddElement(*doc\bodyOrder())
  *doc\bodyOrder()\type = #PbDoc_ELEM_TABLE
  *doc\bodyOrder()\paraIndex = -1
  *doc\bodyOrder()\tableIndex = *doc\tableCount
  *doc\tableCount + 1
  ProcedureReturn *tbl
EndProcedure

; 设置表格样式
; 参数 style: 样式名称 (如 "TableGrid" 表示带边框的表格)
Procedure PbDoc_Table_SetStyle(*tbl.PbDocTable, style.s)
  If *tbl : *tbl\style = style : EndIf
EndProcedure

; 设置表格对齐方式
Procedure PbDoc_Table_SetAlignment(*tbl.PbDocTable, alignment.l)
  If *tbl : *tbl\alignment = alignment : EndIf
EndProcedure

; 设置列宽
; 参数 colIndex: 列索引 (从 0 开始)
; 参数 widthEmu: 列宽 (EMU 单位)
Procedure PbDoc_Table_SetColumnWidth(*tbl.PbDocTable, colIndex.l, widthEmu.q)
  If *tbl = 0 : ProcedureReturn : EndIf
  If colIndex < 0 Or colIndex >= *tbl\cols : ProcedureReturn : EndIf
  If SelectElement(*tbl\colWidths(), colIndex)
    *tbl\colWidths() = PbDoc_EmuToTwips(widthEmu)
  EndIf
EndProcedure

; 获取单元格引用
; 返回值: 单元格对象指针, 用于 PbDoc_Cell_SetText
Procedure PbDoc_Table_GetCell(*tbl.PbDocTable, row.l, col.l)
  If *tbl = 0 : ProcedureReturn 0 : EndIf
  If row < 0 Or row >= *tbl\rows Or col < 0 Or col >= *tbl\cols : ProcedureReturn 0 : EndIf
  g_pbdocCell\table = *tbl
  g_pbdocCell\row = row
  g_pbdocCell\col = col
  ProcedureReturn @g_pbdocCell
EndProcedure

; 设置单元格文本
Procedure PbDoc_Cell_SetText(*cell.PbDocCell, text.s)
  Protected key.s
  If *cell = 0 : ProcedureReturn : EndIf
  key = Str(*cell\row) + "," + Str(*cell\col)
  *cell\table\cellText(key) = text
EndProcedure

; ============================================================================
; 二十一、节与页面设置
; ============================================================================
; 这些函数用于设置文档的页面属性 (大小、边距、方向等)

; 设置页面大小 (缇单位)
Procedure PbDoc_Section_SetPageSize(*doc.PbDocument, widthTwips.q, heightTwips.q)
  If *doc = 0 : ProcedureReturn : EndIf
  *doc\section\pageWidth = widthTwips
  *doc\section\pageHeight = heightTwips
EndProcedure

; 设置页面边距 (缇单位)
Procedure PbDoc_Section_SetMargins(*doc.PbDocument, top.q, bottom.q, left.q, right.q)
  If *doc = 0 : ProcedureReturn : EndIf
  *doc\section\marginTop = top
  *doc\section\marginBottom = bottom
  *doc\section\marginLeft = left
  *doc\section\marginRight = right
EndProcedure

; 设置页面方向
Procedure PbDoc_Section_SetOrientation(*doc.PbDocument, orientation.l)
  If *doc = 0 : ProcedureReturn : EndIf
  *doc\section\orientation = orientation
  ; 横向时交换宽高
  If orientation = #PbDoc_ORIENT_LANDSCAPE
    If *doc\section\pageWidth < *doc\section\pageHeight
      Protected tmp.q
      tmp = *doc\section\pageWidth
      *doc\section\pageWidth = *doc\section\pageHeight
      *doc\section\pageHeight = tmp
    EndIf
  EndIf
EndProcedure

; 设置页眉距离 (缇单位)
Procedure PbDoc_Section_SetHeaderDistance(*doc.PbDocument, dist.q)
  If *doc : *doc\section\headerDist = dist : EndIf
EndProcedure

; 设置页脚距离 (缇单位)
Procedure PbDoc_Section_SetFooterDistance(*doc.PbDocument, dist.q)
  If *doc : *doc\section\footerDist = dist : EndIf
EndProcedure

; 设置装订线 (缇单位)
Procedure PbDoc_Section_SetGutter(*doc.PbDocument, gutter.q)
  If *doc : *doc\section\gutter = gutter : EndIf
EndProcedure

; ============================================================================
; 二十二、文档统计
; ============================================================================

; 获取段落数量
Procedure.l PbDoc_Document_GetParagraphCount(*doc.PbDocument)
  If *doc : ProcedureReturn *doc\paragraphCount : EndIf
  ProcedureReturn 0
EndProcedure

; 获取表格数量
Procedure.l PbDoc_Document_GetTableCount(*doc.PbDocument)
  If *doc : ProcedureReturn *doc\tableCount : EndIf
  ProcedureReturn 0
EndProcedure

; ============================================================================
; 二十三、XML 生成 - Run 与段落
; ============================================================================

; 生成 Run XML
; 将 Run 数据结构转换为 Open XML 格式的 <w:r> 元素
Procedure.s PbDoc_GenerateRunXml(*run.PbDocRun)
  Protected xml.s
  xml = "<w:r>"
  ; Run 属性 <w:rPr>
  If *run\bold Or *run\italic Or *run\underline > 0 Or *run\strike Or *run\doubleStrike Or *run\superscript Or *run\subscript Or *run\fontName <> "" Or *run\colorHex <> "" Or *run\highlight <> "" Or *run\fontSize > 0
    xml + "<w:rPr>"
    If *run\bold
      xml + "<w:b/><w:bCs/>"
    EndIf
    If *run\italic
      xml + "<w:i/><w:iCs/>"
    EndIf
    If *run\underline > 0
      xml + "<w:u" + PbDoc_A("w:val", PbDoc_UnderlineToXml(*run\underline)) + "/>"
    EndIf
    If *run\strike
      xml + "<w:strike/>"
    EndIf
    If *run\doubleStrike
      xml + "<w:dstrike/>"
    EndIf
    If *run\superscript
      xml + "<w:vertAlign" + PbDoc_A("w:val", "superscript") + "/>"
    EndIf
    If *run\subscript
      xml + "<w:vertAlign" + PbDoc_A("w:val", "subscript") + "/>"
    EndIf
    If *run\fontName <> ""
      xml + "<w:rFonts" + PbDoc_A("w:ascii", *run\fontName) + PbDoc_A("w:hAnsi", *run\fontName) + PbDoc_A("w:eastAsia", *run\fontName) + "/>"
    EndIf
    If *run\colorHex <> ""
      xml + "<w:color" + PbDoc_A("w:val", *run\colorHex) + "/>"
    EndIf
    If *run\highlight <> ""
      xml + "<w:highlight" + PbDoc_A("w:val", *run\highlight) + "/>"
    EndIf
    If *run\fontSize > 0
      xml + "<w:sz" + PbDoc_A("w:val", Str(*run\fontSize)) + "/>"
      xml + "<w:szCs" + PbDoc_A("w:val", Str(*run\fontSize)) + "/>"
    EndIf
    xml + "</w:rPr>"
  EndIf
  ; 断行符
  If *run\breakType > 0
    If *run\breakType = #PbDoc_BREAK_LINE
      xml + "<w:br/>"
    ElseIf *run\breakType = #PbDoc_BREAK_PAGE
      xml + "<w:br" + PbDoc_A("w:type", "page") + "/>"
    ElseIf *run\breakType = #PbDoc_BREAK_COLUMN
      xml + "<w:br" + PbDoc_A("w:type", "column") + "/>"
    EndIf
  EndIf
  ; Tab 字符
  If *run\hasTab
    xml + "<w:tab/>"
  EndIf
  ; 文本内容
  xml + "<w:t xml:space=" + g_q + "preserve" + g_q + ">" + PbDoc_XmlEscape(*run\text) + "</w:t>"
  xml + "</w:r>"
  ProcedureReturn xml
EndProcedure

; 生成段落 XML
; 将段落数据结构转换为 Open XML 格式的 <w:p> 元素
Procedure.s PbDoc_GenerateParagraphXml(*para.PbDocParagraph)
  Protected xml.s
  Protected pPrXml.s
  Protected hasPPr.l
  hasPPr = #False
  pPrXml = ""
  ; 段落样式
  If *para\style <> ""
    pPrXml + "<w:pStyle" + PbDoc_A("w:val", *para\style) + "/>"
    hasPPr = #True
  EndIf
  ; 对齐方式
  If *para\alignment > 0
    pPrXml + "<w:jc" + PbDoc_A("w:val", PbDoc_AlignmentToXml(*para\alignment)) + "/>"
    hasPPr = #True
  EndIf
  ; 缩进
  If *para\firstLineIndent <> 0 Or *para\leftIndent <> 0 Or *para\rightIndent <> 0 Or *para\hangingIndent <> 0
    pPrXml + "<w:ind"
    If *para\leftIndent <> 0
      pPrXml + PbDoc_A("w:left", Str(*para\leftIndent))
    EndIf
    If *para\rightIndent <> 0
      pPrXml + PbDoc_A("w:right", Str(*para\rightIndent))
    EndIf
    If *para\firstLineIndent <> 0
      pPrXml + PbDoc_A("w:firstLine", Str(*para\firstLineIndent))
    EndIf
    If *para\hangingIndent <> 0
      pPrXml + PbDoc_A("w:hanging", Str(*para\hangingIndent))
    EndIf
    pPrXml + "/>"
    hasPPr = #True
  EndIf
  ; 间距
  If *para\spaceBefore <> 0 Or *para\spaceAfter <> 0 Or *para\lineSpacing <> 0
    pPrXml + "<w:spacing"
    If *para\spaceBefore <> 0
      pPrXml + PbDoc_A("w:before", Str(*para\spaceBefore))
    EndIf
    If *para\spaceAfter <> 0
      pPrXml + PbDoc_A("w:after", Str(*para\spaceAfter))
    EndIf
    If *para\lineSpacing <> 0
      pPrXml + PbDoc_A("w:line", Str(*para\lineSpacing))
      If *para\lineSpacingRule > 0
        pPrXml + PbDoc_A("w:lineRule", PbDoc_LineSpacingRuleToXml(*para\lineSpacingRule))
      EndIf
    EndIf
    pPrXml + "/>"
    hasPPr = #True
  EndIf
  ; 分页控制
  If *para\keepNext
    pPrXml + "<w:keepNext/>"
    hasPPr = #True
  EndIf
  If *para\keepLines
    pPrXml + "<w:keepLines/>"
    hasPPr = #True
  EndIf
  If *para\widowControl
    pPrXml + "<w:widowControl/>"
    hasPPr = #True
  EndIf
  ; Tab Stops
  If ListSize(*para\tabStops()) > 0
    pPrXml + "<w:tabs>"
    ForEach *para\tabStops()
      pPrXml + "<w:tab" + PbDoc_A("w:val", PbDoc_TabAlignToXml(*para\tabStops()\alignment))
      pPrXml + PbDoc_A("w:pos", Str(*para\tabStops()\position))
      If *para\tabStops()\leader <> #PbDoc_TAB_LEADER_NONE
        pPrXml + PbDoc_A("w:leader", PbDoc_TabLeaderToXml(*para\tabStops()\leader))
      EndIf
      pPrXml + "/>"
    Next
    pPrXml + "</w:tabs>"
    hasPPr = #True
  EndIf
  ; 超链接段落
  If *para\isHyperlink
    xml = "<w:p>"
    If hasPPr
      xml + "<w:pPr>" + pPrXml + "</w:pPr>"
    EndIf
    xml + "<w:hyperlink" + PbDoc_A("r:id", *para\hyperlinkRid) + ">"
    ForEach *para\runs()
      xml + PbDoc_GenerateRunXml(*para\runs())
    Next
    xml + "</w:hyperlink>"
    xml + "</w:p>"
    ProcedureReturn xml
  EndIf
  ; 普通段落
  xml = "<w:p>"
  If hasPPr
    xml + "<w:pPr>" + pPrXml + "</w:pPr>"
  EndIf
  ForEach *para\runs()
    xml + PbDoc_GenerateRunXml(*para\runs())
  Next
  xml + "</w:p>"
  ProcedureReturn xml
EndProcedure

; ============================================================================
; 二十四、XML 生成 - 表格
; ============================================================================

; 生成表格 XML
; 将表格数据结构转换为 Open XML 格式的 <w:tbl> 元素
Procedure.s PbDoc_GenerateTableXml(*tbl.PbDocTable)
  Protected xml.s
  Protected i.l, j.l
  Protected key.s
  Protected cellText.s
  Protected colW.q
  xml = "<w:tbl>"
  ; 表格属性
  xml + "<w:tblPr>"
  If *tbl\style <> ""
    xml + "<w:tblStyle" + PbDoc_A("w:val", *tbl\style) + "/>"
  EndIf
  If *tbl\alignment > 0
    xml + "<w:jc" + PbDoc_A("w:val", PbDoc_AlignmentToXml(*tbl\alignment)) + "/>"
  EndIf
  xml + "<w:tblW" + PbDoc_A("w:w", "0") + PbDoc_A("w:type", "auto") + "/>"
  xml + "</w:tblPr>"
  ; 表格网格
  xml + "<w:tblGrid>"
  For j = 0 To *tbl\cols - 1
    colW = 2000
    If SelectElement(*tbl\colWidths(), j)
      colW = *tbl\colWidths()
    EndIf
    xml + "<w:gridCol" + PbDoc_A("w:w", Str(colW)) + "/>"
  Next
  xml + "</w:tblGrid>"
  ; 表格行
  For i = 0 To *tbl\rows - 1
    xml + "<w:tr>"
    For j = 0 To *tbl\cols - 1
      key = Str(i) + "," + Str(j)
      cellText = ""
      If FindMapElement(*tbl\cellText(), key)
        cellText = *tbl\cellText()
      EndIf
      colW = 2000
      If SelectElement(*tbl\colWidths(), j)
        colW = *tbl\colWidths()
      EndIf
      xml + "<w:tc>"
      xml + "<w:tcPr><w:tcW" + PbDoc_A("w:w", Str(colW)) + PbDoc_A("w:type", "dxa") + "/></w:tcPr>"
      xml + "<w:p><w:r><w:t xml:space=" + g_q + "preserve" + g_q + ">" + PbDoc_XmlEscape(cellText) + "</w:t></w:r></w:p>"
      xml + "</w:tc>"
    Next
    xml + "</w:tr>"
  Next
  xml + "</w:tbl>"
  ProcedureReturn xml
EndProcedure

; ============================================================================
; 二十五、XML 生成 - 分页符与断行
; ============================================================================

; 生成分页符 XML
Procedure.s PbDoc_GeneratePageBreakXml()
  Protected xml.s
  xml = "<w:p>"
  xml + "<w:r>"
  xml + "<w:br" + PbDoc_A("w:type", "page") + "/>"
  xml + "</w:r>"
  xml + "</w:p>"
  ProcedureReturn xml
EndProcedure

; ============================================================================
; 二十六、XML 生成 - 文档主体
; ============================================================================

; 生成文档主体 XML
; 按顺序遍历 bodyOrder 列表，生成所有段落、表格、分页符的 XML
Procedure.s PbDoc_GenerateBodyXml(*doc.PbDocument)
  Protected xml.s
  Protected paraIdx.l, tblIdx.l
  xml = ""
  ForEach *doc\bodyOrder()
    Select *doc\bodyOrder()\type
      Case #PbDoc_ELEM_PARAGRAPH
        paraIdx = *doc\bodyOrder()\paraIndex
        If SelectElement(*doc\paragraphs(), paraIdx)
          xml + PbDoc_GenerateParagraphXml(@*doc\paragraphs())
        EndIf
      Case #PbDoc_ELEM_TABLE
        tblIdx = *doc\bodyOrder()\tableIndex
        If SelectElement(*doc\tables(), tblIdx)
          xml + PbDoc_GenerateTableXml(@*doc\tables())
        EndIf
      Case #PbDoc_ELEM_PAGE_BREAK
        xml + PbDoc_GeneratePageBreakXml()
    EndSelect
  Next
  ProcedureReturn xml
EndProcedure

; ============================================================================
; 二十七、XML 生成 - 完整文档
; ============================================================================

; 生成完整的 word/document.xml 内容
Procedure.s PbDoc_GenerateDocumentXml(*doc.PbDocument)
  Protected xml.s
  Protected bodyXml.s
  Protected orientAttr.s
  bodyXml = PbDoc_GenerateBodyXml(*doc)
  xml = ~"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
  xml + "<w:document xmlns:wpc=" + g_q + "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" + g_q + " "
  xml + "xmlns:mc=" + g_q + "http://schemas.openxmlformats.org/markup-compatibility/2006" + g_q + " "
  xml + "xmlns:o=" + g_q + "urn:schemas-microsoft-com:office:office" + g_q + " "
  xml + "xmlns:r=" + g_q + "http://schemas.openxmlformats.org/officeDocument/2006/relationships" + g_q + " "
  xml + "xmlns:m=" + g_q + "http://schemas.openxmlformats.org/officeDocument/2006/math" + g_q + " "
  xml + "xmlns:v=" + g_q + "urn:schemas-microsoft-com:vml" + g_q + " "
  xml + "xmlns:wp=" + g_q + "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" + g_q + " "
  xml + "xmlns:w10=" + g_q + "urn:schemas-microsoft-com:office:word" + g_q + " "
  xml + "xmlns:w=" + g_q + "http://schemas.openxmlformats.org/wordprocessingml/2006/main" + g_q + " "
  xml + "xmlns:w14=" + g_q + "http://schemas.microsoft.com/office/word/2010/wordml" + g_q + " "
  xml + "xmlns:wpg=" + g_q + "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" + g_q + " "
  xml + "xmlns:wpi=" + g_q + "http://schemas.microsoft.com/office/word/2010/wordprocessingInk" + g_q + " "
  xml + "xmlns:wne=" + g_q + "http://schemas.microsoft.com/office/word/2006/wordml" + g_q + " "
  xml + "xmlns:wps=" + g_q + "http://schemas.microsoft.com/office/word/2010/wordprocessingShape" + g_q + " "
  xml + "mc:Ignorable=" + g_q + "w14" + g_q + ">"
  xml + "<w:body>"
  xml + bodyXml
  ; 节属性
  xml + "<w:sectPr>"
  orientAttr = ""
  If *doc\section\orientation = #PbDoc_ORIENT_LANDSCAPE
    orientAttr = PbDoc_A("w:orient", "landscape")
  EndIf
  xml + "<w:pgSz" + PbDoc_A("w:w", Str(*doc\section\pageWidth)) + PbDoc_A("w:h", Str(*doc\section\pageHeight)) + orientAttr + "/>"
  xml + "<w:pgMar" + PbDoc_A("w:top", Str(*doc\section\marginTop)) + PbDoc_A("w:right", Str(*doc\section\marginRight)) + PbDoc_A("w:bottom", Str(*doc\section\marginBottom)) + PbDoc_A("w:left", Str(*doc\section\marginLeft)) + PbDoc_A("w:header", Str(*doc\section\headerDist)) + PbDoc_A("w:footer", Str(*doc\section\footerDist)) + PbDoc_A("w:gutter", Str(*doc\section\gutter)) + "/>"
  xml + "<w:cols" + PbDoc_A("w:space", "720") + "/>"
  xml + "<w:docGrid" + PbDoc_A("w:linePitch", "360") + "/>"
  xml + "</w:sectPr>"
  xml + "</w:body>"
  xml + "</w:document>"
  ProcedureReturn xml
EndProcedure

; ============================================================================
; 二十八、XML 生成 - 关系文件
; ============================================================================

; 生成 _rels/.rels
Procedure.s PbDoc_GenerateRelsXml(*doc.PbDocument)
  Protected xml.s
  xml = ~"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
  xml + "<Relationships xmlns=" + g_q + "http://schemas.openxmlformats.org/package/2006/relationships" + g_q + ">"
  xml + "<Relationship Id=" + g_q + "rId3" + g_q + " Type=" + g_q + "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" + g_q + " Target=" + g_q + "docProps/core.xml" + g_q + "/>"
  xml + "<Relationship Id=" + g_q + "rId4" + g_q + " Type=" + g_q + "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" + g_q + " Target=" + g_q + "docProps/app.xml" + g_q + "/>"
  xml + "<Relationship Id=" + g_q + "rId1" + g_q + " Type=" + g_q + "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" + g_q + " Target=" + g_q + "word/document.xml" + g_q + "/>"
  ForEach *doc\rels()
    xml + "<Relationship Id=" + g_q + *doc\rels()\rid + g_q + " Type=" + g_q + *doc\rels()\relType + g_q + " Target=" + g_q + PbDoc_XmlEscape(*doc\rels()\target) + g_q + "/>"
  Next
  xml + "</Relationships>"
  ProcedureReturn xml
EndProcedure

; 生成 word/_rels/document.xml.rels
Procedure.s PbDoc_GenerateDocRelsXml(*doc.PbDocument)
  Protected xml.s
  xml = ~"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
  xml + "<Relationships xmlns=" + g_q + "http://schemas.openxmlformats.org/package/2006/relationships" + g_q + ">"
  xml + "<Relationship Id=" + g_q + "rId3" + g_q + " Type=" + g_q + "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" + g_q + " Target=" + g_q + "styles.xml" + g_q + "/>"
  xml + "<Relationship Id=" + g_q + "rId5" + g_q + " Type=" + g_q + "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" + g_q + " Target=" + g_q + "settings.xml" + g_q + "/>"
  xml + "<Relationship Id=" + g_q + "rId6" + g_q + " Type=" + g_q + "http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings" + g_q + " Target=" + g_q + "webSettings.xml" + g_q + "/>"
  xml + "<Relationship Id=" + g_q + "rId7" + g_q + " Type=" + g_q + "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" + g_q + " Target=" + g_q + "fontTable.xml" + g_q + "/>"
  xml + "<Relationship Id=" + g_q + "rId8" + g_q + " Type=" + g_q + "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" + g_q + " Target=" + g_q + "theme/theme1.xml" + g_q + "/>"
  xml + "<Relationship Id=" + g_q + "rId2" + g_q + " Type=" + g_q + "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" + g_q + " Target=" + g_q + "numbering.xml" + g_q + "/>"
  ForEach *doc\rels()
    If *doc\rels()\relType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
      xml + "<Relationship Id=" + g_q + *doc\rels()\rid + g_q + " Type=" + g_q + *doc\rels()\relType + g_q + " Target=" + g_q + PbDoc_XmlEscape(*doc\rels()\target) + g_q + " TargetMode=" + g_q + "External" + g_q + "/>"
    EndIf
  Next
  xml + "</Relationships>"
  ProcedureReturn xml
EndProcedure

; ============================================================================
; 二十九、XML 生成 - 内容类型
; ============================================================================

Procedure.s PbDoc_GenerateContentTypesXml()
  Protected xml.s
  xml = ~"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
  xml + "<Types xmlns=" + g_q + "http://schemas.openxmlformats.org/package/2006/content-types" + g_q + ">"
  xml + "<Default Extension=" + g_q + "xml" + g_q + " ContentType=" + g_q + "application/xml" + g_q + "/>"
  xml + "<Default Extension=" + g_q + "rels" + g_q + " ContentType=" + g_q + "application/vnd.openxmlformats-package.relationships+xml" + g_q + "/>"
  xml + "<Default Extension=" + g_q + "jpeg" + g_q + " ContentType=" + g_q + "image/jpeg" + g_q + "/>"
  xml + "<Override PartName=" + g_q + "/word/document.xml" + g_q + " ContentType=" + g_q + "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml" + g_q + "/>"
  xml + "<Override PartName=" + g_q + "/word/numbering.xml" + g_q + " ContentType=" + g_q + "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml" + g_q + "/>"
  xml + "<Override PartName=" + g_q + "/word/styles.xml" + g_q + " ContentType=" + g_q + "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml" + g_q + "/>"
  xml + "<Override PartName=" + g_q + "/word/settings.xml" + g_q + " ContentType=" + g_q + "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml" + g_q + "/>"
  xml + "<Override PartName=" + g_q + "/word/webSettings.xml" + g_q + " ContentType=" + g_q + "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml" + g_q + "/>"
  xml + "<Override PartName=" + g_q + "/word/fontTable.xml" + g_q + " ContentType=" + g_q + "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml" + g_q + "/>"
  xml + "<Override PartName=" + g_q + "/word/theme/theme1.xml" + g_q + " ContentType=" + g_q + "application/vnd.openxmlformats-officedocument.theme+xml" + g_q + "/>"
  xml + "<Override PartName=" + g_q + "/docProps/core.xml" + g_q + " ContentType=" + g_q + "application/vnd.openxmlformats-package.core-properties+xml" + g_q + "/>"
  xml + "<Override PartName=" + g_q + "/docProps/app.xml" + g_q + " ContentType=" + g_q + "application/vnd.openxmlformats-officedocument.extended-properties+xml" + g_q + "/>"
  xml + "</Types>"
  ProcedureReturn xml
EndProcedure

; ============================================================================
; 三十、XML 生成 - 核心属性
; ============================================================================

; 生成 docProps/core.xml
; 包含文档的标题、作者、创建时间等核心元数据
Procedure.s PbDoc_GenerateCoreXml(*doc.PbDocument)
  Protected xml.s
  Protected dateStr.s
  dateStr = FormatDate("%yyyy-%mm-%dd", Date()) + "T" + FormatDate("%hh:%ii:%ss", Date()) + "Z"
  xml = ~"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
  xml + "<cp:coreProperties xmlns:cp=" + g_q + "http://schemas.openxmlformats.org/package/2006/metadata/core-properties" + g_q + " "
  xml + "xmlns:dc=" + g_q + "http://purl.org/dc/elements/1.1/" + g_q + " "
  xml + "xmlns:dcterms=" + g_q + "http://purl.org/dc/terms/" + g_q + " "
  xml + "xmlns:dcmitype=" + g_q + "http://purl.org/dc/dcmitype/" + g_q + " "
  xml + "xmlns:xsi=" + g_q + "http://www.w3.org/2001/XMLSchema-instance" + g_q + ">"
  xml + "<dc:title>" + PbDoc_XmlEscape(*doc\title) + "</dc:title>"
  xml + "<dc:subject>" + PbDoc_XmlEscape(*doc\subject) + "</dc:subject>"
  xml + "<dc:creator>" + PbDoc_XmlEscape(*doc\author) + "</dc:creator>"
  xml + "<cp:keywords/>"
  xml + "<dc:description>由 PbDoc 生成</dc:description>"
  xml + "<cp:lastModifiedBy>" + PbDoc_XmlEscape(*doc\author) + "</cp:lastModifiedBy>"
  xml + "<cp:revision>1</cp:revision>"
  xml + "<dcterms:created xsi:type=" + g_q + "dcterms:W3CDTF" + g_q + ">" + dateStr + "</dcterms:created>"
  xml + "<dcterms:modified xsi:type=" + g_q + "dcterms:W3CDTF" + g_q + ">" + dateStr + "</dcterms:modified>"
  xml + "<cp:category/>"
  xml + "</cp:coreProperties>"
  ProcedureReturn xml
EndProcedure

; ============================================================================
; 三十一、XML 生成 - 扩展属性
; ============================================================================

; 生成 docProps/app.xml
; 包含应用程序信息、公司名称等扩展元数据
Procedure.s PbDoc_GenerateAppXml(*doc.PbDocument)
  Protected xml.s
  xml = ~"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
  xml + "<Properties xmlns=" + g_q + "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" + g_q + " "
  xml + "xmlns:vt=" + g_q + "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes" + g_q + ">"
  xml + "<Template>Normal.dotm</Template>"
  xml + "<TotalTime>0</TotalTime>"
  xml + "<Pages>1</Pages>"
  xml + "<Words>0</Words>"
  xml + "<Characters>0</Characters>"
  xml + "<Application>PbDoc</Application>"
  xml + "<DocSecurity>0</DocSecurity>"
  xml + "<Lines>0</Lines>"
  xml + "<Paragraphs>0</Paragraphs>"
  xml + "<ScaleCrop>false</ScaleCrop>"
  xml + "<HeadingPairs>"
  xml + "<vt:vector size=" + g_q + "2" + g_q + " baseType=" + g_q + "variant" + g_q + ">"
  xml + "<vt:variant><vt:lpstr>Title</vt:lpstr></vt:variant>"
  xml + "<vt:variant><vt:i4>1</vt:i4></vt:variant>"
  xml + "</vt:vector>"
  xml + "</HeadingPairs>"
  xml + "<TitlesOfParts>"
  xml + "<vt:vector size=" + g_q + "1" + g_q + " baseType=" + g_q + "lpstr" + g_q + ">"
  xml + "<vt:lpstr/>"
  xml + "</vt:vector>"
  xml + "</TitlesOfParts>"
  xml + "<Manager/>"
  xml + "<Company>" + PbDoc_XmlEscape(*doc\company) + "</Company>"
  xml + "<LinksUpToDate>false</LinksUpToDate>"
  xml + "<CharactersWithSpaces>0</CharactersWithSpaces>"
  xml + "<SharedDoc>false</SharedDoc>"
  xml + "<HyperlinkBase/>"
  xml + "<HyperlinksChanged>false</HyperlinksChanged>"
  xml + "<AppVersion>14.0000</AppVersion>"
  xml + "</Properties>"
  ProcedureReturn xml
EndProcedure

; ============================================================================
; 三十二、XML 生成 - 样式定义
; ============================================================================
; 生成 word/styles.xml，包含 Normal、Heading1-9、TableGrid 等必要样式

Procedure.s PbDoc_GenerateStylesXml()
  Protected xml.s
  Protected i.l
  xml = ~"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
  xml + "<w:styles xmlns:mc=" + g_q + "http://schemas.openxmlformats.org/markup-compatibility/2006" + g_q + " "
  xml + "xmlns:r=" + g_q + "http://schemas.openxmlformats.org/officeDocument/2006/relationships" + g_q + " "
  xml + "xmlns:w=" + g_q + "http://schemas.openxmlformats.org/wordprocessingml/2006/main" + g_q + " "
  xml + "xmlns:w14=" + g_q + "http://schemas.microsoft.com/office/word/2010/wordml" + g_q + " "
  xml + "mc:Ignorable=" + g_q + "w14" + g_q + ">"
  ; 文档默认值
  xml + "<w:docDefaults>"
  xml + "<w:rPrDefault><w:rPr>"
  xml + "<w:rFonts w:asciiTheme=" + g_q + "minorHAnsi" + g_q + " w:eastAsiaTheme=" + g_q + "minorEastAsia" + g_q + " w:hAnsiTheme=" + g_q + "minorHAnsi" + g_q + " w:cstheme=" + g_q + "minorBidi" + g_q + "/>"
  xml + "<w:sz" + PbDoc_A("w:val", "22") + "/><w:szCs" + PbDoc_A("w:val", "22") + "/>"
  xml + "<w:lang w:val=" + g_q + "en-US" + g_q + " w:eastAsia=" + g_q + "en-US" + g_q + " w:bidi=" + g_q + "ar-SA" + g_q + "/>"
  xml + "</w:rPr></w:rPrDefault>"
  xml + "<w:pPrDefault><w:pPr>"
  xml + "<w:spacing" + PbDoc_A("w:after", "200") + PbDoc_A("w:line", "276") + PbDoc_A("w:lineRule", "auto") + "/>"
  xml + "</w:pPr></w:pPrDefault>"
  xml + "</w:docDefaults>"
  ; 潜在样式
  xml + "<w:latentStyles" + PbDoc_A("w:defLockedState", "0") + PbDoc_A("w:defUIPriority", "99") + PbDoc_A("w:defSemiHidden", "1") + PbDoc_A("w:defUnhideWhenUsed", "1") + PbDoc_A("w:defQFormat", "0") + PbDoc_A("w:count", "276") + ">"
  xml + "<w:lsdException w:name=" + g_q + "Normal" + g_q + " w:semiHidden=" + g_q + "0" + g_q + " w:uiPriority=" + g_q + "0" + g_q + " w:unhideWhenUsed=" + g_q + "0" + g_q + " w:qFormat=" + g_q + "1" + g_q + "/>"
  xml + "<w:lsdException w:name=" + g_q + "heading 1" + g_q + " w:semiHidden=" + g_q + "0" + g_q + " w:uiPriority=" + g_q + "9" + g_q + " w:unhideWhenUsed=" + g_q + "0" + g_q + " w:qFormat=" + g_q + "1" + g_q + "/>"
  xml + "<w:lsdException w:name=" + g_q + "heading 2" + g_q + " w:uiPriority=" + g_q + "9" + g_q + " w:qFormat=" + g_q + "1" + g_q + "/>"
  xml + "<w:lsdException w:name=" + g_q + "heading 3" + g_q + " w:uiPriority=" + g_q + "9" + g_q + " w:qFormat=" + g_q + "1" + g_q + "/>"
  xml + "<w:lsdException w:name=" + g_q + "Title" + g_q + " w:semiHidden=" + g_q + "0" + g_q + " w:uiPriority=" + g_q + "10" + g_q + " w:unhideWhenUsed=" + g_q + "0" + g_q + " w:qFormat=" + g_q + "1" + g_q + "/>"
  xml + "<w:lsdException w:name=" + g_q + "Table Grid" + g_q + " w:semiHidden=" + g_q + "0" + g_q + " w:uiPriority=" + g_q + "59" + g_q + " w:unhideWhenUsed=" + g_q + "0" + g_q + "/>"
  xml + "</w:latentStyles>"
  ; Normal 样式
  xml + "<w:style w:type=" + g_q + "paragraph" + g_q + " w:default=" + g_q + "1" + g_q + " w:styleId=" + g_q + "Normal" + g_q + ">"
  xml + "<w:name" + PbDoc_A("w:val", "Normal") + "/>"
  xml + "<w:qFormat/>"
  xml + "</w:style>"
  ; Heading1 样式
  xml + "<w:style w:type=" + g_q + "paragraph" + g_q + " w:styleId=" + g_q + "Heading1" + g_q + ">"
  xml + "<w:name" + PbDoc_A("w:val", "heading 1") + "/>"
  xml + "<w:basedOn" + PbDoc_A("w:val", "Normal") + "/>"
  xml + "<w:next" + PbDoc_A("w:val", "Normal") + "/>"
  xml + "<w:uiPriority" + PbDoc_A("w:val", "9") + "/>"
  xml + "<w:qFormat/>"
  xml + "<w:pPr>"
  xml + "<w:keepNext/><w:keepLines/>"
  xml + "<w:spacing" + PbDoc_A("w:before", "480") + PbDoc_A("w:after", "0") + "/>"
  xml + "<w:outlineLvl" + PbDoc_A("w:val", "0") + "/>"
  xml + "</w:pPr>"
  xml + "<w:rPr>"
  xml + "<w:rFonts w:asciiTheme=" + g_q + "majorHAnsi" + g_q + " w:eastAsiaTheme=" + g_q + "majorEastAsia" + g_q + " w:hAnsiTheme=" + g_q + "majorHAnsi" + g_q + " w:cstheme=" + g_q + "majorBidi" + g_q + "/>"
  xml + "<w:b/><w:bCs/>"
  xml + "<w:color" + PbDoc_A("w:val", "365F91") + " w:themeColor=" + g_q + "accent1" + g_q + " w:themeShade=" + g_q + "BF" + g_q + "/>"
  xml + "<w:sz" + PbDoc_A("w:val", "28") + "/><w:szCs" + PbDoc_A("w:val", "28") + "/>"
  xml + "</w:rPr>"
  xml + "</w:style>"
  ; Heading2 样式
  xml + "<w:style w:type=" + g_q + "paragraph" + g_q + " w:styleId=" + g_q + "Heading2" + g_q + ">"
  xml + "<w:name" + PbDoc_A("w:val", "heading 2") + "/>"
  xml + "<w:basedOn" + PbDoc_A("w:val", "Normal") + "/>"
  xml + "<w:next" + PbDoc_A("w:val", "Normal") + "/>"
  xml + "<w:uiPriority" + PbDoc_A("w:val", "9") + "/>"
  xml + "<w:unhideWhenUsed/>"
  xml + "<w:qFormat/>"
  xml + "<w:pPr>"
  xml + "<w:keepNext/><w:keepLines/>"
  xml + "<w:spacing" + PbDoc_A("w:before", "200") + PbDoc_A("w:after", "0") + "/>"
  xml + "<w:outlineLvl" + PbDoc_A("w:val", "1") + "/>"
  xml + "</w:pPr>"
  xml + "<w:rPr>"
  xml + "<w:rFonts w:asciiTheme=" + g_q + "majorHAnsi" + g_q + " w:eastAsiaTheme=" + g_q + "majorEastAsia" + g_q + " w:hAnsiTheme=" + g_q + "majorHAnsi" + g_q + " w:cstheme=" + g_q + "majorBidi" + g_q + "/>"
  xml + "<w:b/><w:bCs/>"
  xml + "<w:color" + PbDoc_A("w:val", "4F81BD") + " w:themeColor=" + g_q + "accent1" + g_q + "/>"
  xml + "<w:sz" + PbDoc_A("w:val", "26") + "/><w:szCs" + PbDoc_A("w:val", "26") + "/>"
  xml + "</w:rPr>"
  xml + "</w:style>"
  ; Heading3 样式
  xml + "<w:style w:type=" + g_q + "paragraph" + g_q + " w:styleId=" + g_q + "Heading3" + g_q + ">"
  xml + "<w:name" + PbDoc_A("w:val", "heading 3") + "/>"
  xml + "<w:basedOn" + PbDoc_A("w:val", "Normal") + "/>"
  xml + "<w:next" + PbDoc_A("w:val", "Normal") + "/>"
  xml + "<w:uiPriority" + PbDoc_A("w:val", "9") + "/>"
  xml + "<w:unhideWhenUsed/>"
  xml + "<w:qFormat/>"
  xml + "<w:pPr>"
  xml + "<w:keepNext/><w:keepLines/>"
  xml + "<w:spacing" + PbDoc_A("w:before", "200") + PbDoc_A("w:after", "0") + "/>"
  xml + "<w:outlineLvl" + PbDoc_A("w:val", "2") + "/>"
  xml + "</w:pPr>"
  xml + "<w:rPr>"
  xml + "<w:rFonts w:asciiTheme=" + g_q + "majorHAnsi" + g_q + " w:eastAsiaTheme=" + g_q + "majorEastAsia" + g_q + " w:hAnsiTheme=" + g_q + "majorHAnsi" + g_q + " w:cstheme=" + g_q + "majorBidi" + g_q + "/>"
  xml + "<w:b/><w:bCs/>"
  xml + "<w:color" + PbDoc_A("w:val", "4F81BD") + " w:themeColor=" + g_q + "accent1" + g_q + "/>"
  xml + "</w:rPr>"
  xml + "</w:style>"
  ; Heading4-9 样式
  For i = 4 To 9
    xml + "<w:style w:type=" + g_q + "paragraph" + g_q + " w:styleId=" + g_q + "Heading" + Str(i) + g_q + ">"
    xml + "<w:name" + PbDoc_A("w:val", "heading " + Str(i)) + "/>"
    xml + "<w:basedOn" + PbDoc_A("w:val", "Normal") + "/>"
    xml + "<w:next" + PbDoc_A("w:val", "Normal") + "/>"
    xml + "<w:uiPriority" + PbDoc_A("w:val", "9") + "/>"
    xml + "<w:unhideWhenUsed/>"
    xml + "<w:qFormat/>"
    xml + "<w:pPr>"
    xml + "<w:keepNext/><w:keepLines/>"
    xml + "<w:spacing" + PbDoc_A("w:before", "200") + PbDoc_A("w:after", "0") + "/>"
    xml + "<w:outlineLvl" + PbDoc_A("w:val", Str(i - 1)) + "/>"
    xml + "</w:pPr>"
    xml + "<w:rPr>"
    xml + "<w:color" + PbDoc_A("w:val", "4F81BD") + " w:themeColor=" + g_q + "accent1" + g_q + "/>"
    xml + "</w:rPr>"
    xml + "</w:style>"
  Next
  ; Title 样式
  xml + "<w:style w:type=" + g_q + "paragraph" + g_q + " w:styleId=" + g_q + "Title" + g_q + ">"
  xml + "<w:name" + PbDoc_A("w:val", "Title") + "/>"
  xml + "<w:basedOn" + PbDoc_A("w:val", "Normal") + "/>"
  xml + "<w:next" + PbDoc_A("w:val", "Normal") + "/>"
  xml + "<w:uiPriority" + PbDoc_A("w:val", "10") + "/>"
  xml + "<w:qFormat/>"
  xml + "<w:pPr>"
  xml + "<w:spacing" + PbDoc_A("w:after", "0") + "/>"
  xml + "</w:pPr>"
  xml + "<w:rPr>"
  xml + "<w:rFonts w:asciiTheme=" + g_q + "majorHAnsi" + g_q + " w:eastAsiaTheme=" + g_q + "majorEastAsia" + g_q + " w:hAnsiTheme=" + g_q + "majorHAnsi" + g_q + " w:cstheme=" + g_q + "majorBidi" + g_q + "/>"
  xml + "<w:b/><w:bCs/>"
  xml + "<w:sz" + PbDoc_A("w:val", "56") + "/><w:szCs" + PbDoc_A("w:val", "56") + "/>"
  xml + "</w:rPr>"
  xml + "</w:style>"
  ; TableGrid 样式
  xml + "<w:style w:type=" + g_q + "table" + g_q + " w:styleId=" + g_q + "TableGrid" + g_q + ">"
  xml + "<w:name" + PbDoc_A("w:val", "Table Grid") + "/>"
  xml + "<w:basedOn" + PbDoc_A("w:val", "TableNormal") + "/>"
  xml + "<w:uiPriority" + PbDoc_A("w:val", "59") + "/>"
  xml + "<w:tblPr>"
  xml + "<w:tblInd" + PbDoc_A("w:w", "0") + PbDoc_A("w:type", "dxa") + "/>"
  xml + "<w:tblBorders>"
  xml + "<w:top" + PbDoc_A("w:val", "single") + PbDoc_A("w:sz", "4") + PbDoc_A("w:space", "0") + PbDoc_A("w:color", "auto") + "/>"
  xml + "<w:left" + PbDoc_A("w:val", "single") + PbDoc_A("w:sz", "4") + PbDoc_A("w:space", "0") + PbDoc_A("w:color", "auto") + "/>"
  xml + "<w:bottom" + PbDoc_A("w:val", "single") + PbDoc_A("w:sz", "4") + PbDoc_A("w:space", "0") + PbDoc_A("w:color", "auto") + "/>"
  xml + "<w:right" + PbDoc_A("w:val", "single") + PbDoc_A("w:sz", "4") + PbDoc_A("w:space", "0") + PbDoc_A("w:color", "auto") + "/>"
  xml + "<w:insideH" + PbDoc_A("w:val", "single") + PbDoc_A("w:sz", "4") + PbDoc_A("w:space", "0") + PbDoc_A("w:color", "auto") + "/>"
  xml + "<w:insideV" + PbDoc_A("w:val", "single") + PbDoc_A("w:sz", "4") + PbDoc_A("w:space", "0") + PbDoc_A("w:color", "auto") + "/>"
  xml + "</w:tblBorders>"
  xml + "<w:tblCellMar>"
  xml + "<w:top" + PbDoc_A("w:w", "0") + PbDoc_A("w:type", "dxa") + "/>"
  xml + "<w:left" + PbDoc_A("w:w", "108") + PbDoc_A("w:type", "dxa") + "/>"
  xml + "<w:bottom" + PbDoc_A("w:w", "0") + PbDoc_A("w:type", "dxa") + "/>"
  xml + "<w:right" + PbDoc_A("w:w", "108") + PbDoc_A("w:type", "dxa") + "/>"
  xml + "</w:tblCellMar>"
  xml + "</w:tblPr>"
  xml + "</w:style>"
  ; TableNormal 样式
  xml + "<w:style w:type=" + g_q + "table" + g_q + " w:default=" + g_q + "1" + g_q + " w:styleId=" + g_q + "TableNormal" + g_q + ">"
  xml + "<w:name" + PbDoc_A("w:val", "Normal Table") + "/>"
  xml + "<w:uiPriority" + PbDoc_A("w:val", "99") + "/>"
  xml + "<w:unhideWhenUsed/>"
  xml + "<w:tblPr>"
  xml + "<w:tblInd" + PbDoc_A("w:w", "0") + PbDoc_A("w:type", "dxa") + "/>"
  xml + "<w:tblCellMar>"
  xml + "<w:top" + PbDoc_A("w:w", "0") + PbDoc_A("w:type", "dxa") + "/>"
  xml + "<w:left" + PbDoc_A("w:w", "108") + PbDoc_A("w:type", "dxa") + "/>"
  xml + "<w:bottom" + PbDoc_A("w:w", "0") + PbDoc_A("w:type", "dxa") + "/>"
  xml + "<w:right" + PbDoc_A("w:w", "108") + PbDoc_A("w:type", "dxa") + "/>"
  xml + "</w:tblCellMar>"
  xml + "</w:tblPr>"
  xml + "</w:style>"
  xml + "</w:styles>"
  ProcedureReturn xml
EndProcedure

; ============================================================================
; 三十三、XML 生成 - 其他静态文件
; ============================================================================

Procedure.s PbDoc_GenerateSettingsXml()
  Protected xml.s
  xml = ~"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
  xml + "<w:settings xmlns:mc=" + g_q + "http://schemas.openxmlformats.org/markup-compatibility/2006" + g_q + " "
  xml + "xmlns:o=" + g_q + "urn:schemas-microsoft-com:office:office" + g_q + " "
  xml + "xmlns:r=" + g_q + "http://schemas.openxmlformats.org/officeDocument/2006/relationships" + g_q + " "
  xml + "xmlns:w=" + g_q + "http://schemas.openxmlformats.org/wordprocessingml/2006/main" + g_q + " "
  xml + "mc:Ignorable=" + g_q + "w14" + g_q + ">"
  xml + "<w:defaultTabStop" + PbDoc_A("w:val", "720") + "/>"
  xml + "<w:characterSpacingControl" + PbDoc_A("w:val", "doNotCompress") + "/>"
  xml + "<w:compat>"
  xml + "<w:compatSetting" + PbDoc_A("w:name", "compatibilityMode") + PbDoc_A("w:uri", "http://schemas.microsoft.com/office/word") + PbDoc_A("w:val", "14") + "/>"
  xml + "</w:compat>"
  xml + "<w:themeFontLang w:val=" + g_q + "en-US" + g_q + " w:eastAsia=" + g_q + "zh-CN" + g_q + "/>"
  xml + "</w:settings>"
  ProcedureReturn xml
EndProcedure

Procedure.s PbDoc_GenerateWebSettingsXml()
  Protected xml.s
  xml = ~"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
  xml + "<w:webSettings xmlns:mc=" + g_q + "http://schemas.openxmlformats.org/markup-compatibility/2006" + g_q + " "
  xml + "xmlns:w=" + g_q + "http://schemas.openxmlformats.org/wordprocessingml/2006/main" + g_q + " "
  xml + "mc:Ignorable=" + g_q + "w14" + g_q + ">"
  xml + "<w:allowPNG/>"
  xml + "</w:webSettings>"
  ProcedureReturn xml
EndProcedure

Procedure.s PbDoc_GenerateFontTableXml()
  Protected xml.s
  xml = ~"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
  xml + "<w:fonts xmlns:mc=" + g_q + "http://schemas.openxmlformats.org/markup-compatibility/2006" + g_q + " "
  xml + "xmlns:w=" + g_q + "http://schemas.openxmlformats.org/wordprocessingml/2006/main" + g_q + " "
  xml + "mc:Ignorable=" + g_q + "w14" + g_q + ">"
  xml + "<w:font w:name=" + g_q + "Calibri" + g_q + ">"
  xml + "<w:panose1" + PbDoc_A("w:val", "020F0502020204030204") + "/>"
  xml + "<w:charset" + PbDoc_A("w:val", "00") + "/>"
  xml + "<w:family" + PbDoc_A("w:val", "auto") + "/>"
  xml + "<w:pitch" + PbDoc_A("w:val", "variable") + "/>"
  xml + "</w:font>"
  xml + "<w:font w:name=" + g_q + "Times New Roman" + g_q + ">"
  xml + "<w:panose1" + PbDoc_A("w:val", "02020603050405020304") + "/>"
  xml + "<w:charset" + PbDoc_A("w:val", "00") + "/>"
  xml + "<w:family" + PbDoc_A("w:val", "auto") + "/>"
  xml + "<w:pitch" + PbDoc_A("w:val", "variable") + "/>"
  xml + "</w:font>"
  xml + "</w:fonts>"
  ProcedureReturn xml
EndProcedure

Procedure.s PbDoc_GenerateNumberingXml()
  Protected xml.s
  xml = ~"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
  xml + "<w:numbering xmlns:w=" + g_q + "http://schemas.openxmlformats.org/wordprocessingml/2006/main" + g_q + ">"
  xml + "</w:numbering>"
  ProcedureReturn xml
EndProcedure

Procedure.s PbDoc_GenerateThemeXml()
  Protected xml.s
  xml = ~"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
  xml + "<a:theme xmlns:a=" + g_q + "http://schemas.openxmlformats.org/drawingml/2006/main" + g_q + " name=" + g_q + "Office Theme" + g_q + ">"
  xml + "<a:themeElements>"
  xml + "<a:clrScheme name=" + g_q + "Office" + g_q + ">"
  xml + "<a:dk1><a:sysClr val=" + g_q + "windowText" + g_q + " lastClr=" + g_q + "000000" + g_q + "/></a:dk1>"
  xml + "<a:lt1><a:sysClr val=" + g_q + "window" + g_q + " lastClr=" + g_q + "FFFFFF" + g_q + "/></a:lt1>"
  xml + "<a:dk2><a:srgbClr val=" + g_q + "44546A" + g_q + "/></a:dk2>"
  xml + "<a:lt2><a:srgbClr val=" + g_q + "E7E6E6" + g_q + "/></a:lt2>"
  xml + "<a:accent1><a:srgbClr val=" + g_q + "4472C4" + g_q + "/></a:accent1>"
  xml + "<a:accent2><a:srgbClr val=" + g_q + "ED7D31" + g_q + "/></a:accent2>"
  xml + "<a:accent3><a:srgbClr val=" + g_q + "A5A5A5" + g_q + "/></a:accent3>"
  xml + "<a:accent4><a:srgbClr val=" + g_q + "FFC000" + g_q + "/></a:accent4>"
  xml + "<a:accent5><a:srgbClr val=" + g_q + "5B9BD5" + g_q + "/></a:accent5>"
  xml + "<a:accent6><a:srgbClr val=" + g_q + "70AD47" + g_q + "/></a:accent6>"
  xml + "<a:hlink><a:srgbClr val=" + g_q + "0563C1" + g_q + "/></a:hlink>"
  xml + "<a:folHlink><a:srgbClr val=" + g_q + "954F72" + g_q + "/></a:folHlink>"
  xml + "</a:clrScheme>"
  xml + "<a:fontScheme name=" + g_q + "Office" + g_q + ">"
  xml + "<a:majorFont><a:latin typeface=" + g_q + "Calibri Light" + g_q + "/></a:majorFont>"
  xml + "<a:minorFont><a:latin typeface=" + g_q + "Calibri" + g_q + "/></a:minorFont>"
  xml + "</a:fontScheme>"
  xml + "<a:fmtScheme name=" + g_q + "Office" + g_q + ">"
  xml + "<a:fillStyleLst>"
  xml + "<a:solidFill><a:schemeClr val=" + g_q + "phClr" + g_q + "/></a:solidFill>"
  xml + "<a:solidFill><a:schemeClr val=" + g_q + "phClr" + g_q + "/></a:solidFill>"
  xml + "<a:solidFill><a:schemeClr val=" + g_q + "phClr" + g_q + "/></a:solidFill>"
  xml + "</a:fillStyleLst>"
  xml + "<a:lnStyleLst>"
  xml + "<a:ln w=" + g_q + "6350" + g_q + " cap=" + g_q + "flat" + g_q + " cmpd=" + g_q + "sng" + g_q + " algn=" + g_q + "ctr" + g_q + "><a:solidFill><a:schemeClr val=" + g_q + "phClr" + g_q + "/></a:solidFill></a:ln>"
  xml + "<a:ln w=" + g_q + "12700" + g_q + " cap=" + g_q + "flat" + g_q + " cmpd=" + g_q + "sng" + g_q + " algn=" + g_q + "ctr" + g_q + "><a:solidFill><a:schemeClr val=" + g_q + "phClr" + g_q + "/></a:solidFill></a:ln>"
  xml + "<a:ln w=" + g_q + "19050" + g_q + " cap=" + g_q + "flat" + g_q + " cmpd=" + g_q + "sng" + g_q + " algn=" + g_q + "ctr" + g_q + "><a:solidFill><a:schemeClr val=" + g_q + "phClr" + g_q + "/></a:solidFill></a:ln>"
  xml + "</a:lnStyleLst>"
  xml + "<a:effectStyleLst>"
  xml + "<a:effectStyle><a:effectLst/></a:effectStyle>"
  xml + "<a:effectStyle><a:effectLst/></a:effectStyle>"
  xml + "<a:effectStyle><a:effectLst/></a:effectStyle>"
  xml + "</a:effectStyleLst>"
  xml + "<a:bgFillStyleLst>"
  xml + "<a:solidFill><a:schemeClr val=" + g_q + "phClr" + g_q + "/></a:solidFill>"
  xml + "<a:solidFill><a:schemeClr val=" + g_q + "phClr" + g_q + "/></a:solidFill>"
  xml + "<a:solidFill><a:schemeClr val=" + g_q + "phClr" + g_q + "/></a:solidFill>"
  xml + "</a:bgFillStyleLst>"
  xml + "</a:fmtScheme>"
  xml + "</a:themeElements>"
  xml + "<a:objectDefaults/>"
  xml + "<a:extraClrSchemeLst/>"
  xml + "</a:theme>"
  ProcedureReturn xml
EndProcedure

; ============================================================================
; 三十四、ZIP 包辅助函数
; ============================================================================

; 将字符串内容添加到 ZIP 包中
; 参数 packNum: 包编号
; 参数 content: 要添加的字符串内容
; 参数 entryName: ZIP 内的文件路径 (如 "word/document.xml")
Procedure PbDoc_AddStringToPack(packNum.l, content.s, entryName.s)
  Protected *buffer
  Protected size.l
  Protected result.l
  size = StringByteLength(content, #PB_UTF8)
  If size <= 0
    size = 1
  EndIf
  *buffer = AllocateMemory(size + 2)
  If *buffer = 0
    ProcedureReturn #False
  EndIf
  PokeS(*buffer, content, -1, #PB_UTF8)
  result = AddPackMemory(packNum, *buffer, size, entryName)
  FreeMemory(*buffer)
  ProcedureReturn result
EndProcedure

; ============================================================================
; 三十五、文档保存
; ============================================================================

; 保存文档到 .docx 文件
; 参数 filePath: 输出文件路径
; 返回值: #True 成功, #False 失败
Procedure.l PbDocument_Save(*doc.PbDocument, filePath.s)
  Protected packNum.l
  Protected result.l
  If *doc = 0
    PbDoc_SetError("文档指针为空")
    ProcedureReturn #False
  EndIf
  ; 删除已存在的文件
  If FileSize(filePath) > 0
    DeleteFile(filePath)
  EndIf
  ; 创建 ZIP 包
  packNum = 1
  result = CreatePack(packNum, filePath, #PB_PackerPlugin_Zip)
  If result = 0
    PbDoc_SetError("无法创建 ZIP 文件: " + filePath)
    ProcedureReturn #False
  EndIf
  ; 添加所有必需的 XML 文件
  PbDoc_AddStringToPack(packNum, PbDoc_GenerateContentTypesXml(), "[Content_Types].xml")
  PbDoc_AddStringToPack(packNum, PbDoc_GenerateRelsXml(*doc), "_rels/.rels")
  PbDoc_AddStringToPack(packNum, PbDoc_GenerateDocumentXml(*doc), "word/document.xml")
  PbDoc_AddStringToPack(packNum, PbDoc_GenerateDocRelsXml(*doc), "word/_rels/document.xml.rels")
  PbDoc_AddStringToPack(packNum, PbDoc_GenerateStylesXml(), "word/styles.xml")
  PbDoc_AddStringToPack(packNum, PbDoc_GenerateSettingsXml(), "word/settings.xml")
  PbDoc_AddStringToPack(packNum, PbDoc_GenerateWebSettingsXml(), "word/webSettings.xml")
  PbDoc_AddStringToPack(packNum, PbDoc_GenerateFontTableXml(), "word/fontTable.xml")
  PbDoc_AddStringToPack(packNum, PbDoc_GenerateNumberingXml(), "word/numbering.xml")
  PbDoc_AddStringToPack(packNum, PbDoc_GenerateThemeXml(), "word/theme/theme1.xml")
  PbDoc_AddStringToPack(packNum, PbDoc_GenerateCoreXml(*doc), "docProps/core.xml")
  PbDoc_AddStringToPack(packNum, PbDoc_GenerateAppXml(*doc), "docProps/app.xml")
  ClosePack(packNum)
  ProcedureReturn #True
EndProcedure

; ============================================================================
; IDE Options = PureBasic 6.40 (Windows - x86)
; EnableThread
; EnableXP
; DPIAware
; CompileSourceDirectory
