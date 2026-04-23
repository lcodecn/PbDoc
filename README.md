# PbDoc Library v1.3

PbDoc

\- PureBasic Word (.docx) 文档处理库

- 作者  : lcode.cn
- 版本  : 1.3
- 许可证  : Apache 2.0
- 编译器  : PureBasic 6.40 (Windows - x86)

***

## 简介

PbDoc 是一个纯 PureBasic 实现的 Word 文档操作库，无需安装 Microsoft Office 或任何第三方依赖，即可创建符合 Office Open XML 标准的 .docx 文件。

该库参考 Python 的 python-docx 项目编写，使用 PureBasic 内置的 Packer（ZIP压缩）库实现。

## 主要功能

- 创建Word文档  : 从零创建符合 Office Open XML 标准的 .docx 文件
- 文档属性  : 设置标题、作者、主题、公司等文档元数据
- 段落操作  : 添加段落、设置对齐方式、设置段落样式
- 标题操作  : 添加1-9级标题，自动应用内置标题样式
- 字体格式  : 粗体、斜体、下划线（7种）、删除线、上下标、字体名、颜色、高亮、字号
- 段落格式  : 首行缩进、左/右缩进、悬挂缩进、段前/段后间距、行间距、分页控制
- Tab制表符  : 左/右/居中/小数点对齐Tab，5种前导符类型
- 表格操作  : 创建表格、设置样式/对齐/列宽、填充单元格
- 超链接  : 添加外部超链接
- 分页符  : 插入分页符、换行符、分栏符
- 页面设置  : 纸张大小、页边距、页面方向（纵向/横向）
- 单位转换  : EMU/缇/磅/厘米/英寸互转

## 系统要求

- 本项目在PureBasic 6.40 （Windows x86）中编译通过，其他环境请自行测试。

## 快速开始

具体可参考开发文档：docs\PbDoc\_Help.html

### 创建Word文档

```purebasic
XIncludeFile "PbDoc.pb"

; 初始化库
PbDoc_Init()

; 创建新文档
*doc.PbDocument = PbDocument_Create()

; 设置文档属性
PbDocument_SetTitle(*doc, "我的文档")
PbDocument_SetAuthor(*doc, "PbDoc 用户")

; 添加标题和段落
PbDoc_Document_AddHeading(*doc, "第一章 引言", 1)
PbDoc_Document_AddParagraph(*doc, "这是文档的第一段内容。")

; 保存文件
PbDocument_Save(*doc, "output.docx")

; 释放资源
PbDocument_Free(*doc)
PbDoc_Cleanup()
```

### 添加带格式的文本

```purebasic
; 添加多Run段落（不同格式）
*para.PbDocParagraph = PbDoc_Document_AddParagraph(*doc, "")
*run.PbDocRun = PbDoc_Paragraph_AddRun(*para, "粗体文字 ")
PbDoc_Font_SetBold(*run, #True)
*run = PbDoc_Paragraph_AddRun(*para, "红色文字 ")
PbDoc_Font_SetColorHex(*run, "FF0000")
*run = PbDoc_Paragraph_AddRun(*para, "18pt文字")
PbDoc_Font_SetSize(*run, 18)
```

### 创建表格

```purebasic
; 创建3行4列表格
*table.PbDocTable = PbDoc_Document_AddTable(*doc, 3, 4)
PbDoc_Table_SetStyle(*table, "TableGrid")

; 填充表头
PbDoc_Cell_SetText(PbDoc_Table_GetCell(*table, 0, 0), "姓名")
PbDoc_Cell_SetText(PbDoc_Table_GetCell(*table, 0, 1), "年龄")
PbDoc_Cell_SetText(PbDoc_Table_GetCell(*table, 0, 2), "城市")
PbDoc_Cell_SetText(PbDoc_Table_GetCell(*table, 0, 3), "职业")

; 填充数据
PbDoc_Cell_SetText(PbDoc_Table_GetCell(*table, 1, 0), "张三")
PbDoc_Cell_SetText(PbDoc_Table_GetCell(*table, 1, 1), "28")
PbDoc_Cell_SetText(PbDoc_Table_GetCell(*table, 1, 2), "北京")
PbDoc_Cell_SetText(PbDoc_Table_GetCell(*table, 1, 3), "工程师")
```

## API 文档

### 初始化与清理

| 函数 | 说明 |
| --- | --- |
| `PbDoc_Init()` | 初始化 PbDoc 库，使用前必须调用 |
| `PbDoc_Cleanup()` | 清理库资源，程序结束前调用 |

### 文档操作

| 函数 | 说明 |
| --- | --- |
| `PbDocument_Create()` | 创建新的空白文档，返回文档指针 |
| `PbDocument_Save(*doc, filePath.s)` | 保存文档到 .docx 文件 |
| `PbDocument_Free(*doc)` | 释放文档内存 |
| `PbDocument_SetTitle(*doc, title.s)` | 设置文档标题 |
| `PbDocument_SetAuthor(*doc, author.s)` | 设置文档作者 |
| `PbDocument_SetSubject(*doc, subject.s)` | 设置文档主题 |
| `PbDocument_SetCompany(*doc, company.s)` | 设置公司名称 |

### 段落操作

| 函数 | 说明 |
| --- | --- |
| `PbDoc_Document_AddParagraph(*doc, text.s)` | 添加段落，返回段落指针 |
| `PbDoc_Paragraph_GetText(*para)` | 获取段落文本 |
| `PbDoc_Paragraph_SetAlignment(*para, alignment)` | 设置段落对齐方式 |
| `PbDoc_Paragraph_SetStyle(*para, styleName.s)` | 设置段落样式 |
| `PbDoc_Document_AddHeading(*doc, text.s, level.l)` | 添加标题（1-9级） |

### Run 操作

| 函数 | 说明 |
| --- | --- |
| `PbDoc_Paragraph_AddRun(*para, text.s)` | 添加文本运行，返回 Run 指针 |
| `PbDoc_Run_AddTab(*run)` | 添加 Tab 制表符 |
| `PbDoc_Run_AddBreak(*run, breakType.l)` | 添加断行符（换行/分页/分栏） |

### 字体格式

| 函数 | 说明 |
| --- | --- |
| `PbDoc_Font_SetBold(*run, bold.l)` | 设置粗体 |
| `PbDoc_Font_SetItalic(*run, italic.l)` | 设置斜体 |
| `PbDoc_Font_SetUnderline(*run, underline.l)` | 设置下划线类型 |
| `PbDoc_Font_SetStrike(*run, strike.l)` | 设置删除线 |
| `PbDoc_Font_SetDoubleStrike(*run, doubleStrike.l)` | 设置双删除线 |
| `PbDoc_Font_SetSuperscript(*run, superscript.l)` | 设置上标 |
| `PbDoc_Font_SetSubscript(*run, subscript.l)` | 设置下标 |
| `PbDoc_Font_SetName(*run, fontName.s)` | 设置字体名称 |
| `PbDoc_Font_SetColorHex(*run, colorHex.s)` | 设置字体颜色（十六进制） |
| `PbDoc_Font_SetColorRGB(*run, r.a, g.a, b.a)` | 设置字体颜色（RGB） |
| `PbDoc_Font_SetHighlight(*run, highlight.s)` | 设置高亮颜色 |
| `PbDoc_Font_SetSize(*run, size.l)` | 设置字号（磅值） |

### 段落格式

| 函数 | 说明 |
| --- | --- |
| `PbDoc_ParaFormat_SetFirstLineIndent(*para, emu.q)` | 设置首行缩进 |
| `PbDoc_ParaFormat_SetLeftIndent(*para, emu.q)` | 设置左缩进 |
| `PbDoc_ParaFormat_SetRightIndent(*para, emu.q)` | 设置右缩进 |
| `PbDoc_ParaFormat_SetHangingIndent(*para, emu.q)` | 设置悬挂缩进 |
| `PbDoc_ParaFormat_SetSpaceBefore(*para, emu.q)` | 设置段前间距 |
| `PbDoc_ParaFormat_SetSpaceAfter(*para, emu.q)` | 设置段后间距 |
| `PbDoc_ParaFormat_SetLineSpacing(*para, spacing.q, rule.l)` | 设置行间距 |
| `PbDoc_ParaFormat_SetKeepNext(*para, keepNext.l)` | 与下段同页 |
| `PbDoc_ParaFormat_SetKeepLines(*para, keepLines.l)` | 段中不分页 |
| `PbDoc_ParaFormat_SetWidowControl(*para, widowControl.l)` | 孤行控制 |
| `PbDoc_ParaFormat_AddTabStop(*para, position.q, alignment.l, leader.l)` | 添加制表符 |

### 表格操作

| 函数 | 说明 |
| --- | --- |
| `PbDoc_Document_AddTable(*doc, rows.l, cols.l)` | 创建表格 |
| `PbDoc_Table_SetStyle(*tbl, style.s)` | 设置表格样式 |
| `PbDoc_Table_SetAlignment(*tbl, alignment.l)` | 设置表格对齐 |
| `PbDoc_Table_SetColumnWidth(*tbl, colIndex.l, widthEmu.q)` | 设置列宽 |
| `PbDoc_Table_GetCell(*tbl, row.l, col.l)` | 获取单元格引用 |
| `PbDoc_Cell_SetText(*cell, text.s)` | 设置单元格文本 |

### 超链接与分页

| 函数 | 说明 |
| --- | --- |
| `PbDoc_Document_AddHyperlink(*doc, url.s, text.s)` | 添加超链接 |
| `PbDoc_Document_AddPageBreak(*doc)` | 添加分页符 |

### 页面设置

| 函数 | 说明 |
| --- | --- |
| `PbDoc_Section_SetPageSize(*doc, widthTwips.q, heightTwips.q)` | 设置页面大小 |
| `PbDoc_Section_SetMargins(*doc, top.q, bottom.q, left.q, right.q)` | 设置页边距 |
| `PbDoc_Section_SetOrientation(*doc, orientation.l)` | 设置页面方向 |
| `PbDoc_Section_SetHeaderDistance(*doc, dist.q)` | 设置页眉距离 |
| `PbDoc_Section_SetFooterDistance(*doc, dist.q)` | 设置页脚距离 |
| `PbDoc_Section_SetGutter(*doc, gutter.q)` | 设置装订线 |

### 单位转换

| 函数 | 说明 |
| --- | --- |
| `PbDoc_InchesToEmu(inches.d)` | 英寸转 EMU |
| `PbDoc_CmToEmu(cm.d)` | 厘米转 EMU |
| `PbDoc_MmToEmu(mm.d)` | 毫米转 EMU |
| `PbDoc_PtToEmu(pt.d)` | 磅转 EMU |
| `PbDoc_PtToTwips(pt.d)` | 磅转缇 |
| `PbDoc_CmToTwips(cm.d)` | 厘米转缇 |
| `PbDoc_InchesToTwips(inches.d)` | 英寸转缇 |

### 枚举常量

| 枚举 | 值 | 说明 |
| --- | --- | --- |
| `#PbDoc_ALIGN_LEFT` | 1 | 左对齐 |
| `#PbDoc_ALIGN_CENTER` | 2 | 居中对齐 |
| `#PbDoc_ALIGN_RIGHT` | 3 | 右对齐 |
| `#PbDoc_ALIGN_JUSTIFY` | 4 | 两端对齐 |
| `#PbDoc_ALIGN_DISTRIBUTE` | 5 | 分散对齐 |
| `#PbDoc_TAB_LEFT` | 1 | 左对齐Tab |
| `#PbDoc_TAB_CENTER` | 2 | 居中Tab |
| `#PbDoc_TAB_RIGHT` | 3 | 右对齐Tab |
| `#PbDoc_TAB_DECIMAL` | 4 | 小数点对齐Tab |
| `#PbDoc_TAB_LEADER_NONE` | 1 | 无前导符 |
| `#PbDoc_TAB_LEADER_DOTS` | 2 | 点线前导符 |
| `#PbDoc_TAB_LEADER_DASHES` | 4 | 短划线前导符 |
| `#PbDoc_UNDERLINE_SINGLE` | 1 | 单下划线 |
| `#PbDoc_UNDERLINE_DOUBLE` | 2 | 双下划线 |
| `#PbDoc_BREAK_LINE` | 1 | 换行符 |
| `#PbDoc_BREAK_PAGE` | 2 | 分页符 |
| `#PbDoc_BREAK_COLUMN` | 3 | 分栏符 |
| `#PbDoc_LINE_SPACING_AUTO` | 1 | 自动行距 |
| `#PbDoc_LINE_SPACING_EXACT` | 2 | 精确行距 |
| `#PbDoc_ORIENT_PORTRAIT` | 1 | 纵向 |
| `#PbDoc_ORIENT_LANDSCAPE` | 2 | 横向 |

## 文件结构

PbDoc.pb 文件按照功能模块分为以下35个版块：

| 版块 | 内容 |
| --- | --- |
| 一 | 编译器指令与初始化 |
| 二 | 常量定义（EMU/缇换算、版本号） |
| 三 | 枚举定义（对齐、Tab、下划线、断行、行距、页面方向） |
| 四 | 数据结构定义（Run、段落、表格、单元格、节、文档等） |
| 五 | 全局变量 |
| 六 | 工具函数 - 单位转换 |
| 七 | 工具函数 - 错误处理 |
| 八 | 工具函数 - XML 辅助 |
| 九 | 初始化与清理 |
| 十 | 关系管理 |
| 十一 | 文档创建与释放 |
| 十二 | 文档属性设置 |
| 十三 | 段落操作 |
| 十四 | Run 操作 |
| 十五 | 字体格式操作 |
| 十六 | 段落格式操作 |
| 十七 | 标题操作 |
| 十八 | 超链接操作 |
| 十九 | 分页符与断行操作 |
| 二十 | 表格操作 |
| 二十一 | 节与页面设置 |
| 二十二 | 文档统计 |
| 二十三~三十三 | XML 生成（Run/段落/表格/分页/文档/关系/内容类型/属性/样式/静态文件） |
| 三十四 | ZIP 包辅助函数 |
| 三十五 | 文档保存 |

## 示例文件

| 文件 | 说明 |
| --- | --- |
| `01_基础文档操作.pb` | 创建文档、属性、段落/标题、对齐 |
| `02_段落格式设置.pb` | 缩进、间距、行距、分页控制 |
| `03_字体格式设置.pb` | 粗体/斜体/下划线/删除线/上下标/颜色/高亮/字号 |
| `04_表格操作.pb` | 创建表格、样式、对齐、列宽、单元格 |
| `05_超链接与分页符.pb` | 超链接、分页符、换行符、分栏符 |
| `06_TabStop与页面设置.pb` | Tab制表符、前导符、页面大小/边距/方向 |
| `07_控制台全面测试.pb` | 控制台输出测试 |
| `08_全面测试.pb` | 完整功能测试 |

## 版本历史

### v1.3 (2026-04-23)

- \[新增] 控制台全面测试示例 (07\_控制台全面测试.pb)
- \[新增] 全面测试示例 (08\_全面测试.pb)
- \[新增] README.md 项目文档
- \[新增] HTML 帮助文档 (docs\PbDoc\_Help.html)
- \[优化] 完善代码中文注释
- \[优化] 文件版块分区，35个版块清晰划分

### v1.2 (2026-04-23)

- \[新增] 下划线类型 (7种: 单线/双线/点线/短划线/波浪/粗线)
- \[新增] 删除线/双删除线 (PbDoc\_Font\_SetStrike/SetDoubleStrike)
- \[新增] 上下标 (PbDoc\_Font\_SetSuperscript/SetSubscript)
- \[新增] 字体名称设置 (PbDoc\_Font\_SetName)
- \[新增] RGB颜色设置 (PbDoc\_Font\_SetColorRGB)
- \[新增] 高亮颜色 (PbDoc\_Font\_SetHighlight)
- \[新增] 左缩进/右缩进/悬挂缩进
- \[新增] 行间距设置 (自动/精确/最小三种规则)
- \[新增] 分页控制 (与下段同页/段中不分页/孤行控制)
- \[新增] 段落样式设置 (PbDoc\_Paragraph\_SetStyle)
- \[新增] 断行符 (换行/分页/分栏)
- \[新增] 表格对齐 (PbDoc\_Table\_SetAlignment)
- \[新增] 表格列宽设置 (PbDoc\_Table\_SetColumnWidth)
- \[新增] 页面设置 (纸张大小/边距/方向/页眉页脚距离/装订线)
- \[新增] Title 样式定义
- \[新增] 6个详细示例文件 (01-06)
- \[优化] 描述文字更新: "基于python-docx移植" → "参考python-docx项目编写"
- \[优化] 提示信息中文化
- \[优化] 代码版块分区，35个版块目录索引

### v1.1 (2026-04-22)

- \[修复] 修复无效内存访问错误 (地址 1101004800)
- \[修复] 修复 #PB\_Packer\_Zip 常量不存在 (改用 #PB\_PackerPlugin\_Zip)
- \[修复] 修复 PbDoc\_Cell\_SetText 参数不匹配
- \[修复] 修复 demon-test.pb 的 EnableExplicit 变量声明问题
- \[优化] 放弃 MSXML COM，改用字符串构建 XML，避免内存管理问题
- \[优化] 使用 PureBasic 内置 ZipPacker 替代手动 ZIP 操作
- \[优化] 数据结构驱动设计，所有内容存储在结构体中

### v1.0 (2026-04-21)

- \[新增] 初始项目创建，参考 python-docx 项目编写
- \[新增] 文档创建与保存 (PbDocument\_Create/Save/Free)
- \[新增] 文档属性设置 (标题/作者/主题/公司)
- \[新增] 段落操作 (添加段落/获取文本/设置对齐)
- \[新增] Run 操作 (添加Run/Tab)
- \[新增] 字体格式 (粗体/斜体/颜色/字号)
- \[新增] 段落格式 (首行缩进/段前段后间距/TabStop)
- \[新增] 标题操作 (1-9级标题)
- \[新增] 超链接操作
- \[新增] 分页符操作
- \[新增] 表格操作 (创建/样式/单元格)
- \[新增] 单位转换函数 (EMU/缇/磅/厘米/英寸)
- \[新增] 完整 XML 生成 (document/styles/settings/fontTable/numbering/theme/core/app)
- \[新增] ZIP 打包保存
- \[新增] 枚举定义 (对齐/Tab/Tab前导符/Body元素类型)
- \[新增] 数据结构定义 (Run/段落/表格/单元格/关系/文档)

## 许可证

本库采用 Apache 2.0 许可证。

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

本库参考的 python-docx 项目采用 MIT 许可证。

```
This software is under the MIT Licence
======================================

Copyright (c) 2010 python-docx

Permission is hereby granted, free of charge, to any person obtaining a
copy of this software and associated documentation files (the
"Software"), to deal in the Software without restriction, including
without limitation the rights to use, copy, modify, merge, publish,
distribute, sublicense, and/or sell copies of the Software, and to
permit persons to whom the Software is furnished to do so, subject to
the following conditions:

The above copyright notice and this permission notice shall be included
in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS
OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY
CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT,
TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE
SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

```

## 致谢

- 感谢 python-docx 项目提供了优秀的参考实现
- 感谢 PureBasic QQ群的支持
