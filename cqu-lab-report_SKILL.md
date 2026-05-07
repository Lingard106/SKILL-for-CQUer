---
name: cqu-lab-report
description: "通用实验报告生成器，适用于重庆大学微电子与通信工程学院所有课程的实验报告。支持多种实验类型（软件仿真、实验室实物操作），支持新建报告和增量续写两种模式。Triggers on: 实验报告, 续写实验报告, 实验数据, 生成报告, lab report, 这次实验, 实验做完, 帮我写报告, 实验记录"
---

# 重庆大学实验报告通用生成器

适用于微电子与通信工程学院所有实验课程，使用统一的学院电子化模板。

## 核心原则

1. **编辑优先**：永远优先在已有文档上编辑，不轻易从零创建
2. **能力边界清晰**：先判断报告所需数据从哪来，明确分工
3. **续写增量**：长期实验每次只追加新章节，保留已有内容不变
4. **成本敏感**：按需加载，不一次性加载不相关内容

## Phase 1：识别与路由

接到实验报告需求时，先判断：

### 1.1 实验类型

| 类型 | 典型工具 | 数据来源 |
|------|----------|----------|
| 纯软件仿真 | MATLAB, Python, ModelSim TCL | 用户提供代码/脚本，按用户指示在用户环境中执行并收集输出 |
| 半自动化软件 | Quartus, Vivado, Multisim | 用户提供截图、关键数据 |
| 实验室实物 | 示波器, 信号发生器, 面包板, 万用表 | 用户提供数据和照片 |

### 1.2 报告模式

| 模式 | 特征 | 处理方式 |
|------|------|----------|
| 一次性报告 | 一次作业 = 一份独立报告 | 从模板创建新文档 → 填充 → 输出 |
| 增量续写 | 一个学期一个文档，每次实验追加 | 打开已有文档 → 定位续写点 → 追加新章节 → 保存 |

## Phase 2：数据收集

### 读取实验要求文件

支持的格式：
- `.docx`：使用 `python-docx`
- `.pdf`：使用 `pdfplumber`（已安装，v0.11.9）。对于 PDF 中的表格，使用 `page.extract_tables()`；对于正文文本，使用 `page.extract_text()`。如果 PDF 是扫描版（图片型），`extract_text()` 可能为空，此时需要用户手动提供文字内容。

### 分工清单

以下事项在用户明确指示且提供文件后可代为执行：
- 编写 MATLAB `.m` 脚本
- 编写 Python 数据处理脚本
- 从 MATLAB 控制台输出解析统计量
- 从已有代码文件生成源代码附录

### 需要用户提供的内容

- 实验的具体题目和要求（如果没有文档文件）
- GUI 软件的截图（Quartus 综合报告、Multisim 电路图、ModelSim 波形等）
- 实验室实物照片（电路连接、示波器屏幕、仪器读数）
- 手工测量的数据表格
- 实验过程中遇到的问题和解决记录

### 数据收集模板

当判断为需要用户提供数据时，发送以下清单：

```
本次实验需要你提供：
1. [必须] 实验名称和实验目的
2. [必须] 实验数据（测量值、波形图、截图等）
3. [可选] 实验过程中遇到的问题和解决方式
4. [可选] 实物连接照片（至少1张）
5. [如果用到代码] 源码文件
```

## Phase 3：报告生成 / 续写

### 3.1 新建报告流程

1. **确定模板文档**：优先级如下
   - 用户提供的该课程已有 .docx 报告（优先编辑）
   - 桌面上的 `随机信号第一次作业.docx`（标准封面和样式）
   - 如以上都不存在，从 `重庆大学微电子与通信工程学院课程实验报告电子化模板.doc` 转换
2. **获取学生信息**：检查记忆中是否有用户的学生信息。如缺失，询问用户后 pin 到记忆。
3. **打开并编辑**：`doc = Document(template_path)`
4. **修改封面**：更新实验名称、课程名称、学生信息、日期
5. **清除旧正文**（如果编辑已有文档）：保留封面，删除 Heading 1 之后的所有内容
6. **构建正文**（见 Phase 4）
7. **添加/更新 TOC**
8. **保存**为 `实验名称_实验报告_姓名.docx`

### 3.2 增量续写流程

1. **打开已有报告**：`doc = Document(existing_report_path)`
2. **定位续写点**：查找文档中最后一个 `<w:bookmarkStart w:name="append_point"/>` 书签
   - 如果存在，在书签所在段落之后插入新内容
   - 如果不存在（旧版报告），在最后一个 Heading 2 之后插入
3. **追加新章节**：
   - 新的 Heading 2（如 `2.3 实验三：xxx`）
   - 按 Phase 4 的结构填充该节内容
4. **更新续写点书签**：在文档末尾添加新的续写点书签
5. **更新 TOC 域**：提示用户在 Word 中右键更新目录
6. **保存**（覆盖原文件或另存为新版本）

**续写点书签实现参考**：
```python
from docx.oxml.ns import qn
from docx.oxml import parse_xml

def add_append_bookmark(paragraph):
    run = paragraph.add_run()
    bm_start = parse_xml(
        '<w:bookmarkStart xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'w:id="99" w:name="append_point"/>'
    )
    run._element.append(bm_start)
    bm_end = parse_xml(
        '<w:bookmarkEnd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'w:id="99"/>'
    )
    run._element.append(bm_end)
    return run

def find_append_point(doc):
    body = doc.element.body
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    bookmarks = body.findall('.//w:bookmarkStart', ns)
    for bm in bookmarks:
        if bm.get(qn('w:name')) == 'append_point':
            parent = bm.getparent()
            while parent is not None and parent.tag != \
                    f'{{http://schemas.openxmlformats.org/wordprocessingml/2006/main}}p':
                parent = parent.getparent()
            if parent is not None:
                for i, p in enumerate(doc.paragraphs):
                    if p._element == parent:
                        return i
    return -1
```

## Phase 4：正文结构

所有实验报告（无论课程）统一使用以下章节结构。用 Heading 1 和 Heading 2 构建。

```
封面（保留模板封面）

目录（TOC 域）

Heading 1: "N 实验名称"
  Heading 2: N.1 实验目的
  Heading 2: N.2 使用仪器、材料
  Heading 2: N.3 实验原理
  Heading 2: N.4 实验步骤
  Heading 2: N.5 实验过程原始记录（数据、截图、照片）
  Heading 2: N.6 实验结果及分析
  Heading 2: N.7 实验总结
```

**多节实验的编号**：一次报告中包含多次实验时，Heading 1 为总标题，每次独立实验用 Heading 2（如 `2.1 实验一：xxx`），子标题用手工加粗段落。

## Phase 5：格式规范

### 5.1 OMML 公式

所有数学表达式使用 OMML（Office Math Markup Language），详见 `cqu-random-signal-report` skill 中的 OMML Equation Formatting 章节。本 skill 引用该规范但不重复。

### 5.2 表格

- 边框：全边框（单线，4/8 pt），不使用预设样式名
- 表头：黑体 10pt 加粗居中
- 表体：宋体 10pt 居中
- 表注：表上方，黑体 10.5pt 加粗居中，"表N  标题"

### 5.3 图片

- 嵌入方式：`run.add_picture(path, width=Inches(5.5))`
- 居中对齐
- 图注：图下方，宋体 10.5pt 居中，"图N  标题"
- 禁止构造 DrawingML anchor XML 做文字环绕

### 5.4 正文格式

- 字体：宋体/Times New Roman 12pt（小四号）
- 行距：固定值 20pt
- 首行缩进：0.74cm
- 页边距：上下 2.54cm，左右 3.17cm（从模板继承，通常不需修改）
- 页眉：宋体 9pt 居中

## Phase 6：输出后处理

生成完成后的标准通知：

1. 报告路径
2. 提示用户：
   - 在 Word/WPS 中右键目录 → 更新域
   - 检查封面信息是否正确
   - 如有图片显示问题，用桌面版 Word 打开
3. 增量续写后额外提示：本次新增了哪些章节

## Pitfalls

### 通用坑（继承自 cqu-random-signal-report）
- 不要用预设表格样式名（`Table Grid` 等）
- 不要用 Heading 3（用加粗段落替代）
- Python 脚本设置 `PYTHONIOENCODING=utf-8`
- 旧文件被 Word 占用时先另存
- 封面名字编码损坏时手动修复 run 文本

### 新增坑
- **模板格式**：.doc 格式需先转换为 .docx 才能编辑。如果无法自动转换，告知用户提供 .docx 版本
- **续写冲突**：续写前检查文件是否被占用（`~$` 文件），如被占用通知用户关闭 Word
- **TOC 失效**：大量修改后 TOC 域可能错位，建议每次生成后重新插入 TOC 域
- **书签 ID 冲突**：续写点书签使用固定 w:id="99"，如果文档中已有同名书签需先清理