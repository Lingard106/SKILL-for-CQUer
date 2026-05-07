---
name: cqu-random-signal-report
description: "Generate experiment reports for 重庆大学微电子与通信工程学院 随机信号分析 course following strict department template format. Use when the user provides experiment requirements in a folder and asks for a report. Triggers on: 随机信号分析实验报告, 随机信号 Matlab 实验, 生成实验报告 + 文件夹路径, or when the user mentions this specific course experiment workflow."
---

# 随机信号分析实验报告生成器

Generates Word (.docx) experiment reports for 重庆大学 微电子与通信工程学院 随机信号分析 course, following the official department electronic template format.

## Core Workflow

### Step 1: Read Experiment Requirements

Read the experiment requirements file. Supported formats:
- `.docx`: use `python-docx`
- `.pdf`: use `pdfplumber` (installed). For tables in PDF, use `page.extract_tables()`. For body text, use `page.extract_text()`.

Extract:
- Experiment chapter/number (e.g., "第一章 MATLAB实验")
- Probability distribution types and parameters (Binomial, Poisson, Normal, Rayleigh etc.)
- Grouping method and parameters for the user's student ID tail number
- Sub-tasks required (PDF/CDF plotting, sample generation, statistics computation, covariance analysis)
- Student info (name, student ID, class, grade; ask if not provided)

**Parameter extraction priority**:
1. First, extract the parameter table from the requirements docx. Use `<w:t>` element extraction via lxml if python-docx `cell.text` returns empty.
2. If the parameter table in the requirements docx is genuinely empty (only column headers, no values), do NOT loop on extraction attempts. Immediately switch to the **reference document** (`随机信号第一次作业.docx` on Desktop). The reference document contains a complete parameter table (distribution type, parameter name, value) that maps to the student ID tail groups.
3. If neither source has parameters, ask the user.

### Step 2: Handle MATLAB Code

If user provides a `.m` file: use it directly, verify parameters match requirements.

If NO `.m` file: write the MATLAB script from scratch based on the experiment requirements:
- Start with `clear; clc; close all;`
- Fix random seed `rng(37)` for reproducibility
- Use MATLAB Statistics Toolbox: `pdf`, `cdf`, `random`, `mean`, `var`, `std`, `cov`, `histogram`, `ecdf`
- Save figures as PNG: `saveas(fig1, 'fig1_xxx.png')`
- Output statistics via `fprintf` with clear headers for later parsing
- Separate sub-tasks with `%%` comment blocks

### Step 3: Run MATLAB

```bash
cd "experiment-folder" && matlab -batch "run('script.m'); exit;"
```

MATLAB is at a standard Windows location in PATH. Wait up to 120 seconds. Capture console output to extract statistics values. Confirm generated PNG files exist.

Console output may have garbled encoding — parse numeric values by their table position and structure.

### Step 4: Generate Word Report

Two hardcoded resource paths (do NOT change these):
- **Reference document**: Search the user's Desktop for `随机信号第一次作业.docx`.
- **Department template**: Located in the experiment folder as a `.doc` file with "电子化模板" in its name, typically under a `Mat_exp` or similar folder on Desktop.

**Strategy**: Open the reference document directly for editing. Do NOT strip and rebuild — modify the existing content in place. This preserves all formatting, styles, and the TOC field without risking style loss.

**Editing priority** (mandatory):
1. Always `doc = Document(REF)` and edit in place, then `doc.save(OUT)` to a new path.
2. Only strip-and-rebuild as a last resort when the reference document structure is fundamentally incompatible with the current experiment.
3. The reference document already has a TOC field, Heading styles, page header, and correct margins — editing in place keeps them all intact.

**Key edit targets** in the reference document:
- **Cover page**: Update date (paragraphs ~11–12). Fix student name encoding if corrupted.
- **Section 1 heading**: Update experiment number and title if different.
- **Section 1.2**: Update tail number and parameter reference.
- **Table 1 (parameters)**: Replace with current tail group's parameters.
- **Section 1.3.x**: Replace distribution formulas with OMML versions if needed.
- **Section 1.5**: Replace old figures and statistics tables with new ones.
- **Section 1.6**: Rewrite analysis text with new data.
- **Section 1.7**: Update conclusions if needed.
- **Appendix**: Replace source code.
- **TOC**: Update the TOC field to reflect any heading changes, or add a fresh TOC after the cover.

**Report body structure** (use built-in Word styles):
- `Heading 1`: "N 实验名称" (e.g., "1 随机信号分析第一次MATLAB实验")
- `Heading 2`: N.1 实验目的
- `Heading 2`: N.2 使用仪器、材料
- `Heading 2`: N.3 实验原理 (with manual sub-sections like N.3.1, N.3.2 for each distribution)
- `Heading 2`: N.4 实验步骤
- `Heading 2`: N.5 实验过程原始记录 (figures, tables, analysis)
- `Heading 2`: N.6 实验结果及分析
- `Heading 2`: N.7 实验总结

**Appendix**: `Heading 1` "附录：MATLAB源程序清单", followed by source code in Courier New 7.5pt.

**Formatting rules** (inherited from reference document):
- Body text: 宋体/Times New Roman 12pt (小四号), line spacing exactly 20pt, first-line indent 0.74cm
- Heading 1: 黑体 12pt bold (小四号), using `doc.add_heading(text, level=1)`
- Heading 2: 黑体 12pt bold (小四号), using `doc.add_heading(text, level=2)`
- Margins: top/bottom 2.54cm, left/right 3.17cm
- Page header: "重庆大学微电子与通信工程学院实验报告" in 宋体 9pt centered

**Tables**: Use python-docx tables with full borders. Header: 黑体 10pt bold centered. Body: 宋体 10pt centered. Caption above: "表N  标题" in 黑体 10.5pt bold centered.

**Images**: Use native `run.add_picture()` — inline embedded type is the most reliable. Centered, width 5.5 inches. Caption below: "图N  标题" in 宋体 10.5pt centered. NEVER construct custom DrawingML anchor XML for image wrapping — it corrupts the file.

**Unicode**: Use escape sequences for Greek letters and math symbols (`\u03bb`, `\u03bc`, `\u03c3`, `\u03c0`, `\u2212`, `\u00b2`, `\u2248`, `\u2265`).

**Output filename**: "实验名称_实验报告_姓名.docx" in the experiment folder.

### Step 5: Report to User

After generation, tell the user:
- Full path of the generated report
- Suggest opening with WPS/Word desktop (not web version) if images have issues
- Offer to make adjustments

## Key Implementation Notes

**Cover extraction (fallback only)** — only use if the edit-in-place strategy fails because the reference document structure is fundamentally incompatible:
```python
shutil.copy(REF, OUT)
doc = Document(OUT)
body = doc.element.body
w_ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
cover = []
found = False
for c in body:
    if c.tag == f'{{{w_ns}}}p':
        pPr = c.find(f'{{{w_ns}}}pPr')
        if pPr is not None:
            s = pPr.find(f'{{{w_ns}}}pStyle')
            if s is not None and s.get(f'{{{w_ns}}}val','') in ('1','Heading1'):
                found = True
        if not found: cover.append(c)
    elif c.tag == f'{{{w_ns}}}sectPr': cover.append(c)
for c in list(body): body.remove(c)
for c in cover: body.append(c)
```

**Border helper**:
```python
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml
def add_borders(table):
    tblPr = table._tbl.tblPr
    if tblPr is None:
        tblPr = etree.SubElement(table._tbl, f'{{{w_ns}}}tblPr')
    bxml = '<w:tblBorders %s><w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/></w:tblBorders>' % nsdecls('w')
    tblPr.append(parse_xml(bxml))
```

## OMML Equation Formatting

All mathematical expressions in the report must use Word's native equation engine (OMML — Office Math Markup Language), not plain text. This ensures formulas render as proper math objects identical to those created via Alt+= in Word.

### OMML Namespace
```python
MATH_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
```

### Helper: Insert an OMML equation into a paragraph
```python
from lxml import etree

def add_omml_equation(paragraph, omml_xml_string):
    """Inject an OMML equation into an existing paragraph.
    omml_xml_string must be a valid <m:oMath> or <m:oMathPara> element."""
    run = paragraph.add_run()
    omml_el = etree.fromstring(omml_xml_string)
    run._element.append(omml_el)
    return run
```

### Common OMML templates

**Fraction (a/b)**:
```xml
<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <m:f>
    <m:fPr><m:ctrlPr><w:rPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr></m:ctrlPr></m:fPr>
    <m:num><m:r><m:t>NUMERATOR</m:t></m:r></m:num>
    <m:den><m:r><m:t>DENOMINATOR</m:t></m:r></m:den>
  </m:f>
</m:oMath>
```

**Superscript (x^n)**:
```xml
<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <m:sSup>
    <m:e><m:r><m:t>BASE</m:t></m:r></m:e>
    <m:sup><m:r><m:t>EXP</m:t></m:r></m:sup>
  </m:sSup>
</m:oMath>
```

**Subscript (x_n)**: use `<m:sSub>` instead of `<m:sSup>`.

**Radical / Square root**:
```xml
<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <m:rad>
    <m:radPr><m:degHide m:val="1"/></m:radPr>
    <m:e><m:r><m:t>EXPRESSION</m:t></m:r></m:e>
  </m:rad>
</m:oMath>
```
For nth root, remove `<m:degHide/>` and add `<m:deg><m:r><m:t>n</m:t></m:r></m:deg>`.

**Bracketed / Parenthesized expression**:
```xml
<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <m:d>
    <m:dPr><m:begChr m:val="("/><m:endChr m:val=")"/></m:dPr>
    <m:e><m:r><m:t>EXPRESSION</m:t></m:r></m:e>
  </m:d>
</m:oMath>
```

**Integral**:
```xml
<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <m:nary>
    <m:naryPr><m:chr m:val="∫"/><m:limLoc m:val="undOvr"/></m:naryPr>
    <m:sub><m:r><m:t>−∞</m:t></m:r></m:sub>
    <m:sup><m:r><m:t>+∞</m:t></m:r></m:sup>
    <m:e><m:r><m:t>f(x)dx</m:t></m:r></m:e>
  </m:nary>
</m:oMath>
```

**Greek letters in OMML**: Use Unicode directly in `<m:t>` tags:
- α ← \u03b1, β ← \u03b2, λ ← \u03bb, μ ← \u03bc
- σ ← \u03c3, π ← \u03c0, Σ ← \u03a3, Π ← \u03a0
- ∞ ← \u221e, − ← \u2212, ≈ ← \u2248, ≥ ← \u2265
- ⋅ (dot operator) ← \u22c5, √ ← \u221a

**Inline vs display mode**: Equations in `<m:oMath>` are inline (in-line with text). For display (centered, on their own line), place the equation in its own paragraph with center alignment and no first-line indent.

### Example: Gaussian PDF formula in OMML
```python
def add_gaussian_pdf_formula(paragraph):
    omml = '''<m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
      <m:oMath>
        <m:r><m:t>f(x)</m:t></m:r>
        <m:r><m:t>=</m:t></m:r>
        <m:f>
          <m:fPr><m:ctrlPr><w:rPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr></m:ctrlPr></m:fPr>
          <m:num><m:r><m:t>1</m:t></m:r></m:num>
          <m:den>
            <m:r><m:t>\u03c3</m:t></m:r>
            <m:rad>
              <m:radPr><m:degHide m:val="1"/></m:radPr>
              <m:e><m:r><m:t>2\u03c0</m:t></m:r></m:e>
            </m:rad>
          </m:den>
        </m:f>
        <m:r><m:t>exp</m:t></m:r>
        <m:d>
          <m:dPr><m:begChr m:val="("/><m:endChr m:val=")"/></m:dPr>
          <m:e>
            <m:r><m:t>−</m:t></m:r>
            <m:f>
              <m:num>
                <m:sSup>
                  <m:e><m:r><m:t>(x−\u03bc)</m:t></m:r></m:e>
                  <m:sup><m:r><m:t>2</m:t></m:r></m:sup>
                </m:sSup>
              </m:num>
              <m:den><m:r><m:t>2\u03c3</m:t><m:sSup><m:e/><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup></m:r></m:den>
            </m:f>
          </m:e>
        </m:d>
      </m:oMath>
    </m:oMathPara>'''
    add_omml_equation(paragraph, omml)
```

## Table of Contents (TOC)

Every generated report must include an automatically-updatable table of contents after the cover page.

### Adding a TOC field
Insert a TOC field after the cover page (before the first Heading 1). The TOC collects all Heading 1 and Heading 2 entries.

```python
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml

def add_toc(doc, insert_index=None):
    """Insert a TOC field into the document.
    If insert_index is provided, insert before that paragraph index.
    Otherwise append to the end of body."""
    w_ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    
    # Create TOC paragraph
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT  # noqa
    pf = p.paragraph_format
    pf.space_before = Pt(6)
    pf.space_after = Pt(6)
    
    # Field begin
    run_begin = p.add_run()
    run_begin._element.append(parse_xml(
        f'<w:fldChar {nsdecls("w")} w:fldCharType="begin"/>'
    ))
    
    # Field instruction: TOC with Heading 1-2, hyperlinks, hide page numbers in Web view
    run_instr = p.add_run()
    run_instr._element.append(parse_xml(
        f'<w:instrText {nsdecls("w")} xml:space="preserve"> TOC \\o "1-2" \\h \\z \\u </w:instrText>'
    ))
    
    # Field separator
    run_sep = p.add_run()
    run_sep._element.append(parse_xml(
        f'<w:fldChar {nsdecls("w")} w:fldCharType="separate"/>'
    ))
    
    # Placeholder text (visible until user updates field in Word)
    run_text = p.add_run('（请在Word中右键点击此处，选择"更新域"以生成目录）')
    run_text.font.name = '宋体'
    run_text._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    run_text.font.size = Pt(10.5)
    run_text.font.color.rgb = RGBColor(0x80, 0x80, 0x80)
    
    # Field end
    run_end = p.add_run()
    run_end._element.append(parse_xml(
        f'<w:fldChar {nsdecls("w")} w:fldCharType="end"/>'
    ))
    
    # Move TOC paragraph to correct position (after cover, before first Heading 1)
    if insert_index is not None:
        body = doc.element.body
        body.remove(p._element)
        ref_element = doc.paragraphs[insert_index]._element
        body.insert(list(body).index(ref_element), p._element)
```

**TOC placement**: Insert the TOC immediately after the last cover paragraph and before the first Heading 1 paragraph. If using the edit-in-place strategy, the reference document already has a TOC field — just verify it exists and update its heading level range if needed.

**Note**: The TOC will only display actual content after the user opens the file in Word and selects "Update Field" (or presses F9). This is standard Word behavior and cannot be automated by python-docx.

## Edge Cases

- If student info is missing → ask the user before proceeding
- If MATLAB fails to run → show error output, offer to debug the code
- If generated images are missing → check MATLAB saveas paths, regenerate
- If distribution parameters differ from a previous experiment → always use the current experiment's parameters, never assume

## Common Pitfalls (from field experience)

### Pitfall 1: Preset table styles don't exist
After stripping body content from the reference document, the surviving style set is minimal. `Table Grid` and similar preset table styles are often absent.
**Fix**: NEVER use `style='Table Grid'` or any named table style. Use `doc.add_table(rows, cols)` without a style argument, then manually apply borders via the `add_borders()` helper. Same applies to paragraph styles beyond `Heading 1` and `Heading 2`.

### Pitfall 2: Heading 3 doesn't exist
Only `Heading 1` and `Heading 2` styles survive the cover-stripping process in the reference document. `Heading 3` and deeper are not guaranteed.
**Fix**: For sub-sections like 1.3.1~1.3.4, use manually formatted bold paragraphs (黑体 12pt bold with first-line indent) rather than `doc.add_heading(text, level=3)`.

### Pitfall 3: Console encoding garbles Chinese output
Python's default console encoding on Windows (GBK) cannot render many Unicode characters, especially math symbols. Print statements will appear garbled even though files are written correctly.
**Fix**: When running the generation script, prepend `PYTHONIOENCODING=utf-8` or pipe output to a file. Do NOT trust console output to judge correctness — verify by checking the output file existence and size.

### Pitfall 4: Student name corruption on cover page
When copying the reference document's cover, student names containing Chinese characters may become U+FFFD replacement characters due to encoding round-tripping through the filesystem or shell.
**Fix**: After generating the report, locate the name paragraph on the cover (typically paragraph index ~11, the run containing U+FFFD characters after "学生姓名：") and overwrite the run text with the correct Chinese name. Then save as the final output filename.

### Pitfall 5: Output file locked by Word
If the user has the old report file open in Word (including the `~$` lock file), `os.remove()` on the old file will fail with PermissionError.
**Fix**: Save the corrected file with a new name first, then attempt removal of the old file. If removal fails, leave both files and tell the user to delete the old one manually.
