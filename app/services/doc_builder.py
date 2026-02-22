"""
DOC-X 系列核心业务引擎。
从 main.py 迁移并增强的 Markdown → Docx 渲染器。
"""

import re
import logging
from io import BytesIO
from typing import Optional, Any

import mistune
import requests
from PIL import Image
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

logger = logging.getLogger(__name__)


# =====================================================
#  MarkdownToDocx 渲染器 (从 main.py 迁移)
# =====================================================

class MarkdownToDocx:
    """将 mistune 3.x AST 渲染为 python-docx Document"""

    def __init__(self, doc: Document):
        self.doc = doc
        self.table = None
        self.row = None
        self.cell = None
        self.list_style = None

    def set_font(self, run, bold=False, italic=False):
        run.font.name = '宋体'
        run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
        run.font.name = 'Times New Roman'
        run._element.rPr.rFonts.set(qn('w:ascii'), 'Times New Roman')
        if bold:
            run.bold = True
        if italic:
            run.italic = True

    def render(self, ast):
        for node in ast:
            self.dispatch(node)

    def dispatch(self, node):
        method_name = f"visit_{node['type']}"
        method = getattr(self, method_name, self.visit_unknown)
        method(node)

    def visit_unknown(self, node):
        if 'children' in node:
            for child in node['children']:
                self.dispatch(child)

    def visit_heading(self, node):
        level = node.get('attrs', {}).get('level', 1)
        text = self.get_text(node)
        heading = self.doc.add_paragraph(style=f'Heading {level}')
        run = heading.add_run(text)
        self.set_font(run, bold=True)
        size_map = {1: 16, 2: 15, 3: 14}
        run.font.size = Pt(size_map.get(level, 12))
        heading.paragraph_format.space_before = Pt(22)
        heading.paragraph_format.space_after = Pt(11)

    def visit_paragraph(self, node):
        if len(node.get('children', [])) == 1 and node['children'][0]['type'] == 'image':
            self.visit_image(node['children'][0])
            return
        p = self.doc.add_paragraph()
        if 'children' in node:
            self.render_inline(p, node['children'])
        p.paragraph_format.line_spacing = Pt(25)
        p.paragraph_format.first_line_indent = Cm(0.74)

    def visit_block_code(self, node):
        code = node.get('raw', '')
        p = self.doc.add_paragraph()
        p.style = 'No Spacing'
        run = p.add_run(code)
        run.font.name = 'Courier New'
        run.font.size = Pt(10)
        p.paragraph_format.left_indent = Cm(1)

    def visit_list(self, node):
        ordered = node.get('attrs', {}).get('ordered', False)
        self.list_style = 'List Number' if ordered else 'List Bullet'
        for child in node.get('children', []):
            self.visit_list_item(child)
        self.list_style = None

    def visit_list_item(self, node):
        p = self.doc.add_paragraph(style=self.list_style)
        if 'children' in node:
            for child in node['children']:
                if child['type'] in ('paragraph', 'block_text'):
                    self.render_inline(p, child.get('children', []))
                elif child['type'] == 'block_code':
                    run = p.add_run('\n' + child.get('raw', ''))
                    run.font.name = 'Courier New'
                elif child['type'] == 'list':
                    self.visit_list(child)
                elif 'children' in child:
                    self.render_inline(p, child.get('children', []))

    def visit_table(self, node):
        children = node.get('children', [])
        thead = next((c for c in children if c['type'] == 'table_head'), None)
        tbody = next((c for c in children if c['type'] == 'table_body'), None)

        all_rows = []
        if thead:
            header_cells = thead.get('children', [])
            if header_cells and header_cells[0].get('type') == 'table_cell':
                all_rows.append(header_cells)
            else:
                for row in header_cells:
                    all_rows.append(row.get('children', []))
        if tbody:
            for row in tbody.get('children', []):
                all_rows.append(row.get('children', []))

        if not all_rows:
            return

        col_count = max(len(cells) for cells in all_rows)
        self.table = self.doc.add_table(rows=len(all_rows), cols=col_count)
        self.table.style = "Table Grid"

        for i, cells in enumerate(all_rows):
            self.row = self.table.rows[i]
            for j, cell_node in enumerate(cells):
                self.cell = self.row.cells[j]
                self.cell._element.clear_content()
                p = self.cell.add_paragraph()
                self.render_inline(p, cell_node.get('children', []))
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in p.runs:
                    self.set_font(run)

    def visit_image(self, node):
        url = node.get('attrs', {}).get('url', '')
        if not url:
            return

        try:
            response = requests.get(url, timeout=30, headers={
                'User-Agent': 'SGA-Office/1.0'
            })
            if response.status_code != 200:
                self.doc.add_paragraph(f"[图片下载失败: HTTP {response.status_code}]")
                return

            image_stream = BytesIO(response.content)
            img = Image.open(image_stream)
            img_width, img_height = img.size

            # RGBA/P → RGB 转换
            if img.mode in ('RGBA', 'P', 'LA'):
                rgb_img = Image.new('RGB', img.size, (255, 255, 255))
                if img.mode == 'P':
                    img = img.convert('RGBA')
                rgb_img.paste(img, mask=img.split()[-1] if img.mode == 'RGBA' else None)
                img = rgb_img
            elif img.mode != 'RGB':
                img = img.convert('RGB')

            output_stream = BytesIO()
            img.save(output_stream, format='JPEG', quality=95)
            output_stream.seek(0)

            # 计算合适的宽度
            max_width_cm = 14
            aspect_ratio = img_width / img_height if img_height > 0 else 1
            final_width_cm = max_width_cm
            expected_height_cm = max_width_cm / aspect_ratio
            if expected_height_cm > 18:
                final_width_cm = 18 * aspect_ratio
            final_width_cm = final_width_cm * 0.6

            paragraph = self.doc.add_paragraph()
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            paragraph.paragraph_format.space_before = Pt(12)
            paragraph.paragraph_format.space_after = Pt(12)

            run = paragraph.add_run()
            inline_shape = run.add_picture(output_stream, width=Cm(final_width_cm))

            # 设置上下型环绕布局
            inline = inline_shape._inline
            anchor = OxmlElement('wp:anchor')
            for attr, val in [
                ('wp:distT', '0'), ('wp:distB', '0'),
                ('wp:distL', '114300'), ('wp:distR', '114300'),
                ('wp:simplePos', '0'), ('wp:relativeHeight', '0'),
                ('wp:behindDoc', '0'), ('wp:locked', '0'),
                ('wp:layoutInCell', '1'), ('wp:allowOverlap', '1'),
            ]:
                anchor.set(qn(attr), val)

            simplePos = OxmlElement('wp:simplePos')
            simplePos.set('x', '0')
            simplePos.set('y', '0')
            anchor.append(simplePos)

            positionH = OxmlElement('wp:positionH')
            positionH.set('relativeFrom', 'column')
            align_elem = OxmlElement('wp:align')
            align_elem.text = 'center'
            positionH.append(align_elem)
            anchor.append(positionH)

            positionV = OxmlElement('wp:positionV')
            positionV.set('relativeFrom', 'paragraph')
            posOffset = OxmlElement('wp:posOffset')
            posOffset.text = '0'
            positionV.append(posOffset)
            anchor.append(positionV)

            for child in inline:
                anchor.append(child)

            wrapTopAndBottom = OxmlElement('wp:wrapTopAndBottom')
            anchor.append(wrapTopAndBottom)

            drawing = inline.getparent()
            drawing.remove(inline)
            drawing.append(anchor)

        except Exception as e:
            logger.error(f"图片处理失败: {e}")
            self.doc.add_paragraph(f"[图片处理失败: {str(e)}]")

    def render_inline(self, paragraph, nodes):
        for node in nodes:
            if node['type'] == 'text':
                run = paragraph.add_run(node.get('raw', ''))
                self.set_font(run)
            elif node['type'] == 'strong':
                run = paragraph.add_run(self.get_text(node))
                self.set_font(run, bold=True)
            elif node['type'] == 'emphasis':
                run = paragraph.add_run(self.get_text(node))
                self.set_font(run, italic=True)
            elif node['type'] == 'codespan':
                run = paragraph.add_run(node.get('raw', ''))
                run.font.name = 'Courier New'
                self.set_font(run)
            elif node['type'] == 'image':
                pass  # inline image skip

    def get_text(self, node):
        if 'raw' in node:
            return node['raw']
        if 'children' in node:
            return "".join([self.get_text(child) for child in node['children']])
        return ""


# =====================================================
#  预处理工具函数
# =====================================================

def _convert_tab_tables_to_markdown(content: str) -> str:
    """将 Tab 分隔的表格转换为 Markdown 表格"""
    lines = content.split('\n')
    result = []
    i = 0
    while i < len(lines):
        line = lines[i]
        if '\t' in line and line.count('\t') >= 1:
            table_lines = []
            while i < len(lines) and '\t' in lines[i]:
                table_lines.append(lines[i])
                i += 1
            if len(table_lines) >= 2:
                for j, tline in enumerate(table_lines):
                    cells = tline.split('\t')
                    md_row = '| ' + ' | '.join(cells) + ' |'
                    result.append(md_row)
                    if j == 0:
                        separator = '| ' + ' | '.join(['---'] * len(cells)) + ' |'
                        result.append(separator)
            else:
                result.extend(table_lines)
        else:
            result.append(line)
            i += 1
    return '\n'.join(result)


def _repair_markdown_table(content: str) -> str:
    """修复 Markdown 表格格式"""
    lines = content.split('\n')
    has_separator = any(
        re.match(r'^\s*\|?[\s:-]{3,}\|[\s:-]{3,}\|?.*$', line)
        for line in lines
    )
    if not has_separator:
        return content

    new_lines = []
    for line in lines:
        stripped = line.strip()
        if '|' in stripped and len(stripped) > 1:
            if not stripped.startswith('|'):
                stripped = '| ' + stripped
            if not stripped.endswith('|'):
                stripped = stripped + ' |'
            new_lines.append(stripped)
        else:
            new_lines.append(line)
    return '\n'.join(new_lines)


# =====================================================
#  DOC-01: render_markdown_to_docx
# =====================================================

def render_markdown_to_docx(
    markdown_content: str,
) -> BytesIO:
    """
    将 Markdown 文本渲染为标准版式 Word 文档。

    Args:
        markdown_content: Markdown 原文

    Returns:
        BytesIO 对象，包含生成的 .docx 数据
    """
    # 预处理
    content = markdown_content
    if content.startswith('\\#'):
        content = content[1:]
    content = content.replace('\\n', '\n')
    content = _convert_tab_tables_to_markdown(content)
    content = _repair_markdown_table(content)

    doc = Document()

    # 页面设置 (A4)
    section = doc.sections[0]
    section.page_height = Cm(29.7)
    section.page_width = Cm(21)
    section.top_margin = Cm(3.7)
    section.bottom_margin = Cm(3.5)
    section.left_margin = Cm(2.8)
    section.right_margin = Cm(2.8)

    # 全局字体
    style = doc.styles['Normal']
    style.font.name = '宋体'
    style._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    style.font.size = Pt(12)
    style.paragraph_format.line_spacing = Pt(25)
    style.paragraph_format.first_line_indent = Cm(0.74)

    # 解析 & 渲染
    markdown = mistune.create_markdown(renderer=None, plugins=['table'])
    ast = markdown(content)
    renderer = MarkdownToDocx(doc)
    renderer.render(ast)

    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output


# =====================================================
#  DOC-02: fill_docx_template
# =====================================================

def fill_docx_template(
    template_bytes: bytes,
    variables: dict[str, Any],
) -> BytesIO:
    """
    将 Word 模板中的 {{ 占位符 }} 替换为实际变量值。

    Args:
        template_bytes: 模板 .docx 文件的原始字节
        variables: key-value 变量字典，key 对应模板中的占位符名称

    Returns:
        BytesIO 对象，包含替换后的 .docx 数据
    """
    doc = Document(BytesIO(template_bytes))

    for paragraph in doc.paragraphs:
        _replace_placeholders_in_paragraph(paragraph, variables)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    _replace_placeholders_in_paragraph(paragraph, variables)

    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output


def _replace_placeholders_in_paragraph(
    paragraph,
    variables: dict[str, Any],
) -> None:
    """在单个段落中替换 {{ key }} 占位符。"""
    for key, value in variables.items():
        placeholder = "{{" + key + "}}"
        if placeholder in paragraph.text:
            for run in paragraph.runs:
                if placeholder in run.text:
                    run.text = run.text.replace(placeholder, str(value))
