from flask import Flask, request, jsonify, send_file
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from datetime import datetime
import os
import re
from urllib.parse import quote
import json
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill
from io import BytesIO
from qcloud_cos import CosConfig, CosS3Client
import mistune
import requests

app = Flask(__name__)

# 腾讯云COS配置
SECRET_ID = os.getenv('COS_SECRET_ID')
SECRET_KEY = os.getenv('COS_SECRET_KEY')
REGION = os.getenv('COS_REGION', 'ap-guangzhou')
BUCKET_NAME = os.getenv('COS_BUCKET_NAME', 'difyfordoc-1323080521')

if not all([SECRET_ID, SECRET_KEY]):
    print("Warning: COS credentials not found in environment variables.")

config = CosConfig(Region=REGION, SecretId=SECRET_ID, SecretKey=SECRET_KEY)
cos_client = CosS3Client(config)

####################################
#          Markdown 渲染器
####################################
class MarkdownToDocx:
    def __init__(self, doc):
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
        if level == 1:
            run.font.size = Pt(16)
        elif level == 2:
            run.font.size = Pt(15)
        elif level == 3:
            run.font.size = Pt(14)
        heading.paragraph_format.space_before = Pt(22)
        heading.paragraph_format.space_after = Pt(11)

    def visit_paragraph(self, node):
        # 如果段落里只有图片，特殊处理
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
                 if child['type'] == 'paragraph':
                     self.render_inline(p, child['children'])
                 elif child['type'] == 'block_code':
                     run = p.add_run('\n' + child.get('raw', ''))
                     run.font.name = 'Courier New'
                 else:
                     pass

    def visit_table(self, node):
        children = node.get('children', [])
        thead = next((c for c in children if c['type'] == 'table_head'), None)
        tbody = next((c for c in children if c['type'] == 'table_body'), None)
        
        rows = []
        if thead:
            rows.extend(thead.get('children', []))
        if tbody:
            rows.extend(tbody.get('children', []))
            
        if not rows:
            return

        # Calculate max columns to ensure all data fits
        col_count = max(len(r.get('children', [])) for r in rows)
        self.table = self.doc.add_table(rows=len(rows), cols=col_count)
        self.table.style = "Table Grid"
        
        for i, row_node in enumerate(rows):
            self.row = self.table.rows[i]
            for j, cell_node in enumerate(row_node.get('children', [])):
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
            response = requests.get(url, timeout=10)
            if response.status_code == 200:
                image_stream = BytesIO(response.content)
                p = self.doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = p.add_run()
                run.add_picture(image_stream, width=Cm(15)) 
        except Exception as e:
            print(f"Failed to download image: {e}")
            p = self.doc.add_paragraph(f"[图片下载失败: {url}]")

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
            elif node['type'] == 'image':
                 # Inline image - simplified
                 pass
            elif node['type'] == 'codespan':
                 run = paragraph.add_run(node.get('raw', ''))
                 run.font.name = 'Courier New'
                 self.set_font(run)

    def get_text(self, node):
        if 'raw' in node:
            return node['raw']
        if 'children' in node:
            return "".join([self.get_text(child) for child in node['children']])
        return ""

#         Docx 生成接口
####################################
@app.route('/')
def home():
    return "Flask 服务器运行成功！"

@app.route('/generate-doc', methods=['POST'])
def generate_doc():
    data = request.json
    if not data:
        return jsonify({'error': '没有提供数据'}), 400
    
    filename_input = data.get("filename", "默认文档").strip()
    content = data.get("content", "")
    
    if not content:
         return jsonify({'error': '内容不能为空'}), 400

    try:
        doc = Document()

        # ========== 页面设置 ==========
        section = doc.sections[0]
        section.page_height = Cm(29.7)
        section.page_width = Cm(21)
        section.top_margin = Cm(3.7)
        section.bottom_margin = Cm(3.5)
        section.left_margin = Cm(2.8)
        section.right_margin = Cm(2.8)

        # ========== 全局字体设置 ==========
        style = doc.styles['Normal']
        style.font.name = '宋体'
        style._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
        style.font.size = Pt(12)
        style.paragraph_format.line_spacing = Pt(25)
        style.paragraph_format.first_line_indent = Cm(0.74)

        # ========== Markdown 解析与渲染 ==========
        markdown = mistune.create_markdown(renderer=None, plugins=['table'])
        ast = markdown(content)
        
        renderer = MarkdownToDocx(doc)
        renderer.render(ast)

        # ========== 生成文件名 ==========
        clean_title = re.sub(r'[\\/:*?"<>|\s]', '', filename_input)[:30]
        if not clean_title:
            clean_title = "无标题文档"
        current_date = datetime.now().strftime("%Y%m%d")
        filename = f"{clean_title}_{current_date}.docx"

        # ========== 保存并上传到COS ==========
        temp_file = f"temp_{datetime.now().strftime('%Y%m%d%H%M%S')}.docx"
        doc.save(temp_file)
        
        cos_path = f"documents/{filename}"
        try:
            cos_client.upload_file(
                Bucket=BUCKET_NAME,
                LocalFilePath=temp_file,
                Key=cos_path
            )
        except Exception as upload_error:
            app.logger.error(f"上传失败: {str(upload_error)}")
            return jsonify({'error': '文件上传失败'}), 500
        finally:
            if os.path.exists(temp_file):
                os.remove(temp_file)
        
        file_url = f"https://{BUCKET_NAME}.cos.{REGION}.myqcloud.com/{quote(cos_path)}"
        return jsonify({"message": "生成成功", "file_url": file_url})
    
    except Exception as e:
        app.logger.error(f"生成文档错误: {str(e)}")
        return jsonify({'error': str(e)}), 500

####################################
#         Excel 生成接口
####################################
def validate_json_data(json_data):
    """验证JSON数据结构"""
    required_fields = ['title', 'data']
    if not all(field in json_data for field in required_fields):
        raise ValueError("Missing required fields in JSON data")
    
    if not json_data['data'] or not isinstance(json_data['data'], list):
        raise ValueError("Data must be a non-empty array")
    
    if not json_data['data'][0] or not isinstance(json_data['data'][0], list):
        raise ValueError("Headers must be a non-empty array")

def calculate_title_style(title):
    """根据标题长度计算字体大小和行高"""
    title_length = len(title)
    if title_length <= 15:
        return {'size': 16, 'height': 25}
    elif 15 < title_length <= 25:
        return {'size': 14, 'height': 35}
    elif 25 < title_length <= 40:
        return {'size': 12, 'height': 45}
    else:
        return {'size': 11, 'height': 55}

def apply_cell_style(cell, is_header=False, is_summary=False):
    """应用单元格样式"""
    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    cell.border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    if is_header:
        cell.font = Font(bold=True, size=11)
        cell.fill = PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid')
    if is_summary:
        cell.font = Font(bold=True, size=11)
        cell.fill = PatternFill(start_color='E6F3FF', end_color='E6F3FF', fill_type='solid')

def generate_excel(json_data):
    """生成Excel文件"""
    validate_json_data(json_data)
    
    wb = openpyxl.Workbook()
    sheet = wb.active
    title = json_data['title']
    title_style = calculate_title_style(title)
    sheet['A1'] = title
    title_cell = sheet['A1']
    title_cell.font = Font(size=title_style['size'], bold=True)
    title_cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    data = json_data['data']
    headers = data[0]
    data_rows = data[1:]
    end_col = get_column_letter(len(headers))
    sheet.merge_cells(f'A1:{end_col}1')
    sheet.row_dimensions[1].height = title_style['height']
    for col_num, header in enumerate(headers, 1):
        cell = sheet.cell(row=2, column=col_num, value=header)
        apply_cell_style(cell, is_header=True)
    for row_num, row_data in enumerate(data_rows, 3):
        is_summary = row_data[0] in ['合计', '总计', 'Total']
        for col_num, value in enumerate(row_data, 1):
            cell = sheet.cell(row=row_num, column=col_num, value=value)
            apply_cell_style(cell, is_summary=is_summary)
    for col_num in range(1, len(headers) + 1):
        max_len = 0
        col_letter = get_column_letter(col_num)
        for cell in sheet[col_letter]:
            try:
                if cell.value:
                    cell_len = len(str(cell.value))
                    max_len = max(max_len, cell_len)
            except:
                pass
        adjusted_width = (max_len + 2) * 1.2
        sheet.column_dimensions[col_letter].width = adjusted_width
    if 'metadata' in json_data:
        metadata = json_data['metadata']
        metadata_row = sheet.max_row + 2
        if 'summary' in metadata:
            sheet.cell(row=metadata_row, column=1, value=metadata['summary'])
            sheet.merge_cells(f'A{metadata_row}:{end_col}{metadata_row}')
        if 'timestamp' in metadata:
            timestamp_row = metadata_row + 1
            try:
                timestamp_date = datetime.strptime(metadata['timestamp'], "%Y-%m-%dT%H:%M:%SZ").strftime("%Y-%m-%d")
            except:
                timestamp_date = datetime.now().strftime("%Y-%m-%d")
            sheet.cell(row=timestamp_row, column=1, value=f"生成时间：{timestamp_date}")
            sheet.merge_cells(f'A{timestamp_row}:{end_col}{timestamp_row}')
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

@app.route('/generate_excel', methods=['POST'])
def generate_excel_route():
    try:
        json_data = request.get_json()
        if not json_data:
            return jsonify({"error": "No JSON data received"}), 400
        
        filename_input = json_data.get("filename", "未命名表格").strip()
        content = json_data.get("content", "")

        if not content:
             return jsonify({'error': '内容不能为空'}), 400

        # Parse Markdown to find table
        markdown = mistune.create_markdown(renderer=None, plugins=['table'])
        ast = markdown(content)
        
        table_node = None
        # Find the first table node
        # Mistune 3 AST is a list of nodes
        for node in ast:
            if node['type'] == 'table':
                table_node = node
                break
        
        if not table_node:
            return jsonify({'error': '未找到表格内容'}), 400

        # Extract data from table node
        headers = []
        rows = []
        
        children = table_node.get('children', [])
        thead = next((c for c in children if c['type'] == 'table_head'), None)
        tbody = next((c for c in children if c['type'] == 'table_body'), None)

        def extract_text(node):
            if 'raw' in node:
                return node['raw']
            if 'children' in node:
                return "".join([extract_text(child) for child in node['children']])
            return ""

        if thead:
            # Assuming single row in thead for now
            for row in thead.get('children', []):
                for cell in row.get('children', []):
                    headers.append(extract_text(cell))
        
        if tbody:
            for row in tbody.get('children', []):
                current_row = []
                for cell in row.get('children', []):
                    current_row.append(extract_text(cell))
                rows.append(current_row)

        if not headers and not rows:
             return jsonify({'error': '表格为空'}), 400

        # Construct data for generate_excel
        # Ensure headers is a list, even if empty (though valid markdown table usually has headers)
        excel_input_data = {
            'title': filename_input,
            'data': [headers] + rows
        }

        print(f"Received request at {datetime.now()}")
        # print("Parsed Excel Data:", json.dumps(excel_input_data, ensure_ascii=False, indent=2))

        excel_data = generate_excel(excel_input_data)
        
        clean_title = re.sub(r'[\\/:*?"<>|\s]', '', filename_input)[:30]
        if not clean_title:
            clean_title = "未命名表格"
        current_date = datetime.now().strftime("%Y%m%d")
        filename = f"{clean_title}_{current_date}.xlsx"
        temp_file = f"temp_{datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx"
        
        with open(temp_file, 'wb') as f:
            f.write(excel_data.getvalue())
            
        try:
            cos_path = f"excel_documents/{filename}"
            cos_client.upload_file(
                Bucket=BUCKET_NAME,
                LocalFilePath=temp_file,
                Key=cos_path
            )
            file_url = f"https://{BUCKET_NAME}.cos.{REGION}.myqcloud.com/{quote(cos_path)}"
            return jsonify({
                "message": "Excel文件生成成功",
                "file_url": file_url,
                "filename": filename
            })
        except Exception as upload_error:
            app.logger.error(f"上传失败: {str(upload_error)}")
            return jsonify({'error': '文件上传失败'}), 500
        finally:
            if os.path.exists(temp_file):
                os.remove(temp_file)
    except ValueError as ve:
        return jsonify({"error": str(ve)}), 400
    except Exception as e:
        print(f"Error: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({"error": "Internal server error"}), 500

@app.route('/health', methods=['GET'])
def health_check():
    """健康检查接口"""
    return jsonify({
        "status": "healthy",
        "timestamp": datetime.now().isoformat(),
        "service": "Excel & Docx Generator"
    })

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5101, debug=True)
