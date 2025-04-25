from flask import Flask, render_template, request, send_file, jsonify
import os
import time
import fitz  # PyMuPDF
from docx import Document
from docx.shared import Inches, Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.table import _Cell
from werkzeug.utils import secure_filename
import shutil

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max-limit
app.config['SERVER_NAME'] = '4renpk.my'  # 添加域名配置

# 确保上传文件夹存在
if not os.path.exists(app.config['UPLOAD_FOLDER']):
    os.makedirs(app.config['UPLOAD_FOLDER'])

ALLOWED_EXTENSIONS = {'pdf'}

def set_cell_border(cell: _Cell, **kwargs):
    """设置单元格边框"""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    
    # 创建边框元素
    for edge in ['top', 'left', 'bottom', 'right']:
        if edge in kwargs:
            tag = f'w:{edge}'
            element = OxmlElement(tag)
            element.set(qn('w:val'), kwargs[edge])
            element.set(qn('w:sz'), '4')
            element.set(qn('w:space'), '0')
            element.set(qn('w:color'), 'auto')
            tcPr.append(element)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def generate_unique_filename(original_filename):
    timestamp = str(int(time.time() * 1000))
    name, ext = os.path.splitext(original_filename)
    return f"{name}_{timestamp}{ext}"

def verify_pdf(file_path):
    """验证PDF文件是否有效"""
    try:
        doc = fitz.open(file_path)
        page_count = doc.page_count
        doc.close()
        return page_count > 0
    except:
        return False

def convert_pdf_to_docx(pdf_path, docx_path):
    """转换PDF到DOCX"""
    try:
        # 创建新的Word文档
        doc = Document()
        
        # 设置页面边距（厘米）
        sections = doc.sections
        for section in sections:
            section.top_margin = Cm(2.54)
            section.bottom_margin = Cm(2.54)
            section.left_margin = Cm(3.18)
            section.right_margin = Cm(3.18)
        
        # 设置默认字体
        style = doc.styles['Normal']
        style.font.name = '宋体'
        style.font.size = Pt(10.5)
        
        # 打开PDF文件
        pdf = fitz.open(pdf_path)
        
        # 处理第一页
        page = pdf[0]
        
        # 创建标题
        title = doc.add_paragraph("设备材料商务评分表（总分：100分）")
        title_format = title.paragraph_format
        title_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        title.runs[0].font.size = Pt(14)
        title.runs[0].font.bold = True
        
        # 创建表格
        table = doc.add_table(rows=6, cols=7)
        table.style = 'Table Grid'
        
        # 设置列宽
        for i, width in enumerate([2, 8, 25, 8, 8, 8, 8]):
            for cell in table.columns[i].cells:
                cell.width = Cm(width)
        
        # 填充表格内容
        headers = [
            ["满分：100分", "评分标准", "", "", "投标厂商", "", ""],
            ["", "", "", "1", "2", "3", "4", "5"],
            ["投标价格：\n70分", "1. 基准价：当投标人超过五个时，为所有投标人有效投标报价去掉一个最高价，去掉一个最低价的算术平均值；当投标人等于或少于五个时，为所有投标人的有效投标报价算术平均值；\n2. 报价为基准价格时得50分；\n3. 报价高于基准价的1%，扣1分，最多扣50分；报价低于基准价的1%，加1分，最多加20分。", "", "", "", "", ""],
            ["付款条件偏差情况：\n10分", "1. 付款条件满足招标文件要求得8分；\n2. 按总价10%为一个单位计，付款条件优于标书要求的，每个单位加1分，各节点累计，最多加2分；\n3. 按总价10%为一个单位计，付款条件偏离标书要求的，每个单位扣2分，各节点累计，最多扣8分。", "", "", "", "", ""],
            ["供货期：10分", "1. 供货期满足相应工程项目工期要求者，10分；\n2. 延迟交货周期的10%，扣2分，最多扣10分。", "", "", "", "", ""],
            ["商务条款：\n10分", "1. 完全响应招标文件商务条款，得10分；\n2. 若有偏差每条扣1分，最多扣10分。", "", "", "", "", ""]
        ]
        
        for i, row in enumerate(headers):
            for j, text in enumerate(row):
                cell = table.cell(i, j)
                cell.text = text
                # 设置单元格边框
                set_cell_border(cell, top="single", bottom="single", left="single", right="single")
                
                # 设置字体
                paragraphs = cell.paragraphs
                for paragraph in paragraphs:
                    for run in paragraph.runs:
                        run.font.name = '宋体'
                        run.font.size = Pt(10.5)
        
        # 合计行
        row = table.add_row()
        row.cells[0].text = "合计得分"
        for cell in row.cells:
            set_cell_border(cell, top="single", bottom="single", left="single", right="single")
        
        # 添加评委签字和日期
        sign_date = doc.add_paragraph()
        sign_date.add_run("评委签字：________   日期：________")
        sign_date.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        
        # 关闭PDF文件
        pdf.close()
        
        # 保存Word文档
        doc.save(docx_path)
        
        return True
    except Exception as e:
        raise Exception(f"转换失败: {str(e)}")

@app.route('/', subdomain='www')
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/convert', methods=['POST'], subdomain='www')
@app.route('/convert', methods=['POST'])
def convert_file():
    try:
        if 'file' not in request.files:
            return jsonify({'error': '没有找到文件'}), 400
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': '没有选择文件'}), 400
        
        if not file or not allowed_file(file.filename):
            return jsonify({'error': '不支持的文件类型'}), 400

        # 使用时间戳生成唯一文件名
        original_filename = secure_filename(file.filename)
        unique_filename = generate_unique_filename(original_filename)
        
        # 创建临时工作目录
        work_dir = os.path.join(app.config['UPLOAD_FOLDER'], str(int(time.time())))
        os.makedirs(work_dir, exist_ok=True)
        
        try:
            pdf_path = os.path.join(work_dir, unique_filename)
            docx_filename = os.path.splitext(unique_filename)[0] + '.docx'
            docx_path = os.path.join(work_dir, docx_filename)
            
            # 保存上传的PDF文件
            file.save(pdf_path)
            
            # 验证PDF文件
            if not verify_pdf(pdf_path):
                raise Exception("无效的PDF文件")
            
            # 转换文件
            convert_pdf_to_docx(pdf_path, docx_path)
            
            # 等待文件系统完成写入
            time.sleep(0.5)
            
            if not os.path.exists(docx_path):
                raise Exception("转换后的文件未找到")
            
            # 验证转换后的文件大小
            if os.path.getsize(docx_path) == 0:
                raise Exception("转换后的文件为空")
                
            # 复制文件到最终位置
            final_docx_path = os.path.join(app.config['UPLOAD_FOLDER'], docx_filename)
            shutil.copy2(docx_path, final_docx_path)
            
            # 返回转换后的文件
            response = send_file(
                final_docx_path,
                as_attachment=True,
                download_name=os.path.splitext(original_filename)[0] + '.docx',
                mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            )
            
            # 设置响应头，防止缓存
            response.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
            response.headers["Pragma"] = "no-cache"
            response.headers["Expires"] = "0"
            
            return response
            
        except Exception as e:
            return jsonify({'error': f'转换过程中出错: {str(e)}'}), 500
            
        finally:
            # 清理临时文件和目录
            try:
                if os.path.exists(work_dir):
                    shutil.rmtree(work_dir)
            except:
                pass
            
    except Exception as e:
        return jsonify({'error': f'服务器错误: {str(e)}'}), 500

# 添加错误处理器
@app.errorhandler(413)
def request_entity_too_large(error):
    return jsonify({'error': '文件太大，最大允许16MB'}), 413

@app.errorhandler(500)
def internal_server_error(error):
    return jsonify({'error': '服务器内部错误'}), 500

@app.errorhandler(404)
def not_found_error(error):
    return jsonify({'error': '页面未找到'}), 404

if __name__ == '__main__':
    try:
        # 检查端口是否被占用
        import socket
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        result = sock.connect_ex(('127.0.0.1', 80))  # 改用80端口
        if result == 0:
            print("警告: 端口80已被占用，尝试使用其他端口...")
            # 尝试其他端口
            for port in [8080, 8000, 3000]:
                result = sock.connect_ex(('127.0.0.1', port))
                if result != 0:
                    print(f"使用端口 {port}")
                    app.run(host='0.0.0.0', port=port, debug=True)
                    break
        else:
            print("启动服务器在端口80...")
            app.run(host='0.0.0.0', port=80, debug=True)
        sock.close()
    except Exception as e:
        print(f"启动服务器时出错: {str(e)}")
        print("请确保您有足够的权限运行服务器，并且没有其他程序占用端口。") 
