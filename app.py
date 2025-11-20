我来帮你开发这个岗前综合技能培训报告生成工具。我会创建一个完整的Streamlit应用，能够生成符合模板要求的Word文档。

<function_calls>
<invoke name="create_file">
<parameter name="path">app.py</parameter>
<parameter name="content">import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from io import BytesIO
import datetime

def set_cell_border(cell, **kwargs):
    """设置单元格边框"""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
  
    tcBorders = OxmlElement('w:tcBorders')
    for edge in ('top', 'left', 'bottom', 'right'):
        edge_data = kwargs.get(edge)
        if edge_data:
            tag = 'w:{}'.format(edge)
            element = OxmlElement(tag)
            element.set(qn('w:val'), 'single')
            element.set(qn('w:sz'), '4')
            element.set(qn('w:space'), '0')
            element.set(qn('w:color'), '000000')
            tcBorders.append(element)
  
    tcPr.append(tcBorders)

def set_cell_shading(cell, fill):
    """设置单元格背景色"""
    shading_elm = OxmlElement('w:shd')
    shading_elm.set(qn('w:fill'), fill)
    cell._element.get_or_add_tcPr().append(shading_elm)

def add_paragraph_with_format(doc, text, font_name='宋体', font_size=12, bold=False, 
                              alignment=WD_ALIGN_PARAGRAPH.LEFT, space_before=0, space_after=0):
    """添加格式化段落"""
    paragraph = doc.add_paragraph()
    paragraph.alignment = alignment
    paragraph.paragraph_format.space_before = Pt(space_before)
    paragraph.paragraph_format.space_after = Pt(space_after)
    paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
  
    run = paragraph.add_run(text)
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
  
    return paragraph

def create_cover_page(doc, student_info):
    """创建封面页"""
    # 标题上方的星号和学号
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f"★            学号：{student_info['学号']}")
    run.font.name = '宋体'
    run.font.size = Pt(14)
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
  
    # 主标题
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(36)
    run = p.add_run('岗前综合技能培训报告书')
    run.font.name = '宋体'
    run.font.size = Pt(22)
    run.font.bold = True
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
  
    # 空行
    for _ in range(3):
        doc.add_paragraph()
  
    # 信息部分
    info_lines = [
        f"          {student_info['学院']}      ",
        f"专业：{student_info.get('专业', '                        ')}",
        f"班    级：{student_info['班级']}",
        f"学生姓名：{student_info['姓名']}"
    ]
  
    for line in info_lines:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(12)
        run = p.add_run(line)
        run.font.name = '宋体'
        run.font.size = Pt(16)
        run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
  
    # 底部学院名称
    for _ in range(3):
        doc.add_paragraph()
  
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run('海南软件职业技术学院')
    run.font.name = '宋体'
    run.font.size = Pt(18)
    run.font.bold = True
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
  
    doc.add_page_break()

def create_task_book(doc, student_info):
    """创建任务书页面"""
    # 标题
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run('海南软件职业技术学院岗前综合技能培训任务书')
    run.font.name = '宋体'
    run.font.size = Pt(16)
    run.font.bold = True
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
  
    doc.add_paragraph()
  
    # 创建表格
    table = doc.add_table(rows=6, cols=4)
    table.style = 'Table Grid'
  
    # 设置列宽
    table.columns[0].width = Cm(3)
    table.columns[1].width = Cm(5)
    table.columns[2].width = Cm(4)
    table.columns[3].width = Cm(5)
  
    # 第一行
    cells = table.rows[0].cells
    cells[0].text = '学院'
    cells[1].text = student_info['学院']
    cells[2].text = '学号'
    cells[3].text = student_info['学号']
  
    # 第二行
    cells = table.rows[1].cells
    cells[0].text = '姓名'
    cells[1].text = student_info['姓名']
    cells[2].text = '岗前综合技能\n培训指导教师'
    cells[3].text = student_info['指导教师']
  
    # 第三行
    cells = table.rows[2].cells
    cells[0].text = '项目名称'
    cells[1].merge(cells[3]).text = student_info['项目名称']
  
    # 第四行
    cells = table.rows[3].cells
    cells[0].text = '起止时间'
    cells[1].merge(cells[3]).text = '20   年   月   日至   20   年   月   日'
  
    # 第五行
    cells = table.rows[4].cells
    cells[0].text = '岗前综合技能培训内容及培养目标'
    cells[0].merge(cells[3])
  
    # 第六行
    cells = table.rows[5].cells
    cells[0].text = '岗前综合技能培训形式'
    cells[0].merge(cells[3])
  
    # 设置表格字体
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.name = '宋体'
                    run.font.size = Pt(12)
                    run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
  
    doc.add_paragraph()
    add_paragraph_with_format(doc, '岗前综合技能培训指导教师签名：', font_size=12)
    doc.add_paragraph()
    add_paragraph_with_format(doc, '岗前综合技能培训领导小组审查意见：', font_size=12)
    doc.add_paragraph()
    doc.add_paragraph()
    add_paragraph_with_format(doc, '备注：此表回收后交院部按班级为单位装订存档。', font_size=10.5)
  
    doc.add_page_break()

def create_guidance_record(doc, student_info):
    """创建指导记录表"""
    # 标题
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run('岗前综合技能培训指导记录表')
    run.font.name = '宋体'
    run.font.size = Pt(16)
    run.font.bold = True
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
  
    doc.add_paragraph()
  
    # 创建表格
    table = doc.add_table(rows=12, cols=4)
    table.style = 'Table Grid'
  
    # 第一行
    cells = table.rows[0].cells
    cells[0].text = '学号'
    cells[1].text = student_info['学号']
    cells[2].text = '岗前综合技能\n培训指导教师'
    cells[3].text = student_info['指导教师']
  
    # 第二行
    cells = table.rows[1].cells
    cells[0].text = '专    业'
    cells[1].text = student_info.get('专业', '')
    cells[2].text = '指导教师专业'
    cells[3].text = ''
  
    # 指导记录行
    for i in range(2, 10):
        cells = table.rows[i].cells
        cells[0].text = f'    月    日'
        cells[1].merge(cells[3])
  
    # 指导教师签名行
    cells = table.rows[10].cells
    cells[0].text = '指导教师签名（每次需签名）：'
    cells[0].merge(cells[3])
  
    # 备注行
    cells = table.rows[11].cells
    cells[0].text = '备注：此表由学生根据老师每次指导的内容填写，指导教师签字后，学生保存，待上交文档时交学院，学院按班级为单位装订存档。'
    cells[0].merge(cells[3])
  
    # 设置表格字体
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.name = '宋体'
                    run.font.size = Pt(12)
                    run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
  
    doc.add_page_break()

def create_evaluation_form(doc, student_info):
    """创建成绩评定表"""
    # 标题
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run('岗前综合技能培训成绩评定表')
    run.font.name = '宋体'
    run.font.size = Pt(16)
    run.font.bold = True
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
  
    doc.add_paragraph()
  
    # 创建表格
    table = doc.add_table(rows=10, cols=2)
    table.style = 'Table Grid'
  
    # 设置列宽
    table.columns[0].width = Cm(5)
    table.columns[1].width = Cm(12)
  
    # 填充内容
    table.rows[0].cells[0].text = '岗前综合技能培训成绩评定'
    table.rows[0].cells[1].text = f'岗前综合技能培训项目名称：{student_info["项目名称"]}'
  
    table.rows[1].cells[0].text = ''
    table.rows[1].cells[1].text = '岗前综合技能培训成果：□软件作品  □影视动漫作品  □电子工艺产品  □综合实训报告'
  
    table.rows[2].cells[0].text = ''
    table.rows[2].cells[1].text = '审查时间：        年    月    日'
  
    table.rows[3].cells[0].text = '岗前综合技能培训初评评语'
    table.rows[3].cells[1].text = '\n\n\n指导教师签名：          '
  
    table.rows[4].cells[0].text = '初评成绩（满分100分）'
    table.rows[4].cells[1].text = '1．岗前综合技能培训过程及成果评价（满分 80分）\nA. 有创新性结果，全面完成了训练任务所规定的各项要求。(71-80分)\nB. 有创新性结果，基本完成了训练任务所规定的各项要求。\nC. 有一定的创新性结果，基本完成了训练任务所规定的各项要求。\nD. 基本没有创新性结果，没有完成训练任务所规定的各项要求。'
  
    table.rows[5].cells[0].text = '评分'
    table.rows[5].cells[1].text = ''
  
    table.rows[6].cells[0].text = ''
    table.rows[6].cells[1].text = '2．答辩材料准备与答辩表现（满分20分）\n项目成果。回答问题正确，概念清楚，知识掌握'
  
    table.rows[7].cells[0].text = '评分'
    table.rows[7].cells[1].text = ''
  
    table.rows[8].cells[0].text = '答辩评语'
    table.rows[8].cells[1].text = '\n\n答辩成绩：        '
  
    table.rows[9].cells[0].text = '岗前综合技能培训最终成绩评定'
    table.rows[9].cells[1].text = '成绩评定（在"□"中划" √")\n优□    良□    中□    及格□    不及格□\n\n学院（签章）：              年    月    日'
  
    # 设置表格字体
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.name = '宋体'
                    run.font.size = Pt(10.5)
                    run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
  
    doc.add_paragraph()
    add_paragraph_with_format(doc, '注：1.如果初评成绩<90分，则"岗前综合技能培训最终成绩评定"栏由指导教师直接依据初评成绩填写，并确定', font_size=9)
    add_paragraph_with_format(doc, '2.初评成绩≥90分（优秀）才进行答辩，其他的不需答辩。', font_size=9)
    add_paragraph_with_format(doc, '3.此表学院需复印一份以班级为单位装订存档。', font_size=9)
  
    doc.add_page_break()

def create_topic_summary(doc, student_info):
    """创建选题汇总表"""
    # 标题
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run('海南软件职业技术学院岗前综合技能培训选题汇总表')
    run.font.name = '宋体'
    run.font.size = Pt(16)
    run.font.bold = True
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
  
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run(f'学院：{student_info["学院"]}')
    run.font.name = '宋体'
    run.font.size = Pt(12)
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
  
    # 创建表格
    table = doc.add_table(rows=2, cols=6)
    table.style = 'Table Grid'
  
    # 表头
    headers = ['序号', '学号', '姓名', '项目名称', '指导教师', '所在学院']
    for i, header in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text =
