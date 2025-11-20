<function_calls>
<invoke name="delete_file">
<parameter name="path">app.py</parameter>
</invoke>
<invoke name="create_file">
<parameter name="path">training_report_generator.py</parameter>
<parameter name="content">import streamlit as st
from docx import Document
from docx.shared import Pt, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from io import BytesIO


def set_font(run, font_name='宋体', font_size=12, bold=False):
    """设置字体格式"""
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)


def generate_report(student_info):
    """生成报告文档"""
    doc = Document()
  
    # 第1页：封面
    create_cover_page(doc, student_info)
  
    # 第2页：任务书
    create_task_book(doc, student_info)
  
    # 第3页：指导记录表
    create_guidance_record(doc, student_info)
  
    # 第4页：成绩评定表
    create_evaluation_form(doc, student_info)
  
    # 第5页：选题汇总表
    create_topic_summary(doc, student_info)
  
    # 第6-7页：报告撰写说明
    create_report_instructions(doc)
  
    return doc


def create_cover_page(doc, info):
    """创建封面页"""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f"★            学号：{info['学号']}")
    set_font(run, font_size=14)
  
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(36)
    run = p.add_run('岗前综合技能培训报告书')
    set_font(run, font_size=22, bold=True)
  
    for _ in range(3):
        doc.add_paragraph()
  
    info_texts = [
        f"          {info['学院']}      ",
        f"专业：{info.get('专业', '                        ')}",
        f"班    级：{info['班级']}",
        f"学生姓名：{info['姓名']}"
    ]
  
    for text in info_texts:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(12)
        run = p.add_run(text)
        set_font(run, font_size=16)
  
    for _ in range(3):
        doc.add_paragraph()
  
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run('海南软件职业技术学院')
    set_font(run, font_size=18, bold=True)
  
    doc.add_page_break()


def create_task_book(doc, info):
    """创建任务书"""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run('海南软件职业技术学院岗前综合技能培训任务书')
    set_font(run, font_size=16, bold=True)
  
    doc.add_paragraph()
  
    table = doc.add_table(rows=6, cols=4)
    table.style = 'Table Grid'
  
    table.rows[0].cells[0].text = '学院'
    table.rows[0].cells[1].text = info['学院']
    table.rows[0].cells[2].text = '学号'
    table.rows[0].cells[3].text = info['学号']
  
    table.rows[1].cells[0].text = '姓名'
    table.rows[1].cells[1].text = info['姓名']
    table.rows[1].cells[2].text = '岗前综合技能培训指导教师'
    table.rows[1].cells[3].text = info['指导教师']
  
    table.rows[2].cells[0].text = '项目名称'
    table.rows[2].cells[1].merge(table.rows[2].cells[3]).text = info['项目名称']
  
    table.rows[3].cells[0].text = '起止时间'
    table.rows[3].cells[1].merge(table.rows[3].cells[3]).text = '20   年   月   日至   20   年   月   日'
  
    table.rows[4].cells[0].merge(table.rows[4].cells[3]).text = '岗前综合技能培训内容及培养目标'
  
    table.rows[5].cells[0].merge(table.rows[5].cells[3]).text = '岗前综合技能培训形式'
  
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    set_font(run, font_size=12)
  
    doc.add_paragraph()
    p = doc.add_paragraph()
    run = p.add_run('岗前综合技能培训指导教师签名：')
    set_font(run, font_size=12)
  
    doc.add_paragraph()
    p = doc.add_paragraph()
    run = p.add_run('岗前综合技能培训领导小组审查意见：')
    set_font(run, font_size=12)
  
    doc.add_paragraph()
    doc.add_paragraph()
  
    p = doc.add_paragraph()
    run = p.add_run('备注：此表回收后交院部按班级为单位装订存档。')
    set_font(run, font_size=10.5)
  
    doc.add_page_break()


def create_guidance_record(doc, info):
    """创建指导记录表"""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run('岗前综合技能培训指导记录表')
    set_font(run, font_size=16, bold=True)
  
    doc.add_paragraph()
  
    table = doc.add_table(rows=12, cols=4)
    table.style = 'Table Grid'
  
    table.rows[0].cells[0].text = '学号'
    table.rows[0].cells[1].text = info['学号']
    table.rows[0].cells[2].text = '岗前综合技能培训指导教师'
    table.rows[0].cells[3].text = info['指导教师']
  
    table.rows[1].cells[0].text = '专    业'
    table.rows[1].cells[1].text = info.get('专业', '')
    table.rows[1].cells[2].text = '指导教师专业'
    table.rows[1].cells[3].text = ''
  
    for i in range(2, 10):
        table.rows[i].cells[0].text = '    月    日'
        table.rows[i].cells[1].merge(table.rows[i].cells[3])
  
    table.rows[10].cells[0].merge(table.rows[10].cells[3]).text = '指导教师签名（每次需签名）：'
  
    table.rows[11].cells[0].merge(table.rows[11].cells[3]).text = '备注：此表由学生根据老师每次指导的内容填写，指导教师签字后，学生保存，待上交文档时交学院，学院按班级为单位装订存档。'
  
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    set_font(run, font_size=12)
  
    doc.add_page_break()


def create_evaluation_form(doc, info):
    """创建成绩评定表"""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run('岗前综合技能培训成绩评定表')
    set_font(run, font_size=16, bold=True)
  
    doc.add_paragraph()
  
    table = doc.add_table(rows=10, cols=2)
    table.style = 'Table Grid'
  
    table.rows[0].cells[0].text = '岗前综合技能培训成绩评定'
    table.rows[0].cells[1].text = f'岗前综合技能培训项目名称：{info["项目名称"]}'
  
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
  
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    set_font(run, font_size=10.5)
  
    doc.add_paragraph()
    p = doc.add_paragraph()
    run = p.add_run('注：1.如果初评成绩<90分，则"岗前综合技能培训最终成绩评定"栏由指导教师直接依据初评成绩填写，并确定')
    set_font(run, font_size=9)
  
    p = doc.add_paragraph()
    run = p.add_run('2.初评成绩≥90分（优秀）才进行答辩，其他的不需答辩。')
    set_font(run, font_size=9)
  
    p = doc.add_paragraph()
    run = p.add_run('3.此表学院需复印一份以班级为单位装订存档。')
    set_font(run, font_size=9)
  
    doc.add_page_break()


def create_topic_summary(doc, info):
    """创建选题汇总表"""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run('海南软件职业技术学院岗前综合技能培训选题汇总表')
    set_font(run, font_size=16, bold=True)
  
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run(f'学院：{info["学院"]}')
    set_font(run, font_size=12)
  
    table = doc.add_table(rows=2, cols=6)
    table.style = 'Table Grid'
  
    headers = ['序号', '学号', '姓名', '项目名称', '指导教师', '所在学院']
    for i, header in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = header
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                set_font(run, font_size=12, bold=True)
  
    cells = table.rows[1].cells
    cells[0].text = '1'
    cells[1].text = info['学号']
    cells[2].text = info['姓名']
    cells[3].text = info['项目名称']
    cells[4].text = info['指导教师']
    cells[5].text = info['学院']
  
    for cell in table.rows[1].cells:
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                set_font(run, font_size=12)
  
    doc.add_paragraph()
    p = doc.add_paragraph()
    run = p.add_run('备注：此表回收后交学院按班级为单位装订存档。')
    set_font(run, font_size=10.5)
  
    doc.add_page_break()


def create_report_instructions(doc):
    """创建报告撰写说明"""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run('岗前综合技能培训报告')
    set_font(run, font_size=16, bold=True)
  
    doc.add_paragraph()
  
    p = doc.add_paragraph()
    run = p.add_run('撰写说明：')
    set_font(run, font_size=12, bold=True)
    run = p.add_run('报告分为四大部分，段落要求1.5倍行距，整个报告内容不少于5页，具体内容及格式要求如下：')
    set_font(run, font_size=12)
  
    doc.add_paragraph()
  
    p = doc.add_paragraph()
    run = p.add_run('一、岗前培训目的（宋体，加粗，四号字，左对齐）')
    set_font(run, font_size=14, bold=True)
  
    p = doc.add_paragraph()
    run = p.add_run('（介绍岗前培训目的和意义，岗前培训单位的发展情况及学习要求等）')
    set_font(run, font_size=12)
  
    p = doc.add_paragraph()
    run = p.add_run('1、小标题（宋体，加粗，小四号字）')
    set_font(run, font_size=12, bold=True)
  
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    run = p.add_run(' XXXXXXX（正文：宋体，小四号字）')
    set_font(run, font_size=12)
  
    p = doc.add_paragraph()
    run = p.add_run('2、小标题（宋体，加粗，小四号字）')
    set_font(run, font_size=12, bold=True)
  
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    run = p.add_run('XXXXXXX（正文：宋体，小四号字）')
    set_font(run, font_size=12)
  
    p = doc.add_paragraph()
    run = p.add_run('………')
    set_font(run, font_size=12)
  
    p = doc.add_paragraph()
    run = p.add_run('二、培训内容（宋体，加粗，四号字，左对齐）')
    set_font(run, font_size=14, bold=True)
  
    p = doc.add_paragraph()
    run = p.add_run('1、小标题（宋体，加粗，小四号字）')
    set_font(run, font_size=12, bold=True)
  
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    run = p.add_run(' XXXXXXX（正文：宋体，小四号字）')
    set_font(run, font_size=12)
  
    p = doc.add_paragraph()
    run = p.add_run('2、小标题（宋体，加粗，小四号字）')
    set_font(run, font_size=12, bold=True)
  
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    run = p.add_run('XXXXXXX（正文：宋体，小四号字）')
    set_font(run, font_size=12)
