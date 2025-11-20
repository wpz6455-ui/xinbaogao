import streamlit as st
from docx import Document
from docx.shared import Pt, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.shared import RGBColor
import io
import datetime

# --- 核心工具函数 ---

def set_font(run, font_name='宋体', font_size=12, bold=False):
    """设置中文字体和字号的核心函数"""
    run.font.name = 'Times New Roman' # 西文
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name) # 中文
    run.font.size = Pt(font_size)
    run.font.bold = bold

def add_paragraph_field(doc, label, value, font_size=12, bold_label=False):
    """添加形如 '姓名：XXX' 的段落"""
    p = doc.add_paragraph()
    run = p.add_run(f"{label}：")
    set_font(run, font_size=font_size, bold=bold_label)
    
    run = p.add_run(f" {value} ")
    set_font(run, font_size=font_size, bold=False)
    run.font.underline = True

def set_cell_text(cell, text, align=WD_ALIGN_PARAGRAPH.CENTER, font_size=12, bold=False):
    """设置表格单元格文字"""
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = align
    run = p.add_run(text)
    set_font(run, font_size=font_size, bold=bold)

# --- 页面生成逻辑 ---

def create_cover(doc, data, logo_file):
    """生成封面 (Page 1)"""
    # 调整页边距
    section = doc.sections[0]
    section.top_margin = Cm(2.54)
    section.bottom_margin = Cm(2.54)
    section.left_margin = Cm(3.17)
    section.right_margin = Cm(3.17)

    # 学号行
    p_top = doc.add_paragraph()
    p_top.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r = p_top.add_run("★") # 原始模板标记
    set_font(r, font_size=10)
    r = p_top.add_run(f"\t\t\t\t\t\t学号：{data['student_id']}")
    set_font(r, font_size=12)

    # 间隔
    doc.add_paragraph()

    # 校徽与标题
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if logo_file:
        try:
            p_title.add_run().add_picture(logo_file, width=Cm(10))
            doc.add_paragraph().alignment = WD_ALIGN_PARAGRAPH.CENTER
        except:
            pass
    
    run_title = p_title.add_run("岗前综合技能培训报告书")
    set_font(run_title, font_size=36, bold=True) # 一号/小初

    doc.add_paragraph()
    doc.add_paragraph()

    # 项目名称
    p_proj = doc.add_paragraph()
    p_proj.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_proj = p_proj.add_run(f"项目：{data['project_name']}")
    set_font(r_proj, font_size=16, bold=True)
    r_proj.font.underline = True

    doc.add_paragraph()
    doc.add_paragraph()
    doc.add_paragraph()

    # 封面信息表格化排版 (为了对齐更好，使用无边框表格模拟)
    table = doc.add_table(rows=5, cols=2)
    table.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    fields = [
        ("学    院：", data['college']),
        ("专    业：", data['major']),
        ("班    级：", data['class_name']),
        ("学生姓名：", data['name']),
        ("指导教师：", data['teacher'])
    ]

    for i, (label, value) in enumerate(fields):
        cell_label = table.cell(i, 0)
        cell_val = table.cell(i, 1)
        
        p = cell_label.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        r = p.add_run(label)
        set_font(r, font_size=16, bold=True)
        
        p2 = cell_val.paragraphs[0]
        p2.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r2 = p2.add_run(value)
        set_font(r2, font_size=16)
        r2.font.underline = True

    doc.add_paragraph()
    doc.add_paragraph()

    # 起止时间
    p_date = doc.add_paragraph()
    p_date.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_date = p_date.add_run(f"起止时间： {data['start_date']} 至 {data['end_date']}")
    set_font(r_date, font_size=14)

    doc.add_paragraph()
    
    p_school = doc.add_paragraph()
    p_school.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_school = p_school.add_run("海南软件职业技术学院")
    set_font(r_school, font_size=22, bold=True)

    doc.add_page_break()

def create_task_sheet(doc, data):
    """生成任务书 (Page 2)"""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("海南软件职业技术学院  岗前综合技能培训任务书")
    set_font(r, font_size=16, bold=True)

    table = doc.add_table(rows=8, cols=6)
    table.style = 'Table Grid'
    table.autofit = False
    
    # 设置列宽 (近似值)
    widths = [Cm(2.5), Cm(3), Cm(2), Cm(2.5), Cm(2), Cm(3)]
    for row in table.rows:
        for idx, width in enumerate(widths):
            row.cells[idx].width = width

    # 第一行：学院、专业
    set_cell_text(table.cell(0, 0), "学院")
    set_cell_text(table.cell(0, 1), data['college'])
    set_cell_text(table.cell(0, 2), "专业")
    table.cell(0, 3).merge(table.cell(0, 5)) # 合并后3列
    set_cell_text(table.cell(0, 3), data['major'])

    # 第二行：班级、学号、姓名
    set_cell_text(table.cell(1, 0), "班级")
    set_cell_text(table.cell(1, 1), data['class_name'])
    set_cell_text(table.cell(1, 2), "学号")
    set_cell_text(table.cell(1, 3), data['student_id'])
    set_cell_text(table.cell(1, 4), "姓名")
    set_cell_text(table.cell(1, 5), data['name'])

    # 第三行：指导教师、题目
    set_cell_text(table.cell(2, 0), "岗前综合技能\n培训指导教师")
    set_cell_text(table.cell(2, 1), data['teacher'])
    set_cell_text(table.cell(2, 2), "题目")
    table.cell(2, 3).merge(table.cell(2, 5))
    set_cell_text(table.cell(2, 3), data['project_name'])

    # 第四行：时间
    set_cell_text(table.cell(3, 0), "起止时间")
    table.cell(3, 1).merge(table.cell(3, 5))
    set_cell_text(table.cell(3, 1), f"{data['start_date']} 至 {data['end_date']}")

    # 内容行 (合并首列，内容列合并)
    labels = ["项目的意义\n及培养目标", "岗前综合技能\n培训成果形式", "技能训练\n基本要求", "岗前综合技能\n培训主要任务"]
    contents = [data['meaning'], data['output_form'], data['requirements'], data['main_tasks']]

    for i, (label, content) in enumerate(zip(labels, contents)):
        row_idx = 4 + i
        # 第一列
        set_cell_text(table.cell(row_idx, 0), label)
        # 合并后面所有列
        table.cell(row_idx, 1).merge(table.cell(row_idx, 5))
        cell = table.cell(row_idx, 1)
        set_cell_text(cell, content, align=WD_ALIGN_PARAGRAPH.LEFT)
        # 增加行高
        table.rows[row_idx].height = Cm(2.5)

    doc.add_paragraph()
    
    # 签名区
    p_sign = doc.add_paragraph()
    r = p_sign.add_run("岗前综合技能培训指导教师签名：")
    set_font(r, font_size=12)
    
    doc.add_paragraph()
    p_group = doc.add_paragraph()
    r = p_group.add_run("岗前综合技能培训领导小组审查意见：")
    set_font(r, font_size=12)
    doc.add_paragraph()
    p_group_sign = doc.add_paragraph()
    p_group_sign.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = p_group_sign.add_run("组长签名：____________   年   月   日    ")
    set_font(r, font_size=12)

    doc.add_page_break()

def create_guidance_record(doc, data):
    """生成指导记录表 (Page 3)"""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("岗前综合技能培训指导记录表")
    set_font(r, font_size=16, bold=True)

    table = doc.add_table(rows=12, cols=4)
    table.style = 'Table Grid'
    
    # 表头信息
    headers = [
        ("学  号", data['student_id'], "指导教师", data['teacher']),
        ("姓  名", data['name'], "专    业", data['major']),
        ("班  级", data['class_name'], "项目名称", data['project_name'])
    ]

    for i, row_data in enumerate(headers):
        set_cell_text(table.cell(i, 0), row_data[0])
        set_cell_text(table.cell(i, 1), row_data[1])
        set_cell_text(table.cell(i, 2), row_data[2])
        set_cell_text(table.cell(i, 3), row_data[3])

    # 记录列表头
    set_cell_text(table.cell(3, 0), "指导时间")
    table.cell(3, 1).merge(table.cell(3, 3))
    set_cell_text(table.cell(3, 1), "指导内容")

    # 生成8行记录 (模拟)
    current_date = datetime.datetime.strptime(data['start_date'].split('至')[0].strip(), "%Y年%m月%d日") if '年' in data['start_date'] else datetime.datetime.now()
    
    for i in range(8):
        row_idx = 4 + i
        date_str = ""
        # 简单的日期递增模拟，实际应由用户填
        sim_date = current_date + datetime.timedelta(days=i*7) 
        date_str = f"{sim_date.month}月{sim_date.day}日"
        
        set_cell_text(table.cell(row_idx, 0), date_str)
        table.cell(row_idx, 1).merge(table.cell(row_idx, 3))
        set_cell_text(table.cell(row_idx, 1), f"指导内容记录 {i+1} ...", align=WD_ALIGN_PARAGRAPH.LEFT)
        table.rows[row_idx].height = Cm(1.2)

    doc.add_paragraph()
    p_sign = doc.add_paragraph()
    p_sign.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = p_sign.add_run("指导教师签名：___________   年   月   日")
    set_font(r, font_size=12)

    doc.add_page_break()

def create_assessment(doc, data):
    """生成成绩评定表 (Page 4)"""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("海南软件职业技术学院\n岗前综合技能培训成绩评定表")
    set_font(r, font_size=16, bold=True)

    # 信息行
    p_info = doc.add_paragraph()
    r = p_info.add_run(f"学院：{data['college']}   专业：{data['major']}   班级：{data['class_name']}")
    set_font(r, font_size=10.5) # 小五/五号

    table = doc.add_table(rows=9, cols=5)
    table.style = 'Table Grid'

    # 行1：标题
    table.cell(0, 0).merge(table.cell(0, 4))
    set_cell_text(table.cell(0, 0), "岗前综合技能培训成绩评定", bold=True)

    # 行2：项目名称
    set_cell_text(table.cell(1, 0), "项目名称")
    table.cell(1, 1).merge(table.cell(1, 4))
    set_cell_text(table.cell(1, 1), data['project_name'])

    # 行3：成果形式
    set_cell_text(table.cell(2, 0), "成果形式")
    table.cell(2, 1).merge(table.cell(2, 4))
    set_cell_text(table.cell(2, 1), data['output_form'], align=WD_ALIGN_PARAGRAPH.LEFT)

    # 行4：指导教师评语 (合并左侧标题)
    set_cell_text(table.cell(3, 0), "指导教师评语")
    table.cell(3, 1).merge(table.cell(3, 4))
    set_cell_text(table.cell(3, 1), "（此处由指导教师填写评语）\n\n\n\n指导教师签名：          年   月   日", align=WD_ALIGN_PARAGRAPH.LEFT)

    # 行5：初评成绩
    set_cell_text(table.cell(4, 0), "初评成绩")
    table.cell(4, 1).merge(table.cell(4, 4))
    
    # 行6：答辩成绩 (复杂结构)
    table.cell(5, 0).merge(table.cell(7, 0))
    set_cell_text(table.cell(5, 0), "答辩成绩")
    
    # 答辩子项1
    table.cell(5, 1).merge(table.cell(5, 3))
    set_cell_text(table.cell(5, 1), "1. 成果水平和工作量评价 (满分80分)\nA. 创新，完成各项要求 (71-80)\nB. 有创新，基本完成 (61-70)\nC. 无创新，基本完成 (51-60)\nD. 未完成 (0-50)", align=WD_ALIGN_PARAGRAPH.LEFT, font_size=9)
    set_cell_text(table.cell(5, 4), "评分：")

    # 答辩子项2
    table.cell(6, 1).merge(table.cell(6, 3))
    set_cell_text(table.cell(6, 1), "2. 答辩表现 (满分20分)\nA. 准备充分，概念清楚 (15-20)\nB. 表现较好 (10-15)\nC. 表现一般 (5-10)\nD. 表现很差 (0-5)", align=WD_ALIGN_PARAGRAPH.LEFT, font_size=9)
    set_cell_text(table.cell(6, 4), "评分：")

    # 答辩小组签名
    table.cell(7, 1).merge(table.cell(7, 4))
    set_cell_text(table.cell(7, 1), "答辩小组负责人签名：                     年   月   日", align=WD_ALIGN_PARAGRAPH.RIGHT)

    # 最终成绩
    set_cell_text(table.cell(8, 0), "最终成绩")
    table.cell(8, 1).merge(table.cell(8, 4))
    set_cell_text(table.cell(8, 1), "优□   良□   中□   及格□   不及格□\n\n学院（签章）：             年   月   日", align=WD_ALIGN_PARAGRAPH.LEFT)

    doc.add_page_break()

def create_approval_form(doc, data):
    """生成选题审批表 (Page 5/6)"""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("海南软件职业技术学院\n岗前综合技能培训选题审批表")
    set_font(r, font_size=16, bold=True)

    table = doc.add_table(rows=5, cols=6)
    table.style = 'Table Grid'

    # 基础信息
    set_cell_text(table.cell(0, 0), "学号")
    set_cell_text(table.cell(0, 1), data['student_id'])
    set_cell_text(table.cell(0, 2), "姓名")
    set_cell_text(table.cell(0, 3), data['name'])
    set_cell_text(table.cell(0, 4), "班级")
    set_cell_text(table.cell(0, 5), data['class_name'])

    # 项目与教师
    set_cell_text(table.cell(1, 0), "项目名称")
    table.cell(1, 1).merge(table.cell(1, 3))
    set_cell_text(table.cell(1, 1), data['project_name'])
    set_cell_text(table.cell(1, 4), "指导教师")
    set_cell_text(table.cell(1, 5), data['teacher'])

    # 选题理由
    set_cell_text(table.cell(2, 0), "选题理由及\n准备情况")
    table.cell(2, 1).merge(table.cell(2, 5))
    set_cell_text(table.cell(2, 1), data['reason'], align=WD_ALIGN_PARAGRAPH.LEFT)
    table.rows[2].height = Cm(4)

    # 教师意见
    set_cell_text(table.cell(3, 0), "指导教师\n意见")
    table.cell(3, 1).merge(table.cell(3, 5))
    set_cell_text(table.cell(3, 1), "\n\n签字：             年   月   日", align=WD_ALIGN_PARAGRAPH.RIGHT)
    table.rows[3].height = Cm(3)

    # 学院意见
    set_cell_text(table.cell(4, 0), "学院意见")
    table.cell(4, 1).merge(table.cell(4, 5))
    set_cell_text(table.cell(4, 1), "\n\n签字（盖章）：             年   月   日", align=WD_ALIGN_PARAGRAPH.RIGHT)
    table.rows[4].height = Cm(3)

    doc.add_page_break()

def create_report_body_template(doc, data):
    """生成报告正文模板 (Page 7+)"""
    
    # 设置正文格式：宋体小四，1.5倍行距
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    font.size = Pt(12)
    style.paragraph_format.line_spacing = 1.5

    # 一、目的
    p = doc.add_paragraph()
    r = p.add_run("一、岗前培训目的")
    set_font(r, font_size=14, bold=True)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT

    p = doc.add_paragraph("（在此处撰写岗前培训的目的和意义，岗前培训单位的发展情况及学习要求等，不少于300字。）")
    p.paragraph_format.first_line_indent = Cm(0.74) # 首行缩进2字符

    # 二、内容
    doc.add_paragraph() # 空行
    p = doc.add_paragraph()
    r = p.add_run("二、岗前培训内容")
    set_font(r, font_size=14, bold=True)

    p = doc.add_paragraph("1、小标题一")
    set_font(p.runs[0], font_size=12, bold=True)
    doc.add_paragraph("（正文内容...）")

    p = doc.add_paragraph("2、小标题二")
    set_font(p.runs[0], font_size=12, bold=True)
    doc.add_paragraph("（正文内容...）")

    # 三、结果
    doc.add_paragraph()
    p = doc.add_paragraph()
    r = p.add_run("三、岗前培训结果")
    set_font(r, font_size=14, bold=True)
    doc.add_paragraph("（展示作品截图、代码片段或实物照片等。）")

    # 四、总结
    doc.add_paragraph()
    p = doc.add_paragraph()
    r = p.add_run("四、培训总结或体会")
    set_font(r, font_size=14, bold=True)
    doc.add_paragraph("（总结培训过程中的收获、不足以及对未来职业生涯的规划，字数建议不少于500字。）")

# --- Streamlit UI ---

st.set_page_config(page_title="岗前培训报告生成工具", layout="wide")
st.title("📝 岗前综合技能培训报告生成工具")
st.markdown("**适用院校：** 海南软件职业技术学院 | **输出格式：** 标准Word模板 (.docx)")

with st.sidebar:
    st.header("1. 基本信息输入")
    name = st.text_input("学生姓名", "张三")
    student_id = st.text_input("学号", "20220001")
    college = st.text_input("学院", "机电工程学院")
    major = st.text_input("专业", "软件技术")
    class_name = st.text_input("班级", "23软件技术1班")
    teacher = st.text_input("指导教师", "李四")
    
    st.header("2. 项目信息")
    project_name = st.text_input("项目名称", "基于Python的企业网站开发")
    # 默认日期处理
    today = datetime.date.today()
    start_date_obj = st.date_input("开始时间", datetime.date(2025, 7, 1))
    end_date_obj = st.date_input("结束时间", datetime.date(2025, 8, 31))
    start_date = start_date_obj.strftime("%Y年%m月%d日")
    end_date = end_date_obj.strftime("%Y年%m月%d日")
    
    st.header("3. 详细内容 (用于任务书)")
    meaning = st.text_area("项目的意义及培养目标", "通过本项目训练，掌握Web开发全流程，提升编码能力...", height=100)
    output_form = st.selectbox("成果形式", ["软件作品", "项目文档", "综述报告", "电子工艺产品", "其他"])
    requirements = st.text_area("技能训练基本要求", "1. 代码规范\n2. 功能完整\n3. 文档齐全", height=100)
    main_tasks = st.text_area("主要任务", "1. 需求分析\n2. 数据库设计\n3. 前端页面开发\n4. 后端接口实现", height=100)
    reason = st.text_area("选题理由 (审批表用)", "该项目符合专业培养目标，且能结合实习岗位实际...", height=80)

    st.header("4. 附件")
    logo_file = st.file_uploader("上传校徽 (可选)", type=['png', 'jpg', 'jpeg'])

# 数据打包
data = {
    "name": name, "student_id": student_id, "college": college,
    "major": major, "class_name": class_name, "teacher": teacher,
    "project_name": project_name, "start_date": start_date, "end_date": end_date,
    "meaning": meaning, "output_form": output_form, 
    "requirements": requirements, "main_tasks": main_tasks,
    "reason": reason
}

# 主界面预览区
st.info("👈 请在左侧侧边栏填写报告所需的详细信息。完成后点击下方按钮生成Word文档。")

col1, col2 = st.columns(2)
with col1:
    st.write("### 📄 包含页面预览")
    st.markdown("""
    1. **封面** (自动排版，含校徽)
    2. **岗前综合技能培训任务书** (自动填充任务详情)
    3. **指导记录表** (生成8周记录模板)
    4. **成绩评定表** (标准评分标准布局)
    5. **选题审批表** (含选题理由)
    6. **报告正文模板** (预设小四宋体、1.5倍行距、大纲)
    """)

with col2:
    st.write("### ⚙️ 操作")
    if st.button("🚀 生成报告 (.docx)", type="primary"):
        # 生成文档
        doc = Document()
        
        # 依次生成各页面
        create_cover(doc, data, logo_file)
        create_task_sheet(doc, data)
        create_guidance_record(doc, data)
        create_assessment(doc, data)
        create_approval_form(doc, data) # 放在正文前或后均可，此处按常见逻辑放前
        create_report_body_template(doc, data)
        
        # 保存到内存
        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        
        st.success("报告生成成功！请下载。")
        file_name = f"{student_id}_{name}_岗前综合技能培训报告.docx"
        st.download_button(
            label="📥 点击下载 Word 文档",
            data=buffer,
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

st.markdown("---")
st.caption("注：本工具仅生成格式规范的文档模板，正文具体内容及手写签名需下载后自行补充。")
