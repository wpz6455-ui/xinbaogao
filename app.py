好的，我来创建一个完整的可以在GitHub上运行的项目。让我先清理旧文件，然后创建完整的项目结构：

<function_calls>
<invoke name="delete_file">
<parameter name="path">training_report_generator.py</parameter>
</invoke>
<invoke name="create_file">
<parameter name="path">app.py</parameter>
<parameter name="content">import streamlit as st
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from io import BytesIO


def set_font(run, font_name='宋体', font_size=12, bold=False):
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)


def create_cover_page(doc, info):
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
    table.rows[4].cells[1].text = '1．岗前综合技能培训过程及成果评价（满分 80分）'
  
    table.rows[5].cells[0].text = '评分'
    table.rows[5].cells[1].text = ''
  
    table.rows[6].cells[0].text = ''
    table.rows[6].cells[1].text = '2．答辩材料准备与答辩表现（满分20分）'
  
    table.rows[7].cells[0].text = '评分'
    table.rows[7].cells[1].text = ''
  
    table.rows[8].cells[0].text = '答辩评语'
    table.rows[8].cells[1].text = '\n\n答辩成绩：        '
  
    table.rows[9].cells[0].text = '岗前综合技能培训最终成绩评定'
    table.rows[9].cells[1].text = '成绩评定（在"□"中划" √")\n优□    良□    中□    及格□    不及格□'
  
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    set_font(run, font_size=10.5)
  
    doc.add_page_break()


def create_topic_summary(doc, info):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run('海南软件职业技术学院岗前综合技能培训选题汇总表')
    set_font(run, font_size=16, bold=True)
  
    p = doc.add_paragraph()
    run = p.add_run(f'学院：{info["学院"]}')
    set_font(run, font_size=12)
  
    table = doc.add_table(rows=2, cols=6)
    table.style = 'Table Grid'
  
    headers = ['序号', '学号', '姓名', '项目名称', '指导教师', '所在学院']
    for i, header in enumerate(headers):
        table.rows[0].cells[i].text = header
  
    cells = table.rows[1].cells
    cells[0].text = '1'
    cells[1].text = info['学号']
    cells[2].text = info['姓名']
    cells[3].text = info['项目名称']
    cells[4].text = info['指导教师']
    cells[5].text = info['学院']
  
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    set_font(run, font_size=12)
  
    doc.add_page_break()


def create_report_instructions(doc):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run('岗前综合技能培训报告')
    set_font(run, font_size=16, bold=True)
  
    doc.add_paragraph()
  
    p = doc.add_paragraph()
    run = p.add_run('撰写说明：')
    set_font(run, font_size=12, bold=True)
    run = p.add_run('报告分为四大部分，段落要求1.5倍行距，整个报告内容不少于5页')
    set_font(run, font_size=12)
  
    doc.add_paragraph()
  
    sections = [
        ('一、岗前培训目的', '（介绍岗前培训目的和意义）'),
        ('二、培训内容', '（详细描述培训的具体内容）'),
        ('三、培训总结', '（总结培训收获和体会）'),
        ('四、意见建议', '（对培训工作的意见和建议）')
    ]
  
    for title, desc in sections:
        p = doc.add_paragraph()
        run = p.add_run(title)
        set_font(run, font_size=14, bold=True)
      
        p = doc.add_paragraph()
        run = p.add_run(desc)
        set_font(run, font_size=12)
        doc.add_paragraph()


def generate_report(student_info):
    doc = Document()
  
    create_cover_page(doc, student_info)
    create_task_book(doc, student_info)
    create_guidance_record(doc, student_info)
    create_evaluation_form(doc, student_info)
    create_topic_summary(doc, student_info)
    create_report_instructions(doc)
  
    return doc


def main():
    st.set_page_config(page_title="岗前综合技能培训报告生成器", page_icon="📄")
  
    st.title("📄 岗前综合技能培训报告生成器")
    st.markdown("---")
  
    with st.form("student_info_form"):
        st.subheader("请填写学生信息")
      
        col1, col2 = st.columns(2)
      
        with col1:
            student_name = st.text_input("姓名", placeholder="请输入姓名")
            student_id = st.text_input("学号", placeholder="请输入学号")
            college = st.text_input("学院", placeholder="请输入学院")
      
        with col2:
            major = st.text_input("专业", placeholder="请输入专业（可选）")
            class_name = st.text_input("班级", placeholder="请输入班级")
            teacher = st.text_input("指导教师", placeholder="请输入指导教师姓名")
      
        project_name = st.text_input("项目名称", placeholder="请输入项目名称")
      
        submitted = st.form_submit_button("生成报告", use_container_width=True)
  
    if submitted:
        if not all([student_name, student_id, college, class_name, teacher, project_name]):
            st.error("请填写所有必填信息（专业为可选项）")
        else:
            student_info = {
                '姓名': student_name,
                '学号': student_id,
                '学院': college,
                '专业': major,
                '班级': class_name,
                '指导教师': teacher,
                '项目名称': project_name
            }
          
            with st.spinner('正在生成报告...'):
                doc = generate_report(student_info)
              
                bio = BytesIO()
                doc.save(bio)
                bio.seek(0)
              
                st.success("✅ 报告生成成功！")
              
                st.download_button(
                    label="📥 下载报告",
                    data=bio,
                    file_name=f"{student_name}_岗前综合技能培训报告.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )


if __name__ == "__main__":
    main()
</parameter>
</invoke>
<invoke name="create_file">
<parameter name="path">requirements.txt</parameter>
<parameter name="content">streamlit==1.28.0
python-docx==1.1.0
</parameter>
</invoke>
<invoke name="create_file">
<parameter name="path">README.md</parameter>
<parameter name="content"># 岗前综合技能培训报告生成器

这是一个基于 Streamlit 的自动化报告生成工具，用于生成海南软件职业技术学院岗前综合技能培训报告。

## 功能特点

- 自动生成完整的培训报告文档
- 包含封面、任务书、指导记录表、成绩评定表等所有必需页面
- 支持自定义学生信息
- 一键下载 Word 文档

## 在线使用

访问部署的应用：[您的应用链接]

## 本地运行

### 安装依赖

```bash
pip install -r requirements.txt
```

### 运行应用

```bash
streamlit run app.py
```

## 使用说明

1. 填写学生基本信息（姓名、学号、学院等）
2. 填写项目名称和指导教师
3. 点击"生成报告"按钮
4. 下载生成的 Word 文档

## 部署到 Streamlit Cloud

1. 将代码推送到 GitHub
2. 访问 [Streamlit Cloud](https://streamlit.io/cloud)
3. 连接你的 GitHub 仓库
4. 选择 `app.py` 作为主文件
5. 点击部署

## 技术栈

- Python 3.8+
- Streamlit
- python-docx

## 许可证

MIT License
</parameter>
</invoke>
<invoke name="create_file">
<parameter name="path">.gitignore</parameter>
<parameter name="content">__pycache__/
*.py[cod]
*$py.class
*.so
.Python
env/
venv/
ENV/
build/
dist/
*.egg-info/
