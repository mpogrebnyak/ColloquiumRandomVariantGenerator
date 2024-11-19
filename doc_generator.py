from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Inches, Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT


def add_page_break(document):
    page_break = OxmlElement('w:br')
    page_break.set(qn('w:type'), 'page')
    document.add_paragraph()._element.append(page_break)


def set_cell_border(cell, **kwargs):
    tc = cell._element
    tcPr = tc.get_or_add_tcPr()
    for border_name in ["top", "left", "bottom", "right"]:
        if border_name in kwargs:
            border = OxmlElement(f"w:{border_name}")
            for key, value in kwargs[border_name].items():
                border.set(qn(f"w:{key}"), str(value))
            tcPr.append(border)


def create_word_document(students_questions, questions_text, output_path):
    document = Document()

    # Header and footer setup
    for section in document.sections:
        header = section.header
        paragraph = header.paragraphs[0]
        paragraph.text = "Коллоквиум по курсу \"Основы программирования\""
        paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT

        footer = section.footer
        paragraph = footer.paragraphs[0]
        paragraph.text = "Преподаватель: Погребняк Максим Анатольевич"
        paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

    # Narrow margins
    for section in document.sections:
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)

    # Generate content
    for count, (student, question_numbers) in enumerate(students_questions.items(), start=1):
        document.add_paragraph()

        title_paragraph = document.add_paragraph()
        title_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        run = title_paragraph.add_run(f"Бланк заданий/ответов №{count}")
        run.bold = True
        run.font.size = Pt(16)

        fio_group_paragraph = document.add_paragraph()
        fio_group_paragraph.add_run("ФИО: _______________________________________  Группа: ____________").bold = True

        # Table setup
        table = document.add_table(rows=2, cols=12)
        table.alignment = WD_TABLE_ALIGNMENT.LEFT
        table.autofit = False

        for i in range(10):
            cell = table.cell(0, i)
            cell.text = str(i + 1)
            cell.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

        table.cell(0, 10).text = "к/р"
        table.cell(0, 11).text = "всего"

        for row in table.rows:
            for cell in row.cells:
                set_cell_border(cell, top={"sz": 6, "val": "single", "color": "000000"},
                                bottom={"sz": 6, "val": "single", "color": "000000"},
                                left={"sz": 6, "val": "single", "color": "000000"},
                                right={"sz": 6, "val": "single", "color": "000000"})

        explanation_text = (
            "Таблица заполняется преподавателем. Самостоятельно заполнять её не нужно. "
            "За каждое задание можно получить от 0 до 1 балла.\n"
            "Ответы на задания необходимо писать на этом бланке с передней и обратной стороны."
        )
        document.add_paragraph(explanation_text, style="Normal").alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

        for i, q_num in enumerate(question_numbers, start=1):
            document.add_paragraph(f"{i}. {questions_text[q_num]}").paragraph_format.line_spacing = Pt(9)

        add_page_break(document)

    document.save(output_path)
