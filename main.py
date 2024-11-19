from doc_generator import create_word_document
from variant_generator import read_students, read_questions_from_excel, generate_student_questions, save_to_excel


def main():
    students_file = 'students.txt'
    questions_file = 'questions.xlsx'
    output_word_file = 'students_assignments.docx'
    output_excel_file = 'students_questions.xlsx'

    students = read_students(students_file)
    questions_text = read_questions_from_excel(questions_file)

    students_questions = generate_student_questions(students, total_questions=250, questions_per_student=10)
    save_to_excel(students_questions, output_excel_file)

    create_word_document(students_questions, questions_text, output_word_file)
    print(f"Документы успешно созданы: {output_excel_file}, {output_word_file}")


if __name__ == "__main__":
    main()
