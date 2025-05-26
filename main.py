import os
from doc_generator import create_word_document
from variant_generator import read_questions_from_excel, generate_variants, save_to_excel


def read_students(file_path):
    with open(file_path, 'r', encoding='utf-8') as file:
        return [' '.join(line.strip().split()) for line in file.readlines() if line.strip()]


def main():
    # Configurations
    students_file = 'students.txt'
    questions_file = 'new_questions.xlsx'
    output_file = 'students_questions.xlsx'
    output_word_file = 'students_assignments.docx'
    max_semester = 2
    questions_per_topic = {
        "Common": 1,
        "Metaprogramming": 1,
        "Preprocessor&compilation": 1,
        "C++": 5,
        "TwoPointers": 1,
        "Algorithms": 1
    }

    personalized_questions = False

    if personalized_questions:
        students = read_students(students_file)
    else:
        if os.path.exists(students_file):
            with open(students_file, 'r', encoding='utf-8') as file:
                num_students = len(file.readlines())
        else:
            num_students = 10

        students = [f"Student {i + 1}" for i in range(num_students)]

    students.sort()
    questions_by_topic = read_questions_from_excel(questions_file, max_semester)

    try:
        student_variants = generate_variants(students, questions_by_topic, questions_per_topic)
    except ValueError as e:
        print(f"Error: {e}")
        return

    save_to_excel(student_variants, output_file)
    create_word_document(student_variants, output_word_file)
    print(f"Документы успешно созданы: {output_file}, {output_word_file}")

if __name__ == "__main__":
    main()
