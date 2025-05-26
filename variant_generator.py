import random
import math
from collections import defaultdict
import pandas as pd


def read_questions_from_excel(file_path, max_semester):
    """
    Считывает вопросы из Excel файла с учетом семестра и типа вопросов.

    :param file_path: путь к Excel файлу
    :param max_semester: максимальный номер семестра для фильтрации вопросов
    :return: словарь {тема: [(вопрос, тип)]}
    """
    df = pd.read_excel(file_path, header=None, names=["Question", "Topic", "Semester", "Type"])
    filtered_df = df[df["Semester"] <= max_semester]
    questions_by_topic = defaultdict(list)
    for _, row in filtered_df.iterrows():
        questions_by_topic[row["Topic"]].append((row["Question"], row["Type"]))
    return questions_by_topic


def generate_variants(students, all_questions_by_topic, questions_per_topic):
    """
    Генерирует варианты для студентов с учетом теоретических и практических вопросов.

    :param students: список студентов
    :param all_questions_by_topic: словарь {тема: [(вопрос, тип)]}
    :param questions_per_topic: словарь {тема: количество вопросов}
    :return: словарь {студент: [(вопрос, тип)]}
    """
    student_variants = defaultdict(list)
    question_count = defaultdict(int)

    max_repeats_by_topic = {
        topic: math.ceil(len(students) / len(all_questions_by_topic[topic])) + 2
        for topic in all_questions_by_topic
    }

    for student in students:
        assigned_questions = set()
        for topic, num_questions in questions_per_topic.items():
            questions_by_topic = set()

            while len(questions_by_topic) < num_questions:
                question, question_type = random.choice(all_questions_by_topic[topic])

                # Проверяем, не превышено ли максимальное количество повторений для вопроса
                if question_count[question] >= max_repeats_by_topic[topic]:
                    all_questions_by_topic[topic] = [
                        q for q in all_questions_by_topic[topic] if q[0] != question
                    ]

                if not all_questions_by_topic[topic]:
                    raise ValueError(f"Not enough questions in the topic '{topic}' to assign {num_questions} questions.")

                questions_by_topic.add((question, question_type))
                question_count[question] += 1

            assigned_questions.update(questions_by_topic)

        student_variants[student] = list(assigned_questions)

    return student_variants


def save_to_excel(student_variants, output_path):
    df = pd.DataFrame.from_dict(student_variants, orient='index').reset_index()
    df.columns = ['Student'] + [f'Question {i+1}' for i in range(len(df.columns) - 1)]
    df.to_excel(output_path, index=False)
    print(f"Variants saved to {output_path}")
