import asyncio
import os
from tortoise import Tortoise

from models.models import (
    Teacher, Group, Student, Discipline,
    ControlWork, Literature,
    Publication, Template
)

from core.config import settings


async def init():
    await Tortoise.init(
        db_url=settings.DATABASE_URL,
        modules={"models": ["models.models"]},
    )
    await Tortoise.generate_schemas()

    if await Teacher.exists():
        print("База данных уже содержит данные. Проверяем наличие шаблонов...")

        await import_templates()

        await Tortoise.close_connections()
        return

    print("Начинаем импорт данных...")

    teachers_data = [
        {"full_name": "Александрова Виктория Петровна", "email": "viktoria.aleksandrova@example.com"},
        {"full_name": "Громов Сергей Юрьевич", "email": "sergey.gromov@example.com"},
        {"full_name": "Лебедева Наталья Станиславовна", "email": "natalya.lebedyeva@example.com"},
        {"full_name": "Ковалев Игорь Алексеевич", "email": "igor.kovalev@example.com"},
        {"full_name": "Морозова Дарья Владимировна", "email": "daria.morozova@example.com"},
        {"full_name": "Чернов Николай Семёнович", "email": "nikolai.chernov@example.com"},
        {"full_name": "Соколова Елена Андреевна", "email": "elena.sokolova@example.com"},
        {"full_name": "Романов Денис Олегович", "email": "denis.romanov@example.com"},
        {"full_name": "Федорова Светлана Артемовна", "email": "svetlana.fedorova@example.com"},
        {"full_name": "Зайцева Ксения Валерьевна", "email": "ksenia.zaytseva@example.com"}
    ]

    for teacher_data in teachers_data:
        await Teacher.create(**teacher_data)

    print(f"Импортировано {len(teachers_data)} преподавателей")

    groups_data = [
        {"code": "КМБО-05-21", "course": "Программирование и алгоритмы"},
        {"code": "КМБО-02-21", "course": "Электротехника и электроника"},
        {"code": "КМБО-05-22", "course": "Техническая физика"},
        {"code": "КМБО-02-22", "course": "Теория автоматического управления"},
        {"code": "КМБО-05-23", "course": "Материаловедение"},
        {"code": "КМБО-02-23", "course": "Машиностроение"}
    ]

    teachers = await Teacher.all()
    for i, group_data in enumerate(groups_data):
        teacher = teachers[i % len(teachers)]
        group_data["teacher_id"] = teacher.id
        await Group.create(**group_data)

    print(f"Импортировано {len(groups_data)} групп")

    students_data = [
        # КМБО-05-21
        {"full_name": "Петрова Марина Викторовна", "email": "marina.petrova@example.com", "group_id": 1},
        {"full_name": "Сергеев Алексей Андреевич", "email": "aleksey.sergeyev@example.com", "group_id": 1},
        {"full_name": "Кузнецова Ольга Николаевна", "email": "olga.kuznetsova@example.com", "group_id": 1},
        {"full_name": "Фролов Дмитрий Павлович", "email": "dmitriy.frolov@example.com", "group_id": 1},
        {"full_name": "Тихонова Светлана Юрьевна", "email": "svetlana.tikhonova@example.com", "group_id": 1},
        {"full_name": "Лебедев Игорь Анатольевич", "email": "igor.lebedev@example.com", "group_id": 1},
        {"full_name": "Михайлова Анна Сергеевна", "email": "anna.mikhaylova@example.com", "group_id": 1},
        {"full_name": "Сидоров Василий Тимофеевич", "email": "vasiliy.sidorov@example.com", "group_id": 1},
        {"full_name": "Громова Екатерина Валерьевна", "email": "ekaterina.gromova@example.com", "group_id": 1},
        {"full_name": "Зайцев Константин Олегович", "email": "konstantin.zaytsev@example.com", "group_id": 1},

        # КМБО-02-21
        {"full_name": "Андреев Артем Сергеевич", "email": "artem.andreev@example.com", "group_id": 2},
        {"full_name": "Семенова Кристина Валерьевна", "email": "kristina.semenova@example.com", "group_id": 2},
        {"full_name": "Волков Павел Викторович", "email": "pavel.volkov@example.com", "group_id": 2},
        {"full_name": "Шевченко Дарья Александровна", "email": "darya.shevchenko@example.com", "group_id": 2},
        {"full_name": "Кириллова Наталья Дмитриевна", "email": "natalya.kirillova@example.com", "group_id": 2},
        {"full_name": "Белов Николай Федорович", "email": "nikolay.belov@example.com", "group_id": 2},
        {"full_name": "Соболева Елена Ивановна", "email": "elena.soboleva@example.com", "group_id": 2},
        {"full_name": "Григорьев Илья Андреевич", "email": "ilya.grigoryev@example.com", "group_id": 2},
        {"full_name": "Ларина Юлия Степановна", "email": "yulia.larina@example.com", "group_id": 2},
        {"full_name": "Федосеева Анна Борисовна", "email": "anna.fedoseeva@example.com", "group_id": 2},
        {"full_name": "Тимофеев Максим Алексеевич", "email": "maksim.timofeev@example.com", "group_id": 2},

        # Другие группы...
        # КМБО-05-22
        {"full_name": "Ковалев Сергей Дмитриевич", "email": "sergey.kovalev@example.com", "group_id": 3},
        {"full_name": "Павлова Оксана Александровна", "email": "oksana.pavlova@example.com", "group_id": 3},

        # КМБО-02-22
        {"full_name": "Семенов Виктор Игоревич", "email": "viktor.semenov@example.com", "group_id": 4},
        {"full_name": "Кузнецова Вероника Александровна", "email": "veronika.kuznetsova@example.com", "group_id": 4},

        # КМБО-05-23
        {"full_name": "Ковалев Алексей Викторович", "email": "aleksey.kovalev@example.com", "group_id": 5},
        {"full_name": "Мартынова Татьяна Сергеевна", "email": "tatyana.martynova@example.com", "group_id": 5},

        # КМБО-02-23
        {"full_name": "Семенов Николай Викторович", "email": "nikolay.semenov@example.com", "group_id": 6},
        {"full_name": "Лебедева Ксения Андреевна", "email": "kseniya.lebedeva@example.com", "group_id": 6}
    ]

    for student_data in students_data:
        await Student.create(**student_data)

    print(f"Импортировано {len(students_data)} студентов")

    disciplines_data = [
        {
            "name": "Инженерная механика",
            "education_level": "Бакалавариат",
            "department": "БК №536",
            "hours_lecture": 10,
            "hours_practice": 0,
            "hours_lab": 25
        },
        {
            "name": "Электротехника и электроника",
            "education_level": "Магистратура",
            "department": "",
            "hours_lecture": 5,
            "hours_practice": 1,
            "hours_lab": 40
        },
        {
            "name": "Программирование и алгоритмы",
            "education_level": "Бакалавариат",
            "department": "",
            "hours_lecture": 20,
            "hours_practice": 0,
            "hours_lab": 60
        },
        {
            "name": "Теория автоматического управления",
            "education_level": "Магистратура",
            "department": "",
            "hours_lecture": 15,
            "hours_practice": 2,
            "hours_lab": 35
        },
        {
            "name": "Материаловедение",
            "education_level": "Бакалавариат",
            "department": "",
            "hours_lecture": 8,
            "hours_practice": 1,
            "hours_lab": 50
        }
    ]

    for discipline_data in disciplines_data:
        await Discipline.create(**discipline_data)

    print(f"Импортировано {len(disciplines_data)} дисциплин")

    control_works_data = []
    disciplines = await Discipline.all()

    for discipline in disciplines:
        for i in range(1, 6):  # 5 контрольных работ для каждой дисциплины
            control_format = "очно"
            if i == 4 and discipline.id in [1, 5]:  # Для некоторых дисциплин 4-я работа - СДО
                control_format = "СДО"
            elif i == 5 and discipline.id in [3, 4]:  # Для некоторых дисциплин 5-я работа - СДО
                control_format = "СДО"

            control_works_data.append({
                "number": i,
                "max_score": 10 if i <= 3 else 5,  # Первые три по 10 баллов, остальные по 5
                "week": [3, 4, 7, 10, 14][i - 1],  # Недели из таблицы
                "semester": 1,
                "format": control_format,
                "discipline_id": discipline.id
            })

    for control_work_data in control_works_data:
        await ControlWork.create(**control_work_data)

    print(f"Импортировано {len(control_works_data)} контрольных работ")

    literature_data = [
        {
            "title": "Фейнмановские лекции по физике",
            "authors": "Фейнман, Р. П.",
            "publisher": "Наука",
            "year": 1965,
            "discipline_id": 1  # Инженерная механика
        },
        {
            "title": "Алгоритмы: построение и анализ",
            "authors": "Бен-Ор, М.",
            "publisher": "Вильямс",
            "year": 2005,
            "discipline_id": 3  # Программирование и алгоритмы
        },
        {
            "title": "Основы теории автоматического управления",
            "authors": "Дьяконов, А. А.",
            "publisher": "Высшая школа",
            "year": 2000,
            "discipline_id": 4  # Теория автоматического управления
        }
    ]

    for lit_data in literature_data:
        await Literature.create(**lit_data)

    print(f"Импортировано {len(literature_data)} записей литературы")

    # Импорт данных о публикациях
    publications_data = [
        {
            "title": "Разработка и оптимизация алгоритмов машинного обучения для обработки больших данных.",
            "student_id": 1  # Петрова Марина Викторовна
        },
        {
            "title": "Сравнительный анализ языков программирования для веб-разработки: производительность, безопасность и удобство.",
            "student_id": 2  # Сергеев Алексей Андреевич
        },
        {
            "title": "Исследование подходов к разработке программного обеспечения с использованием методологий Agile и DevOps.",
            "student_id": 3  # Кузнецова Ольга Николаевна
        }
    ]

    for pub_data in publications_data:
        await Publication.create(**pub_data)

    print(f"Импортировано {len(publications_data)} публикаций")

    await import_templates()

    print("Инициализация базы данных завершена успешно.")
    await Tortoise.close_connections()


async def import_templates():
    """Импортирует все шаблоны из директории templates в базу данных"""

    if not os.path.exists(settings.TEMPLATE_DIR):
        print(f"Директория шаблонов {settings.TEMPLATE_DIR} не найдена. Создаем...")
        os.makedirs(settings.TEMPLATE_DIR, exist_ok=True)
        return

    template_files = [f for f in os.listdir(settings.TEMPLATE_DIR)
                      if os.path.isfile(os.path.join(settings.TEMPLATE_DIR, f))
                      and (f.endswith('.docx') or f.endswith('.xlsx'))]

    template_descriptions = {
        "БРС_Методы_и_стандарты_программирования_1_536_Б.docx": "Балльно-рейтинговая система по дисциплине",
        "Журнал A4.xlsx": "Журнал оценок формата A4",
        "Журнал A5.xlsx": "Журнал оценок формата A5",
        "Загруженность аудиторий.docx": "Шаблон загруженности аудиторий",
        "ЗАДАНИЕ 4 курс.docx": "Шаблон задания для 4 курса",
        "ЗАДАНИЕ 6 курс.docx": "Шаблон задания для 6 курса",
        "Отчет по практике 6 семестр.docx": "Шаблон отчета по практике за 6 семестр",
        "Отчет по практике 7 семестр.docx": "Шаблон отчета по практике за 7 семестр",
        "Отчет по технологической практике.docx": "Шаблон отчета по технологической практике",
        "Программа секции ИТиПОРЭА.docx": "Шаблон программы секции ИТиПОРЭА",
        "Расписание преподавателей.docx": "Шаблон расписания преподавателей",
        "Список вопросом для экзамена или зачета.docx": "Шаблон списка вопросов к экзамену или зачету",
        "Список группы.xlsx": "Шаблон списка группы в Excel",
        "Список литературу.docx": "Шаблон списка литературы",
        "Список публикаций.docx": "Шаблон списка публикаций",
        "Темы практических работ.docx": "Шаблон тем практических работ",
        "Титульный лист Курсовая работа.docx": "Шаблон титульного листа для курсовой работы",
        "Титульный лист Курсовой проект.docx": "Шаблон титульного листа для курсового проекта",
        "Титульный лист Лабораторная работа.docx": "Шаблон титульного листа для лабораторной работы",
        "Титульный лист Реферат.docx": "Шаблон титульного листа для реферата",
        "Шаблон билета.docx": "Шаблон экзаменационного билета",
        "Шаблон титул бак.docx": "Шаблон титульного листа для бакалавра",
        "Шаблон титул маг.docx": "Шаблон титульного листа для магистра"
    }

    added_count = 0

    for file_name in template_files:
        file_type = file_name.split('.')[-1].lower()
        template_name = file_name.rsplit('.', 1)[0]
        existing_template = await Template.filter(file_path=file_name).first()
        if not existing_template:
            description = template_descriptions.get(file_name, f"Шаблон {template_name}")

            await Template.create(
                name=template_name,
                file_path=file_name,
                file_type=file_type,
                description=description
            )
            added_count += 1

    print(f"Импортировано {added_count} шаблонов из директории {settings.TEMPLATE_DIR}")

    if added_count == 0 and not template_files:
        print("Директория шаблонов пуста. Шаблоны не импортированы.")


if __name__ == "__main__":
    asyncio.run(init())