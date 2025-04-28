import os
import re
import random
from typing import Dict, Any, List, Optional
from docx import Document
from docx.shared import Pt

from services.base_document_service import BaseDocumentService
from models.models import (
    Group, Student, Teacher, Discipline, ExamQuestion,
    ControlWork, ScheduleItem, Classroom
)


class DocxService(BaseDocumentService):
    """
    Сервис для обработки DOCX-документов
    """

    async def load_document(self, template_path: str) -> Document:
        """
        Загружает DOCX-документ из файла шаблона

        Args:
            template_path: Путь к файлу шаблона

        Returns:
            Объект документа Word
        """
        return Document(template_path)

    async def save_document(self, document: Document, output_path: str) -> None:
        """
        Сохраняет DOCX-документ в файл

        Args:
            document: Объект документа Word
            output_path: Путь для сохранения
        """
        document.save(output_path)

    async def apply_times_new_roman(self, document: Document) -> None:
        """
        Устанавливает шрифт Times New Roman для всех элементов документа

        Args:
            document: Объект документа Word
        """
        # Для всех параграфов
        for paragraph in document.paragraphs:
            for run in paragraph.runs:
                run.font.name = "Times New Roman"
                # Сохраняем оригинальный размер, если он есть
                if run.font.size:
                    continue
                # Иначе устанавливаем стандартный размер 12pt
                run.font.size = Pt(12)

        # Для всех таблиц
        for table in document.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            run.font.name = "Times New Roman"
                            # Сохраняем оригинальный размер, если он есть
                            if run.font.size:
                                continue
                            # Иначе устанавливаем стандартный размер 12pt
                            run.font.size = Pt(12)

    async def replace_yellow_highlights(self, document: Document, highlights_replacements: Dict[str, Any]) -> None:
        """
        Заменяет текст с желтым выделением на соответствующие значения

        Args:
            document: документ Word
            highlights_replacements: словарь с заменами
        """
        # Для всех параграфов
        for paragraph in document.paragraphs:
            for run in paragraph.runs:
                # Проверяем наличие выделения
                if hasattr(run, 'font') and hasattr(run.font, 'highlight_color') and run.font.highlight_color:
                    # Заменяем текст в зависимости от его содержимого
                    if 'discipline_name' in highlights_replacements and run.text.strip():
                        run.text = highlights_replacements['discipline_name']
                    elif 'institute' in highlights_replacements and run.text.lower().strip() == 'ипи':
                        run.text = highlights_replacements['institute']
                    elif 'department' in highlights_replacements and 'бк' in run.text.lower().strip():
                        run.text = highlights_replacements['department']
                    elif 'education_level' in highlights_replacements and run.text.lower().strip() == 'бакалавриат':
                        run.text = highlights_replacements['education_level']
                    elif 'hours' in highlights_replacements and '/' in run.text.strip():
                        run.text = highlights_replacements['hours']

        # Для всех таблиц
        for table in document.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            # Проверяем наличие выделения
                            if hasattr(run, 'font') and hasattr(run.font,
                                                                'highlight_color') and run.font.highlight_color:
                                # Производим специальные замены для таблиц БРС
                                if 'control_works' in highlights_replacements and run.text.strip():
                                    # Определяем, какой элемент контрольной работы это
                                    text_lower = run.text.lower().strip()
                                    if 'контрольная работа' in text_lower or 'тестирование' in text_lower:
                                        # Это название контрольной работы, находим её номер
                                        for i, work in enumerate(highlights_replacements['control_works']):
                                            if f"№{i + 1}" in run.text or (i == 3 and 'тест' in text_lower):
                                                run.text = work['name']
                                                break
                                    elif 'очно' in text_lower or 'сдо' in text_lower:
                                        # Это формат контрольной работы
                                        for i, work in enumerate(highlights_replacements['control_works']):
                                            if i == int(paragraph._p.xpath('./preceding-sibling::w:p')[
                                                            0].text_content().strip()) - 1:
                                                run.text = work['format']
                                                break
                                    elif run.text.isdigit() and int(run.text) == 10:
                                        # Это баллы за контрольную работу
                                        for i, work in enumerate(highlights_replacements['control_works']):
                                            if i == int(paragraph._p.xpath('./preceding-sibling::w:p')[
                                                            0].text_content().strip()) - 1:
                                                run.text = str(work['max_score'])
                                                break
                                    elif run.text.isdigit() and int(run.text) in [6, 10, 14, 16]:
                                        # Это период проведения
                                        for i, work in enumerate(highlights_replacements['control_works']):
                                            if i == int(paragraph._p.xpath('./preceding-sibling::w:p')[
                                                            0].text_content().strip()) - 1:
                                                run.text = f"{work['week']} неделя"
                                                break

    async def process_exam_ticket(self, document: Document, discipline: Discipline,
                                  ticket_number: Optional[int] = None) -> Document:
        """
        Специальная обработка для экзаменационного билета

        Args:
            document: объект документа Word
            discipline: объект дисциплины
            ticket_number: номер билета (если не указан, выбирается случайно)

        Returns:
            Обработанный документ Word
        """
        # Получаем вопросы к экзамену для данной дисциплины
        questions = await ExamQuestion.filter(discipline_id=discipline.id).order_by('number')

        if not questions:
            # Если нет вопросов, создаем тестовые
            questions_data = [
                {"number": 1, "text": "Основные принципы и концепции дисциплины"},
                {"number": 2, "text": "Ключевые методы и подходы в изучаемой области"},
                {"number": 3, "text": "Практическое применение теоретических знаний"}
            ]
        else:
            # Преобразуем в формат для удобства использования
            questions_data = [{"number": q.number, "text": q.text} for q in questions]

        # Если не указан номер билета, выбираем случайно от 1 до 30
        if not ticket_number:
            ticket_number = random.randint(1, 30)

        # Выбираем случайные вопросы (если их меньше 3, берем все имеющиеся)
        if len(questions_data) <= 3:
            selected_questions = questions_data
        else:
            # Выбираем 3 случайных вопроса
            selected_questions = random.sample(questions_data, 3)

        # Определяем информацию для замены в шаблоне
        replacements = {
            "ticket_number": str(ticket_number),
            "discipline_name": discipline.name,
            "course": "Программирование и алгоритмы",  # Можно получить из дисциплины или группы
            "semester": "1",  # Можно получить из параметров
            "questions": selected_questions
        }

        # Обрабатываем таблицы в документе для замены содержимого
        for table in document.tables:
            for row_idx, row in enumerate(table.rows):
                for col_idx, cell in enumerate(row.cells):
                    # Обрабатываем все параграфы в ячейке
                    for paragraph in cell.paragraphs:
                        # Заменяем номер билета и название дисциплины
                        if "БИЛЕТ №" in paragraph.text:
                            # Сохраняем форматирование
                            original_text = paragraph.text
                            # Создаем новый текст с заменой
                            new_text = f"БИЛЕТ № {ticket_number} {discipline.name}"
                            # Заменяем текст, сохраняя форматирование
                            paragraph.text = new_text

                        # Заменяем "Курс N" на "Курс [название]"
                        if "Курс N" in paragraph.text:
                            paragraph.text = paragraph.text.replace("Курс N", f"Курс {replacements['course']}")

                        # Заменяем "Семестр N" на "Семестр [номер]"
                        if "Семестр N" in paragraph.text:
                            paragraph.text = paragraph.text.replace("Семестр N", f"Семестр {replacements['semester']}")

                        # Заменяем вопросы для билета
                        for i, question in enumerate(selected_questions):
                            question_marker = f"{i + 1}. M"
                            if question_marker in paragraph.text:
                                paragraph.text = paragraph.text.replace(question_marker, f"{i + 1}. {question['text']}")

                        # Точечно ищем другие замены по контексту
                        if "Дисциплина:" in paragraph.text:
                            # Находим следующую ячейку и заменяем текст
                            if col_idx + 1 < len(row.cells):
                                next_cell = row.cells[col_idx + 1]
                                if next_cell.text:
                                    next_cell.text = discipline.name

        return document

    async def process_teacher_schedule(self, document: Document, teacher_id: int,
                                       day_of_week: Optional[str] = None) -> Document:
        """
        Специальная обработка для расписания преподавателя

        Args:
            document: объект документа Word
            teacher_id: ID преподавателя
            day_of_week: день недели (если не указан, берется всё расписание)

        Returns:
            Обработанный документ Word
        """
        # Получаем преподавателя
        teacher = await Teacher.get(id=teacher_id)

        # Получаем расписание преподавателя
        schedule_query = ScheduleItem.filter(teacher_id=teacher_id).prefetch_related(
            'time_slot', 'discipline', 'group', 'classroom'
        )

        # Если указан день недели, фильтруем по нему
        if day_of_week:
            schedule_query = schedule_query.filter(day_of_week=day_of_week)

        # Получаем отсортированное расписание
        schedule_items = await schedule_query.order_by('day_of_week', 'time_slot__number')

        # Группируем расписание по дням недели
        schedule_by_day = {}
        for item in schedule_items:
            day = item.day_of_week.value
            if day not in schedule_by_day:
                schedule_by_day[day] = []
            schedule_by_day[day].append(item)

        # Поиск таблицы для расписания в документе
        schedule_table = None
        for table in document.tables:
            # Ищем таблицу, в которой есть нужные заголовки
            first_row_text = ' '.join([cell.text for cell in table.rows[0].cells])
            if "День" in first_row_text and "Время" in first_row_text:
                schedule_table = table
                break

        # Если нашли таблицу расписания, заполняем её
        if schedule_table:
            # Определяем количество строк в таблице (исключая заголовок)
            num_rows = len(schedule_table.rows)

            # Если дней больше, чем строк (минус заголовок), добавляем строки
            days_count = len(schedule_by_day)
            if days_count > num_rows - 1:
                for _ in range(days_count - (num_rows - 1)):
                    schedule_table.add_row()

            # Заполняем таблицу данными
            row_index = 1  # Начинаем с первой строки после заголовка
            for day, items in schedule_by_day.items():
                if row_index < len(schedule_table.rows):
                    row = schedule_table.rows[row_index]

                    # Заполняем день недели
                    if len(row.cells) > 0:
                        row.cells[0].text = day

                    # Формируем расписание на день
                    schedule_text = ""
                    for item in items:
                        time_slot = f"{item.time_slot.start_time}-{item.time_slot.end_time}"
                        discipline = item.discipline.name
                        group = item.group.code
                        classroom = item.classroom.name if item.classroom else "Нет аудитории"

                        schedule_text += f"{time_slot}: {discipline}, {group}, {classroom}\n"

                    # Заполняем расписание
                    if len(row.cells) > 1:
                        row.cells[1].text = schedule_text.strip()

                    row_index += 1

        # Обновляем информацию о преподавателе в документе
        for paragraph in document.paragraphs:
            if "Расписание преподавателя:" in paragraph.text:
                paragraph.text = f"Расписание преподавателя: {teacher.full_name}"
                break

        # Если расписание пустое, добавляем сообщение
        if not schedule_items:
            for paragraph in document.paragraphs:
                if "Расписание" in paragraph.text:
                    document.add_paragraph("У преподавателя нет занятий в расписании.")
                    break

        return document

    async def process_classroom_schedule(self, document: Document, classroom_id: int,
                                         day_of_week: Optional[str] = None) -> Document:
        """
        Специальная обработка для расписания аудитории

        Args:
            document: объект документа Word
            classroom_id: ID аудитории
            day_of_week: день недели (если не указан, берется всё расписание)

        Returns:
            Обработанный документ Word
        """
        # Получаем аудиторию
        classroom = await Classroom.get(id=classroom_id)

        # Получаем расписание аудитории
        schedule_query = ScheduleItem.filter(classroom_id=classroom_id).prefetch_related(
            'time_slot', 'discipline', 'group', 'teacher'
        )

        # Если указан день недели, фильтруем по нему
        if day_of_week:
            schedule_query = schedule_query.filter(day_of_week=day_of_week)

        # Получаем отсортированное расписание
        schedule_items = await schedule_query.order_by('day_of_week', 'time_slot__number')

        # Группируем расписание по дням недели
        schedule_by_day = {}
        for item in schedule_items:
            day = item.day_of_week.value
            if day not in schedule_by_day:
                schedule_by_day[day] = []
            schedule_by_day[day].append(item)

        # Поиск таблицы для расписания в документе
        schedule_table = None
        for table in document.tables:
            # Ищем таблицу, в которой есть нужные заголовки
            first_row_text = ' '.join([cell.text for cell in table.rows[0].cells])
            if "День" in first_row_text and "Время" in first_row_text:
                schedule_table = table
                break

        # Если нашли таблицу расписания, заполняем её
        if schedule_table:
            # Определяем количество строк в таблице (исключая заголовок)
            num_rows = len(schedule_table.rows)

            # Если дней больше, чем строк (минус заголовок), добавляем строки
            days_count = len(schedule_by_day)
            if days_count > num_rows - 1:
                for _ in range(days_count - (num_rows - 1)):
                    schedule_table.add_row()

            # Заполняем таблицу данными
            row_index = 1  # Начинаем с первой строки после заголовка
            for day, items in schedule_by_day.items():
                if row_index < len(schedule_table.rows):
                    row = schedule_table.rows[row_index]

                    # Заполняем день недели
                    if len(row.cells) > 0:
                        row.cells[0].text = day

                    # Формируем расписание на день
                    schedule_text = ""
                    for item in items:
                        time_slot = f"{item.time_slot.start_time}-{item.time_slot.end_time}"
                        discipline = item.discipline.name
                        group = item.group.code
                        teacher = item.teacher.full_name

                        schedule_text += f"{time_slot}: {discipline}, {group}, {teacher}\n"

                    # Заполняем расписание
                    if len(row.cells) > 1:
                        row.cells[1].text = schedule_text.strip()

                    row_index += 1

        # Обновляем информацию об аудитории в документе
        for paragraph in document.paragraphs:
            if "Загруженность аудитории:" in paragraph.text:
                paragraph.text = f"Загруженность аудитории: {classroom.name}"
                break

        # Если расписание пустое, добавляем сообщение
        if not schedule_items:
            for paragraph in document.paragraphs:
                if "Загруженность" in paragraph.text:
                    document.add_paragraph("В аудитории нет занятий в расписании.")
                    break

        return document

    async def process_document(self, document: Document, params: Dict[str, Any]) -> Document:
        """
        Обрабатывает DOCX-документ, заменяя плейсхолдеры на данные

        Args:
            document: Объект документа Word
            params: Параметры для обработки документа

        Returns:
            Обработанный документ Word
        """
        self.replacements = {}

        # Извлекаем параметры
        template_id = params.get('template_id')
        group_id = params.get('group_id')
        student_id = params.get('student_id')
        teacher_id = params.get('teacher_id')
        discipline_id = params.get('discipline_id')
        ticket_number = params.get('ticket_number')
        day_of_week = params.get('day_of_week')
        classroom_id = params.get('classroom_id')

        # Получаем шаблон для определения его типа
        template = await self.get_template(template_id)

        # Определяем тип документа
        is_exam_ticket = False
        is_teacher_schedule = False
        is_classroom_schedule = False

        if "билет" in template.name.lower() or "ticket" in template.name.lower():
            is_exam_ticket = True
        elif "расписание_преподавателя" in template.name.lower() or "teacher_schedule" in template.name.lower():
            is_teacher_schedule = True
        elif "загруженность_аудитории" in template.name.lower() or "classroom_schedule" in template.name.lower():
            is_classroom_schedule = True

        # Собираем данные для обычной замены плейсхолдеров
        # Данные группы
        if group_id:
            group = await Group.get(id=group_id).prefetch_related('students')
            self.replacements.update({
                "group": group.code,
                "course": group.course,
                "students": "\n".join([f"{student.full_name} ({student.email})" for student in group.students]),
                "S": group.code  # Для одиночного плейсхолдера
            })

        # Данные студента
        if student_id:
            student = await Student.get(id=student_id).prefetch_related('group')
            self.replacements.update({
                "student_name": student.full_name,
                "student_email": student.email,
                "G": student.full_name,  # Для шаблона
                "M": student.full_name  # Для одиночного плейсхолдера
            })

            if student.group:
                self.replacements.update({
                    "group": student.group.code,
                    "S": student.group.code,  # Для одиночного плейсхолдера
                })

        # Данные преподавателя
        if teacher_id:
            teacher = await Teacher.get(id=teacher_id)
            self.replacements.update({
                "teacher_name": teacher.full_name,
                "teacher_email": teacher.email,
                "P": teacher.full_name,  # Для шаблона
                "B": teacher.full_name  # Для одиночного плейсхолдера
            })

        # Данные дисциплины
        discipline = None
        highlights_replacements = None

        if discipline_id:
            discipline = await Discipline.get(id=discipline_id)
            self.replacements.update({
                "discipline": discipline.name,
                "N": discipline.name,  # Для одиночного плейсхолдера
            })

            # Данные для желтых выделений
            highlights_replacements = {
                "discipline_name": discipline.name,
                "institute": "ИПИ",  # Можно получать из модели или параметров
                "department": f"БК №{discipline.department}" if discipline.department else "БК №536",
                "education_level": discipline.education_level if discipline.education_level else "бакалавриат",
                "hours": f"{discipline.hours_lecture}/{discipline.hours_practice}/{discipline.hours_lab}"
            }

            # Данные контрольных работ, если они есть
            control_works = await ControlWork.filter(discipline_id=discipline_id).order_by('number')
            if control_works:
                highlights_replacements["control_works"] = [
                    {
                        "name": f"Контрольная работа №{work.number}" if work.number < 4 else "Тестирование",
                        "format": work.format,
                        "week": work.week,
                        "max_score": work.max_score
                    }
                    for work in control_works[:4]  # Берем только первые 4 работы для БРС
                ]

        # Выполняем специальную обработку в зависимости от типа документа
        if is_exam_ticket and discipline_id:
            # Специальная обработка для экзаменационного билета
            document = await self.process_exam_ticket(document, discipline, ticket_number)
        elif is_teacher_schedule and teacher_id:
            # Специальная обработка для расписания преподавателя
            document = await self.process_teacher_schedule(document, teacher_id, day_of_week)
        elif is_classroom_schedule and classroom_id:
            # Специальная обработка для загруженности аудитории
            document = await self.process_classroom_schedule(document, classroom_id, day_of_week)
        else:
            # Стандартная обработка для других типов документов
            # Обрабатываем текст в параграфах
            for paragraph in document.paragraphs:
                if any(placeholder in paragraph.text for placeholder in self.replacements.keys()) or \
                        any("{{" + placeholder + "}}" in paragraph.text for placeholder in self.replacements.keys()):
                    # Сохраняем выравнивание
                    alignment = paragraph.alignment

                    # Получаем текст с замененными плейсхолдерами
                    new_text = await self.replace_placeholders_in_text(paragraph.text, self.replacements)

                    # Заменяем текст параграфа
                    paragraph.clear()
                    run = paragraph.add_run(new_text)

                    # Устанавливаем шрифт Times New Roman
                    run.font.name = "Times New Roman"

                    # Восстанавливаем выравнивание
                    paragraph.alignment = alignment

            # Обрабатываем текст в таблицах (если есть)
            for table in document.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            if any(placeholder in paragraph.text for placeholder in self.replacements.keys()) or \
                                    any("{{" + placeholder + "}}" in paragraph.text for placeholder in
                                        self.replacements.keys()):
                                # Сохраняем выравнивание
                                alignment = paragraph.alignment

                                # Получаем текст с замененными плейсхолдерами
                                new_text = await self.replace_placeholders_in_text(paragraph.text, self.replacements)

                                # Заменяем текст параграфа
                                paragraph.clear()
                                run = paragraph.add_run(new_text)

                                # Устанавливаем шрифт Times New Roman
                                run.font.name = "Times New Roman"

                                # Восстанавливаем выравнивание
                                paragraph.alignment = alignment

            # Если есть данные дисциплины, обрабатываем желтые выделения
            if discipline_id and highlights_replacements:
                try:
                    # Проверяем, является ли шаблон БРС документом
                    if "бр" in template.name.lower() or "рейтинг" in template.name.lower():
                        await self.replace_yellow_highlights(document, highlights_replacements)
                except Exception as e:
                    print(f"Ошибка при обработке желтых выделений: {str(e)}")

        # Применяем Times New Roman ко всему документу
        await self.apply_times_new_roman(document)

        return document
