import os
import re
from typing import Dict, Any, List, Optional
from datetime import datetime, timedelta
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.styles.colors import Color

from services.base_document_service import BaseDocumentService
from models.models import (
    Group, Student, Teacher, Discipline, Grade, ControlWork,
    ScheduleItem, TimeSlot, Classroom, DayOfWeek
)
from core.config import settings


class XlsxService(BaseDocumentService):
    """
    Сервис для обработки XLSX-документов
    """

    async def load_document(self, template_path: str) -> openpyxl.Workbook:
        """
        Загружает XLSX-документ из файла шаблона

        Args:
            template_path: Путь к файлу шаблона

        Returns:
            Объект книги Excel
        """
        return openpyxl.load_workbook(template_path)

    async def save_document(self, document: openpyxl.Workbook, output_path: str) -> None:
        """
        Сохраняет XLSX-документ в файл

        Args:
            document: Объект книги Excel
            output_path: Путь для сохранения
        """
        document.save(output_path)

    async def fill_grades_journal(self, worksheet: openpyxl.worksheet.worksheet.Worksheet, group_id: int,
                                  discipline_id: int, start_date: Optional[datetime] = None,
                                  period_weeks: int = 16) -> None:
        """
        Заполняет журнал оценок для группы по дисциплине

        Args:
            worksheet: Лист Excel
            group_id: ID группы
            discipline_id: ID дисциплины
            start_date: Начальная дата (по умолчанию текущая)
            period_weeks: Количество недель (по умолчанию 16)
        """
        # Получаем группу и студентов
        group = await Group.get(id=group_id).prefetch_related('students')

        # Получаем дисциплину и контрольные работы
        discipline = await Discipline.get(id=discipline_id).prefetch_related('control_works')

        # Определяем начальную дату, если не указана
        if start_date is None:
            start_date = datetime.now()

        # Находим ячейки с N (имена студентов) и d (даты)
        student_cells = []
        date_cells = []

        # Сначала получаем все ячейки со значением точно "N" или "d"
        for row in range(1, worksheet.max_row + 1):
            for col in range(1, worksheet.max_column + 1):
                cell = worksheet.cell(row=row, column=col)
                if cell.value == 'N':
                    student_cells.append((row, col))
                elif cell.value == 'd':
                    date_cells.append((row, col))

        # Если не найдены ячейки с маркерами, ищем в первой строке для дат
        # и в первом столбце для студентов
        if not date_cells and not student_cells:
            # Проверяем первую строку на наличие повторяющихся элементов (возможно дат)
            header_row = next(
                (i for i in range(1, 11) if all(worksheet.cell(row=i, column=j).value for j in range(2, 6))), 1)

            # Проверяем в каком столбце начинаются имена студентов
            student_col = 1
            for row in range(header_row + 1, min(header_row + 20, worksheet.max_row)):
                if worksheet.cell(row=row, column=student_col).value:
                    student_cells.append((row, student_col))

            # Добавляем ячейки для дат в первой строке
            for col in range(2, min(worksheet.max_column, 20)):
                date_cells.append((header_row, col))

        # Если не найдены ячейки, используем стандартную структуру журнала
        if not student_cells:
            # Берем первый столбец начиная с третьей строки как места для студентов
            for row in range(3, min(3 + len(group.students), 30)):
                student_cells.append((row, 1))

        if not date_cells:
            # Берем первую строку начиная со второго столбца для дат
            for col in range(2, min(2 + period_weeks, 20)):
                date_cells.append((1, col))

        # Заполняем заголовок журнала
        if worksheet.cell(row=1, column=1).value:
            # Устанавливаем название дисциплины и группы в верхней части
            worksheet.cell(row=1, column=1).value = f"Журнал оценок - {discipline.name} - Группа {group.code}"

        # Заполняем имена студентов
        for i, student in enumerate(group.students):
            if i < len(student_cells):
                row, col = student_cells[i]
                cell = worksheet.cell(row=row, column=col)
                cell.value = student.full_name
                # Применяем форматирование: полужирный шрифт
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal='left', vertical='center')

        # Заполняем даты
        current_date = start_date
        for i in range(min(period_weeks, len(date_cells))):
            if i < len(date_cells):
                row, col = date_cells[i]
                date_str = current_date.strftime("%d.%m")
                cell = worksheet.cell(row=row, column=col)
                cell.value = date_str
                # Применяем форматирование: полужирный шрифт и выравнивание по центру
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal='center', vertical='center')
                # Задаем ширину столбца
                worksheet.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 10
                current_date += timedelta(days=7)  # Увеличиваем на неделю

        # Заполняем оценки (если есть)
        for i, student in enumerate(group.students):
            if i < len(student_cells):
                student_row, student_col = student_cells[i]

                # Получаем оценки студента
                grades = await Grade.filter(student_id=student.id).prefetch_related('control_work')

                for grade in grades:
                    # Находим контрольную работу
                    control_work = grade.control_work

                    # Проверяем, относится ли контрольная работа к текущей дисциплине
                    if control_work.discipline_id == discipline_id:
                        # Находим соответствующую колонку для недели контрольной работы
                        week = control_work.week - 1  # Индексация с 0

                        if week < len(date_cells):
                            _, date_col = date_cells[week]
                            cell = worksheet.cell(row=student_row, column=date_col)
                            cell.value = grade.score
                            # Выравнивание по центру
                            cell.alignment = Alignment(horizontal='center', vertical='center')

        # Добавляем форматирование для пустых ячеек в таблице
        for i in range(len(student_cells)):
            if i < len(group.students):
                student_row, _ = student_cells[i]
                for j in range(len(date_cells)):
                    _, date_col = date_cells[j]
                    cell = worksheet.cell(row=student_row, column=date_col)
                    if cell.value is None:
                        # Заполняем пустые ячейки дефисом или оставляем пустыми
                        cell.value = "-"
                        # Выравнивание по центру
                        cell.alignment = Alignment(horizontal='center', vertical='center')

    async def fill_student_list(self, worksheet: openpyxl.worksheet.worksheet.Worksheet, group_id: int) -> None:
        """
        Заполняет список студентов группы в Excel

        Args:
            worksheet: Лист Excel
            group_id: ID группы
        """
        # Получаем группу и студентов
        group = await Group.get(id=group_id).prefetch_related('students')

        # Сначала полностью обработаем все ячейки с маркерами
        for row in range(1, worksheet.max_row + 1):
            for col in range(1, worksheet.max_column + 1):
                cell = worksheet.cell(row=row, column=col)
                cell_value = cell.value

                # Заменяем маркеры в заголовках и прочих местах
                if cell_value and isinstance(cell_value, str):
                    # Заменяем КМБО-05 S на код группы
                    if 'S' in cell_value:
                        cell.value = cell_value.replace('S', group.code)

                    # Если ячейка содержит только M или N (маркеры для данных студентов),
                    # оставляем их для последующей обработки
                    if cell_value not in ['M', 'N']:
                        continue

        # Находим строки с маркерами M и N для вставки данных студентов
        data_rows = []
        student_col = None
        group_col = None
        email_col = None

        # Ищем расположение колонок для данных студентов
        for row in range(1, worksheet.max_row + 1):
            row_has_markers = False

            for col in range(1, worksheet.max_column + 1):
                cell_value = worksheet.cell(row=row, column=col).value

                if cell_value == 'M':
                    student_col = col
                    row_has_markers = True
                elif cell_value == 'N':
                    group_col = col
                    row_has_markers = True
                # Ищем колонку для Email по заголовку
                elif isinstance(cell_value, str) and 'email' in cell_value.lower():
                    email_col = col

            if row_has_markers:
                data_rows.append(row)

        # Если не нашли маркеры, ищем по заголовкам
        if not data_rows:
            header_row = None

            for row in range(1, min(5, worksheet.max_row + 1)):
                for col in range(1, worksheet.max_column + 1):
                    cell_value = worksheet.cell(row=row, column=col).value

                    if isinstance(cell_value, str):
                        if 'студент' in cell_value.lower():
                            student_col = col
                            header_row = row
                        elif 'группа' in cell_value.lower():
                            group_col = col
                            header_row = row
                        elif 'email' in cell_value.lower():
                            email_col = col
                            header_row = row

                if header_row:
                    # Данные начинаются со следующей строки после заголовка
                    for i in range(len(group.students)):
                        data_rows.append(header_row + 1 + i)
                    break

        # Если все еще не нашли, используем стандартную структуру
        if not data_rows:
            start_row = 2  # Начинаем со второй строки
            student_col = student_col or 2  # Колонка B
            group_col = group_col or 1  # Колонка A
            email_col = email_col or 3  # Колонка C

            # Создаем строки для данных
            for i in range(len(group.students)):
                data_rows.append(start_row + i)

        # Теперь заполняем данные студентов
        for i, student in enumerate(group.students):
            if i < len(data_rows):
                row = data_rows[i]

                # Заполняем имя студента
                if student_col:
                    cell = worksheet.cell(row=row, column=student_col)
                    # Если ячейка содержит маркер M, заменяем его
                    if cell.value == 'M':
                        cell.value = student.full_name
                    # Иначе просто вставляем имя
                    else:
                        cell.value = student.full_name
                    cell.alignment = Alignment(horizontal='left', vertical='center')

                # Заполняем группу
                if group_col:
                    cell = worksheet.cell(row=row, column=group_col)
                    # Если ячейка содержит маркер N, заменяем его
                    if cell.value == 'N':
                        cell.value = group.code
                    # Иначе просто вставляем код группы
                    else:
                        cell.value = group.code
                    cell.alignment = Alignment(horizontal='left', vertical='center')

                # Заполняем email
                if email_col:
                    cell = worksheet.cell(row=row, column=email_col)
                    cell.value = student.email
                    cell.alignment = Alignment(horizontal='left', vertical='center')

    async def process_teacher_schedule_xlsx(self, document: openpyxl.Workbook, teacher_id: int,
                                            day_of_week: Optional[str] = None) -> openpyxl.Workbook:
        """
        Специальная обработка для расписания преподавателя в Excel

        Args:
            document: объект книги Excel
            teacher_id: ID преподавателя
            day_of_week: день недели (если не указан, берется всё расписание)

        Returns:
            Обработанный документ Excel
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

        # Сгруппируем расписание по дням недели
        schedule_by_day = {}
        for item in schedule_items:
            day = item.day_of_week.value
            if day not in schedule_by_day:
                schedule_by_day[day] = []
            schedule_by_day[day].append(item)

        # Берем первый лист (или создаем новый, если нет листов)
        if len(document.worksheets) == 0:
            worksheet = document.create_sheet("Расписание")
        else:
            worksheet = document.worksheets[0]

        # Очищаем лист
        for row in range(1, worksheet.max_row + 1):
            for col in range(1, worksheet.max_column + 1):
                worksheet.cell(row=row, column=col).value = None

        # Добавляем заголовок
        worksheet.cell(row=1, column=1).value = f"Расписание преподавателя: {teacher.full_name}"
        worksheet.cell(row=1, column=1).font = Font(bold=True, size=14)
        worksheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=5)

        # Добавляем заголовки таблицы
        headers = ["День недели", "№ пары", "Время", "Дисциплина", "Группа", "Аудитория"]
        for i, header in enumerate(headers):
            cell = worksheet.cell(row=3, column=i + 1)
            cell.value = header
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='center', vertical='center')
            # Устанавливаем ширину столбца
            worksheet.column_dimensions[openpyxl.utils.get_column_letter(i + 1)].width = 15

        # Заполняем таблицу данными
        row_index = 4  # Начинаем с 4-й строки (после заголовков)

        if schedule_items:
            for day, items in schedule_by_day.items():
                # Добавляем день недели
                worksheet.cell(row=row_index, column=1).value = day
                worksheet.cell(row=row_index, column=1).font = Font(bold=True)

                # Заполняем занятия на этот день
                for item in items:
                    worksheet.cell(row=row_index, column=2).value = item.time_slot.number
                    worksheet.cell(row=row_index,
                                   column=3).value = f"{item.time_slot.start_time}-{item.time_slot.end_time}"
                    worksheet.cell(row=row_index, column=4).value = item.discipline.name
                    worksheet.cell(row=row_index, column=5).value = item.group.code
                    worksheet.cell(row=row_index,
                                   column=6).value = item.classroom.name if item.classroom else "Нет аудитории"

                    # Выравнивание всех ячеек по центру
                    for col in range(1, 7):
                        worksheet.cell(row=row_index, column=col).alignment = Alignment(horizontal='center',
                                                                                        vertical='center')

                    row_index += 1

                # Добавляем пустую строку между днями
                row_index += 1
        else:
            # Если расписание пустое, добавляем сообщение
            worksheet.cell(row=4, column=1).value = "У преподавателя нет занятий в расписании."
            worksheet.merge_cells(start_row=4, start_column=1, end_row=4, end_column=6)

        return document

    async def process_classroom_schedule_xlsx(self, document: openpyxl.Workbook, classroom_id: int,
                                              day_of_week: Optional[str] = None) -> openpyxl.Workbook:
        """
        Специальная обработка для расписания аудитории в Excel

        Args:
            document: объект книги Excel
            classroom_id: ID аудитории
            day_of_week: день недели (если не указан, берется всё расписание)

        Returns:
            Обработанный документ Excel
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

        # Сгруппируем расписание по дням недели
        schedule_by_day = {}
        for item in schedule_items:
            day = item.day_of_week.value
            if day not in schedule_by_day:
                schedule_by_day[day] = []
            schedule_by_day[day].append(item)

        # Берем первый лист (или создаем новый, если нет листов)
        if len(document.worksheets) == 0:
            worksheet = document.create_sheet("Загруженность аудитории")
        else:
            worksheet = document.worksheets[0]

        # Очищаем лист
        for row in range(1, worksheet.max_row + 1):
            for col in range(1, worksheet.max_column + 1):
                worksheet.cell(row=row, column=col).value = None

        # Добавляем заголовок
        worksheet.cell(row=1, column=1).value = f"Загруженность аудитории: {classroom.name}"
        worksheet.cell(row=1, column=1).font = Font(bold=True, size=14)
        worksheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=5)

        # Добавляем заголовки таблицы
        headers = ["День недели", "№ пары", "Время", "Дисциплина", "Группа", "Преподаватель"]
        for i, header in enumerate(headers):
            cell = worksheet.cell(row=3, column=i + 1)
            cell.value = header
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='center', vertical='center')
            # Устанавливаем ширину столбца
            worksheet.column_dimensions[openpyxl.utils.get_column_letter(i + 1)].width = 15

        # Заполняем таблицу данными
        row_index = 4  # Начинаем с 4-й строки (после заголовков)

        if schedule_items:
            for day, items in schedule_by_day.items():
                # Добавляем день недели
                worksheet.cell(row=row_index, column=1).value = day
                worksheet.cell(row=row_index, column=1).font = Font(bold=True)

                # Заполняем занятия на этот день
                for item in items:
                    worksheet.cell(row=row_index, column=2).value = item.time_slot.number
                    worksheet.cell(row=row_index,
                                   column=3).value = f"{item.time_slot.start_time}-{item.time_slot.end_time}"
                    worksheet.cell(row=row_index, column=4).value = item.discipline.name
                    worksheet.cell(row=row_index, column=5).value = item.group.code
                    worksheet.cell(row=row_index, column=6).value = item.teacher.full_name

                    # Выравнивание всех ячеек по центру
                    for col in range(1, 7):
                        worksheet.cell(row=row_index, column=col).alignment = Alignment(horizontal='center',
                                                                                        vertical='center')

                    row_index += 1

                # Добавляем пустую строку между днями
                row_index += 1
        else:
            # Если расписание пустое, добавляем сообщение
            worksheet.cell(row=4, column=1).value = "В аудитории нет занятий в расписании."
            worksheet.merge_cells(start_row=4, start_column=1, end_row=4, end_column=6)

        return document

    async def process_document(self, document: openpyxl.Workbook, params: Dict[str, Any]) -> openpyxl.Workbook:
        """
        Обрабатывает XLSX-документ, заменяя плейсхолдеры на данные

        Args:
            document: Объект книги Excel
            params: Параметры для обработки документа

        Returns:
            Обработанный документ Excel
        """
        self.replacements = {}

        # Извлекаем параметры
        template_id = params.get('template_id')
        group_id = params.get('group_id')
        student_id = params.get('student_id')
        teacher_id = params.get('teacher_id')
        discipline_id = params.get('discipline_id')
        start_date_str = params.get('start_date')
        day_of_week = params.get('day_of_week')
        classroom_id = params.get('classroom_id')

        # Преобразуем start_date в объект datetime, если указан
        start_date = None
        if start_date_str:
            try:
                start_date = datetime.strptime(start_date_str, "%Y-%m-%d")
            except (ValueError, TypeError):
                start_date = datetime.now()

        # Получаем шаблон для определения его типа
        template = await self.get_template(template_id)

        # Определяем тип документа
        is_journal = False
        is_student_list = False
        is_teacher_schedule = False
        is_classroom_schedule = False

        # Проверяем по имени файла
        if 'журнал' in template.name.lower() or 'journal' in template.name.lower():
            is_journal = True
        elif 'список' in template.name.lower() or 'list' in template.name.lower() or 'студент' in template.name.lower():
            is_student_list = True
        elif 'расписание_преподавателя' in template.name.lower() or 'teacher_schedule' in template.name.lower():
            is_teacher_schedule = True
        elif 'загруженность_аудитории' in template.name.lower() or 'classroom_schedule' in template.name.lower():
            is_classroom_schedule = True

        # Если тип не определен по имени, проверяем содержимое
        if not is_journal and not is_student_list and not is_teacher_schedule and not is_classroom_schedule:
            # Проверяем наличие маркеров
            for sheet_name in document.sheetnames:
                worksheet = document[sheet_name]
                for row in range(1, min(10, worksheet.max_row + 1)):  # Проверяем только первые 10 строк
                    for col in range(1, min(10, worksheet.max_column + 1)):  # И первые 10 колонок
                        cell_value = worksheet.cell(row=row, column=col).value
                        if cell_value == 'd':
                            is_journal = True
                            break
                        elif cell_value in ['M', 'N']:
                            is_student_list = True
                            break

        # Собираем данные для обычной замены плейсхолдеров
        # Данные группы
        if group_id:
            group = await Group.get(id=group_id).prefetch_related('students')
            self.replacements.update({
                "group": group.code,
                "course": group.course,
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
        if discipline_id:
            discipline = await Discipline.get(id=discipline_id)
            self.replacements.update({
                "discipline": discipline.name,
                "N": discipline.name,  # Для одиночного плейсхолдера
            })

        # Применяем специальную обработку в зависимости от типа документа
        if is_teacher_schedule and teacher_id:
            # Специальная обработка для расписания преподавателя
            document = await self.process_teacher_schedule_xlsx(document, teacher_id, day_of_week)
        elif is_classroom_schedule and classroom_id:
            # Специальная обработка для загруженности аудитории
            document = await self.process_classroom_schedule_xlsx(document, classroom_id, day_of_week)
        else:
            # Стандартная обработка
            for sheet_name in document.sheetnames:
                worksheet = document[sheet_name]

                # Выполняем специализированную обработку в зависимости от типа файла
                if is_journal and group_id and discipline_id:
                    # Обрабатываем журнал оценок
                    await self.fill_grades_journal(worksheet, group_id, discipline_id, start_date)
                elif is_student_list and group_id:
                    # Обрабатываем список студентов
                    await self.fill_student_list(worksheet, group_id)
                else:
                    # Стандартная обработка для других типов документов
                    # Сначала делаем общие замены для всех типов файлов
                    for row in range(1, worksheet.max_row + 1):
                        for col in range(1, worksheet.max_column + 1):
                            cell = worksheet.cell(row=row, column=col)

                            if cell.value and isinstance(cell.value, str) and cell.value not in ['M', 'N', 'd']:
                                # Проверяем, есть ли в ячейке плейсхолдеры
                                if any(placeholder in cell.value for placeholder in self.replacements.keys()) or \
                                        any("{{" + placeholder + "}}" in cell.value for placeholder in
                                            self.replacements.keys()):
                                    # Заменяем плейсхолдеры
                                    cell.value = await self.replace_placeholders_in_text(cell.value, self.replacements)

        return document