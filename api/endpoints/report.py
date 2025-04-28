from fastapi import APIRouter, Query, HTTPException
from fastapi.responses import FileResponse
import os
from datetime import datetime
from typing import Optional, List

from models.models import Group, DayOfWeek, Classroom, Teacher, Template
from services.report_service import ReportService

router = APIRouter()


@router.get("/officesvc/report")
async def create_report(
        template_id: int = Query(..., description="ID шаблона"),
        group_id: Optional[int] = Query(None, description="ID группы"),
        student_id: Optional[int] = Query(None, description="ID студента"),
        teacher_id: Optional[int] = Query(None, description="ID преподавателя"),
        discipline_id: Optional[int] = Query(None, description="ID дисциплины"),
        classroom_id: Optional[int] = Query(None, description="ID аудитории"),
        start_date: Optional[str] = Query(None, description="Начальная дата для журнала (формат: YYYY-MM-DD)"),
        ticket_number: Optional[int] = Query(None, description="Номер билета (для экзаменационных билетов)"),
        day_of_week: Optional[str] = Query(None, description="День недели (для расписания)"),
        download: bool = Query(False, description="Скачать файл напрямую")
):
    """
    Создает отчет на основе шаблона и предоставленных данных

    Примеры запросов:
    - /officesvc/report?template_id=1&group_id=1
    - /officesvc/report?template_id=2&student_id=1&teacher_id=1
    - /officesvc/report?template_id=3&group_id=1&discipline_id=1&start_date=2025-04-01
    - /officesvc/report?template_id=4&teacher_id=1&day_of_week=Понедельник
    - /officesvc/report?template_id=5&classroom_id=1&day_of_week=Вторник
    """
    # Проверяем наличие обязательных параметров
    if not template_id:
        raise HTTPException(
            status_code=400,
            detail="Не указан параметр 'template_id'"
        )

    # Проверяем валидность day_of_week
    if day_of_week and day_of_week not in [e.value for e in DayOfWeek]:
        raise HTTPException(
            status_code=400,
            detail=f"Неверный день недели. Допустимые значения: {', '.join([e.value for e in DayOfWeek])}"
        )

    # Преобразуем start_date в объект datetime, если указан
    parsed_start_date = None
    if start_date:
        try:
            parsed_start_date = datetime.strptime(start_date, "%Y-%m-%d")
        except ValueError:
            raise HTTPException(
                status_code=400,
                detail="Неверный формат даты. Используйте формат YYYY-MM-DD"
            )

    # Генерируем отчет
    result = await ReportService.generate_report(
        template_id=template_id,
        group_id=group_id,
        student_id=student_id,
        teacher_id=teacher_id,
        discipline_id=discipline_id,
        classroom_id=classroom_id,
        start_date=parsed_start_date,
        ticket_number=ticket_number,
        day_of_week=day_of_week
    )

    # Если запрошено скачивание, возвращаем файл
    if download and os.path.exists(result["file_path"]):
        return FileResponse(
            path=result["file_path"],
            filename=os.path.basename(result["file_path"]),
            media_type=f"application/{'vnd.openxmlformats-officedocument.spreadsheetml.sheet' if result['file_type'] == 'xlsx' else 'vnd.openxmlformats-officedocument.wordprocessingml.document'}"
        )

    # Иначе возвращаем информацию о файле
    return result


@router.get("/officesvc/templates", response_model=List[dict])
async def list_templates():
    """Возвращает список доступных шаблонов"""
    templates = await Template.all()
    return [
        {
            "id": template.id,
            "name": template.name,
            "file_type": template.file_type,
            "description": template.description
        }
        for template in templates
    ]


@router.get("/officesvc/teachers", response_model=List[dict])
async def list_teachers():
    """Возвращает список преподавателей"""
    teachers = await Teacher.all()
    return [
        {
            "id": teacher.id,
            "full_name": teacher.full_name,
            "email": teacher.email
        }
        for teacher in teachers
    ]


@router.get("/officesvc/classrooms", response_model=List[dict])
async def list_classrooms():
    """Возвращает список аудиторий"""
    classrooms = await Classroom.all()
    return [
        {
            "id": classroom.id,
            "name": classroom.name,
            "capacity": classroom.capacity
        }
        for classroom in classrooms
    ]


@router.get("/officesvc/days", response_model=List[str])
async def list_days():
    """Возвращает список доступных дней недели"""
    return [day.value for day in DayOfWeek]


@router.get("/officesvc/groups", response_model=List[dict])
async def list_groups():
    """Возвращает список доступных групп"""
    groups = await Group.all()
    return [
        {
            "id": group.id,
            "code": group.code,
            "course": group.course
        }
        for group in groups
    ]


@router.get("/officesvc/students")
async def list_students(group_id: Optional[int] = Query(None, description="ID группы для фильтрации")):
    """Возвращает список студентов, опционально фильтруя по группе"""
    from models.models import Student

    if group_id:
        students = await Student.filter(group_id=group_id)
    else:
        students = await Student.all()

    return {
        "students": [
            {
                "id": student.id,
                "full_name": student.full_name,
                "email": student.email,
                "group_id": student.group_id
            }
            for student in students
        ]
    }


@router.get("/officesvc/teachers")
async def list_teachers():
    """Возвращает список преподавателей"""
    from models.models import Teacher

    teachers = await Teacher.all()
    return {
        "teachers": [
            {
                "id": teacher.id,
                "full_name": teacher.full_name,
                "email": teacher.email
            }
            for teacher in teachers
        ]
    }


@router.get("/officesvc/disciplines")
async def list_disciplines():
    """Возвращает список дисциплин"""
    from models.models import Discipline

    disciplines = await Discipline.all()
    return {
        "disciplines": [
            {
                "id": discipline.id,
                "name": discipline.name,
                "education_level": discipline.education_level,
                "department": discipline.department,
                "hours_lecture": discipline.hours_lecture,
                "hours_practice": discipline.hours_practice,
                "hours_lab": discipline.hours_lab
            }
            for discipline in disciplines
        ]
    }