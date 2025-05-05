from fastapi import HTTPException
from typing import Dict, Any, Optional
from datetime import datetime

from services.base_document_service import DocumentServiceFactory
from models.models import Template


class ReportService:
    """
    Сервис для генерации отчетов
    """

    @staticmethod
    async def generate_report(
            template_id: int,
            group_id: Optional[int] = None,
            student_id: Optional[int] = None,
            teacher_id: Optional[int] = None,
            discipline_id: Optional[int] = None,
            classroom_id: Optional[int] = None,
            start_date: Optional[datetime] = None,
            ticket_number: Optional[int] = None,
            day_of_week: Optional[str] = None,
            group_id_2: Optional[int] = None
    ) -> Dict[str, Any]:
        """
        Генерирует отчет на основе шаблона и предоставленных данных

        Args:
            template_id: ID шаблона
            group_id: ID группы (опционально)
            student_id: ID студента (опционально)
            teacher_id: ID преподавателя (опционально)
            discipline_id: ID дисциплины (опционально)
            classroom_id: ID аудитории (опционально)
            start_date: Начальная дата для журнала (опционально)
            ticket_number: Номер билета для экзаменационных билетов (опционально)
            day_of_week: День недели для расписания (опционально)
            group_id_2: ID второй группы для загруженности аудиторий (опционально)

        Returns:
            Dict с информацией о сгенерированном файле
        """
        try:
            template = await Template.get(id=template_id)

            params = {
                'template_id': template_id,
                'group_id': group_id,
                'student_id': student_id,
                'teacher_id': teacher_id,
                'discipline_id': discipline_id,
                'classroom_id': classroom_id,
                'start_date': start_date,
                'ticket_number': ticket_number,
                'day_of_week': day_of_week,
                'group_id_2': group_id_2
            }

            service = await DocumentServiceFactory.get_service(template_id)
            file_path = await service.generate_document(template_id, params)

            return {
                "message": "Отчет успешно сгенерирован и сохранён!",
                "file_path": file_path,
                "file_type": template.file_type
            }

        except FileNotFoundError as e:
            raise HTTPException(
                status_code=404,
                detail=str(e)
            )

        except ValueError as e:
            raise HTTPException(
                status_code=400,
                detail=str(e)
            )

        except Exception as e:
            raise HTTPException(
                status_code=500,
                detail=f"Ошибка при генерации отчета: {str(e)}"
            )
