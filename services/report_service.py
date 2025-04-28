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
            day_of_week: Optional[str] = None
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

        Returns:
            Dict с информацией о сгенерированном файле
        """
        try:
            # Получаем шаблон
            template = await Template.get(id=template_id)

            # Подготавливаем параметры для генерации документа
            params = {
                'template_id': template_id,
                'group_id': group_id,
                'student_id': student_id,
                'teacher_id': teacher_id,
                'discipline_id': discipline_id,
                'classroom_id': classroom_id,
                'start_date': start_date,
                'ticket_number': ticket_number,
                'day_of_week': day_of_week
            }

            # Используем фабрику для получения нужного сервиса
            service = await DocumentServiceFactory.get_service(template_id)

            # Генерируем документ
            file_path = await service.generate_document(template_id, params)

            # Возвращаем информацию о созданном файле
            return {
                "message": "Отчет успешно сгенерирован и сохранён!",
                "file_path": file_path,
                "file_type": template.file_type
            }

        except Template.DoesNotExist:
            raise HTTPException(
                status_code=404,
                detail=f"Шаблон с ID {template_id} не найден"
            )

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
