import os
import re
from abc import ABC, abstractmethod
from typing import Dict, Any

from core.config import settings
from models.models import Template


class BaseDocumentService(ABC):
    """
    Базовый класс для сервисов генерации документов.
    Определяет общие методы и интерфейс для всех типов документов.
    """

    def __init__(self):
        """Инициализация сервиса"""
        self.replacements = {}

    async def get_template(self, template_id: int) -> Template:
        """
        Получает шаблон из базы данных по ID

        Args:
            template_id: ID шаблона

        Returns:
            Объект шаблона

        Raises:
            FileNotFoundError: Если файл шаблона не найден
        """
        template = await Template.get(id=template_id)
        template_path = os.path.join(settings.TEMPLATE_DIR, template.file_path)

        if not os.path.exists(template_path):
            raise FileNotFoundError(f"Шаблон {template_path} не найден")

        return template

    async def generate_output_filename(self, template: Template, params: Dict[str, Any]) -> str:
        """
        Генерирует имя выходного файла на основе шаблона и параметров

        Args:
            template: Объект шаблона
            params: Словарь параметров (group_id, student_id, и т.д.)

        Returns:
            Имя выходного файла с расширением
        """
        output_filename = f"{template.name}"

        for key, value in params.items():
            if value and key != 'template_id' and key != 'download':
                output_filename += f"_{key}_{value}"

        output_filename += f".{template.file_type}"
        return output_filename

    async def replace_placeholders_in_text(self, text: str, replacements: Dict[str, str]) -> str:
        """
        Заменяет все плейсхолдеры в тексте на соответствующие значения

        Args:
            text: Исходный текст с плейсхолдерами
            replacements: Словарь замен {placeholder: value}

        Returns:
            Текст с замененными плейсхолдерами
        """
        if not text or not isinstance(text, str):
            return text

        pattern = r"\{\{([^}]+)\}\}"

        def replace_match(match):
            placeholder = match.group(1).strip()
            return str(replacements.get(placeholder, match.group(0)))

        text = re.sub(pattern, replace_match, text)

        single_chars = {'S', 'M', 'B', 'G', 'P', 'N'}
        for char in single_chars:
            if char in replacements:
                text = re.sub(r'\b' + char + r'\b', str(replacements[char]), text)

        return text

    @abstractmethod
    async def load_document(self, template_path: str) -> Any:
        """
        Загружает документ из файла шаблона

        Args:
            template_path: Путь к файлу шаблона

        Returns:
            Загруженный документ
        """
        pass

    @abstractmethod
    async def save_document(self, document: Any, output_path: str) -> None:
        """
        Сохраняет документ в файл

        Args:
            document: Объект документа
            output_path: Путь для сохранения
        """
        pass

    @abstractmethod
    async def process_document(self, document: Any, params: Dict[str, Any]) -> Any:
        """
        Обрабатывает документ, заменяя плейсхолдеры на данные

        Args:
            document: Объект документа
            params: Параметры для обработки документа

        Returns:
            Обработанный документ
        """
        pass

    async def generate_document(self, template_id: int, params: Dict[str, Any]) -> str:
        """
        Основной метод генерации документа

        Args:
            template_id: ID шаблона
            params: Параметры для генерации документа

        Returns:
            Путь к сгенерированному файлу
        """
        # Получаем шаблон
        template = await self.get_template(template_id)
        template_path = os.path.join(settings.TEMPLATE_DIR, template.file_path)

        # Загружаем документ
        document = await self.load_document(template_path)

        # Обрабатываем документ
        document = await self.process_document(document, params)

        # Создаем директорию для выходных файлов, если она не существует
        os.makedirs(settings.OUTPUT_DIR, exist_ok=True)

        # Генерируем имя выходного файла
        output_filename = await self.generate_output_filename(template, params)
        output_path = os.path.join(settings.OUTPUT_DIR, output_filename)

        # Сохраняем результат
        await self.save_document(document, output_path)

        return output_path


class DocumentServiceFactory:
    """
    Фабрика для создания сервисов обработки документов в зависимости от типа файла
    """

    @staticmethod
    async def get_service(template_id: int) -> BaseDocumentService:
        """
        Возвращает сервис для обработки документа в зависимости от типа файла

        Args:
            template_id: ID шаблона

        Returns:
            Экземпляр соответствующего сервиса

        Raises:
            ValueError: Если тип файла не поддерживается
        """
        from services.docx_service import DocxService
        from services.xlsx_service import XlsxService

        template = await Template.get(id=template_id)

        if template.file_type.lower() == 'docx':
            return DocxService()
        elif template.file_type.lower() == 'xlsx':
            return XlsxService()
        else:
            raise ValueError(f"Неподдерживаемый тип файла: {template.file_type}")