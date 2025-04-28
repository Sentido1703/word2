import os
from pydantic_settings import BaseSettings


class Settings(BaseSettings):
    """Настройки приложения"""
    DATABASE_URL: str = os.getenv("DATABASE_URL", "postgres://postgres:12345678@localhost:5432/sentido")

    TEMPLATE_DIR: str = os.getenv("TEMPLATE_DIR", "./templates")
    OUTPUT_DIR: str = os.getenv("OUTPUT_DIR", "./output")

    APP_NAME: str = "Document Service"
    APP_VERSION: str = "1.0.0"
    APP_DESCRIPTION: str = "Сервис для генерации документов на основе шаблонов"

    CORS_ORIGINS: list = ["*"]
    CORS_ALLOW_CREDENTIALS: bool = True
    CORS_ALLOW_METHODS: list = ["*"]
    CORS_ALLOW_HEADERS: list = ["*"]

    class Config:
        env_file = ".env"


settings = Settings()
