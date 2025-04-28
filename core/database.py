from tortoise import Tortoise
from core.config import settings

TORTOISE_ORM = {
    "connections": {"default": settings.DATABASE_URL},
    "apps": {
        "models": {
            "models": ["models.models", "aerich.models"],
            "default_connection": "default",
        },
    },
    "use_tz": False,
}


async def init_db():
    """Инициализация базы данных"""
    await Tortoise.init(
        db_url=settings.DATABASE_URL,
        modules={"models": ["models.models"]},
    )
    await Tortoise.generate_schemas()


async def close_db():
    """Закрытие соединения с базой данных"""
    await Tortoise.close_connections()
    