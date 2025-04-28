from contextlib import asynccontextmanager

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
import os
from fastadmin import fastapi_app as admin_app

from core.config import settings
from core.database import TORTOISE_ORM
from api.endpoints.report import router as report_router
from models.models import AdminUser

# Создаем директории, если они не существуют
os.makedirs(settings.TEMPLATE_DIR, exist_ok=True)
os.makedirs(settings.OUTPUT_DIR, exist_ok=True)

@asynccontextmanager
async def lifespan(app: FastAPI):
    os.makedirs("static", exist_ok=True)
    from tortoise import Tortoise

    await Tortoise.init(config=TORTOISE_ORM)

    user = await AdminUser.get_or_none(username='admin')

    if not user:
        user = await AdminUser.create(
            username='admin',
            is_superuser=True,
            password='adminpassword'
        )

        user.set_password(user.password.encode())
        await user.save()

    yield

    await Tortoise.close_connections()


# Создаем приложение FastAPI
app = FastAPI(
    title=settings.APP_NAME,
    version=settings.APP_VERSION,
    description=settings.APP_DESCRIPTION,
    lifespan=lifespan
)

# Добавляем CORS middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=settings.CORS_ORIGINS,
    allow_credentials=settings.CORS_ALLOW_CREDENTIALS,
    allow_methods=settings.CORS_ALLOW_METHODS,
    allow_headers=settings.CORS_ALLOW_HEADERS,
)

# Настраиваем статические файлы для доступа к сгенерированным документам
app.mount("/output", StaticFiles(directory=settings.OUTPUT_DIR), name="output")
app.mount("/admin", admin_app)

# Регистрируем роутеры
app.include_router(report_router, tags=["reports"])

# Корневой маршрут
@app.get("/", tags=["root"])
async def root():
    return {
        "message": "Добро пожаловать в сервис генерации документов!",
        "docs_url": "/docs",
        "admin_url": "/admin"
    }

# Для запуска через uvicorn
if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)
