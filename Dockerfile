FROM python:3.11-slim

WORKDIR /app

# Установка системных зависимостей
RUN apt-get update \
  # dependencies for building Python packages
  && apt-get install -y build-essential \
  # psycopg2 dependencies
  && apt-get install -y libpq-dev \
  # Additional dependencies
  && apt-get install -y telnet netcat-openbsd \
  # cleaning up unused files
  && apt-get purge -y --auto-remove -o APT::AutoRemove::RecommendsImportant=false \
  && rm -rf /var/lib/apt/lists/*

# Копирование зависимостей и установка их
COPY requirements.txt .
RUN pip install -r requirements.txt
RUN pip cache purge

# Копирование исходного кода приложения
COPY . .

# Создание файла для пользовательских настроек
RUN touch .env

# Открытие порта
EXPOSE 8000

# Настройка переменной среды
ENV PYTHONUNBUFFERED=1

# Делаем скрипт запуска исполняемым
RUN chmod +x entrypoint.sh

# Команда для запуска приложения
CMD ["./entrypoint.sh"]
