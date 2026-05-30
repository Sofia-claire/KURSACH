# Итоговый легковесный образ Python
FROM python:3.10-slim

# Установка системных зависимостей для корректной сборки RapidFuzz и Tkinter
RUN apt-get update && apt-get install -y \
    build-essential \
    python3-tk \
    && rm -rf /var/lib/apt/lists/*

# Создаем рабочую папку в контейнере
WORKDIR /app

# Копируем список зависимостей и ставим их
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Копируем все остальные файлы проекта
COPY . .

# Команда, которая запустится при старте контейнера
CMD ["python", "main.py"]
