FROM python:3.10-slim

WORKDIR /app

# Устанавливаем системные зависимости для PyQt
RUN apt-get update && apt-get install -y \
    gcc \
    g++ \
    make \
    libgl1-mesa-glx \
    libglib2.0-0 \
    && rm -rf /var/lib/apt/lists/*

# Копируем файлы проекта
COPY . .

# Устанавливаем зависимости Python
RUN pip install --no-cache-dir -r requirements.txt
RUN pip install pyinstaller

# Собираем exe файл
RUN pyinstaller --onefile --windowed --name=AgiAnalytics  --icon="icon.ico" main.py

# Результирующий файл будет в /app/dist/AgiAnalytics.exe