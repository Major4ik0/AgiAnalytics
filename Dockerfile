FROM python:3.10

WORKDIR /app

# Для Windows сборки лучше использовать wine
RUN apt-get update && apt-get install -y \
    wine64 \
    mingw-w64 \
    gcc-mingw-w64-x86-64 \
    g++-mingw-w64-x86-64 \
    && rm -rf /var/lib/apt/lists/*

# Настраиваем wine
ENV WINEPREFIX=/root/.wine
ENV WINEARCH=win64
RUN wine wineboot --init

# Копируем файлы проекта
COPY requirements.txt .
COPY *.py ./
COPY *.ui ./

# Устанавливаем Python зависимости в wine
RUN wine python -m pip install --upgrade pip
RUN wine python -m pip install pyinstaller PyQt6 pandas openpyxl

# Создаем сборку
RUN wine python -m PyInstaller \
    --onefile \
    --windowed \
    --name="AgiAnalytics" \
    --icon=icon.ico \
    main.py

# Копируем результат
RUN mkdir -p /output && cp /root/.wine/drive_c/users/root/AppData/Local/Programs/Python/Python310/dist/AgiAnalytics.exe /output/