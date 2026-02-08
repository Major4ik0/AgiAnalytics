# -*- coding: utf-8 -*-
"""
Модуль для работы с ресурсами в PyInstaller и разработке
"""
import sys
import os


def resource_path(relative_path):
    """Получить абсолютный путь к ресурсу, работает для dev и для PyInstaller"""
    try:
        # PyInstaller создает временную папку и сохраняет путь в _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    path = os.path.join(base_path, relative_path)

    # Пробуем найти файл
    if not os.path.exists(path):
        # Проверяем альтернативные пути
        alternative_paths = [
            path,
            os.path.join(os.path.dirname(sys.executable), relative_path),
            os.path.join(os.path.dirname(__file__), relative_path),
            os.path.join(os.path.dirname(__file__), "..", relative_path),
        ]

        for alt_path in alternative_paths:
            if os.path.exists(alt_path):
                return alt_path

    return path


def get_icon_path(icon_name="icon.ico"):
    """Получить путь к иконке"""
    icon_paths = [
        icon_name,  # Просто имя файла
        f"icons/{icon_name}",  # В папке icons
        f"../icons/{icon_name}",  # На уровень выше
        f"../../icons/{icon_name}",  # На два уровня выше
    ]

    for icon_path in icon_paths:
        full_path = resource_path(icon_path)
        if os.path.exists(full_path):
            return full_path

    # Если иконка не найдена, возвращаем None
    return None


# Функция для отладки
def debug_resources():
    """Отладка поиска ресурсов"""
    print("=" * 50)
    print("Отладка ресурсов")
    print("=" * 50)

    print(f"sys._MEIPASS: {getattr(sys, '_MEIPASS', 'Не доступен')}")
    print(f"Текущий файл: {__file__}")
    print(f"Директория файла: {os.path.dirname(__file__)}")
    print(f"Исполняемый файл: {sys.executable}")
    print(f"Директория исполняемого файла: {os.path.dirname(sys.executable)}")

    # Проверяем иконки
    icon_path = get_icon_path("icon.ico")
    print(f"\nПуть к иконке: {icon_path}")
    print(f"Иконка существует: {os.path.exists(icon_path) if icon_path else False}")

    print("=" * 50)


if __name__ == "__main__":
    debug_resources()