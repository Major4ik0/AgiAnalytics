# hooks/hook-PyQt5.py
from PyInstaller.utils.hooks import collect_submodules, collect_data_files

# Получаем все подмодули PyQt5
hiddenimports = collect_submodules('PyQt5')

# Добавляем основные модули Qt которые могут быть пропущены
critical_imports = [
    'PyQt5.QtCore',
    'PyQt5.QtGui',
    'PyQt5.QtWidgets',
    'PyQt5.sip',
    'PyQt5.QtPrintSupport',
    'PyQt5.QtSvg',
    'PyQt5.QtXml',
]

for imp in critical_imports:
    if imp not in hiddenimports:
        hiddenimports.append(imp)

# Собираем данные PyQt5
datas = collect_data_files('PyQt5', include_py_files=True)