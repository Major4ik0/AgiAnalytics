# hooks/hook-pandas.py
from PyInstaller.utils.hooks import collect_submodules, collect_data_files

# Получаем все подмодули pandas
hiddenimports = collect_submodules('pandas')

# Критические импорты которые могут быть пропущены
critical_imports = [
    # pandas._libs модули
    'pandas._libs.tslibs.np_datetime',
    'pandas._libs.tslibs.timedeltas',
    'pandas._libs.tslibs.nattype',
    'pandas._libs.tslibs.base',
    'pandas._libs.missing',
    'pandas._libs.tslib',
    'pandas._libs.algos',

    # pandas.io модули
    'pandas.io.formats.excel',
    'pandas.io.excel._openpyxl',

    # pandas.core модули
    'pandas.core.computation.ops',
    'pandas.core.dtypes.dtypes',
]

# Добавляем только те которые еще не в hiddenimports
for imp in critical_imports:
    if imp not in hiddenimports:
        hiddenimports.append(imp)

# Собираем данные pandas
datas = collect_data_files('pandas')