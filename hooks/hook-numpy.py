# hooks/hook-numpy.py
from PyInstaller.utils.hooks import collect_submodules, collect_data_files

hiddenimports = collect_submodules('numpy')

# Критические импорты numpy
critical_imports = [
    'numpy.core._dtype_ctypes',
    'numpy.core._multiarray_umath',
    'numpy.lib.format',
]

for imp in critical_imports:
    if imp not in hiddenimports:
        hiddenimports.append(imp)

datas = collect_data_files('numpy')