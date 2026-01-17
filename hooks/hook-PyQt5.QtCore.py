# Custom hook for PyQt5.QtCore to handle Chinese path issues
from PyInstaller.utils.hooks import collect_submodules

hiddenimports = collect_submodules('PyQt5.QtCore')

# Don't try to collect Qt plugins automatically
datas = []
binaries = []