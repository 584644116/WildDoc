# Custom hook for PyQt5.QtWidgets to handle Chinese path issues
from PyInstaller.utils.hooks import collect_submodules

hiddenimports = collect_submodules('PyQt5.QtWidgets')

# Don't try to collect Qt plugins automatically - we'll handle them manually
datas = []
binaries = []