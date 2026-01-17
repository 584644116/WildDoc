# Custom hook for PyQt5.QtGui to handle Chinese path issues
from PyInstaller.utils.hooks import collect_submodules

hiddenimports = collect_submodules('PyQt5.QtGui')

# Don't try to collect Qt plugins automatically
datas = []
binaries = []