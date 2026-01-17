import os
import sys

# 设置Qt插件路径
qt_plugin_path = os.path.join(os.path.dirname(__file__), 'venv', 'Lib', 'site-packages', 'PyQt5', 'Qt5', 'plugins')
if os.path.exists(qt_plugin_path):
    os.environ['QT_PLUGIN_PATH'] = qt_plugin_path

# 设置Qt平台插件路径
qt_platform_path = os.path.join(qt_plugin_path, 'platforms')
if os.path.exists(qt_platform_path):
    os.environ['QT_QPA_PLATFORM_PLUGIN_PATH'] = qt_platform_path

# 设置高DPI支持和缩放
os.environ['QT_AUTO_SCREEN_SCALE_FACTOR'] = '0'
os.environ['QT_SCALE_FACTOR'] = '1.0'
os.environ['QT_SCREEN_SCALE_FACTORS'] = '1.0'
os.environ['QT_DEVICE_PIXEL_RATIO'] = '1'

# 导入并运行主程序
try:
    from PyQt5.QtWidgets import QApplication, QStyleFactory
    from PyQt5.QtCore import Qt
    from PyQt5.QtGui import QFont, QPalette, QColor
    from gui.modern_main_window import ModernMainWindow
    
    def main():
        # 设置高DPI属性（必须在QApplication创建之前）
        QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, False)
        QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
        QApplication.setAttribute(Qt.AA_DisableWindowContextHelpButton, True)
        
        app = QApplication(sys.argv)
        
        # 设置应用程序样式为Windows原生样式
        available_styles = QStyleFactory.keys()
        if 'WindowsVista' in available_styles:
            app.setStyle('WindowsVista')
        elif 'Windows' in available_styles:
            app.setStyle('Windows')
        else:
            app.setStyle('Fusion')
        
        # 设置合适的字体
        font = QFont("Microsoft YaHei UI", 9)
        font.setStyleHint(QFont.SansSerif)
        app.setFont(font)
        
        # 设置调色板以确保良好的对比度
        palette = QPalette()
        palette.setColor(QPalette.Window, QColor(240, 240, 240))
        palette.setColor(QPalette.WindowText, QColor(0, 0, 0))
        palette.setColor(QPalette.Base, QColor(255, 255, 255))
        palette.setColor(QPalette.AlternateBase, QColor(245, 245, 245))
        palette.setColor(QPalette.ToolTipBase, QColor(255, 255, 220))
        palette.setColor(QPalette.ToolTipText, QColor(0, 0, 0))
        palette.setColor(QPalette.Text, QColor(0, 0, 0))
        palette.setColor(QPalette.Button, QColor(240, 240, 240))
        palette.setColor(QPalette.ButtonText, QColor(0, 0, 0))
        palette.setColor(QPalette.BrightText, QColor(255, 0, 0))
        palette.setColor(QPalette.Link, QColor(0, 0, 255))
        palette.setColor(QPalette.Highlight, QColor(0, 120, 215))
        palette.setColor(QPalette.HighlightedText, QColor(255, 255, 255))
        app.setPalette(palette)
        
        # 创建并显示主窗口（现代UI）
        window = ModernMainWindow()

        # 设置窗口属性
        window.setMinimumSize(900, 600)
        window.resize(1200, 800)
        
        window.show()
        
        # 运行应用程序
        sys.exit(app.exec_())
    
    if __name__ == '__main__':
        main()
        
except ImportError as e:
    print(f"导入错误: {e}")
    print("请确保已经激活虚拟环境并安装了所有依赖")
    print("运行以下命令:")
    print("1. .\\venv\\Scripts\\activate")
    print("2. pip install -r requirements.txt")
    input("按回车键退出...")
except Exception as e:
    print(f"运行错误: {e}")
    import traceback
    traceback.print_exc()
    input("按回车键退出...")
