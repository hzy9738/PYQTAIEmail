#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
寻拟邮件工具 - 主程序入口
"""

import sys
import os
from PyQt5.QtWidgets import QApplication
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import Qt
from main_window import MainWindow


def get_resource_path(relative_path):
    """获取资源文件的绝对路径(支持打包后的环境)"""
    try:
        # PyInstaller创建临时文件夹,将路径存储在_MEIPASS中
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


def main():
    """主函数"""
    # 启用高DPI支持
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)

    # 创建应用
    app = QApplication(sys.argv)
    app.setApplicationName("寻拟邮件工具")
    app.setOrganizationName("XunNiEmail")
    app.setStyle("Fusion")  # 使用Fusion风格,界面更现代

    # 设置应用程序图标
    icon_path = get_resource_path("icon.ico")
    if os.path.exists(icon_path):
        app.setWindowIcon(QIcon(icon_path))

    # 创建主窗口
    window = MainWindow()
    window.show()

    # 运行应用
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
