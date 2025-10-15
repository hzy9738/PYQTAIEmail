#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
自动化邮件服务系统 - 主程序入口
"""

import sys
from PyQt5.QtWidgets import QApplication
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import Qt
from main_window import MainWindow


def main():
    """主函数"""
    # 启用高DPI支持
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)

    # 创建应用
    app = QApplication(sys.argv)
    app.setApplicationName("自动化邮件服务系统")
    app.setOrganizationName("EmailService")
    app.setStyle("Fusion")  # 使用Fusion风格,界面更现代

    # 创建主窗口
    window = MainWindow()
    window.show()

    # 运行应用
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
