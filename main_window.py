#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
主窗口模块
提供PyQt5图形界面
"""

from PyQt5.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                            QTabWidget, QPushButton, QLabel, QMessageBox,
                            QSystemTrayIcon, QMenu, QAction, QApplication)
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QIcon
import sys
import os

from account_tab import AccountTab
from send_email_tab import SendEmailTab
from schedule_tab import ScheduleTab
from auto_reply_tab import AutoReplyTab
from config_manager import ConfigManager
from email_sender import EmailSender
from auto_reply import AutoReply, AutoReplyManager
from task_scheduler import ScheduleManager


def get_resource_path(relative_path):
    """获取资源文件的绝对路径(支持打包后的环境)"""
    try:
        # PyInstaller创建临时文件夹,将路径存储在_MEIPASS中
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


class MainWindow(QMainWindow):
    """主窗口类"""

    def __init__(self):
        super().__init__()
        self.config_manager = ConfigManager()
        self.auto_reply_manager = AutoReplyManager()
        self.schedule_manager = ScheduleManager()

        self.init_ui()
        self.setup_tray()

    def init_ui(self):
        """初始化UI"""
        self.setWindowTitle("寻拟邮件工具")

        # 设置窗口初始大小和位置
        self.setGeometry(100, 100, 1000, 700)

        # 设置窗口最小尺寸(允许用户调整大小)
        self.setMinimumSize(800, 600)

        # 设置窗口图标
        icon_path = get_resource_path("icon.ico")
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))

        # 中央部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        # 主布局
        main_layout = QVBoxLayout(central_widget)

        # 标题
        title_label = QLabel("📧 寻拟邮件工具")
        title_label.setStyleSheet("""
            QLabel {
                font-size: 24px;
                font-weight: bold;
                color: #2c3e50;
                padding: 10px;
            }
        """)
        title_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title_label)

        # 标签页
        self.tab_widget = QTabWidget()
        self.tab_widget.setStyleSheet("""
            QTabWidget::pane {
                border: 1px solid #cccccc;
                border-radius: 5px;
            }
            QTabBar::tab {
                background: #ecf0f1;
                padding: 10px 20px;
                margin: 2px;
                border-radius: 5px;
            }
            QTabBar::tab:selected {
                background: #3498db;
                color: white;
            }
        """)

        # 创建各个标签页
        self.account_tab = AccountTab(self.config_manager, self)
        self.send_email_tab = SendEmailTab(self.config_manager, self)
        self.schedule_tab = ScheduleTab(self.config_manager, self.schedule_manager, self)
        self.auto_reply_tab = AutoReplyTab(self.config_manager, self.auto_reply_manager, self)

        # 添加标签页
        self.tab_widget.addTab(self.account_tab, "邮箱账号")
        self.tab_widget.addTab(self.send_email_tab, "发送邮件")
        self.tab_widget.addTab(self.schedule_tab, "定时任务")
        self.tab_widget.addTab(self.auto_reply_tab, "自动回复")

        main_layout.addWidget(self.tab_widget)

        # 底部状态栏
        self.statusBar().showMessage("就绪")

    def setup_tray(self):
        """设置系统托盘"""
        self.tray_icon = QSystemTrayIcon(self)

        # 设置托盘图标
        icon_path = get_resource_path("icon.ico")
        if os.path.exists(icon_path):
            self.tray_icon.setIcon(QIcon(icon_path))

        self.tray_icon.setToolTip("寻拟邮件工具")

        # 创建托盘菜单
        tray_menu = QMenu()

        show_action = QAction("显示主窗口", self)
        show_action.triggered.connect(self.show)
        tray_menu.addAction(show_action)

        quit_action = QAction("退出", self)
        quit_action.triggered.connect(self.quit_application)
        tray_menu.addAction(quit_action)

        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.activated.connect(self.tray_icon_activated)
        self.tray_icon.show()

    def tray_icon_activated(self, reason):
        """托盘图标激活事件"""
        if reason == QSystemTrayIcon.DoubleClick:
            self.show()

    def closeEvent(self, event):
        """关闭事件 - 最小化到托盘"""
        if self.tray_icon.isVisible():
            # 保存发送邮件页面的状态
            self.send_email_tab.save_state()
            
            self.hide()
            self.tray_icon.showMessage(
                "寻拟邮件工具",
                "程序已最小化到系统托盘",
                QSystemTrayIcon.Information,
                2000
            )
            event.ignore()
        else:
            # 保存发送邮件页面的状态
            self.send_email_tab.save_state()
            event.accept()

    def quit_application(self):
        """退出应用"""
        reply = QMessageBox.question(
            self,
            "确认退出",
            "确定要退出程序吗？\n所有定时任务和自动回复将停止。",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            # 保存发送邮件页面的状态
            self.send_email_tab.save_state()
            
            # 停止所有服务
            self.auto_reply_manager.stop_all()
            self.schedule_manager.stop_scheduler()

            # 退出程序
            QApplication.quit()

    def show_message(self, title, message, msg_type="info"):
        """显示消息框"""
        if msg_type == "info":
            QMessageBox.information(self, title, message)
        elif msg_type == "warning":
            QMessageBox.warning(self, title, message)
        elif msg_type == "error":
            QMessageBox.critical(self, title, message)
        elif msg_type == "question":
            return QMessageBox.question(self, title, message)

    def update_status(self, message):
        """更新状态栏"""
        self.statusBar().showMessage(message)

    def refresh_send_email_accounts(self):
        """刷新发送邮件标签页的账号列表"""
        self.send_email_tab.refresh_accounts()

    def refresh_schedule_accounts(self):
        """刷新定时任务标签页的账号列表"""
        self.schedule_tab.refresh_accounts()

    def refresh_auto_reply_accounts(self):
        """刷新自动回复标签页的账号列表"""
        self.auto_reply_tab.refresh_accounts()
