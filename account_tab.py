#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
邮箱账号管理标签页
"""

from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
                            QLabel, QLineEdit, QTableWidget, QTableWidgetItem,
                            QMessageBox, QGroupBox, QFormLayout, QSpinBox,
                            QHeaderView)
from PyQt5.QtCore import Qt
from email_sender import EmailSender


class AccountTab(QWidget):
    """邮箱账号管理标签页"""

    def __init__(self, config_manager, main_window):
        super().__init__()
        self.config_manager = config_manager
        self.main_window = main_window
        self.init_ui()
        self.load_accounts()

    def init_ui(self):
        """初始化UI"""
        layout = QVBoxLayout(self)

        # 添加账号区域
        add_group = QGroupBox("添加邮箱账号")
        add_group.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                border: 2px solid #3498db;
                border-radius: 5px;
                margin-top: 10px;
                padding-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px;
            }
        """)
        add_layout = QFormLayout()

        # 邮箱地址
        self.email_input = QLineEdit()
        self.email_input.setPlaceholderText("example@163.com")
        add_layout.addRow("邮箱地址:", self.email_input)

        # 授权码
        self.password_input = QLineEdit()
        self.password_input.setEchoMode(QLineEdit.Password)
        self.password_input.setPlaceholderText("邮箱授权码(非登录密码)")
        add_layout.addRow("授权码:", self.password_input)

        # SMTP服务器
        smtp_layout = QHBoxLayout()
        self.smtp_server_input = QLineEdit("smtp.163.com")
        self.smtp_port_input = QSpinBox()
        self.smtp_port_input.setRange(1, 65535)
        self.smtp_port_input.setValue(465)
        smtp_layout.addWidget(self.smtp_server_input, 2)
        smtp_layout.addWidget(QLabel("端口:"))
        smtp_layout.addWidget(self.smtp_port_input, 1)
        add_layout.addRow("SMTP服务器:", smtp_layout)

        # 按钮
        button_layout = QHBoxLayout()

        self.test_btn = QPushButton("测试连接")
        self.test_btn.setStyleSheet("""
            QPushButton {
                background-color: #f39c12;
                color: white;
                padding: 8px 15px;
                border-radius: 5px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #e67e22;
            }
        """)
        self.test_btn.clicked.connect(self.test_connection)
        button_layout.addWidget(self.test_btn)

        self.add_btn = QPushButton("添加账号")
        self.add_btn.setStyleSheet("""
            QPushButton {
                background-color: #27ae60;
                color: white;
                padding: 8px 15px;
                border-radius: 5px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #229954;
            }
        """)
        self.add_btn.clicked.connect(self.add_account)
        button_layout.addWidget(self.add_btn)

        add_layout.addRow("", button_layout)
        add_group.setLayout(add_layout)
        layout.addWidget(add_group)

        # 账号列表区域
        list_group = QGroupBox("已添加的邮箱账号")
        list_group.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                border: 2px solid #3498db;
                border-radius: 5px;
                margin-top: 10px;
                padding-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px;
            }
        """)
        list_layout = QVBoxLayout()

        # 账号表格
        self.account_table = QTableWidget()
        self.account_table.setColumnCount(4)
        self.account_table.setHorizontalHeaderLabels(["邮箱地址", "SMTP服务器", "端口", "操作"])
        self.account_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.account_table.setStyleSheet("""
            QTableWidget {
                border: 1px solid #ddd;
                gridline-color: #ddd;
            }
            QTableWidget::item {
                padding: 5px;
            }
            QHeaderView::section {
                background-color: #3498db;
                color: white;
                padding: 5px;
                font-weight: bold;
            }
        """)
        list_layout.addWidget(self.account_table)

        list_group.setLayout(list_layout)
        layout.addWidget(list_group)

        # 说明文字
        help_label = QLabel(
            "提示: 请使用邮箱的授权码,而不是登录密码。\n"
            "163邮箱获取授权码: 设置 → POP3/SMTP/IMAP → 开启IMAP/SMTP服务 → 生成授权码"
        )
        help_label.setStyleSheet("""
            QLabel {
                color: #7f8c8d;
                font-size: 12px;
                padding: 10px;
                background-color: #ecf0f1;
                border-radius: 5px;
            }
        """)
        layout.addWidget(help_label)

    def test_connection(self):
        """测试邮箱连接"""
        email = self.email_input.text().strip()
        password = self.password_input.text().strip()
        smtp_server = self.smtp_server_input.text().strip()
        smtp_port = self.smtp_port_input.value()

        if not email or not password:
            QMessageBox.warning(self, "警告", "请填写邮箱地址和授权码")
            return

        self.main_window.update_status("正在测试连接...")
        self.test_btn.setEnabled(False)

        try:
            sender = EmailSender(email, password, smtp_server, smtp_port)
            success, message = sender.test_connection()

            if success:
                QMessageBox.information(self, "成功", "连接测试成功！\n" + message)
                self.main_window.update_status("连接测试成功")
            else:
                QMessageBox.critical(self, "失败", "连接测试失败！\n" + message)
                self.main_window.update_status("连接测试失败")

        except Exception as e:
            QMessageBox.critical(self, "错误", f"测试连接时出错:\n{str(e)}")
            self.main_window.update_status("测试连接出错")
        finally:
            self.test_btn.setEnabled(True)

    def add_account(self):
        """添加邮箱账号"""
        email = self.email_input.text().strip()
        password = self.password_input.text().strip()
        smtp_server = self.smtp_server_input.text().strip()
        smtp_port = self.smtp_port_input.value()

        if not email or not password:
            QMessageBox.warning(self, "警告", "请填写邮箱地址和授权码")
            return

        # 添加账号
        if self.config_manager.add_email_account(email, password, smtp_server, smtp_port):
            QMessageBox.information(self, "成功", f"账号 {email} 添加成功！")
            self.email_input.clear()
            self.password_input.clear()
            self.load_accounts()

            # 通知其他标签页更新
            self.main_window.refresh_send_email_accounts()
            self.main_window.refresh_schedule_accounts()
            self.main_window.refresh_auto_reply_accounts()

            self.main_window.update_status(f"账号 {email} 添加成功")
        else:
            QMessageBox.warning(self, "失败", f"账号 {email} 已存在或添加失败")
            self.main_window.update_status("添加账号失败")

    def load_accounts(self):
        """加载账号列表"""
        accounts = self.config_manager.get_email_accounts()
        self.account_table.setRowCount(len(accounts))

        for i, account in enumerate(accounts):
            # 邮箱地址
            self.account_table.setItem(i, 0, QTableWidgetItem(account["email"]))

            # SMTP服务器
            self.account_table.setItem(i, 1, QTableWidgetItem(account["smtp_server"]))

            # 端口
            self.account_table.setItem(i, 2, QTableWidgetItem(str(account["smtp_port"])))

            # 删除按钮
            delete_btn = QPushButton("删除")
            delete_btn.setStyleSheet("""
                QPushButton {
                    background-color: #e74c3c;
                    color: white;
                    padding: 5px 10px;
                    border-radius: 3px;
                }
                QPushButton:hover {
                    background-color: #c0392b;
                }
            """)
            delete_btn.clicked.connect(lambda checked, email=account["email"]: self.delete_account(email))
            self.account_table.setCellWidget(i, 3, delete_btn)

    def delete_account(self, email):
        """删除账号"""
        reply = QMessageBox.question(
            self,
            "确认删除",
            f"确定要删除账号 {email} 吗？",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            if self.config_manager.remove_email_account(email):
                QMessageBox.information(self, "成功", f"账号 {email} 已删除")
                self.load_accounts()

                # 通知其他标签页更新
                self.main_window.refresh_send_email_accounts()
                self.main_window.refresh_schedule_accounts()
                self.main_window.refresh_auto_reply_accounts()

                self.main_window.update_status(f"账号 {email} 已删除")
            else:
                QMessageBox.warning(self, "失败", "删除账号失败")
