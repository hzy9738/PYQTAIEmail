#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
自动回复标签页
"""

from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
                            QLabel, QTextEdit, QComboBox, QMessageBox,
                            QGroupBox, QFormLayout, QSpinBox, QTableWidget,
                            QTableWidgetItem, QHeaderView)
from PyQt5.QtCore import Qt
from email_sender import EmailSender
from auto_reply import AutoReply


class AutoReplyTab(QWidget):
    """自动回复标签页"""

    def __init__(self, config_manager, auto_reply_manager, main_window):
        super().__init__()
        self.config_manager = config_manager
        self.auto_reply_manager = auto_reply_manager
        self.main_window = main_window
        self.init_ui()
        self.load_auto_reply_status()

    def init_ui(self):
        """初始化UI"""
        layout = QVBoxLayout(self)

        # 说明
        info_label = QLabel(
            "自动回复功能: 监控收件箱,自动回复新邮件\n"
            "��用于假期、休息时间等场景"
        )
        info_label.setStyleSheet("""
            QLabel {
                color: #34495e;
                font-size: 13px;
                padding: 10px;
                background-color: #ecf0f1;
                border-radius: 5px;
            }
        """)
        layout.addWidget(info_label)

        # 配置区域
        config_group = QGroupBox("自动回复配置")
        config_group.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                border: 2px solid #3498db;
                border-radius: 5px;
                margin-top: 10px;
                padding-top: 10px;
            }
        """)
        config_layout = QFormLayout()

        # 邮箱选择
        self.account_combo = QComboBox()
        self.account_combo.currentTextChanged.connect(self.on_account_changed)
        config_layout.addRow("选择邮箱:", self.account_combo)

        # 回复内容
        self.reply_content_input = QTextEdit()
        self.reply_content_input.setPlaceholderText("请输入自动回复的内容\n例如: 您好,我目前不在办公室,将在X月X日后回复您的邮件。")
        self.reply_content_input.setMinimumHeight(150)
        config_layout.addRow("回复内容:", self.reply_content_input)

        # 检查间隔
        interval_layout = QHBoxLayout()
        self.interval_spinbox = QSpinBox()
        self.interval_spinbox.setRange(30, 3600)
        self.interval_spinbox.setValue(60)
        self.interval_spinbox.setSuffix(" 秒")
        interval_layout.addWidget(self.interval_spinbox)
        interval_layout.addWidget(QLabel("(建议60秒以上,避免频繁检查)"))
        interval_layout.addStretch()
        config_layout.addRow("检查间隔:", interval_layout)

        config_group.setLayout(config_layout)
        layout.addWidget(config_group)

        # 控制按钮
        control_layout = QHBoxLayout()

        self.test_imap_btn = QPushButton("测试IMAP连接")
        self.test_imap_btn.setStyleSheet("""
            QPushButton {
                background-color: #f39c12;
                color: white;
                padding: 10px 20px;
                border-radius: 5px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #e67e22;
            }
        """)
        self.test_imap_btn.clicked.connect(self.test_imap_connection)
        control_layout.addWidget(self.test_imap_btn)

        self.save_btn = QPushButton("保存配置")
        self.save_btn.setStyleSheet("""
            QPushButton {
                background-color: #3498db;
                color: white;
                padding: 10px 20px;
                border-radius: 5px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
        """)
        self.save_btn.clicked.connect(self.save_config)
        control_layout.addWidget(self.save_btn)

        self.start_btn = QPushButton("启动自动回复")
        self.start_btn.setStyleSheet("""
            QPushButton {
                background-color: #27ae60;
                color: white;
                padding: 10px 20px;
                border-radius: 5px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #229954;
            }
        """)
        self.start_btn.clicked.connect(self.start_auto_reply)
        control_layout.addWidget(self.start_btn)

        self.stop_btn = QPushButton("停止自动回复")
        self.stop_btn.setStyleSheet("""
            QPushButton {
                background-color: #e74c3c;
                color: white;
                padding: 10px 20px;
                border-radius: 5px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #c0392b;
            }
        """)
        self.stop_btn.clicked.connect(self.stop_auto_reply)
        self.stop_btn.setEnabled(False)
        control_layout.addWidget(self.stop_btn)

        control_layout.addStretch()
        layout.addLayout(control_layout)

        # 状态显示
        status_group = QGroupBox("自动回复状态")
        status_group.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                border: 2px solid #3498db;
                border-radius: 5px;
                margin-top: 10px;
                padding-top: 10px;
            }
        """)
        status_layout = QVBoxLayout()

        self.status_table = QTableWidget()
        self.status_table.setColumnCount(3)
        self.status_table.setHorizontalHeaderLabels(["邮箱地址", "状态", "回复内容预览"])
        self.status_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.status_table.setStyleSheet("""
            QTableWidget {
                border: 1px solid #ddd;
                gridline-color: #ddd;
            }
            QHeaderView::section {
                background-color: #3498db;
                color: white;
                padding: 5px;
                font-weight: bold;
            }
        """)
        status_layout.addWidget(self.status_table)

        status_group.setLayout(status_layout)
        layout.addWidget(status_group)

        # 日志显示区域
        log_group = QGroupBox("运行日志")
        log_group.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                border: 2px solid #27ae60;
                border-radius: 5px;
                margin-top: 10px;
                padding-top: 10px;
            }
        """)
        log_layout = QVBoxLayout()

        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMinimumHeight(200)
        self.log_text.setMaximumHeight(300)
        self.log_text.setStyleSheet("""
            QTextEdit {
                background-color: #2c3e50;
                color: #ecf0f1;
                font-family: 'Consolas', 'Courier New', monospace;
                font-size: 12px;
                border: 1px solid #34495e;
                padding: 5px;
            }
        """)
        self.log_text.setPlaceholderText("启动自动回复后，日志信息将在此显示...")

        # 日志控制按钮
        log_control_layout = QHBoxLayout()

        clear_log_btn = QPushButton("清空日志")
        clear_log_btn.setStyleSheet("""
            QPushButton {
                background-color: #95a5a6;
                color: white;
                padding: 5px 15px;
                border-radius: 3px;
            }
            QPushButton:hover {
                background-color: #7f8c8d;
            }
        """)
        clear_log_btn.clicked.connect(self.clear_log)
        log_control_layout.addWidget(clear_log_btn)

        log_control_layout.addStretch()

        log_layout.addWidget(self.log_text)
        log_layout.addLayout(log_control_layout)

        log_group.setLayout(log_layout)
        layout.addWidget(log_group)

        # 加载账号
        self.refresh_accounts()

    def refresh_accounts(self):
        """刷新账号列表"""
        self.account_combo.clear()
        accounts = self.config_manager.get_email_accounts()
        for account in accounts:
            self.account_combo.addItem(account["email"])

        # 加载配置
        if accounts:
            self.load_auto_reply_config()

    def on_account_changed(self):
        """账号切换事件"""
        self.load_auto_reply_config()

    def load_auto_reply_config(self):
        """加载自动回复配置"""
        config = self.config_manager.get_auto_reply_config()
        self.reply_content_input.setPlainText(config.get("reply_content", ""))

    def test_imap_connection(self):
        """测试IMAP连接"""
        email = self.account_combo.currentText()
        if not email:
            QMessageBox.warning(self, "警告", "请先选择邮箱账号")
            return

        # 获取账号凭证
        credentials = self.config_manager.get_account_credentials(email)
        if not credentials:
            QMessageBox.critical(self, "错误", "获取邮箱凭证失败")
            return

        self.main_window.update_status("正在测试IMAP连接...")
        self.test_imap_btn.setEnabled(False)

        try:
            # 创建自动回复实例用于测试
            auto_reply = AutoReply(
                email_address=credentials["email"],
                password=credentials.get("imap_password", credentials["password"]),
                imap_server=credentials["imap_server"],
                imap_port=credentials["imap_port"],
                smtp_sender=None
            )

            # 测试连接
            success, message = auto_reply.test_connection()

            if success:
                QMessageBox.information(self, "成功", "IMAP连接测试成功！\n" + message)
                self.main_window.update_status("IMAP连接测试成功")
            else:
                QMessageBox.critical(self, "失败", "IMAP连接测试失败！\n" + message)
                self.main_window.update_status("IMAP连接测试失败")

        except Exception as e:
            QMessageBox.critical(self, "错误", f"测试IMAP连接时出错:\n{str(e)}")
            self.main_window.update_status("测试IMAP连接出错")
        finally:
            self.test_imap_btn.setEnabled(True)

    def save_config(self):
        """保存配置"""
        reply_content = self.reply_content_input.toPlainText().strip()

        if not reply_content:
            QMessageBox.warning(self, "警告", "请输入回复内容")
            return

        self.config_manager.set_auto_reply(
            enabled=False,
            reply_content=reply_content
        )

        QMessageBox.information(self, "成功", "配置已保存")
        self.main_window.update_status("自动回复配置已保存")

    def start_auto_reply(self):
        """启动自动回复"""
        email = self.account_combo.currentText()
        if not email:
            QMessageBox.warning(self, "警告", "请先添加邮箱账号")
            return

        reply_content = self.reply_content_input.toPlainText().strip()
        if not reply_content:
            QMessageBox.warning(self, "警告", "请输入回复内容")
            return

        check_interval = self.interval_spinbox.value()

        # 获取账号凭证
        credentials = self.config_manager.get_account_credentials(email)
        if not credentials:
            QMessageBox.critical(self, "错误", "获取邮箱凭证失败")
            return

        try:
            # 第一步:测试IMAP连接
            self.main_window.update_status("正在验证IMAP连接...")

            test_auto_reply = AutoReply(
                email_address=credentials["email"],
                password=credentials.get("imap_password", credentials["password"]),
                imap_server=credentials["imap_server"],
                imap_port=credentials["imap_port"],
                smtp_sender=None
            )

            success, test_message = test_auto_reply.test_connection()
            if not success:
                QMessageBox.critical(self, "错误", f"IMAP连接失败,无法启动自动回复:\n{test_message}")
                self.main_window.update_status("IMAP连接失败")
                return

            # 第二步:创建邮件发送器和自动回复实例
            sender = EmailSender(
                email=credentials["email"],
                password=credentials["password"],
                smtp_server=credentials["smtp_server"],
                smtp_port=credentials["smtp_port"]
            )

            # 创建自动回复实例 - 使用IMAP授权码
            auto_reply = AutoReply(
                email_address=credentials["email"],
                password=credentials.get("imap_password", credentials["password"]),  # 使用IMAP授权码,如果不存在则使用SMTP授权码
                imap_server=credentials["imap_server"],
                imap_port=credentials["imap_port"],
                smtp_sender=sender,
                log_callback=self.append_log  # 传递日志回调函数
            )

            # 清空日志显示
            self.clear_log()
            self.append_log("="*60)
            self.append_log(f"准备启动自动回复服务: {email}")
            self.append_log("="*60)

            # 添加到管理器
            self.auto_reply_manager.add_auto_reply(email, auto_reply)

            # 启动自动回复
            if self.auto_reply_manager.start_auto_reply(email, reply_content, check_interval):
                QMessageBox.information(self, "成功", f"已为 {email} 启动自动回复\nIMAP连接已验证,服务正在运行")

                # 更新UI状态
                self.start_btn.setEnabled(False)
                self.stop_btn.setEnabled(True)
                self.account_combo.setEnabled(False)

                # 保存配置
                self.config_manager.set_auto_reply(enabled=True, reply_content=reply_content)

                # 更新状态表
                self.load_auto_reply_status()

                self.main_window.update_status(f"{email} 自动回复已启动")
            else:
                QMessageBox.warning(self, "失败", "启动自动回复失败")

        except Exception as e:
            QMessageBox.critical(self, "错误", f"启动自动回复时出错:\n{str(e)}")

    def stop_auto_reply(self):
        """停止自动回复"""
        email = self.account_combo.currentText()
        if not email:
            return

        if self.auto_reply_manager.stop_auto_reply(email):
            QMessageBox.information(self, "成功", f"已停止 {email} 的自动回复")

            # 更新UI状态
            self.start_btn.setEnabled(True)
            self.stop_btn.setEnabled(False)
            self.account_combo.setEnabled(True)

            # 更新配置
            self.config_manager.set_auto_reply(enabled=False)

            # 更新状态表
            self.load_auto_reply_status()

            self.main_window.update_status(f"{email} 自动回复已停止")
        else:
            QMessageBox.warning(self, "失败", "停止自��回复失败")

    def load_auto_reply_status(self):
        """加载自动回复状态"""
        accounts = self.config_manager.get_email_accounts()
        self.status_table.setRowCount(len(accounts))

        config = self.config_manager.get_auto_reply_config()
        reply_content_preview = config.get("reply_content", "")[:50]
        if len(config.get("reply_content", "")) > 50:
            reply_content_preview += "..."

        for i, account in enumerate(accounts):
            email = account["email"]

            # 邮箱地址
            self.status_table.setItem(i, 0, QTableWidgetItem(email))

            # 状态
            is_active = self.auto_reply_manager.get_status(email)
            status_text = "运行中" if is_active else "已停止"
            status_item = QTableWidgetItem(status_text)

            if is_active:
                status_item.setForeground(Qt.darkGreen)
                # 更新UI状态(如果是当前选中的账号)
                if email == self.account_combo.currentText():
                    self.start_btn.setEnabled(False)
                    self.stop_btn.setEnabled(True)
                    self.account_combo.setEnabled(False)
            else:
                status_item.setForeground(Qt.red)

            self.status_table.setItem(i, 1, status_item)

            # 回复内容预览
            self.status_table.setItem(i, 2, QTableWidgetItem(reply_content_preview))

    def clear_log(self):
        """清空日志"""
        self.log_text.clear()

    def append_log(self, message):
        """添加日志消息（线程安全）"""
        from PyQt5.QtCore import QMetaObject, Q_ARG
        # 使用invokeMethod确保在主线程中更新UI
        QMetaObject.invokeMethod(self.log_text, "append",
                                 Qt.QueuedConnection,
                                 Q_ARG(str, message))
        # 自动滚动到底部
        QMetaObject.invokeMethod(self.log_text.verticalScrollBar(), "setValue",
                                 Qt.QueuedConnection,
                                 Q_ARG(int, self.log_text.verticalScrollBar().maximum()))
