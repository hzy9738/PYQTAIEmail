#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
定时任务管理标签页
"""

from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
                            QLabel, QLineEdit, QTextEdit, QComboBox,
                            QMessageBox, QGroupBox, QFormLayout,
                            QTableWidget, QTableWidgetItem, QHeaderView,
                            QTimeEdit, QDialog, QDialogButtonBox)
from PyQt5.QtCore import Qt, QTime
from email_sender import EmailSender


class AddTaskDialog(QDialog):
    """添加定时任务对话框"""

    def __init__(self, config_manager, parent=None):
        super().__init__(parent)
        self.config_manager = config_manager
        self.setWindowTitle("添加定时任务")
        self.setModal(True)
        self.resize(600, 500)
        self.init_ui()

    def init_ui(self):
        """初始化UI"""
        layout = QVBoxLayout(self)

        form_layout = QFormLayout()

        # 任务名称
        self.task_name_input = QLineEdit()
        self.task_name_input.setPlaceholderText("例如: 每日报告")
        form_layout.addRow("任务名称:", self.task_name_input)

        # 发件账号
        self.sender_combo = QComboBox()
        accounts = self.config_manager.get_email_accounts()
        for account in accounts:
            self.sender_combo.addItem(account["email"])
        form_layout.addRow("发件账号:", self.sender_combo)

        # 收件人
        self.recipients_input = QTextEdit()
        self.recipients_input.setPlaceholderText("每行一个邮箱地址")
        self.recipients_input.setMaximumHeight(80)
        form_layout.addRow("收件人:", self.recipients_input)

        # 邮件主题
        self.subject_input = QLineEdit()
        self.subject_input.setPlaceholderText("请输入邮件主题")
        form_layout.addRow("邮件主题:", self.subject_input)

        # 邮件内容
        self.content_input = QTextEdit()
        self.content_input.setPlaceholderText("请输入邮件内容")
        self.content_input.setMinimumHeight(150)
        form_layout.addRow("邮件内容:", self.content_input)

        # 执行时间
        self.time_edit = QTimeEdit()
        self.time_edit.setDisplayFormat("HH:mm")
        self.time_edit.setTime(QTime(9, 0))
        form_layout.addRow("执行时间:", self.time_edit)

        layout.addLayout(form_layout)

        # 按钮
        button_box = QDialogButtonBox(
            QDialogButtonBox.Ok | QDialogButtonBox.Cancel
        )
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)

    def get_task_data(self):
        """获��任务数据"""
        recipients = []
        for line in self.recipients_input.toPlainText().split('\n'):
            email = line.strip()
            if email and '@' in email:
                recipients.append(email)

        return {
            "task_name": self.task_name_input.text().strip(),
            "sender_email": self.sender_combo.currentText(),
            "recipients": recipients,
            "subject": self.subject_input.text().strip(),
            "content": self.content_input.toPlainText().strip(),
            "schedule_time": self.time_edit.time().toString("HH:mm")
        }


class ScheduleTab(QWidget):
    """定时任务管理标签页"""

    def __init__(self, config_manager, schedule_manager, main_window):
        super().__init__()
        self.config_manager = config_manager
        self.schedule_manager = schedule_manager
        self.main_window = main_window
        self.init_ui()
        self.load_tasks()

        # 启动调度器
        self.schedule_manager.start_scheduler()

    def init_ui(self):
        """初始化UI"""
        layout = QVBoxLayout(self)

        # 标题和说明
        info_label = QLabel(
            "定时任务功能: 可以设置在每天指定时间自动发送邮件\n"
            "例如: 每天早上9:00发送日报"
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

        # 调度器状态
        status_layout = QHBoxLayout()

        self.status_label = QLabel("调度器状态: 运行中")
        self.status_label.setStyleSheet("""
            QLabel {
                color: #27ae60;
                font-weight: bold;
                padding: 5px;
            }
        """)
        status_layout.addWidget(self.status_label)

        status_layout.addStretch()

        self.toggle_scheduler_btn = QPushButton("停止调度器")
        self.toggle_scheduler_btn.setStyleSheet("""
            QPushButton {
                background-color: #e67e22;
                color: white;
                padding: 8px 15px;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #d35400;
            }
        """)
        self.toggle_scheduler_btn.clicked.connect(self.toggle_scheduler)
        status_layout.addWidget(self.toggle_scheduler_btn)

        layout.addLayout(status_layout)

        # 任务列表
        task_group = QGroupBox("定时任务列表")
        task_group.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                border: 2px solid #3498db;
                border-radius: 5px;
                margin-top: 10px;
                padding-top: 10px;
            }
        """)
        task_layout = QVBoxLayout()

        self.task_table = QTableWidget()
        self.task_table.setColumnCount(6)
        self.task_table.setHorizontalHeaderLabels([
            "任务名称", "发件人", "收件人数", "执行时间", "下次执行", "操作"
        ])
        self.task_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.task_table.setStyleSheet("""
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
        task_layout.addWidget(self.task_table)

        task_group.setLayout(task_layout)
        layout.addWidget(task_group)

        # 底部按钮
        button_layout = QHBoxLayout()
        button_layout.addStretch()

        add_task_btn = QPushButton("添加定时任务")
        add_task_btn.setStyleSheet("""
            QPushButton {
                background-color: #27ae60;
                color: white;
                padding: 10px 20px;
                border-radius: 5px;
                font-size: 14px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #229954;
            }
        """)
        add_task_btn.clicked.connect(self.add_task)
        button_layout.addWidget(add_task_btn)

        refresh_btn = QPushButton("刷新列表")
        refresh_btn.setStyleSheet("""
            QPushButton {
                background-color: #3498db;
                color: white;
                padding: 10px 20px;
                border-radius: 5px;
                font-size: 14px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
        """)
        refresh_btn.clicked.connect(self.load_tasks)
        button_layout.addWidget(refresh_btn)

        button_layout.addStretch()
        layout.addLayout(button_layout)

    def refresh_accounts(self):
        """刷新账号列表(由主窗口调用)"""
        pass

    def toggle_scheduler(self):
        """切换调度器状态"""
        if self.schedule_manager.is_running():
            self.schedule_manager.stop_scheduler()
            self.status_label.setText("调度器状态: 已停止")
            self.status_label.setStyleSheet("QLabel { color: #e74c3c; font-weight: bold; padding: 5px; }")
            self.toggle_scheduler_btn.setText("启动调度器")
            self.main_window.update_status("定时任务调度器已停止")
        else:
            self.schedule_manager.start_scheduler()
            self.status_label.setText("调度器状态: 运行中")
            self.status_label.setStyleSheet("QLabel { color: #27ae60; font-weight: bold; padding: 5px; }")
            self.toggle_scheduler_btn.setText("停止调度器")
            self.main_window.update_status("定时任务调度器已启动")

    def add_task(self):
        """添加定时任务"""
        accounts = self.config_manager.get_email_accounts()
        if not accounts:
            QMessageBox.warning(self, "警告", "请先添加邮箱账号")
            return

        dialog = AddTaskDialog(self.config_manager, self)
        if dialog.exec_() == QDialog.Accepted:
            task_data = dialog.get_task_data()

            # 验证数据
            if not task_data["task_name"]:
                QMessageBox.warning(self, "警告", "请输入任务名称")
                return

            if not task_data["recipients"]:
                QMessageBox.warning(self, "警告", "请输入收件人")
                return

            if not task_data["subject"] or not task_data["content"]:
                QMessageBox.warning(self, "警告", "请填写邮件主题和内容")
                return

            # 保存到配置
            if self.config_manager.add_scheduled_task(
                task_name=task_data["task_name"],
                recipients=task_data["recipients"],
                subject=task_data["subject"],
                content=task_data["content"],
                schedule_time=task_data["schedule_time"],
                sender_email=task_data["sender_email"]
            ):
                # 添加到调度器
                credentials = self.config_manager.get_account_credentials(task_data["sender_email"])
                sender = EmailSender(
                    email=credentials["email"],
                    password=credentials["password"],
                    smtp_server=credentials["smtp_server"],
                    smtp_port=credentials["smtp_port"]
                )

                # 添加发送器和任务
                self.schedule_manager.add_email_sender(task_data["sender_email"], sender)
                self.schedule_manager.add_scheduled_task(
                    task_name=task_data["task_name"],
                    schedule_time=task_data["schedule_time"],
                    sender_email=task_data["sender_email"],
                    recipients=task_data["recipients"],
                    subject=task_data["subject"],
                    content=task_data["content"]
                )

                QMessageBox.information(self, "成功", f"定时任务 '{task_data['task_name']}' 添加成功")
                self.load_tasks()
                self.main_window.update_status(f"定时任务 '{task_data['task_name']}' 已添加")
            else:
                QMessageBox.warning(self, "失败", "添加任务失败")

    def load_tasks(self):
        """加载任务列表"""
        tasks = self.config_manager.get_scheduled_tasks()
        self.task_table.setRowCount(len(tasks))

        for i, task in enumerate(tasks):
            # 任务名称
            self.task_table.setItem(i, 0, QTableWidgetItem(task["task_name"]))

            # 发件人
            self.task_table.setItem(i, 1, QTableWidgetItem(task["sender_email"]))

            # 收件人数
            self.task_table.setItem(i, 2, QTableWidgetItem(str(len(task["recipients"]))))

            # 执行时间
            self.task_table.setItem(i, 3, QTableWidgetItem(task["schedule_time"]))

            # 下次执行时间
            task_info = self.schedule_manager.get_task_status(task["task_name"])
            next_run = task_info["next_run"] if task_info else "未知"
            self.task_table.setItem(i, 4, QTableWidgetItem(next_run))

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
            delete_btn.clicked.connect(
                lambda checked, name=task["task_name"]: self.delete_task(name)
            )
            self.task_table.setCellWidget(i, 5, delete_btn)

    def delete_task(self, task_name):
        """删除任务"""
        reply = QMessageBox.question(
            self,
            "确认删除",
            f"确定要删除任务 '{task_name}' 吗？",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            # 从调度器中删除
            self.schedule_manager.remove_scheduled_task(task_name)

            # 从配置中删除
            if self.config_manager.remove_scheduled_task(task_name):
                QMessageBox.information(self, "成功", f"任务 '{task_name}' 已删除")
                self.load_tasks()
                self.main_window.update_status(f"任务 '{task_name}' 已删除")
            else:
                QMessageBox.warning(self, "失败", "删除任务失败")
