#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
发送邮件标签页
支持单发和群发邮件
"""

from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
                            QLabel, QLineEdit, QTextEdit, QComboBox,
                            QMessageBox, QGroupBox, QFormLayout, QCheckBox,
                            QFileDialog, QListWidget, QProgressDialog, QTabWidget,
                            QRadioButton, QButtonGroup)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont
from email_sender import EmailSender, BulkEmailSender
from script_executor import ScriptExecutor, ScriptTemplate


class SendEmailWorker(QThread):
    """邮件发送工作线程"""
    finished = pyqtSignal(dict)
    progress = pyqtSignal(str)

    def __init__(self, sender, recipients, subject, content, attachments, is_html):
        super().__init__()
        self.sender = sender
        self.recipients = recipients
        self.subject = subject
        self.content = content
        self.attachments = attachments
        self.is_html = is_html

    def run(self):
        """执行发送"""
        try:
            bulk_sender = BulkEmailSender(self.sender)
            self.progress.emit(f"正在发送邮件到 {len(self.recipients)} 个收件人...")

            result = bulk_sender.send_bulk_email(
                recipients=self.recipients,
                subject=self.subject,
                content=self.content,
                attachments=self.attachments,
                is_html=self.is_html,
                batch_size=10
            )

            self.finished.emit(result)
        except Exception as e:
            self.finished.emit({"error": str(e)})


class SendEmailTab(QWidget):
    """发送邮件标签页"""

    def __init__(self, config_manager, main_window):
        super().__init__()
        self.config_manager = config_manager
        self.main_window = main_window
        self.attachments = []
        self.worker = None
        self.script_executor = ScriptExecutor()
        self.init_ui()

    def init_ui(self):
        """初始化UI"""
        layout = QVBoxLayout(self)

        # 发件人选择
        sender_group = QGroupBox("发件人")
        sender_group.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                border: 2px solid #3498db;
                border-radius: 5px;
                margin-top: 10px;
                padding-top: 10px;
            }
        """)
        sender_layout = QHBoxLayout()

        self.sender_combo = QComboBox()
        self.sender_combo.setMinimumHeight(30)
        sender_layout.addWidget(QLabel("选择发件账号:"))
        sender_layout.addWidget(self.sender_combo, 1)

        sender_group.setLayout(sender_layout)
        layout.addWidget(sender_group)

        # 收件人
        recipient_group = QGroupBox("收件人")
        recipient_group.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                border: 2px solid #3498db;
                border-radius: 5px;
                margin-top: 10px;
                padding-top: 10px;
            }
        """)
        recipient_layout = QVBoxLayout()

        self.recipient_input = QTextEdit()
        self.recipient_input.setPlaceholderText("请输入收件人邮箱,多个邮箱用换行、逗号或分号分隔\n例如:\nuser1@example.com\nuser2@example.com")
        self.recipient_input.setMaximumHeight(100)
        recipient_layout.addWidget(self.recipient_input)

        # 导入收件人按钮
        import_btn = QPushButton("从文件导入收件人")
        import_btn.setStyleSheet("""
            QPushButton {
                background-color: #95a5a6;
                color: white;
                padding: 5px 10px;
                border-radius: 3px;
            }
            QPushButton:hover {
                background-color: #7f8c8d;
            }
        """)
        import_btn.clicked.connect(self.import_recipients)
        recipient_layout.addWidget(import_btn)

        recipient_group.setLayout(recipient_layout)
        layout.addWidget(recipient_group)

        # 邮件内容
        content_group = QGroupBox("邮件内容")
        content_group.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                border: 2px solid #3498db;
                border-radius: 5px;
                margin-top: 10px;
                padding-top: 10px;
            }
        """)
        content_layout = QVBoxLayout()

        # 主题
        subject_layout = QHBoxLayout()
        subject_layout.addWidget(QLabel("主题:"))
        self.subject_input = QLineEdit()
        self.subject_input.setPlaceholderText("请输入邮件主题")
        subject_layout.addWidget(self.subject_input)
        content_layout.addLayout(subject_layout)

        # 内容模式选择
        mode_layout = QHBoxLayout()
        mode_layout.addWidget(QLabel("内容模式:"))

        self.mode_group = QButtonGroup()
        self.text_mode_radio = QRadioButton("普通文本")
        self.script_mode_radio = QRadioButton("Python脚本")
        self.text_mode_radio.setChecked(True)

        self.mode_group.addButton(self.text_mode_radio)
        self.mode_group.addButton(self.script_mode_radio)

        mode_layout.addWidget(self.text_mode_radio)
        mode_layout.addWidget(self.script_mode_radio)
        mode_layout.addStretch()

        # HTML格式复选框
        self.html_checkbox = QCheckBox("使用HTML格式")
        mode_layout.addWidget(self.html_checkbox)

        content_layout.addLayout(mode_layout)

        # 内容标签页(普通文本/脚本)
        self.content_tabs = QTabWidget()

        # 普通文本标签页
        text_widget = QWidget()
        text_layout = QVBoxLayout(text_widget)
        self.content_input = QTextEdit()
        self.content_input.setPlaceholderText("请输入邮件正文")
        self.content_input.setMinimumHeight(200)
        text_layout.addWidget(self.content_input)

        # Python脚本标签页
        script_widget = QWidget()
        script_layout = QVBoxLayout(script_widget)

        # 脚本模板选择
        template_layout = QHBoxLayout()
        template_layout.addWidget(QLabel("脚本模板:"))
        self.template_combo = QComboBox()
        self.template_combo.addItem("-- 选择模板 --", "")
        for template in ScriptTemplate.get_template_list():
            self.template_combo.addItem(template["name"], template)
        self.template_combo.currentIndexChanged.connect(self.load_template)
        template_layout.addWidget(self.template_combo, 1)

        # 测试脚本按钮
        test_script_btn = QPushButton("测试脚本")
        test_script_btn.setStyleSheet("""
            QPushButton {
                background-color: #f39c12;
                color: white;
                padding: 5px 15px;
                border-radius: 3px;
            }
            QPushButton:hover {
                background-color: #e67e22;
            }
        """)
        test_script_btn.clicked.connect(self.test_script)
        template_layout.addWidget(test_script_btn)

        script_layout.addLayout(template_layout)

        # 脚本编辑器
        self.script_input = QTextEdit()
        self.script_input.setPlaceholderText(
            "请输入Python脚本代码...\n\n"
            "使用方式:\n"
            "1. 使用 print() 输出内容\n"
            "2. 定义 result 变量\n"
            "3. 定义 generate_content() 函数(推荐)\n\n"
            "示例:\n"
            "def generate_content():\n"
            "    return '邮件内容'\n"
        )
        self.script_input.setMinimumHeight(200)

        # 设置等宽字体
        font = QFont("Courier New", 10)
        self.script_input.setFont(font)

        script_layout.addWidget(self.script_input)

        # 帮助文本
        help_label = QLabel(
            "💡 提示: 脚本可以读取本地文件(如Excel、CSV等)来动态生成邮件内容。\n"
            "   需要的库请自行安装,如: pip install pandas openpyxl"
        )
        help_label.setStyleSheet("color: #7f8c8d; font-size: 11px; padding: 5px;")
        script_layout.addWidget(help_label)

        # 添加标签页
        self.content_tabs.addTab(text_widget, "普通文本")
        self.content_tabs.addTab(script_widget, "Python脚本")

        # 根据单选按钮切换标签页
        self.text_mode_radio.toggled.connect(lambda checked: self.content_tabs.setCurrentIndex(0) if checked else None)
        self.script_mode_radio.toggled.connect(lambda checked: self.content_tabs.setCurrentIndex(1) if checked else None)

        content_layout.addWidget(self.content_tabs)

        content_group.setLayout(content_layout)
        layout.addWidget(content_group)

        # ���件
        attachment_group = QGroupBox("附件")
        attachment_group.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                border: 2px solid #3498db;
                border-radius: 5px;
                margin-top: 10px;
                padding-top: 10px;
            }
        """)
        attachment_layout = QVBoxLayout()

        self.attachment_list = QListWidget()
        self.attachment_list.setMaximumHeight(80)
        attachment_layout.addWidget(self.attachment_list)

        attachment_btn_layout = QHBoxLayout()

        add_attachment_btn = QPushButton("添加附件")
        add_attachment_btn.setStyleSheet("""
            QPushButton {
                background-color: #3498db;
                color: white;
                padding: 5px 10px;
                border-radius: 3px;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
        """)
        add_attachment_btn.clicked.connect(self.add_attachment)
        attachment_btn_layout.addWidget(add_attachment_btn)

        remove_attachment_btn = QPushButton("移除附件")
        remove_attachment_btn.setStyleSheet("""
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
        remove_attachment_btn.clicked.connect(self.remove_attachment)
        attachment_btn_layout.addWidget(remove_attachment_btn)

        attachment_layout.addLayout(attachment_btn_layout)
        attachment_group.setLayout(attachment_layout)
        layout.addWidget(attachment_group)

        # 发送按钮
        send_btn_layout = QHBoxLayout()
        send_btn_layout.addStretch()

        self.send_btn = QPushButton("立即发送")
        self.send_btn.setStyleSheet("""
            QPushButton {
                background-color: #27ae60;
                color: white;
                padding: 10px 30px;
                border-radius: 5px;
                font-size: 16px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #229954;
            }
            QPushButton:disabled {
                background-color: #95a5a6;
            }
        """)
        self.send_btn.clicked.connect(self.send_email)
        send_btn_layout.addWidget(self.send_btn)

        send_btn_layout.addStretch()
        layout.addLayout(send_btn_layout)

        # 加载账号
        self.refresh_accounts()

    def refresh_accounts(self):
        """刷新账号列表"""
        self.sender_combo.clear()
        accounts = self.config_manager.get_email_accounts()
        for account in accounts:
            self.sender_combo.addItem(account["email"])

    def load_template(self):
        """加载脚本模板"""
        template_data = self.template_combo.currentData()
        if template_data and isinstance(template_data, dict):
            self.script_input.setPlainText(template_data["code"])

    def test_script(self):
        """测试Python脚本"""
        script_code = self.script_input.toPlainText().strip()

        if not script_code:
            QMessageBox.warning(self, "警告", "请先输入脚本代码")
            return

        # 先验证语法
        is_valid, msg = self.script_executor.validate_script(script_code)
        if not is_valid:
            QMessageBox.critical(self, "语法错误", msg)
            return

        # 执行脚本
        self.main_window.update_status("正在测试脚本...")
        success, output = self.script_executor.execute_script(script_code)

        if success:
            # 显示测试结果
            QMessageBox.information(
                self,
                "测试成功",
                f"脚本执行成功!\n\n生成的内容:\n{'-'*40}\n{output[:500]}\n{'-'*40}\n\n"
                f"内容长度: {len(output)} 字符"
            )
            self.main_window.update_status("脚本测试成功")
        else:
            QMessageBox.critical(
                self,
                "执行失败",
                f"脚本执行失败:\n\n{output}"
            )
            self.main_window.update_status("脚本测试失败")

    def import_recipients(self):
        """从文件导入收件人"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择收件人文件",
            "",
            "文本文件 (*.txt);;所有文件 (*)"
        )

        if file_path:
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                    current_text = self.recipient_input.toPlainText()
                    if current_text:
                        self.recipient_input.setPlainText(current_text + "\n" + content)
                    else:
                        self.recipient_input.setPlainText(content)
                QMessageBox.information(self, "成功", "收件人导入成功")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"导入收件人失败:\n{str(e)}")

    def add_attachment(self):
        """添加附件"""
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "选择附件",
            "",
            "所有文件 (*)"
        )

        for file_path in files:
            if file_path not in self.attachments:
                self.attachments.append(file_path)
                self.attachment_list.addItem(file_path)

    def remove_attachment(self):
        """移除附件"""
        current_item = self.attachment_list.currentItem()
        if current_item:
            file_path = current_item.text()
            self.attachments.remove(file_path)
            self.attachment_list.takeItem(self.attachment_list.currentRow())

    def parse_recipients(self, text):
        """解析收件人列表"""
        # 支持换行、逗号、分号分隔
        recipients = []
        for line in text.split('\n'):
            for sep in [',', ';']:
                line = line.replace(sep, ' ')
            for email in line.split():
                email = email.strip()
                if email and '@' in email:
                    recipients.append(email)
        return list(set(recipients))  # 去重

    def send_email(self):
        """发送邮件"""
        # 获取发件人
        sender_email = self.sender_combo.currentText()
        if not sender_email:
            QMessageBox.warning(self, "警告", "请先添加邮箱账号")
            return

        # 获取收件人
        recipient_text = self.recipient_input.toPlainText().strip()
        if not recipient_text:
            QMessageBox.warning(self, "警告", "请输入收件人")
            return

        recipients = self.parse_recipients(recipient_text)
        if not recipients:
            QMessageBox.warning(self, "警告", "没有有效的收件人邮箱")
            return

        # 获取主题
        subject = self.subject_input.text().strip()
        if not subject:
            QMessageBox.warning(self, "警告", "请填写邮件主题")
            return

        # 获取内容(根据模式)
        content = ""
        if self.script_mode_radio.isChecked():
            # Python脚本模式
            script_code = self.script_input.toPlainText().strip()
            if not script_code:
                QMessageBox.warning(self, "警告", "请输入Python脚本")
                return

            # 先验证语法
            is_valid, msg = self.script_executor.validate_script(script_code)
            if not is_valid:
                QMessageBox.critical(self, "语法错误", f"脚本语法错误:\n{msg}")
                return

            # 执行脚本生成内容
            self.main_window.update_status("正在执行脚本生成邮件内容...")
            success, output = self.script_executor.execute_script(script_code)

            if not success:
                QMessageBox.critical(
                    self,
                    "脚本执行失败",
                    f"无法生成邮件内容:\n{output}"
                )
                self.main_window.update_status("脚本执行失败")
                return

            content = output
            self.main_window.update_status("脚本执行成功")

        else:
            # 普通文本模式
            content = self.content_input.toPlainText().strip()
            if not content:
                QMessageBox.warning(self, "警告", "请填写邮件正文")
                return

        # 确认发送
        reply = QMessageBox.question(
            self,
            "确认发送",
            f"确定要发送邮件到 {len(recipients)} 个收件人吗？",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )

        if reply != QMessageBox.Yes:
            return

        # 获取发件人凭证
        credentials = self.config_manager.get_account_credentials(sender_email)
        if not credentials:
            QMessageBox.critical(self, "错误", "获取邮箱凭证失败")
            return

        # 创建发送器
        sender = EmailSender(
            email=credentials["email"],
            password=credentials["password"],
            smtp_server=credentials["smtp_server"],
            smtp_port=credentials["smtp_port"]
        )

        # 创建工作线程
        self.worker = SendEmailWorker(
            sender=sender,
            recipients=recipients,
            subject=subject,
            content=content,
            attachments=self.attachments if self.attachments else None,
            is_html=self.html_checkbox.isChecked()
        )

        # 连接信号
        self.worker.progress.connect(self.update_progress)
        self.worker.finished.connect(self.send_finished)

        # 禁用发送按钮
        self.send_btn.setEnabled(False)
        self.main_window.update_status("正在发送邮件...")

        # 启动线程
        self.worker.start()

    def update_progress(self, message):
        """更新进度"""
        self.main_window.update_status(message)

    def send_finished(self, result):
        """发送完成"""
        self.send_btn.setEnabled(True)

        if "error" in result:
            QMessageBox.critical(self, "错误", f"发送失败:\n{result['error']}")
            self.main_window.update_status("发送失败")
        else:
            success_count = result["success_count"]
            failed_count = result["failed_count"]
            total = result["total"]

            message = f"发送完成！\n\n总计: {total}\n成功: {success_count}\n失败: {failed_count}"

            if failed_count > 0:
                failed_list = "\n".join([f"{item['recipient']}: {item['error']}" for item in result["failed"][:5]])
                message += f"\n\n部分失败的收件人:\n{failed_list}"
                if failed_count > 5:
                    message += f"\n... 还有 {failed_count - 5} 个失败"

            QMessageBox.information(self, "发送结果", message)
            self.main_window.update_status(f"发送完成: {success_count}/{total}")
