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
                            QRadioButton, QButtonGroup, QSpinBox)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont
from email_sender import EmailSender, BulkEmailSender
from script_executor import ScriptExecutor, ScriptTemplate
from batch_data_sender import BatchDataEmailSender


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


class BatchDataEmailWorker(QThread):
    """批量数据邮件发送工作线程"""
    finished = pyqtSignal(dict)
    progress = pyqtSignal(str)

    def __init__(self, batch_sender, recipients, subject_template, script_code,
                 folder_path, is_html, interval):
        super().__init__()
        self.batch_sender = batch_sender
        self.recipients = recipients
        self.subject_template = subject_template
        self.script_code = script_code
        self.folder_path = folder_path
        self.is_html = is_html
        self.interval = interval

    def run(self):
        """执行批量发送"""
        try:
            result = self.batch_sender.send_batch_multi(
                recipients=self.recipients,
                subject_template=self.subject_template,
                script_code=self.script_code,
                folder_path=self.folder_path,
                is_html=self.is_html,
                interval=self.interval,
                progress_callback=self.progress.emit
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
        self.batch_worker = None
        self.excel_files = []
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
        self.batch_data_mode_radio = QRadioButton("批量数据")
        self.text_mode_radio.setChecked(True)

        self.mode_group.addButton(self.text_mode_radio)
        self.mode_group.addButton(self.script_mode_radio)
        self.mode_group.addButton(self.batch_data_mode_radio)

        mode_layout.addWidget(self.text_mode_radio)
        mode_layout.addWidget(self.script_mode_radio)
        mode_layout.addWidget(self.batch_data_mode_radio)
        mode_layout.addStretch()

        # HTML格式复选框
        self.html_checkbox = QCheckBox("使用HTML格式")
        mode_layout.addWidget(self.html_checkbox)

        content_layout.addLayout(mode_layout)

        # 内容标签页(普通文本/脚本/批量数据)
        self.content_tabs = QTabWidget()

        # 普通文本标签页
        text_widget = QWidget()
        text_layout = QVBoxLayout(text_widget)
        self.content_input = QTextEdit()
        self.content_input.setPlaceholderText("请输入邮件正文")
        self.content_input.setMinimumHeight(300)  # 增大内容框高度
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

        # 批量数据标签页
        batch_data_widget = QWidget()
        batch_data_layout = QVBoxLayout(batch_data_widget)

        # 提示信息
        batch_hint_label = QLabel("💡 收件人使用上方的「收件人」输入框(支持多个邮箱)")
        batch_hint_label.setStyleSheet("color: #3498db; font-weight: bold; padding: 5px; background-color: #ecf0f1; border-radius: 3px;")
        batch_data_layout.addWidget(batch_hint_label)

        # 主题模板
        batch_subject_layout = QHBoxLayout()
        batch_subject_layout.addWidget(QLabel("主题模板:"))
        self.batch_subject_template = QLineEdit()
        self.batch_subject_template.setPlaceholderText("支持变量: {filename}, {index}, {total}, {date}")
        self.batch_subject_template.setText("数据报告 - {filename}")
        batch_subject_layout.addWidget(self.batch_subject_template)
        batch_data_layout.addLayout(batch_subject_layout)

        # Excel文件夹选择
        folder_layout = QHBoxLayout()
        folder_layout.addWidget(QLabel("Excel文件夹:"))
        self.batch_folder_input = QLineEdit()
        self.batch_folder_input.setPlaceholderText("选择包含Excel文件的文件夹")
        folder_layout.addWidget(self.batch_folder_input)

        browse_folder_btn = QPushButton("浏览...")
        browse_folder_btn.setStyleSheet("""
            QPushButton {
                background-color: #3498db;
                color: white;
                padding: 5px 15px;
                border-radius: 3px;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
        """)
        browse_folder_btn.clicked.connect(self.browse_batch_folder)
        folder_layout.addWidget(browse_folder_btn)

        scan_btn = QPushButton("扫描")
        scan_btn.setStyleSheet("""
            QPushButton {
                background-color: #27ae60;
                color: white;
                padding: 5px 15px;
                border-radius: 3px;
            }
            QPushButton:hover {
                background-color: #229954;
            }
        """)
        scan_btn.clicked.connect(self.scan_excel_files)
        folder_layout.addWidget(scan_btn)

        batch_data_layout.addLayout(folder_layout)

        # Excel文件列表
        self.batch_file_list = QListWidget()
        self.batch_file_list.setMaximumHeight(80)
        batch_data_layout.addWidget(QLabel("已找到的Excel文件:"))
        batch_data_layout.addWidget(self.batch_file_list)

        # Python脚本编辑器(批量数据专用)
        batch_script_label = QLabel("Python数据处理脚本:")
        batch_script_label.setStyleSheet("font-weight: bold; margin-top: 10px;")
        batch_data_layout.addWidget(batch_script_label)

        # 批量数据脚本模板选择
        batch_template_layout = QHBoxLayout()
        batch_template_layout.addWidget(QLabel("脚本模板:"))
        self.batch_template_combo = QComboBox()
        self.batch_template_combo.addItem("-- 选择批量数据模板 --", "")
        for template in ScriptTemplate.get_batch_data_template_list():
            self.batch_template_combo.addItem(template["name"], template)
        self.batch_template_combo.currentIndexChanged.connect(self.load_batch_template)
        batch_template_layout.addWidget(self.batch_template_combo, 1)

        # 测试批量脚本按钮
        test_batch_script_btn = QPushButton("测试脚本")
        test_batch_script_btn.setStyleSheet("""
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
        test_batch_script_btn.clicked.connect(self.test_batch_script)
        batch_template_layout.addWidget(test_batch_script_btn)

        # 预览邮件按钮
        preview_batch_btn = QPushButton("预览第1封")
        preview_batch_btn.setStyleSheet("""
            QPushButton {
                background-color: #9b59b6;
                color: white;
                padding: 5px 15px;
                border-radius: 3px;
            }
            QPushButton:hover {
                background-color: #8e44ad;
            }
        """)
        preview_batch_btn.clicked.connect(self.preview_batch_email)
        batch_template_layout.addWidget(preview_batch_btn)

        batch_data_layout.addLayout(batch_template_layout)

        self.batch_script_input = QTextEdit()
        self.batch_script_input.setPlaceholderText(
            "请输入Python脚本处理单个Excel文件...\n\n"
            "可用变量:\n"
            "  context['file']          - Excel文件完整路径\n"
            "  context['filename']      - 文件名(不含扩展名)\n"
            "  context['filename_full'] - 文件名(含扩展名)\n"
            "  context['index']         - 当前序号\n"
            "  context['total']         - 总文件数\n\n"
            "示例:\n"
            "def generate_content():\n"
            "    import pandas as pd\n"
            "    df = pd.read_excel(context['file'])\n"
            "    return f\"数据: {df.to_html()}\"\n"
        )
        self.batch_script_input.setMinimumHeight(150)
        font = QFont("Courier New", 10)
        self.batch_script_input.setFont(font)
        batch_data_layout.addWidget(self.batch_script_input)

        # 批量数据发送选项
        batch_options_layout = QHBoxLayout()

        batch_options_layout.addWidget(QLabel("发送间隔:"))
        self.batch_interval_spin = QSpinBox()
        self.batch_interval_spin.setMinimum(0)
        self.batch_interval_spin.setMaximum(60)
        self.batch_interval_spin.setValue(2)
        self.batch_interval_spin.setSuffix(" 秒")
        batch_options_layout.addWidget(self.batch_interval_spin)

        batch_options_layout.addStretch()
        batch_data_layout.addLayout(batch_options_layout)

        # 提示
        batch_help_label = QLabel(
            "💡 提示: 每个收件人会收到所有Excel文件对应的邮件 (N个收件人 × M个Excel = N×M封邮件)"
        )
        batch_help_label.setStyleSheet("color: #7f8c8d; font-size: 11px; padding: 5px;")
        batch_data_layout.addWidget(batch_help_label)

        # 添加标签页
        self.content_tabs.addTab(text_widget, "普通文本")
        self.content_tabs.addTab(script_widget, "Python脚本")
        self.content_tabs.addTab(batch_data_widget, "批量数据")

        # 根据单选按钮切换标签页
        self.text_mode_radio.toggled.connect(lambda checked: self.on_mode_changed(0) if checked else None)
        self.script_mode_radio.toggled.connect(lambda checked: self.on_mode_changed(1) if checked else None)
        self.batch_data_mode_radio.toggled.connect(lambda checked: self.on_mode_changed(2) if checked else None)

        # 根据标签页切换单选按钮（双向同步）
        self.content_tabs.currentChanged.connect(self.on_content_tab_changed)

        content_layout.addWidget(self.content_tabs)

        content_group.setLayout(content_layout)
        layout.addWidget(content_group)

        # 附件
        self.attachment_group = QGroupBox("附件")
        self.attachment_group.setStyleSheet("""
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
        self.attachment_list.setMaximumHeight(60)  # 减小附件框高度
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
        self.attachment_group.setLayout(attachment_layout)
        layout.addWidget(self.attachment_group)

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

    def on_mode_changed(self, mode_index):
        """模式切换时的处理"""
        # 切换标签页
        self.content_tabs.setCurrentIndex(mode_index)

        # 批量数据模式下隐藏附件区域
        if mode_index == 2:  # 批量数据模式
            self.attachment_group.setVisible(False)
        else:
            self.attachment_group.setVisible(True)

    def on_content_tab_changed(self, index):
        """标签页切换时同步单选按钮"""
        if index == 0:
            self.text_mode_radio.setChecked(True)
        elif index == 1:
            self.script_mode_radio.setChecked(True)
        elif index == 2:
            self.batch_data_mode_radio.setChecked(True)

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
        # 如果是批量数据模式,调用专门的方法
        if self.batch_data_mode_radio.isChecked():
            self.send_batch_data_email()
            return

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

    def browse_batch_folder(self):
        """浏览选择Excel文件夹"""
        folder_path = QFileDialog.getExistingDirectory(
            self,
            "选择Excel文件夹",
            "",
            QFileDialog.ShowDirsOnly
        )
        if folder_path:
            self.batch_folder_input.setText(folder_path)
            self.scan_excel_files()

    def scan_excel_files(self):
        """扫描文件夹中的Excel文件"""
        folder_path = self.batch_folder_input.text().strip()
        if not folder_path:
            QMessageBox.warning(self, "警告", "请先选择Excel文件夹")
            return

        # 创建临时的BatchDataEmailSender来扫描文件
        from email_sender import EmailSender
        temp_sender = EmailSender("", "", "", 0)
        batch_sender = BatchDataEmailSender(temp_sender, self.script_executor)

        self.excel_files = batch_sender.scan_excel_files(folder_path)

        # 更新列表
        self.batch_file_list.clear()
        if self.excel_files:
            for file_path in self.excel_files:
                import os
                self.batch_file_list.addItem(os.path.basename(file_path))
            self.main_window.update_status(f"找到 {len(self.excel_files)} 个Excel文件")
        else:
            self.main_window.update_status("未找到Excel文件")
            QMessageBox.information(self, "提示", "该文件夹中没有Excel文件(.xlsx, .xls, .xlsm)")

    def load_batch_template(self):
        """加载批量数据脚本模板"""
        template_data = self.batch_template_combo.currentData()
        if template_data and isinstance(template_data, dict):
            self.batch_script_input.setPlainText(template_data["code"])

    def test_batch_script(self):
        """测试批量数据脚本"""
        script_code = self.batch_script_input.toPlainText().strip()

        if not script_code:
            QMessageBox.warning(self, "警告", "请先输入脚本代码")
            return

        if not self.excel_files:
            QMessageBox.warning(self, "警告", "请先扫描Excel文件")
            return

        # 先验证语法
        is_valid, msg = self.script_executor.validate_script(script_code)
        if not is_valid:
            QMessageBox.critical(self, "语法错误", msg)
            return

        # 使用第一个Excel文件测试
        test_file = self.excel_files[0]
        import os
        context = {
            'file': test_file,
            'filename': os.path.splitext(os.path.basename(test_file))[0],
            'filename_full': os.path.basename(test_file),
            'index': 1,
            'total': len(self.excel_files)
        }

        # 执行脚本
        self.main_window.update_status("正在测试脚本...")
        success, output = self.script_executor.execute_script(script_code, context)

        if success:
            # 显示测试结果
            QMessageBox.information(
                self,
                "测试成功",
                f"脚本执行成功!\n\n测试文件: {os.path.basename(test_file)}\n\n"
                f"生成的内容:\n{'-'*40}\n{output[:500]}\n{'-'*40}\n\n"
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

    def preview_batch_email(self):
        """预览批量数据邮件"""
        if not self.excel_files:
            QMessageBox.warning(self, "警告", "请先扫描Excel文件")
            return

        script_code = self.batch_script_input.toPlainText().strip()
        if not script_code:
            QMessageBox.warning(self, "警告", "请输入Python脚本")
            return

        subject_template = self.batch_subject_template.text().strip()
        if not subject_template:
            QMessageBox.warning(self, "警告", "请输入主题模板")
            return

        # 创建BatchDataEmailSender
        from email_sender import EmailSender
        temp_sender = EmailSender("", "", "", 0)
        batch_sender = BatchDataEmailSender(temp_sender, self.script_executor)

        # 预览第一封邮件
        success, subject, content = batch_sender.preview_email(
            self.excel_files[0],
            subject_template,
            script_code,
            1,
            len(self.excel_files)
        )

        if success:
            import os
            preview_text = f"【预览第1封邮件】\n\n"
            preview_text += f"文件: {os.path.basename(self.excel_files[0])}\n"
            preview_text += f"主题: {subject}\n\n"
            preview_text += f"内容预览:\n{'-'*50}\n{content[:1000]}\n"
            if len(content) > 1000:
                preview_text += f"\n... (还有 {len(content)-1000} 字符)"

            QMessageBox.information(self, "邮件预览", preview_text)
        else:
            QMessageBox.critical(self, "预览失败", content)

    def send_batch_data_email(self):
        """发送批量数据邮件"""
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

        # 检查Excel文件
        if not self.excel_files:
            QMessageBox.warning(self, "警告", "请先扫描Excel文件")
            return

        # 获取主题模板
        subject_template = self.batch_subject_template.text().strip()
        if not subject_template:
            QMessageBox.warning(self, "警告", "请输入主题模板")
            return

        # 获取脚本
        script_code = self.batch_script_input.toPlainText().strip()
        if not script_code:
            QMessageBox.warning(self, "警告", "请输入Python脚本")
            return

        # 验证脚本语法
        is_valid, msg = self.script_executor.validate_script(script_code)
        if not is_valid:
            QMessageBox.critical(self, "语法错误", f"脚本语法错误:\n{msg}")
            return

        # 计算总邮件数
        total_emails = len(recipients) * len(self.excel_files)

        # 确认发送
        reply = QMessageBox.question(
            self,
            "确认发送",
            f"确定要发送邮件吗？\n\n"
            f"收件人: {len(recipients)} 个\n"
            f"Excel文件: {len(self.excel_files)} 个\n"
            f"总邮件数: {total_emails} 封\n\n"
            f"每个收件人将收到所有Excel文件对应的邮件。",
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

        batch_sender = BatchDataEmailSender(sender, self.script_executor)

        # 创建工作线程
        self.batch_worker = BatchDataEmailWorker(
            batch_sender=batch_sender,
            recipients=recipients,
            subject_template=subject_template,
            script_code=script_code,
            folder_path=self.batch_folder_input.text().strip(),
            is_html=self.html_checkbox.isChecked(),
            interval=self.batch_interval_spin.value()
        )

        # 连接信号
        self.batch_worker.progress.connect(self.update_progress)
        self.batch_worker.finished.connect(self.batch_send_finished)

        # 禁用发送按钮
        self.send_btn.setEnabled(False)
        self.main_window.update_status("正在批量发送邮件...")

        # 启动线程
        self.batch_worker.start()

    def batch_send_finished(self, result):
        """批量发送完成"""
        self.send_btn.setEnabled(True)

        if "error" in result:
            QMessageBox.critical(self, "错误", f"批量发送失败:\n{result['error']}")
            self.main_window.update_status("批量发送失败")
        else:
            success_count = result.get("success_count", len(result.get("success", [])))
            failed_count = result.get("failed_count", len(result.get("failed", [])))
            total = result.get("total", 0)

            message = f"批量发送完成！\n\n总计: {total} 封\n成功: {success_count} 封\n失败: {failed_count} 封"

            if failed_count > 0:
                failed_items = result.get("failed", [])
                failed_list = "\n".join([f"{item['file']}: {item['error']}" for item in failed_items[:5]])
                message += f"\n\n部分失败的文件:\n{failed_list}"
                if failed_count > 5:
                    message += f"\n... 还有 {failed_count - 5} 个失败"

            QMessageBox.information(self, "发送结果", message)
            self.main_window.update_status(f"批量发送完成: {success_count}/{total}")
