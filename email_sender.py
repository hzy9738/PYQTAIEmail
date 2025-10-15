#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
邮件发送模块
支持单发和群发邮件功能
"""

import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import os
from typing import List, Optional
from datetime import datetime


class EmailSender:
    """邮件发送器类"""

    def __init__(self, email: str, password: str, smtp_server: str, smtp_port: int):
        """
        初始化邮件发送器

        Args:
            email: 发件人邮箱
            password: 邮箱密码(授权码)
            smtp_server: SMTP服务器地址
            smtp_port: SMTP服务器端口
        """
        self.email = email
        self.password = password
        self.smtp_server = smtp_server
        self.smtp_port = smtp_port

    def send_email(self, recipients: List[str], subject: str, content: str,
                   attachments: Optional[List[str]] = None, is_html: bool = False) -> dict:
        """
        发送邮件

        Args:
            recipients: 收件人列表
            subject: 邮件主题
            content: 邮件内容
            attachments: 附件路径列表
            is_html: 是否为HTML格式

        Returns:
            发送结果字典,包含成功和失败的收件人
        """
        success_list = []
        failed_list = []

        try:
            # 连接SMTP服务器
            server = smtplib.SMTP_SSL(self.smtp_server, self.smtp_port)
            server.login(self.email, self.password)

            # 为每个收件人发送邮件
            for recipient in recipients:
                try:
                    # 创建邮件对象
                    msg = MIMEMultipart()
                    msg['From'] = self.email
                    msg['To'] = recipient
                    msg['Subject'] = subject
                    msg['Date'] = datetime.now().strftime('%a, %d %b %Y %H:%M:%S +0800')

                    # 添加邮件正文
                    if is_html:
                        msg.attach(MIMEText(content, 'html', 'utf-8'))
                    else:
                        msg.attach(MIMEText(content, 'plain', 'utf-8'))

                    # 添加附件
                    if attachments:
                        for file_path in attachments:
                            if os.path.exists(file_path):
                                self._add_attachment(msg, file_path)

                    # 发送邮件
                    server.send_message(msg)
                    success_list.append(recipient)
                    print(f"成功发送邮件到: {recipient}")

                except Exception as e:
                    failed_list.append({"recipient": recipient, "error": str(e)})
                    print(f"发送邮件到 {recipient} 失败: {e}")

            server.quit()

        except Exception as e:
            print(f"SMTP连接错误: {e}")
            failed_list = [{"recipient": r, "error": f"SMTP连接错误: {e}"} for r in recipients]

        return {
            "success": success_list,
            "failed": failed_list,
            "total": len(recipients),
            "success_count": len(success_list),
            "failed_count": len(failed_list)
        }

    def _add_attachment(self, msg: MIMEMultipart, file_path: str):
        """添加附件到邮件"""
        try:
            with open(file_path, 'rb') as f:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(f.read())
                encoders.encode_base64(part)

                filename = os.path.basename(file_path)
                part.add_header('Content-Disposition', f'attachment; filename="{filename}"')
                msg.attach(part)
        except Exception as e:
            print(f"添加附件失败 {file_path}: {e}")

    def test_connection(self) -> tuple:
        """
        测试SMTP连接

        Returns:
            (是否成功, 消息)
        """
        try:
            server = smtplib.SMTP_SSL(self.smtp_server, self.smtp_port, timeout=10)
            server.login(self.email, self.password)
            server.quit()
            return (True, "连接成功")
        except smtplib.SMTPAuthenticationError:
            return (False, "认证失败,请检查邮箱和授权码")
        except smtplib.SMTPException as e:
            return (False, f"SMTP错误: {e}")
        except Exception as e:
            return (False, f"连接错误: {e}")


class BulkEmailSender:
    """批量邮件发送器"""

    def __init__(self, sender: EmailSender):
        self.sender = sender

    def send_bulk_email(self, recipients: List[str], subject: str, content: str,
                       attachments: Optional[List[str]] = None, is_html: bool = False,
                       batch_size: int = 10) -> dict:
        """
        批量发送邮件(分批发送以避免被限流)

        Args:
            recipients: 收件人列表
            subject: 邮件主题
            content: 邮件内容
            attachments: 附件列表
            is_html: 是否为HTML格式
            batch_size: 每批次发送数量

        Returns:
            发送结果统计
        """
        total_success = []
        total_failed = []

        # 分批发送
        for i in range(0, len(recipients), batch_size):
            batch = recipients[i:i + batch_size]
            print(f"正在发送第 {i//batch_size + 1} 批,共 {len(batch)} 封邮件...")

            result = self.sender.send_email(
                recipients=batch,
                subject=subject,
                content=content,
                attachments=attachments,
                is_html=is_html
            )

            total_success.extend(result["success"])
            total_failed.extend(result["failed"])

        return {
            "success": total_success,
            "failed": total_failed,
            "total": len(recipients),
            "success_count": len(total_success),
            "failed_count": len(total_failed)
        }


if __name__ == "__main__":
    # 测试代码
    print("邮件发送模块加载成功")
