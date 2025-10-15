#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
自动回复模块
监控收件箱并自动回复邮件
"""

import imaplib
import email
from email.header import decode_header
from email.mime.text import MIMEText
import time
from datetime import datetime
from typing import List, Dict, Optional
import threading


class AutoReply:
    """自动回复类"""

    def __init__(self, email_address: str, password: str, imap_server: str,
                 imap_port: int, smtp_sender):
        """
        初始化自动回复

        Args:
            email_address: 邮箱地址
            password: 邮箱密码(授权码)
            imap_server: IMAP服务器地址
            imap_port: IMAP服务器端口
            smtp_sender: 邮件发送器对象
        """
        self.email_address = email_address
        self.password = password
        self.imap_server = imap_server
        self.imap_port = imap_port
        self.smtp_sender = smtp_sender
        self.replied_emails = set()  # 记录已回复的邮件ID
        self.is_running = False
        self.thread = None

    def connect_imap(self) -> Optional[imaplib.IMAP4_SSL]:
        """连接到IMAP服务器"""
        try:
            mail = imaplib.IMAP4_SSL(self.imap_server, self.imap_port)
            mail.login(self.email_address, self.password)
            return mail
        except Exception as e:
            print(f"IMAP连接失败: {e}")
            return None

    def decode_email_subject(self, subject_header) -> str:
        """解码邮件主题"""
        try:
            decoded_parts = decode_header(subject_header)
            subject = ''
            for part, encoding in decoded_parts:
                if isinstance(part, bytes):
                    subject += part.decode(encoding or 'utf-8', errors='ignore')
                else:
                    subject += str(part)
            return subject
        except:
            return subject_header

    def get_email_sender(self, msg) -> str:
        """获取发件人地址"""
        try:
            from_header = msg.get('From', '')
            if '<' in from_header and '>' in from_header:
                return from_header.split('<')[1].split('>')[0]
            return from_header
        except:
            return ''

    def check_new_emails(self, mail: imaplib.IMAP4_SSL) -> List[Dict]:
        """检查新邮件"""
        new_emails = []

        try:
            # 选择收件箱
            mail.select('INBOX')

            # 搜索未读邮件
            status, messages = mail.search(None, 'UNSEEN')

            if status != 'OK':
                return new_emails

            # 获取邮件ID列表
            email_ids = messages[0].split()

            for email_id in email_ids:
                try:
                    # 获取邮件
                    status, msg_data = mail.fetch(email_id, '(RFC822)')

                    if status != 'OK':
                        continue

                    # 解析邮件
                    msg = email.message_from_bytes(msg_data[0][1])

                    # 获取邮件信息
                    subject = self.decode_email_subject(msg.get('Subject', ''))
                    sender = self.get_email_sender(msg)
                    message_id = msg.get('Message-ID', '')

                    # 检查是否已回复过
                    if message_id and message_id not in self.replied_emails:
                        new_emails.append({
                            'email_id': email_id,
                            'message_id': message_id,
                            'sender': sender,
                            'subject': subject,
                            'date': msg.get('Date', '')
                        })

                except Exception as e:
                    print(f"处理邮件失败: {e}")
                    continue

        except Exception as e:
            print(f"检查新邮件失败: {e}")

        return new_emails

    def send_auto_reply(self, to_email: str, original_subject: str, reply_content: str) -> bool:
        """发送自动回复"""
        try:
            subject = f"Re: {original_subject}"
            result = self.smtp_sender.send_email(
                recipients=[to_email],
                subject=subject,
                content=reply_content,
                is_html=False
            )

            return len(result['success']) > 0

        except Exception as e:
            print(f"发送自动回复失败: {e}")
            return False

    def auto_reply_loop(self, reply_content: str, check_interval: int = 60):
        """自动回复循环"""
        print(f"自动回复已启动,检查间隔: {check_interval}秒")

        while self.is_running:
            try:
                mail = self.connect_imap()
                if not mail:
                    print("IMAP连接失败,等待重试...")
                    time.sleep(check_interval)
                    continue

                # 检查新邮件
                new_emails = self.check_new_emails(mail)

                # 处理新邮件
                for email_info in new_emails:
                    sender = email_info['sender']
                    subject = email_info['subject']
                    message_id = email_info['message_id']

                    print(f"收到新邮件: {sender} - {subject}")

                    # 发送自动回复
                    if self.send_auto_reply(sender, subject, reply_content):
                        print(f"已自动回复: {sender}")
                        self.replied_emails.add(message_id)
                    else:
                        print(f"自动回复失败: {sender}")

                mail.logout()

            except Exception as e:
                print(f"自动回复循环错误: {e}")

            # 等待下一次检查
            time.sleep(check_interval)

        print("自动回复已停止")

    def start(self, reply_content: str, check_interval: int = 60):
        """启动自动回复"""
        if self.is_running:
            print("自动回复已在运行中")
            return

        self.is_running = True
        self.thread = threading.Thread(
            target=self.auto_reply_loop,
            args=(reply_content, check_interval),
            daemon=True
        )
        self.thread.start()
        print("自动回复线程已启动")

    def stop(self):
        """停止自动回复"""
        if not self.is_running:
            print("自动回复未在运行")
            return

        self.is_running = False
        if self.thread:
            self.thread.join(timeout=5)
        print("自动回复已停止")

    def is_active(self) -> bool:
        """检查自动回复是否激活"""
        return self.is_running


class AutoReplyManager:
    """自动回复管理器"""

    def __init__(self):
        self.auto_replies = {}  # email -> AutoReply对象

    def add_auto_reply(self, email_address: str, auto_reply: AutoReply):
        """添加自动回复实例"""
        self.auto_replies[email_address] = auto_reply

    def start_auto_reply(self, email_address: str, reply_content: str, check_interval: int = 60) -> bool:
        """启动指定邮箱的自动回复"""
        if email_address in self.auto_replies:
            self.auto_replies[email_address].start(reply_content, check_interval)
            return True
        return False

    def stop_auto_reply(self, email_address: str) -> bool:
        """停止指定邮箱的自动回复"""
        if email_address in self.auto_replies:
            self.auto_replies[email_address].stop()
            return True
        return False

    def stop_all(self):
        """停止所有自动回复"""
        for auto_reply in self.auto_replies.values():
            auto_reply.stop()

    def get_status(self, email_address: str) -> bool:
        """获取指定邮箱的自动回复状态"""
        if email_address in self.auto_replies:
            return self.auto_replies[email_address].is_active()
        return False


if __name__ == "__main__":
    # 测试代码
    print("自动回复模块加载成功")
