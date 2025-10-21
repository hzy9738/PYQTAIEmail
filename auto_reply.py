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
from datetime import datetime, timedelta
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
        self.replied_emails = set()  # 记录已��复的邮件ID
        self.is_running = False
        self.thread = None
        # 更新docstring - password应为IMAP授权码
        # IMAP连接时password需要使用IMAP授权码

    def connect_imap(self) -> Optional[imaplib.IMAP4_SSL]:
        """连接到IMAP服务器"""
        try:
            mail = imaplib.IMAP4_SSL(self.imap_server, self.imap_port, timeout=10)
            mail.login(self.email_address, self.password)
            print(f"IMAP连接成功: {self.email_address}")
            return mail
        except imaplib.IMAP4.error as e:
            print(f"IMAP认证失败: {e}")
            print(f"请检查是否使用了正确的IMAP授权码(非登录密码)")
            return None
        except Exception as e:
            print(f"IMAP连接失败: {e}")
            return None

    def test_connection(self) -> tuple:
        """测试IMAP连接

        Returns:
            (是否成功, 错误信息)
        """
        try:
            mail = imaplib.IMAP4_SSL(self.imap_server, self.imap_port, timeout=10)
            mail.login(self.email_address, self.password)
            mail.logout()
            return (True, "IMAP连接成功")
        except imaplib.IMAP4.error as e:
            error_msg = f"IMAP认证失败: {str(e)}\n请确保:\n1. 已开启IMAP服务\n2. 使用的是IMAP授权码(非登录密码)\n3. 授权码输入正确"
            print(error_msg)
            return (False, error_msg)
        except TimeoutError:
            error_msg = f"IMAP连接超时: 无法连接到{self.imap_server}:{self.imap_port}\n请检查网络连接和防火墙设置"
            print(error_msg)
            return (False, error_msg)
        except Exception as e:
            error_msg = f"IMAP连接失败: {str(e)}"
            print(error_msg)
            return (False, error_msg)

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
            status, data = mail.select('INBOX')
            print(f"[调试] INBOX中共有 {data[0].decode()} 封邮件")

            # 搜索所有邮件，然后手动过滤最近的
            # 不使用SINCE命令，因为可能不兼容某些IMAP服务器
            print("[调试] 正在搜索所有邮件...")
            status, messages = mail.search(None, 'ALL')

            if status != 'OK':
                print(f"[调试] 搜索失败: {status}")
                return new_emails

            # 获取邮件ID列表
            email_ids = messages[0].split()
            total_emails = len(email_ids)
            print(f"[调试] 搜索到 {total_emails} 封邮件")

            # 只处理最近的20封邮件（倒序）
            recent_email_ids = email_ids[-20:] if len(email_ids) > 20 else email_ids
            recent_email_ids = list(reversed(recent_email_ids))  # 从最新的开始
            print(f"[调试] 将检查最近的 {len(recent_email_ids)} 封邮件")

            checked_count = 0
            filtered_count = 0

            for email_id in recent_email_ids:
                try:
                    # 获取邮件头信息（更快，不获取整个邮件体）
                    status, msg_data = mail.fetch(email_id, '(BODY.PEEK[HEADER])')

                    if status != 'OK':
                        continue

                    # 解析邮件
                    msg = email.message_from_bytes(msg_data[0][1])

                    # 获取邮件信息
                    subject = self.decode_email_subject(msg.get('Subject', ''))
                    sender = self.get_email_sender(msg)
                    message_id = msg.get('Message-ID', '')
                    date_str = msg.get('Date', '')

                    checked_count += 1

                    # 检查是否已回复过
                    if message_id and message_id not in self.replied_emails:
                        new_emails.append({
                            'email_id': email_id,
                            'message_id': message_id,
                            'sender': sender,
                            'subject': subject,
                            'date': date_str
                        })
                        print(f"[调试] 发现新邮件: {sender} - {subject[:30]}")
                    else:
                        filtered_count += 1
                        print(f"[调试] 已回复过，跳过: {sender} - {subject[:30]}")

                except Exception as e:
                    print(f"[调试] 处理邮件失败: {e}")
                    continue

            print(f"[调试] 检查完成: 检查了{checked_count}封, 过滤了{filtered_count}封, 发现{len(new_emails)}封新邮件")

        except Exception as e:
            print(f"[错误] 检查新邮件失败: {e}")
            import traceback
            traceback.print_exc()

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
        print(f"="*50)
        print(f"自动回复已启动")
        print(f"检查间隔: {check_interval}秒")
        print(f"回复内容: {reply_content[:50]}...")
        print(f"="*50)

        loop_count = 0

        while self.is_running:
            try:
                loop_count += 1
                print(f"\n[循环 #{loop_count}] {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
                print(f"[循环 #{loop_count}] 正在连接IMAP服务器...")

                mail = self.connect_imap()
                if not mail:
                    print(f"[循环 #{loop_count}] IMAP连接失败,等待重试...")
                    time.sleep(check_interval)
                    continue

                print(f"[循环 #{loop_count}] IMAP连接成功，开始检查新邮件...")

                # 检查新邮件
                new_emails = self.check_new_emails(mail)

                if len(new_emails) == 0:
                    print(f"[循环 #{loop_count}] 没有发现需要回复的新邮件")
                else:
                    print(f"[循环 #{loop_count}] 发现 {len(new_emails)} 封需要回复的新邮件")

                # 处理新邮件
                for i, email_info in enumerate(new_emails, 1):
                    sender = email_info['sender']
                    subject = email_info['subject']
                    message_id = email_info['message_id']

                    print(f"[循环 #{loop_count}] [{i}/{len(new_emails)}] 处理邮件: {sender} - {subject}")

                    # 发送自动回复
                    print(f"[循环 #{loop_count}] [{i}/{len(new_emails)}] 正在发送自动回复...")
                    if self.send_auto_reply(sender, subject, reply_content):
                        print(f"[循环 #{loop_count}] [{i}/{len(new_emails)}] ✓ 自动回复成功: {sender}")
                        self.replied_emails.add(message_id)
                    else:
                        print(f"[循环 #{loop_count}] [{i}/{len(new_emails)}] ✗ 自动回复失败: {sender}")

                mail.logout()
                print(f"[循环 #{loop_count}] IMAP连接已关闭")
                print(f"[循环 #{loop_count}] 等待 {check_interval} 秒后进行下一次检查...")

            except Exception as e:
                print(f"[循环 #{loop_count}] 错误: {e}")
                import traceback
                traceback.print_exc()

            # 等待下一次检查
            time.sleep(check_interval)

        print("\n" + "="*50)
        print("自动回复已停止")
        print("="*50)

    def start(self, reply_content: str, check_interval: int = 60):
        """启动自动回复"""
        if self.is_running:
            print("自动回复已在运行中")
            return

        # 清空已回复邮件记录，确保每次启动都是全新状态
        self.replied_emails.clear()
        print(f"已清空之前的回复记录，当前replied_emails有 {len(self.replied_emails)} 条记录")

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
