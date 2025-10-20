#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
配置管理模块
负责读取、保存和加密邮箱配置信息
"""

import json
import os
from cryptography.fernet import Fernet
from typing import List, Dict, Optional


class ConfigManager:
    """配置管理器类"""

    def __init__(self, config_file: str = "config.json"):
        self.config_file = config_file
        self.key_file = ".secret.key"
        self.cipher = self._load_or_create_cipher()
        self.config = self._load_config()

    def _load_or_create_cipher(self) -> Fernet:
        """加载或创建加密密钥"""
        if os.path.exists(self.key_file):
            with open(self.key_file, 'rb') as f:
                key = f.read()
        else:
            key = Fernet.generate_key()
            with open(self.key_file, 'wb') as f:
                f.write(key)
        return Fernet(key)

    def _load_config(self) -> Dict:
        """加载配置文件"""
        if os.path.exists(self.config_file):
            with open(self.config_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        else:
            # 创建默认配置
            default_config = {
                "email_accounts": [],
                "auto_reply": {
                    "enabled": False,
                    "reply_content": "感谢您的来信,我会尽快回复。"
                },
                "scheduled_tasks": [],
                "send_email_state": {
                    "recipients": "",
                    "subject": "",
                    "text_content": "",
                    "script_content": "",
                    "batch_subject_template": "数据报告 - {filename}",
                    "batch_folder_path": "",
                    "batch_script_content": "",
                    "mode": "text",
                    "html_enabled": False
                }
            }
            self._save_config(default_config)
            return default_config

    def _save_config(self, config: Dict = None):
        """保存配置文件"""
        if config is None:
            config = self.config
        with open(self.config_file, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=2)

    def encrypt_password(self, password: str) -> str:
        """加密密码"""
        return self.cipher.encrypt(password.encode()).decode()

    def decrypt_password(self, encrypted_password: str) -> str:
        """解密密码"""
        return self.cipher.decrypt(encrypted_password.encode()).decode()

    def add_email_account(self, email: str, password: str, smtp_server: str = "smtp.163.com",
                         smtp_port: int = 465, imap_password: str = None) -> bool:
        """添加邮箱账号"""
        try:
            # 检查是否已存在
            for account in self.config["email_accounts"]:
                if account["email"] == email:
                    return False

            # 加密密码
            encrypted_pwd = self.encrypt_password(password)

            # 如果没有提供IMAP授权码，则使用SMTP授权码
            if imap_password is None:
                imap_password = password
            encrypted_imap_pwd = self.encrypt_password(imap_password)

            # 添加账号
            self.config["email_accounts"].append({
                "email": email,
                "password": encrypted_pwd,
                "imap_password": encrypted_imap_pwd,
                "smtp_server": smtp_server,
                "smtp_port": smtp_port,
                "imap_server": "imap.163.com",
                "imap_port": 993
            })

            self._save_config()
            return True
        except Exception as e:
            print(f"添加邮箱账号失败: {e}")
            return False

    def remove_email_account(self, email: str) -> bool:
        """删除邮箱账号"""
        try:
            self.config["email_accounts"] = [
                acc for acc in self.config["email_accounts"]
                if acc["email"] != email
            ]
            self._save_config()
            return True
        except Exception as e:
            print(f"删除邮箱账号失败: {e}")
            return False

    def get_email_accounts(self) -> List[Dict]:
        """获取所有邮箱账号"""
        return self.config["email_accounts"]

    def get_account_credentials(self, email: str) -> Optional[Dict]:
        """获取账号凭证(包含解密后的密码和IMAP密码)"""
        for account in self.config["email_accounts"]:
            if account["email"] == email:
                credentials = account.copy()
                credentials["password"] = self.decrypt_password(account["password"])

                # 获取IMAP密码，如果不存在则使用SMTP密码
                if "imap_password" in account:
                    credentials["imap_password"] = self.decrypt_password(account["imap_password"])
                else:
                    credentials["imap_password"] = credentials["password"]

                return credentials
        return None

    def set_auto_reply(self, enabled: bool, reply_content: str = None):
        """设置自动回复"""
        self.config["auto_reply"]["enabled"] = enabled
        if reply_content:
            self.config["auto_reply"]["reply_content"] = reply_content
        self._save_config()

    def get_auto_reply_config(self) -> Dict:
        """获取自动回复配置"""
        return self.config["auto_reply"]

    def add_scheduled_task(self, task_name: str, recipients: List[str], subject: str,
                          content: str, schedule_time: str, sender_email: str) -> bool:
        """添加定时任务"""
        try:
            task = {
                "task_name": task_name,
                "recipients": recipients,
                "subject": subject,
                "content": content,
                "schedule_time": schedule_time,
                "sender_email": sender_email,
                "enabled": True
            }
            self.config["scheduled_tasks"].append(task)
            self._save_config()
            return True
        except Exception as e:
            print(f"添加定时任务失败: {e}")
            return False

    def remove_scheduled_task(self, task_name: str) -> bool:
        """删除定时任务"""
        try:
            self.config["scheduled_tasks"] = [
                task for task in self.config["scheduled_tasks"]
                if task["task_name"] != task_name
            ]
            self._save_config()
            return True
        except Exception as e:
            print(f"删除定时任务失败: {e}")
            return False

    def get_scheduled_tasks(self) -> List[Dict]:
        """获取所有定时任务"""
        return self.config["scheduled_tasks"]

    def update_task_status(self, task_name: str, enabled: bool):
        """更新任务状态"""
        for task in self.config["scheduled_tasks"]:
            if task["task_name"] == task_name:
                task["enabled"] = enabled
        self._save_config()

    def save_send_email_state(self, state: Dict) -> bool:
        """保存发送邮件页面的状态"""
        try:
            # 确保send_email_state字段存在
            if "send_email_state" not in self.config:
                self.config["send_email_state"] = {}
            
            # 更新状态
            self.config["send_email_state"].update(state)
            self._save_config()
            return True
        except Exception as e:
            print(f"保存发送邮件状态失败: {e}")
            return False

    def get_send_email_state(self) -> Dict:
        """获取发送邮件页面的保存状态"""
        default_state = {
            "recipients": "",
            "subject": "",
            "text_content": "",
            "script_content": "",
            "batch_subject_template": "数据报告 - {filename}",
            "batch_folder_path": "",
            "batch_script_content": "",
            "mode": "text",
            "html_enabled": False
        }
        
        if "send_email_state" not in self.config:
            self.config["send_email_state"] = default_state
        
        return self.config.get("send_email_state", default_state)


if __name__ == "__main__":
    # 测试代码
    config = ConfigManager()
    print("配置管理器初始化成功")
