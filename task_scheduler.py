#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
定时任务调度模块
使用schedule库实现邮件定时发送功能
"""

import schedule
import time
import threading
from datetime import datetime
from typing import List, Dict, Callable
from email_sender import EmailSender


class TaskScheduler:
    """定时任务调度器"""

    def __init__(self):
        """初始化调度器"""
        self.is_running = False
        self.thread = None
        self.tasks = {}  # task_name -> schedule.Job
        self.task_callbacks = {}  # 任务执行回调

    def add_task(self, task_name: str, schedule_time: str, sender: EmailSender,
                 recipients: List[str], subject: str, content: str,
                 attachments: List[str] = None, is_html: bool = False,
                 callback: Callable = None):
        """
        添加定时任务

        Args:
            task_name: 任务名称
            schedule_time: 执行时间,格式如 "09:30", "14:00"
            sender: 邮件发送器
            recipients: 收件人列表
            subject: 邮件主题
            content: 邮件内容
            attachments: 附件列表
            is_html: 是否HTML格式
            callback: 任务执行后的回调函数
        """
        def task_function():
            """任务执行函数"""
            print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] 执行定时任务: {task_name}")

            try:
                result = sender.send_email(
                    recipients=recipients,
                    subject=subject,
                    content=content,
                    attachments=attachments,
                    is_html=is_html
                )

                print(f"任务 {task_name} 执行完成: 成功 {result['success_count']}, 失败 {result['failed_count']}")

                # 执行回调
                if callback:
                    callback(task_name, result)

            except Exception as e:
                print(f"任务 {task_name} 执行失败: {e}")
                if callback:
                    callback(task_name, {"error": str(e)})

        # 创建定时任务
        job = schedule.every().day.at(schedule_time).do(task_function)
        self.tasks[task_name] = job

        if callback:
            self.task_callbacks[task_name] = callback

        print(f"已添加定时任务: {task_name}, 执行时间: 每天 {schedule_time}")

    def remove_task(self, task_name: str) -> bool:
        """删除定时任务"""
        if task_name in self.tasks:
            schedule.cancel_job(self.tasks[task_name])
            del self.tasks[task_name]

            if task_name in self.task_callbacks:
                del self.task_callbacks[task_name]

            print(f"已删除定时任务: {task_name}")
            return True
        return False

    def get_task_info(self, task_name: str) -> Dict:
        """获取任务信息"""
        if task_name in self.tasks:
            job = self.tasks[task_name]
            return {
                "task_name": task_name,
                "next_run": str(job.next_run) if job.next_run else "未知",
                "last_run": str(job.last_run) if job.last_run else "从未执行"
            }
        return None

    def get_all_tasks(self) -> List[Dict]:
        """获取所有任务信息"""
        return [self.get_task_info(name) for name in self.tasks.keys()]

    def _run_pending(self):
        """运行待执行的任务循环"""
        print("定时任务调度器已启动")

        while self.is_running:
            schedule.run_pending()
            time.sleep(1)

        print("定时任务调度器已停止")

    def start(self):
        """启动调度器"""
        if self.is_running:
            print("调度器已在运行中")
            return

        self.is_running = True
        self.thread = threading.Thread(target=self._run_pending, daemon=True)
        self.thread.start()
        print("定时任务调度器线程已启动")

    def stop(self):
        """停止调度器"""
        if not self.is_running:
            print("调度器未在运行")
            return

        self.is_running = False
        if self.thread:
            self.thread.join(timeout=3)
        print("定时任务调度器已停止")

    def clear_all_tasks(self):
        """清除所有任务"""
        schedule.clear()
        self.tasks.clear()
        self.task_callbacks.clear()
        print("已清除所有定时任务")

    def is_active(self) -> bool:
        """检查调度器是否激活"""
        return self.is_running


class ScheduleManager:
    """定时任务管理器"""

    def __init__(self):
        self.scheduler = TaskScheduler()
        self.email_senders = {}  # email -> EmailSender

    def add_email_sender(self, email: str, sender: EmailSender):
        """添加邮件发送器"""
        self.email_senders[email] = sender

    def add_scheduled_task(self, task_name: str, schedule_time: str,
                          sender_email: str, recipients: List[str],
                          subject: str, content: str, attachments: List[str] = None,
                          is_html: bool = False, callback: Callable = None) -> bool:
        """添加定时任务"""
        if sender_email not in self.email_senders:
            print(f"发件人邮箱 {sender_email} 未配置")
            return False

        sender = self.email_senders[sender_email]

        self.scheduler.add_task(
            task_name=task_name,
            schedule_time=schedule_time,
            sender=sender,
            recipients=recipients,
            subject=subject,
            content=content,
            attachments=attachments,
            is_html=is_html,
            callback=callback
        )

        return True

    def remove_scheduled_task(self, task_name: str) -> bool:
        """删除定时任务"""
        return self.scheduler.remove_task(task_name)

    def start_scheduler(self):
        """启动调度器"""
        self.scheduler.start()

    def stop_scheduler(self):
        """停止调度器"""
        self.scheduler.stop()

    def get_task_status(self, task_name: str) -> Dict:
        """获取任务状态"""
        return self.scheduler.get_task_info(task_name)

    def get_all_tasks_status(self) -> List[Dict]:
        """获取所有任务状态"""
        return self.scheduler.get_all_tasks()

    def is_running(self) -> bool:
        """检查调度器是否运行"""
        return self.scheduler.is_active()


if __name__ == "__main__":
    # 测试代码
    print("定时任务调度模块加载成功")
