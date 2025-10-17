#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
批量数据邮件发送器模块
支持从文件夹读取多个Excel文件,使用Python脚本处理每个文件并发送邮件
"""

import os
import time
from typing import Dict, List, Optional
from datetime import datetime


class BatchDataEmailSender:
    """批量数据邮件发送器"""

    def __init__(self, email_sender, script_executor):
        """
        初始化批量数据邮件发送器

        Args:
            email_sender: EmailSender实例
            script_executor: ScriptExecutor实例
        """
        self.sender = email_sender
        self.executor = script_executor

    def scan_excel_files(self, folder_path: str) -> List[str]:
        """
        扫描文件夹中的所有Excel文件

        Args:
            folder_path: 文件夹路径

        Returns:
            Excel文件路径列表
        """
        if not os.path.exists(folder_path):
            return []

        if not os.path.isdir(folder_path):
            return []

        excel_files = []
        try:
            for file in os.listdir(folder_path):
                if file.endswith(('.xlsx', '.xls', '.xlsm')):
                    full_path = os.path.join(folder_path, file)
                    excel_files.append(full_path)
        except Exception as e:
            print(f"扫描文件夹失败: {e}")

        # 按文件名排序
        excel_files.sort()
        return excel_files

    def preview_email(self, excel_path: str, subject_template: str,
                     script_code: str, index: int, total: int) -> tuple:
        """
        预览单个邮件内容

        Args:
            excel_path: Excel文件路径
            subject_template: 主题模板
            script_code: Python脚本代码
            index: 当前序号
            total: 总数

        Returns:
            (是否成功, 主题, 内容或错误信息)
        """
        try:
            # 准备上下文
            context = self._prepare_context(excel_path, index, total)

            # 执行脚本生成内容
            success, content = self.executor.execute_script(script_code, context)
            if not success:
                return (False, "", f"脚本执行失败:\n{content}")

            # 渲染主题
            subject = self._render_template(subject_template, context)

            return (True, subject, content)

        except Exception as e:
            return (False, "", f"预览失败: {str(e)}")

    def send_batch(self, recipient: str, subject_template: str, script_code: str,
                   folder_path: str, attach_excel: bool = False, is_html: bool = False,
                   interval: int = 2, progress_callback=None) -> Dict:
        """
        批量发送数据邮件

        Args:
            recipient: 收件人邮箱
            subject_template: 主题模板(支持变量: {filename}, {index}, {total})
            script_code: Python脚本代码
            folder_path: Excel文件夹路径
            attach_excel: 是否将Excel文件作为附件
            is_html: 是否为HTML格式
            interval: 发送间隔(秒)
            progress_callback: 进度回调函数

        Returns:
            发送结果统计字典
        """
        results = {
            'success': [],
            'failed': [],
            'total': 0,
            'success_count': 0,
            'failed_count': 0
        }

        # 扫描Excel文件
        excel_files = self.scan_excel_files(folder_path)
        if not excel_files:
            results['failed'].append({
                'file': folder_path,
                'error': '未找到Excel文件'
            })
            return results

        results['total'] = len(excel_files)

        # 遍历每个Excel文件
        for index, excel_path in enumerate(excel_files, 1):
            filename = os.path.basename(excel_path)

            try:
                # 更新进度
                if progress_callback:
                    progress_callback(f"正在处理 {index}/{len(excel_files)}: {filename}")

                # 准备上下文
                context = self._prepare_context(excel_path, index, len(excel_files))

                # 执行脚本生成邮件内容
                success, content = self.executor.execute_script(script_code, context)
                if not success:
                    results['failed'].append({
                        'file': filename,
                        'error': f'脚本执行失败: {content}'
                    })
                    continue

                # 渲染主题
                subject = self._render_template(subject_template, context)

                # 准备附件
                attachments = [excel_path] if attach_excel else None

                # 发送邮件
                if progress_callback:
                    progress_callback(f"正在发送邮件: {filename}")

                send_result = self.sender.send_email(
                    recipients=[recipient],
                    subject=subject,
                    content=content,
                    attachments=attachments,
                    is_html=is_html
                )

                if send_result['success_count'] > 0:
                    results['success'].append(filename)
                    results['success_count'] += 1
                else:
                    error_msg = send_result['failed'][0]['error'] if send_result['failed'] else '未知错误'
                    results['failed'].append({
                        'file': filename,
                        'error': error_msg
                    })
                    results['failed_count'] += 1

                # 发送间隔(最后一个文件不需要等待)
                if index < len(excel_files) and interval > 0:
                    if progress_callback:
                        progress_callback(f"等待 {interval} 秒...")
                    time.sleep(interval)

            except Exception as e:
                results['failed'].append({
                    'file': filename,
                    'error': str(e)
                })
                results['failed_count'] += 1

        return results

    def _prepare_context(self, excel_path: str, index: int, total: int) -> Dict:
        """
        准备脚本执行上下文

        Args:
            excel_path: Excel文件路径
            index: 当前序号
            total: 总数

        Returns:
            上下文字典
        """
        filename_full = os.path.basename(excel_path)
        filename_noext = os.path.splitext(filename_full)[0]

        context = {
            'file': excel_path,                    # Excel完整路径
            'filename': filename_noext,            # 文件名(无扩展名)
            'filename_full': filename_full,        # 文件名(含扩展名)
            'index': index,                        # 当前序号
            'total': total,                        # 总数
            'date': datetime.now().strftime('%Y-%m-%d'),
            'time': datetime.now().strftime('%H:%M:%S'),
            'datetime': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        }

        return context

    def send_batch_multi(self, recipients: List[str], subject_template: str, script_code: str,
                         folder_path: str, is_html: bool = False,
                         interval: int = 2, progress_callback=None) -> Dict:
        """
        批量发送数据邮件到多个收件人(多对多模式)

        Args:
            recipients: 收件人邮箱列表
            subject_template: 主题模板(支持变量: {filename}, {index}, {total})
            script_code: Python脚本代码
            folder_path: Excel文件夹路径
            is_html: 是否为HTML格式
            interval: 发送间隔(秒)
            progress_callback: 进度回调函数

        Returns:
            发送结果统计字典
        """
        results = {
            'success': [],
            'failed': [],
            'total': 0,
            'success_count': 0,
            'failed_count': 0
        }

        # 扫描Excel文件
        excel_files = self.scan_excel_files(folder_path)
        if not excel_files:
            results['failed'].append({
                'file': folder_path,
                'error': '未找到Excel文件'
            })
            return results

        # 计算总邮件数
        results['total'] = len(recipients) * len(excel_files)

        email_count = 0
        # 遍历每个收件人
        for recipient in recipients:
            # 遍历每个Excel文件
            for index, excel_path in enumerate(excel_files, 1):
                filename = os.path.basename(excel_path)
                email_count += 1

                try:
                    # 更新进度
                    if progress_callback:
                        progress_callback(f"正在发送 {email_count}/{results['total']}: {recipient} - {filename}")

                    # 准备上下文
                    context = self._prepare_context(excel_path, index, len(excel_files))

                    # 执行脚本生成邮件内容
                    success, content = self.executor.execute_script(script_code, context)
                    if not success:
                        results['failed'].append({
                            'file': f"{recipient} - {filename}",
                            'error': f'脚本执行失败: {content}'
                        })
                        results['failed_count'] += 1
                        continue

                    # 渲染主题
                    subject = self._render_template(subject_template, context)

                    # 发送邮件
                    if progress_callback:
                        progress_callback(f"正在发送邮件: {recipient} - {filename}")

                    send_result = self.sender.send_email(
                        recipients=[recipient],
                        subject=subject,
                        content=content,
                        attachments=None,  # 批量数据模式不支持附件
                        is_html=is_html
                    )

                    if send_result['success_count'] > 0:
                        results['success'].append(f"{recipient} - {filename}")
                        results['success_count'] += 1
                    else:
                        error_msg = send_result['failed'][0]['error'] if send_result['failed'] else '未知错误'
                        results['failed'].append({
                            'file': f"{recipient} - {filename}",
                            'error': error_msg
                        })
                        results['failed_count'] += 1

                    # 发送间隔(最后一封邮件不需要等待)
                    if email_count < results['total'] and interval > 0:
                        if progress_callback:
                            progress_callback(f"等待 {interval} 秒...")
                        time.sleep(interval)

                except Exception as e:
                    results['failed'].append({
                        'file': f"{recipient} - {filename}",
                        'error': str(e)
                    })
                    results['failed_count'] += 1

        return results

    def _render_template(self, template: str, context: Dict) -> str:
        """
        渲染模板字符串

        Args:
            template: 模板字符串
            context: 变量字典

        Returns:
            渲染后的字符串
        """
        result = template
        for key, value in context.items():
            placeholder = f"{{{key}}}"
            result = result.replace(placeholder, str(value))
        return result


if __name__ == "__main__":
    # 测试代码
    print("批量数据邮件发送器模块加载成功")
