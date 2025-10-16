#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Python脚本执行器模块
支持执行用户自定义Python脚本来动态生成邮件内容
"""

import sys
import io
import traceback
from typing import Dict, Any


class ScriptExecutor:
    """Python脚本执行器"""

    def __init__(self):
        """初始化执行器"""
        self.last_error = None

    def execute_script(self, script_code: str, context: Dict[str, Any] = None) -> tuple:
        """
        执行Python脚本并返回结果

        Args:
            script_code: 要执行的Python脚本代码
            context: 传递给脚本的上下文变量字典

        Returns:
            (是否成功, 结果内容或错误信息)
        """
        if not script_code or not script_code.strip():
            return (False, "脚本内容为空")

        # 准备执行环境
        if context is None:
            context = {}

        # 添加常用模块到执行环境
        exec_globals = {
            '__builtins__': __builtins__,
            'context': context,  # 用户可以通过context访问传入的变量
        }

        # 捕获标准输出
        old_stdout = sys.stdout
        sys.stdout = captured_output = io.StringIO()

        try:
            # 执行脚本
            exec(script_code, exec_globals)

            # 获取输出内容
            output = captured_output.getvalue()

            # 如果脚本定义了generate_content函数,优先使用其返回值
            if 'generate_content' in exec_globals:
                result = exec_globals['generate_content']()
                if result is not None:
                    output = str(result)

            # 如果脚本定义了result变量,使用该变量
            elif 'result' in exec_globals:
                result = exec_globals['result']
                if result is not None:
                    output = str(result)

            # 恢复标准输出
            sys.stdout = old_stdout

            if not output:
                return (False, "脚本未产生任何输出")

            return (True, output)

        except Exception as e:
            # 恢复标准输出
            sys.stdout = old_stdout

            # 获取详细错误信息
            error_msg = traceback.format_exc()
            self.last_error = error_msg

            return (False, f"脚本执行错误:\n{error_msg}")

    def validate_script(self, script_code: str) -> tuple:
        """
        验证脚本语法

        Args:
            script_code: Python脚本代码

        Returns:
            (是否有效, 错误信息)
        """
        try:
            compile(script_code, '<string>', 'exec')
            return (True, "语法正确")
        except SyntaxError as e:
            return (False, f"语法错误(第{e.lineno}行): {e.msg}")
        except Exception as e:
            return (False, f"验证错误: {str(e)}")


class ScriptTemplate:
    """脚本模板库"""

    @staticmethod
    def get_template_list() -> list:
        """获取所有可用模板"""
        return [
            {
                "name": "基础示例",
                "description": "展示基本的输出方式",
                "code": ScriptTemplate.basic_example()
            },
            {
                "name": "读取Excel文件",
                "description": "读取Excel文件并提取内容",
                "code": ScriptTemplate.read_excel()
            },
            {
                "name": "读取多个Excel文件",
                "description": "读取目录下所有Excel文件并汇总",
                "code": ScriptTemplate.read_multiple_excel()
            },
            {
                "name": "读取CSV文件",
                "description": "读取CSV文件内容",
                "code": ScriptTemplate.read_csv()
            },
            {
                "name": "读取文本文件",
                "description": "读取并处理文本文件",
                "code": ScriptTemplate.read_text_file()
            },
            {
                "name": "动态日期时间",
                "description": "生成包含当前日期时间的内容",
                "code": ScriptTemplate.datetime_example()
            }
        ]

    @staticmethod
    def basic_example() -> str:
        """基础示例模板"""
        return """# 基础示例 - 三种输出方式

# 方式1: 使用print()输出
print("这是通过print输出的内容")

# 方式2: 定义result变量
# result = "这是通过result变量输出的内容"

# 方式3: 定义generate_content()函数(推荐)
# def generate_content():
#     return "这是通过generate_content函数输出的内容"
"""

    @staticmethod
    def read_excel() -> str:
        """读取Excel模板"""
        return """# 读取Excel文件示例
# 需要安装: pip install openpyxl pandas

import pandas as pd
from datetime import datetime

def generate_content():
    # 设置Excel文件路径
    excel_path = r"C:\\Users\\YourName\\Documents\\data.xlsx"

    try:
        # 读取Excel文件
        df = pd.read_excel(excel_path, sheet_name=0)

        # 生成邮件内容
        content = f"数据报表\\n"
        content += f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\\n\\n"
        content += f"数据概览:\\n"
        content += f"总行数: {len(df)}\\n"
        content += f"列名: {', '.join(df.columns)}\\n\\n"

        # 显示前5行数据
        content += "数据预览:\\n"
        content += df.head().to_string()

        return content

    except Exception as e:
        return f"读取Excel文件失败: {str(e)}"
"""

    @staticmethod
    def read_multiple_excel() -> str:
        """读取多个Excel文件模板"""
        return """# 读取目录下所有Excel文件
# 需要安装: pip install openpyxl pandas

import pandas as pd
import os
from datetime import datetime

def generate_content():
    # 设置目录路径
    folder_path = r"C:\\Users\\YourName\\Documents\\excel_files"

    try:
        content = f"Excel文件汇总报告\\n"
        content += f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\\n"
        content += f"目录: {folder_path}\\n\\n"

        # 查找所有Excel文件
        excel_files = [f for f in os.listdir(folder_path)
                      if f.endswith(('.xlsx', '.xls'))]

        content += f"找到 {len(excel_files)} 个Excel文件\\n\\n"

        # 读取每个文件
        for excel_file in excel_files:
            file_path = os.path.join(folder_path, excel_file)
            df = pd.read_excel(file_path)

            content += f"文件: {excel_file}\\n"
            content += f"  行数: {len(df)}\\n"
            content += f"  列数: {len(df.columns)}\\n"
            content += f"  列名: {', '.join(df.columns)}\\n\\n"

        return content

    except Exception as e:
        return f"读取Excel文件失败: {str(e)}"
"""

    @staticmethod
    def read_csv() -> str:
        """读取CSV文件模板"""
        return """# 读取CSV文件示例
# 需要安装: pip install pandas

import pandas as pd
from datetime import datetime

def generate_content():
    # 设置CSV文件路径
    csv_path = r"C:\\Users\\YourName\\Documents\\data.csv"

    try:
        # 读取CSV文件
        df = pd.read_csv(csv_path, encoding='utf-8')

        # 生成邮件内容
        content = f"CSV数据报表\\n"
        content += f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\\n\\n"
        content += f"总记录数: {len(df)}\\n\\n"

        # 统计信息
        content += "数据统计:\\n"
        content += df.describe().to_string()

        return content

    except Exception as e:
        return f"读取CSV文件失败: {str(e)}"
"""

    @staticmethod
    def read_text_file() -> str:
        """读取文本文件模板"""
        return """# 读取文本文件示例

def generate_content():
    # 设置文本文件路径
    file_path = r"C:\\Users\\YourName\\Documents\\report.txt"

    try:
        # 读取文件内容
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()

        # 可以对内容进行处理
        lines = content.split('\\n')

        result = f"文件内容摘要\\n"
        result += f"总行数: {len(lines)}\\n"
        result += f"总字符数: {len(content)}\\n\\n"
        result += "文件内容:\\n"
        result += content

        return result

    except Exception as e:
        return f"读取文件失败: {str(e)}"
"""

    @staticmethod
    def datetime_example() -> str:
        """日期时间示例模板"""
        return """# 动态日期时间示例

from datetime import datetime, timedelta

def generate_content():
    now = datetime.now()

    content = f"亲爱的用户,您好!\\n\\n"
    content += f"这是一封自动生成的邮件\\n\\n"
    content += f"发送时间: {now.strftime('%Y年%m月%d日 %H:%M:%S')}\\n"
    content += f"星期: {['一', '二', '三', '四', '五', '六', '日'][now.weekday()]}\\n\\n"

    # 计算一周后的日期
    next_week = now + timedelta(days=7)
    content += f"下次发送时间: {next_week.strftime('%Y年%m月%d日')}\\n\\n"
    content += f"祝您工作顺利!\\n"

    return content
"""


if __name__ == "__main__":
    # 测试代码
    executor = ScriptExecutor()

    # 测试基础示例
    test_script = """
print("Hello from script!")
result = "This is the result"
"""

    success, output = executor.execute_script(test_script)
    print(f"Success: {success}")
    print(f"Output: {output}")
