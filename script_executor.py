#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Python脚本执行器模块
支持执行用户自定义Python脚本来动态生成邮件内容
"""

import sys
import os
import io
import traceback
import importlib
from typing import Dict, Any


class ScriptExecutor:
    """Python脚本执行器"""

    def __init__(self):
        """初始化执行器"""
        self.last_error = None
        self._modules_cache = {}
        # 预先导入常用模块，避免 numpy 2.x 的 CPU dispatcher 重复初始化问题
        self._preload_modules()

    def _preload_modules(self):
        """
        预先加载常用模块到缓存
        这样可以避免在 exec() 中重复导入导致的 numpy CPU dispatcher 错误
        """
        # 在 PyInstaller 环境中，切换工作目录以避免导入冲突
        original_dir = None
        if hasattr(sys, '_MEIPASS'):
            original_dir = os.getcwd()
            # 切换到临时解压目录，避免 numpy 导入错误
            os.chdir(sys._MEIPASS)

        try:
            # 预先导入常用模块
            import numpy
            import pandas
            import os as os_module
            from datetime import datetime, timedelta

            # 缓存常用模块
            self._modules_cache = {
                'numpy': numpy,
                'np': numpy,
                'pandas': pandas,
                'pd': pandas,
                'os': os_module,
                'datetime': datetime,
                'timedelta': timedelta,
            }
        except Exception as e:
            # 如果导入失败，记录但不中断
            print(f"Warning: Failed to preload modules: {e}")
            import traceback
            traceback.print_exc()
        finally:
            # 恢复原始工作目录
            if original_dir:
                os.chdir(original_dir)

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

        # 添加常用模块到执行环境（使用预加载的模块，避免重复导入）
        exec_globals = {
            '__builtins__': __builtins__,
            'context': context,  # 用户可以通过context访问传入的变量
        }

        # 将预加载的模块添加到执行环境
        exec_globals.update(self._modules_cache)

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
    def get_batch_data_template_list() -> list:
        """获取批量数据模板列表"""
        return [
            {
                "name": "批量数据 - 简单示例",
                "description": "展示如何处理单个Excel文件",
                "code": ScriptTemplate.batch_data_basic()
            },
            {
                "name": "批量数据 - HTML表格",
                "description": "将Excel转换为HTML表格",
                "code": ScriptTemplate.batch_data_html_table()
            },
            {
                "name": "批量数据 - 数据统计",
                "description": "生成Excel数据统计报告",
                "code": ScriptTemplate.batch_data_statistics()
            },
            {
                "name": "批量数据 - 自定义格式",
                "description": "自定义邮件格式模板",
                "code": ScriptTemplate.batch_data_custom()
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

    @staticmethod
    def batch_data_basic() -> str:
        """批量数据基础示例"""
        return """# 批量数据 - 简单示例
# 从context获取当前Excel文件信息并处理

import pandas as pd
from datetime import datetime

def generate_content():
    # 获取当前Excel文件路径
    file_path = context['file']
    filename = context['filename']
    index = context['index']
    total = context['total']

    # 读取Excel文件
    df = pd.read_excel(file_path)

    # 生成邮件内容
    content = f"尊敬的用户,您好!\\n\\n"
    content += f"这是第 {index}/{total} 份数据报告\\n"
    content += f"文件名: {filename}\\n\\n"
    content += f"数据概览:\\n"
    content += f"- 总行数: {len(df)}\\n"
    content += f"- 总列数: {len(df.columns)}\\n"
    content += f"- 列名: {', '.join(df.columns)}\\n\\n"
    content += f"数据预览(前5行):\\n"
    content += df.head().to_string()
    content += f"\\n\\n发送时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\\n"

    return content
"""

    @staticmethod
    def batch_data_html_table() -> str:
        """批量数据HTML表格模板"""
        return """# 批量数据 - HTML表格格式
# 将Excel转换为美观的HTML表格

import pandas as pd
from datetime import datetime

def generate_content():
    # 获取当前Excel文件信息
    file_path = context['file']
    filename = context['filename']
    index = context['index']
    total = context['total']

    # 读取Excel文件
    df = pd.read_excel(file_path)

    # 生成HTML格式的邮件内容
    html_content = f'''
    <html>
    <head>
        <style>
            body {{ font-family: Arial, sans-serif; }}
            h2 {{ color: #2c3e50; }}
            table {{ border-collapse: collapse; width: 100%; margin: 20px 0; }}
            th {{ background-color: #3498db; color: white; padding: 10px; text-align: left; }}
            td {{ border: 1px solid #ddd; padding: 8px; }}
            tr:nth-child(even) {{ background-color: #f2f2f2; }}
            .info {{ color: #7f8c8d; font-size: 12px; }}
        </style>
    </head>
    <body>
        <h2>📊 {filename} 数据报告</h2>
        <p>这是第 <strong>{index}/{total}</strong> 份报告</p>

        <h3>数据概览</h3>
        <ul>
            <li>总记录数: {len(df)}</li>
            <li>数据列: {len(df.columns)}</li>
        </ul>

        <h3>详细数据</h3>
        {df.to_html(index=False, border=0)}

        <p class="info">生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
    </body>
    </html>
    '''

    return html_content
"""

    @staticmethod
    def batch_data_statistics() -> str:
        """批量数据统计模板"""
        return """# 批量数据 - 数据统计报告
# 生成包含统计信息的数据报告

import pandas as pd
from datetime import datetime

def generate_content():
    # 获取当前Excel文件信息
    file_path = context['file']
    filename = context['filename']
    index = context['index']
    total = context['total']

    # 读取Excel文件
    df = pd.read_excel(file_path)

    # 生成统计报告
    content = f'''
<html>
<head>
    <style>
        body {{ font-family: Arial, sans-serif; padding: 20px; }}
        h2 {{ color: #2c3e50; border-bottom: 2px solid #3498db; padding-bottom: 10px; }}
        .stats {{ background-color: #ecf0f1; padding: 15px; border-radius: 5px; margin: 10px 0; }}
        .stats-item {{ margin: 5px 0; }}
        table {{ border-collapse: collapse; width: 100%; margin: 20px 0; }}
        th {{ background-color: #34495e; color: white; padding: 10px; }}
        td {{ border: 1px solid #bdc3c7; padding: 8px; }}
    </style>
</head>
<body>
    <h2>📈 {filename} 数据统计报告</h2>
    <p>报告编号: {index}/{total} | 生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>

    <div class="stats">
        <h3>基本统计</h3>
        <div class="stats-item">📋 总记录数: <strong>{len(df)}</strong></div>
        <div class="stats-item">📊 数据维度: <strong>{len(df.columns)} 列</strong></div>
        <div class="stats-item">📝 列名: {', '.join(df.columns)}</div>
    </div>
'''

    # 如果有数值列,添加统计信息
    numeric_cols = df.select_dtypes(include=['number']).columns
    if len(numeric_cols) > 0:
        content += f'''
    <div class="stats">
        <h3>数值统计</h3>
        {df[numeric_cols].describe().to_html()}
    </div>
'''

    # 添加数据预览
    content += f'''
    <h3>数据预览(前10行)</h3>
    {df.head(10).to_html(index=False)}

    <p style="color: #7f8c8d; font-size: 12px; margin-top: 20px;">
        此报告由寻拟邮件工具自动生成
    </p>
</body>
</html>
'''

    return content
"""

    @staticmethod
    def batch_data_custom() -> str:
        """批量数据自定义格式模板"""
        return """# 批量数据 - 自定义格式
# 根据实际业务需求自定义邮件格式

import pandas as pd
from datetime import datetime

def generate_content():
    # 获取当前Excel文件信息
    file_path = context['file']
    filename = context['filename']
    index = context['index']
    total = context['total']

    # 读取Excel文件
    df = pd.read_excel(file_path)

    # 示例: 假设Excel包含"姓名"、"金额"、"日期"等列
    # 请根据实际Excel结构修改下面的代码

    content = f"<html><body>"
    content += f"<h2>数据报告 - {filename}</h2>"
    content += f"<p>报告序号: {index}/{total}</p>"
    content += f"<hr>"

    # 示例: 遍历每一行数据
    for idx, row in df.iterrows():
        content += f"<p><strong>记录 {idx+1}:</strong></p>"
        content += f"<ul>"

        # 遍历每一列
        for col in df.columns:
            content += f"<li>{col}: {row[col]}</li>"

        content += f"</ul>"

        # 只显示前5条,避免邮件过长
        if idx >= 4:
            content += f"<p>... (还有 {len(df)-5} 条记录)</p>"
            break

    content += f"<hr>"
    content += f"<p style='color: gray;'>生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>"
    content += f"</body></html>"

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
