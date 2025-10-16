#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试Python脚本功能
"""

from script_executor import ScriptExecutor, ScriptTemplate


def test_basic_script():
    """测试基础脚本执行"""
    print("=" * 50)
    print("测试1: 基础print输出")
    print("=" * 50)

    executor = ScriptExecutor()

    script = """
print("Hello, World!")
print("这是第二行")
"""

    success, output = executor.execute_script(script)
    print(f"成功: {success}")
    print(f"输出:\n{output}")
    print()


def test_result_variable():
    """测试result变量"""
    print("=" * 50)
    print("测试2: result变量输出")
    print("=" * 50)

    executor = ScriptExecutor()

    script = """
result = "这是通过result变量输出的内容"
"""

    success, output = executor.execute_script(script)
    print(f"成功: {success}")
    print(f"输出:\n{output}")
    print()


def test_generate_function():
    """测试generate_content函数"""
    print("=" * 50)
    print("测试3: generate_content函数")
    print("=" * 50)

    executor = ScriptExecutor()

    script = """
def generate_content():
    return "这是通过generate_content函数输出的内容"
"""

    success, output = executor.execute_script(script)
    print(f"成功: {success}")
    print(f"输出:\n{output}")
    print()


def test_datetime_script():
    """测试日期时间脚本"""
    print("=" * 50)
    print("测试4: 日期时间脚本")
    print("=" * 50)

    executor = ScriptExecutor()

    script = """
from datetime import datetime

def generate_content():
    now = datetime.now()
    return f"当前时间: {now.strftime('%Y-%m-%d %H:%M:%S')}"
"""

    success, output = executor.execute_script(script)
    print(f"成功: {success}")
    print(f"输出:\n{output}")
    print()


def test_syntax_validation():
    """测试语法验证"""
    print("=" * 50)
    print("测试5: 语法验证")
    print("=" * 50)

    executor = ScriptExecutor()

    # 正确的语法
    valid_script = "print('Hello')"
    is_valid, msg = executor.validate_script(valid_script)
    print(f"正确语法: {is_valid} - {msg}")

    # 错误的语法
    invalid_script = "print('Hello'"  # 缺少右括号
    is_valid, msg = executor.validate_script(invalid_script)
    print(f"错误语法: {is_valid} - {msg}")
    print()


def test_error_handling():
    """测试错误处理"""
    print("=" * 50)
    print("测试6: 错误处理")
    print("=" * 50)

    executor = ScriptExecutor()

    script = """
# 这会导致运行时错误
x = 1 / 0
"""

    success, output = executor.execute_script(script)
    print(f"成功: {success}")
    print(f"错误信息:\n{output}")
    print()


def test_templates():
    """测试模板"""
    print("=" * 50)
    print("测试7: 模板列表")
    print("=" * 50)

    templates = ScriptTemplate.get_template_list()
    print(f"共有 {len(templates)} 个模板:\n")

    for i, template in enumerate(templates, 1):
        print(f"{i}. {template['name']}")
        print(f"   描述: {template['description']}")
        print()


def test_complex_script():
    """测试复杂脚本"""
    print("=" * 50)
    print("测试8: 复杂脚本(模拟读取数据)")
    print("=" * 50)

    executor = ScriptExecutor()

    script = """
def generate_content():
    # 模拟数据
    data = {
        '产品A': 100,
        '产品B': 150,
        '产品C': 80
    }

    content = "销售数据报表\\n\\n"

    total = 0
    for product, sales in data.items():
        content += f"{product}: {sales}件\\n"
        total += sales

    content += f"\\n总计: {total}件"

    return content
"""

    success, output = executor.execute_script(script)
    print(f"成功: {success}")
    print(f"输出:\n{output}")
    print()


if __name__ == "__main__":
    print("\n" + "=" * 50)
    print("Python脚本功能测试")
    print("=" * 50 + "\n")

    test_basic_script()
    test_result_variable()
    test_generate_function()
    test_datetime_script()
    test_syntax_validation()
    test_error_handling()
    test_templates()
    test_complex_script()

    print("=" * 50)
    print("所有测试完成!")
    print("=" * 50)
