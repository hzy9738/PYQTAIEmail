# Python脚本生成邮件内容 - 使用指南

## 功能简介

邮件发送功能现在支持使用Python脚本动态生成邮件内容,这让您可以:

- 📊 读取Excel、CSV等文件数据生成报表邮件
- 📁 扫描目录批量处理文件
- 🕐 生成包含动态日期时间的邮件
- 🔄 根据复杂逻辑动态生成个性化内容

## 快速开始

### 1. 切换到Python脚本模式

在发送邮件界面:
1. 选择 **内容模式** 为 "Python脚本"
2. 在 "Python脚本" 标签页编写代码
3. 点击 "测试脚本" 按钮验证脚本
4. 填写其他信息后点击 "立即发送"

### 2. 三种输出方式

#### 方式1: 使用 `print()` 输出
```python
print("这是邮件内容")
print("可以多行输出")
```

#### 方式2: 定义 `result` 变量
```python
result = "这是邮件内容"
```

#### 方式3: 定义 `generate_content()` 函数(推荐)
```python
def generate_content():
    content = "这是邮件内容"
    return content
```

## 使用示例

### 示例1: 读取Excel文件

```python
# 需要先安装: pip install pandas openpyxl
import pandas as pd
from datetime import datetime

def generate_content():
    # Excel文件路径
    excel_path = r"C:\Users\YourName\Documents\sales_data.xlsx"

    # 读取Excel
    df = pd.read_excel(excel_path)

    # 生成邮件内容
    content = f"销售数据报表\n"
    content += f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n"
    content += f"总记录数: {len(df)}\n"
    content += f"数据预览:\n{df.head().to_string()}"

    return content
```

### 示例2: 读取目录下所有Excel文件

```python
import pandas as pd
import os

def generate_content():
    folder = r"C:\Users\YourName\Documents\reports"

    content = "Excel文件汇总\n\n"

    # 遍历所有Excel文件
    for filename in os.listdir(folder):
        if filename.endswith('.xlsx'):
            filepath = os.path.join(folder, filename)
            df = pd.read_excel(filepath)

            content += f"文件: {filename}\n"
            content += f"  行数: {len(df)}\n\n"

    return content
```

### 示例3: 读取CSV文件

```python
import pandas as pd

def generate_content():
    csv_path = r"C:\Users\YourName\Documents\data.csv"
    df = pd.read_csv(csv_path, encoding='utf-8')

    content = f"CSV数据统计\n\n"
    content += f"总行数: {len(df)}\n"
    content += f"列名: {', '.join(df.columns)}\n\n"
    content += df.describe().to_string()

    return content
```

### 示例4: 读取文本文件

```python
def generate_content():
    file_path = r"C:\Users\YourName\Documents\log.txt"

    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()

    # 统计信息
    lines = content.split('\n')

    result = f"文件分析报告\n\n"
    result += f"总行数: {len(lines)}\n"
    result += f"总字符数: {len(content)}\n\n"
    result += "文件内容:\n" + content

    return result
```

### 示例5: 动态日期时间

```python
from datetime import datetime, timedelta

def generate_content():
    now = datetime.now()

    content = f"定期报告\n\n"
    content += f"报告日期: {now.strftime('%Y年%m月%d日')}\n"
    content += f"星期: {['一','二','三','四','五','六','日'][now.weekday()]}\n\n"

    # 下周的日期
    next_week = now + timedelta(days=7)
    content += f"下次报告: {next_week.strftime('%Y年%m月%d日')}\n"

    return content
```

### 示例6: HTML邮件内容

```python
from datetime import datetime

def generate_content():
    html = f"""
    <html>
    <head>
        <style>
            body {{ font-family: Arial, sans-serif; }}
            .header {{ background-color: #3498db; color: white; padding: 10px; }}
            .content {{ padding: 20px; }}
            table {{ border-collapse: collapse; width: 100%; }}
            th, td {{ border: 1px solid #ddd; padding: 8px; }}
            th {{ background-color: #3498db; color: white; }}
        </style>
    </head>
    <body>
        <div class="header">
            <h2>销售数据报表</h2>
        </div>
        <div class="content">
            <p>生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
            <table>
                <tr>
                    <th>产品</th>
                    <th>销量</th>
                </tr>
                <tr>
                    <td>产品A</td>
                    <td>100</td>
                </tr>
                <tr>
                    <td>产品B</td>
                    <td>150</td>
                </tr>
            </table>
        </div>
    </body>
    </html>
    """
    return html
```
**注意**: 发送HTML邮件时需要勾选 "使用HTML格式" 选项。

## 内置Python库支持

### 打包版本(exe)
打包后的程序已内置以下常用Python库,无需额外安装:

- ✅ **pandas** - Excel、CSV数据处理
- ✅ **openpyxl** - Excel文件读写
- ✅ **numpy** - 数值计算
- ✅ **xlrd/xlwt** - 旧版Excel支持

### 源码运行版本
如果从源码运行,需要安装相关库:

```bash
# Excel文件处理
pip install pandas openpyxl

# CSV文件处理
pip install pandas

# Word文档处理
pip install python-docx

# PDF文件处理
pip install PyPDF2

# JSON文件处理 (Python内置,无需安装)
```

**注意**: 打包版本已内置常用库,可直接使用Excel、CSV读取功能,无需本地Python环境

## 注意事项

### 1. 文件路径
- Windows系统使用原始字符串: `r"C:\Users\..."`
- 或使用双反斜杠: `"C:\\Users\\..."`

### 2. 文件编码
- 读取中文文件时指定编码: `encoding='utf-8'`
- 常见编码: `utf-8`, `gbk`, `gb2312`

### 3. 错误处理
建议添加错误处理:

```python
def generate_content():
    try:
        # 你的代码
        return "成功生成的内容"
    except Exception as e:
        return f"生成失败: {str(e)}"
```

### 4. 测试脚本
- 发送前务必点击 "测试脚本" 按钮
- 确认生成的内容符合预期
- 检查是否有错误信息

### 5. 安全性
- 仅执行您自己编写或信任的脚本
- 不要执行来源不明的代码
- 脚本在本地环境执行,请确保安全

## 模板库

软件内置了多个模板,在 "脚本模板" 下拉框中选择:

- **基础示例**: 展示三种输出方式
- **读取Excel文件**: 读取单个Excel文件
- **读取多个Excel文件**: 批量读取目录下的Excel
- **读取CSV文件**: 读取CSV数据
- **读取文本文件**: 读取文本文件内容
- **动态日期时间**: 生成包含日期时间的邮件

选择模板后,可以根据需要修改代码。

## 高级用法

### 1. 使用环境变量

```python
import os

def generate_content():
    username = os.environ.get('USERNAME')
    content = f"报告提交人: {username}\n"
    return content
```

### 2. 连接数据库

```python
# 需要安装: pip install pymysql
import pymysql

def generate_content():
    conn = pymysql.connect(
        host='localhost',
        user='root',
        password='password',
        database='mydb'
    )

    cursor = conn.cursor()
    cursor.execute("SELECT COUNT(*) FROM users")
    count = cursor.fetchone()[0]

    conn.close()

    return f"用户总数: {count}"
```

### 3. 网络请求

```python
# 需要安装: pip install requests
import requests

def generate_content():
    response = requests.get('https://api.example.com/data')
    data = response.json()

    content = f"API返回数据:\n{data}"
    return content
```

## 故障排除

### 问题1: 模块未找到
**错误**: `ModuleNotFoundError: No module named 'xxx'`

**解决**: 安装缺失的模块
```bash
pip install xxx
```

### 问题2: 文件未找到
**错误**: `FileNotFoundError: [Errno 2] No such file or directory`

**解决**:
- 检查文件路径是否正确
- 使用绝对路径
- 确认文件是否存在

### 问题3: 编码错误
**错误**: `UnicodeDecodeError`

**解决**: 指定正确的编码
```python
with open(file_path, 'r', encoding='gbk') as f:
    content = f.read()
```

### 问题4: 脚本无输出
**解决**:
- 确保使用了三种输出方式之一
- 检查脚本是否有语法错误
- 使用 "测试脚本" 功能查看错误信息

## 技术支持

如有问题,请:
1. 查看错误提示信息
2. 使用 "测试脚本" 功能调试
3. 检查Python环境和依赖库
4. 参考本文档的示例代码

---

**祝您使用愉快!**
