#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
创建测试用的Excel文件
"""

import pandas as pd
import os

# 创建测试目录
test_dir = "test_excel"
if not os.path.exists(test_dir):
    os.makedirs(test_dir)

# 创建第一个测试Excel
data1 = {
    '姓名': ['张三', '李四', '王五'],
    '年龄': [25, 30, 28],
    '城市': ['北京', '上海', '广州'],
    '工资': [8000, 12000, 10000]
}
df1 = pd.DataFrame(data1)
df1.to_excel(os.path.join(test_dir, 'report1.xlsx'), index=False)

# 创建第二个测试Excel
data2 = {
    '产品': ['苹果', '香蕉', '橙子', '葡萄'],
    '数量': [100, 150, 80, 200],
    '单价': [5.5, 3.2, 4.8, 6.5],
    '总价': [550, 480, 384, 1300]
}
df2 = pd.DataFrame(data2)
df2.to_excel(os.path.join(test_dir, 'report2.xlsx'), index=False)

# 创建第三个测试Excel
data3 = {
    '日期': ['2025-01-01', '2025-01-02', '2025-01-03'],
    '销售额': [5000, 6000, 7500],
    '成本': [3000, 3500, 4200],
    '利润': [2000, 2500, 3300]
}
df3 = pd.DataFrame(data3)
df3.to_excel(os.path.join(test_dir, 'report3.xlsx'), index=False)

print(f"成功创建3个测试Excel文件到 {test_dir} 目录!")
print("文件列表:")
for file in os.listdir(test_dir):
    print(f"  - {file}")
