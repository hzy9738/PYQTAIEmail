#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PyInstaller runtime hook for numpy
设置 numpy 环境变量以避免运行时错误
"""

import os
import sys

# 在 PyInstaller 环境中设置 numpy 环境变量
if hasattr(sys, '_MEIPASS'):
    # 禁用实验性数组功能，避免兼容性问题
    os.environ['NUMPY_EXPERIMENTAL_ARRAY_FUNCTION'] = '0'
    # 设置 OpenBLAS 线程数，避免多线程冲突
    os.environ['OPENBLAS_NUM_THREADS'] = '1'

    # 添加 numpy DLL 路径到系统路径
    numpy_libs_path = os.path.join(sys._MEIPASS, 'numpy', '.libs')
    if os.path.exists(numpy_libs_path):
        # Windows 上需要使用 add_dll_directory (Python 3.8+)
        if hasattr(os, 'add_dll_directory'):
            os.add_dll_directory(numpy_libs_path)
        # 同时添加到 PATH 环境变量
        os.environ['PATH'] = numpy_libs_path + os.pathsep + os.environ.get('PATH', '')
