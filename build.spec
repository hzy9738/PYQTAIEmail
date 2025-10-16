# -*- mode: python ; coding: utf-8 -*-

"""
PyInstaller打包配置文件
用于将Python程序打包成Windows可执行文件
"""

block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],  # numpy 2.x 不再需要手动添加 .libs
    datas=[
        ('icon.ico', '.'),  # 将图标文件添加到打包资源中
    ],
    hiddenimports=[
        'PyQt5',
        'PyQt5.QtCore',
        'PyQt5.QtGui',
        'PyQt5.QtWidgets',
        'cryptography',
        'schedule',
        'email',
        'smtplib',
        'imaplib',
        # Python脚本常用库(可选,如果不打包这些库,用户脚本需要本地安装)
        'pandas',
        'pandas._libs',
        'pandas._libs.tslibs',
        'pandas._libs.tslibs.timedeltas',
        'openpyxl',
        'xlrd',
        'xlwt',
        # numpy 2.x 核心模块 (_core 替代了旧的 core)
        'numpy',
        'numpy._core',
        'numpy._core._exceptions',
        'numpy._core._multiarray_umath',
        'numpy._core.multiarray',
        'numpy._core.umath',
        'numpy._core._dtype',
        'numpy._core._internal',
        'numpy._core._methods',
        'numpy._core.numeric',
        'numpy._core.numerictypes',
        'numpy.core',
        'numpy.core._multiarray_umath',
        'numpy.core.multiarray',
        'numpy.linalg',
        'numpy.fft',
        'numpy.random',
        'numpy.random._common',
        'numpy.random._bounded_integers',
        'numpy.random._mt19937',
        'numpy.random._pcg64',
        'numpy.random._philox',
        'numpy.random._sfc64',
        'numpy.random._generator',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=['hook-numpy.py'],  # 添加 numpy runtime hook
    excludes=[
        'matplotlib',  # 排除不需要的大型库
        'scipy',
        'IPython',
        'notebook',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='寻拟邮件工具',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # 不显示控制台窗口
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico',  # 使用自定义图标
)
