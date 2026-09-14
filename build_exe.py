# -*- coding: utf-8 -*-
"""
陈平安资料库 打包配置脚本
使用 PyInstaller 打包成独立可执行程序
"""

import os
import sys

APP_NAME = '陈平安资料库'
APP_SCRIPT = 'launcher.py'

# 需要打包的数据文件
datas = [
    ('templates', 'templates'),
    ('static', 'static'),
    ('step1_generator.py', '.'),
    ('step23_processor.py', '.'),
]

# 需要的隐藏导入
hiddenimports = [
    'openpyxl',
    'xlrd',
    'pandas',
    'numpy',
    'flask',
    'sqlite3',
    're',
    'json',
    'datetime',
    'hashlib',
    'secrets',
    'uuid',
    'urllib.request',
]

# 排除的模块
excludes = [
    'tkinter',
    'matplotlib',
    'scipy',
    'PIL',
    'IPython',
    'jupyter',
    'pytest',
]

if __name__ == '__main__':
    import PyInstaller.__main__

    args = [
        APP_SCRIPT,
        '--name=' + APP_NAME,
        '--noconfirm',
        '--clean',
        '--onefile',
        '--windowed',
        '--add-data=templates;templates',
        '--add-data=static;static',
        '--add-data=step1_generator.py;.',
        '--add-data=step23_processor.py;.',
        '--hidden-import=openpyxl',
        '--hidden-import=xlrd',
        '--hidden-import=pandas',
        '--hidden-import=numpy',
        '--hidden-import=flask',
        '--hidden-import=jinja2',
        '--hidden-import=werkzeug',
        '--hidden-import=sqlite3',
        '--exclude-module=tkinter',
        '--exclude-module=matplotlib',
        '--exclude-module=pytest',
        '--exclude-module=IPython',
    ]

    # 添加图标
    icon_path = os.path.join('static', 'images', 'logo.png')
    if os.path.exists(icon_path):
        args.append('--icon=' + icon_path)

    print("开始打包...")
    print(f"输出目录: {os.path.join(os.getcwd(), 'dist')}")
    PyInstaller.__main__.run(args)
    print("\n打包完成！")
    print(f"可执行文件位于: dist/{APP_NAME}.exe")
