# -*- mode: python ; coding: utf-8 -*-
# PyInstaller 打包：题库自动填表 - 后端服务
# 用法：pyinstaller 题库后端.spec
# 生成在 dist/题库后端/ 目录，可整文件夹发给别人，运行 题库后端.exe 即可

import sys

block_cipher = None

# 需打包的 Python 模块
main_scripts = ['run_server.py']
hidden_imports = [
    'app',  # uvicorn 通过字符串 "app:app" 加载，必须显式加入否则 exe 内找不到
    'word_reader', 'llm_extract', 'submit_api',
    'uvicorn.logging', 'uvicorn.loops', 'uvicorn.loops.auto', 'uvicorn.protocols',
    'uvicorn.protocols.http', 'uvicorn.protocols.http.auto', 'uvicorn.protocols.websockets',
    'uvicorn.protocols.websockets.auto', 'uvicorn.lifespan', 'uvicorn.lifespan.on',
    'multipart', 'python_multipart',
    # 图片处理（Word 图片提取 + TIFF/CMYK 转 JPEG）
    'PIL', 'PIL.Image', 'PIL.JpegImagePlugin', 'PIL.PngImagePlugin',
    'PIL.TiffImagePlugin', 'PIL.BmpImagePlugin', 'PIL.GifImagePlugin',
    'PIL.WebPImagePlugin', 'PIL.ImageFile',
    # JSON 修复（LLM 输出容错）
    'json_repair',
]

# 数据文件：静态页、示例目录、配置模板（与 exe 同目录会放 config.example.yaml）
datas = [
    ('static', 'static'),
    ('samples', 'samples'),
    ('config.example.yaml', '.'),
]

a = Analysis(
    ['run_server.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['tkinter'],
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
    name='题库后端',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,   # 保留控制台，方便看日志和错误
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
