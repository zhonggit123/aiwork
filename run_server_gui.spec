# -*- mode: python ; coding: utf-8 -*-
# PyInstaller 打包配置：题库录入后端 GUI 版 → 单 exe
# 用法：pyinstaller run_server_gui.spec

import os

block_cipher = None

# 需要随 exe 打包的数据文件（会解压到运行时的临时目录 _MEIPASS）
datas = [
    ("config.example.yaml", "."),
    ("static", "static"),
    ("samples", "samples"),
    ("extension.zip", "."),  # 插件压缩包，供用户下载
]

# 可能被动态导入的模块，显式收集避免漏打
hiddenimports = [
    # uvicorn 相关
    "uvicorn.logging",
    "uvicorn.loops",
    "uvicorn.loops.auto",
    "uvicorn.protocols",
    "uvicorn.protocols.http",
    "uvicorn.protocols.http.auto",
    "uvicorn.protocols.websockets",
    "uvicorn.protocols.websockets.auto",
    "uvicorn.lifespan",
    "uvicorn.lifespan.on",
    # edge-tts 相关模块（TTS 语音合成）
    "edge_tts",
    "edge_tts.communicate",
    "edge_tts.submaker",
    "edge_tts.list_voices",
    # aiohttp 相关（edge-tts 依赖）
    "aiohttp",
    "aiohttp.web",
    "certifi",
    # PIL/Pillow 图片处理
    "PIL",
    "PIL.Image",
    "PIL.ImageDraw",
    "PIL.ImageFont",
    # pymupdf/fitz PDF 解析
    "fitz",
    # json_repair JSON 修复
    "json_repair",
    # python-docx Word 解析
    "docx",
    "docx.oxml",
    "docx.oxml.ns",
    # openai API
    "openai",
    # pydantic 数据验证
    "pydantic",
    # yaml 配置
    "yaml",
    # httpx HTTP 客户端
    "httpx",
    # multipart 文件上传
    "multipart",
]

a = Analysis(
    ["run_server_gui.py"],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],  # 不排除 tkinter，GUI 需要
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
    name="题库录入服务",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # GUI 模式，不显示控制台窗口
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,  # 可以添加 icon="icon.ico" 指定图标
)
