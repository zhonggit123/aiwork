# -*- coding: utf-8 -*-
"""
后端服务启动入口。直接运行即启动 Web 服务（端口 8766）。
用于本地开发或 PyInstaller 打包成 exe 后的入口。
"""
import sys
from pathlib import Path

# 打包成 exe 时，确保工作目录为 exe 所在目录，便于找到同目录下的 config.yaml
if getattr(sys, "frozen", False):
    import shutil
    import os
    exe_dir = Path(sys.executable).parent
    config_path = exe_dir / "config.yaml"
    # 若未存在 config.yaml，从打包内嵌资源复制 config.example.yaml
    if not config_path.exists():
        src = Path(sys._MEIPASS) / "config.example.yaml"
        if src.exists():
            shutil.copy(src, config_path)
            print("已生成 config.yaml，请编辑后填写 API 等配置并重新运行。")
    if exe_dir != Path.cwd():
        os.chdir(exe_dir)


def main():
    import uvicorn
    uvicorn.run(
        "app:app",
        host="0.0.0.0",
        port=8766,
        log_level="info",
        reload=not getattr(__import__("sys"), "frozen", False),  # 本地开发自动重载，打包 exe 时禁用
    )


if __name__ == "__main__":
    main()
